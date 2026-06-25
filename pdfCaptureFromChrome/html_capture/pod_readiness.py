"""DOM-readiness profiling for automated POD tracking-page captures."""

from __future__ import annotations

import json
import os
import re
import time
import urllib.parse
from collections import Counter
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

try:
    from ..chrome_devtools import _CdpSession
except ImportError:  # pragma: no cover - script-style import fallback
    from chrome_devtools import _CdpSession  # type: ignore[no-redef]

PROFILE_VERSION = 1
PROFILE_FILENAME = "tracking_page_readiness_profiles.json"
EVENT_LOG_FILENAME = "pod_readiness_events.jsonl"
DEBUG_ENV_KEY = "POD_READINESS_DEBUG"
DEFAULT_QUIET_SECONDS = 5.0
DEFAULT_PROFILE_GRACE_SECONDS = 1.0
MIN_READY_TEXT_CHARS = 80
MAX_SELECTOR_CANDIDATES = 48


def pod_readiness_debug_enabled() -> bool:
    """Return true when the launcher-specific POD readiness debug switch is on."""
    try:
        from shared.settings_store import apply_runtime_settings_from_json

        apply_runtime_settings_from_json()
    except Exception:
        pass
    return (os.getenv(DEBUG_ENV_KEY) or "").strip().lower() in ("1", "true", "yes", "on")


def _project_root() -> Path:
    try:
        from shared.project_paths import ensure_base_dir_in_environ

        return ensure_base_dir_in_environ()
    except Exception:
        raw = (os.getenv("BASE_DIR") or "").strip()
        if raw:
            return Path(raw).expanduser().resolve()
        return Path.cwd().resolve()


def _profile_path() -> Path:
    path = _project_root() / "email_contents" / "json" / PROFILE_FILENAME
    path.parent.mkdir(parents=True, exist_ok=True)
    return path


def _event_log_path() -> Path:
    path = _project_root() / "email_contents" / "logs" / EVENT_LOG_FILENAME
    path.parent.mkdir(parents=True, exist_ok=True)
    return path


def _utc_now() -> str:
    return datetime.now(timezone.utc).isoformat(timespec="seconds").replace("+00:00", "Z")


def _host_for_url(url: str | None) -> str:
    raw = str(url or "").strip()
    if not raw:
        return "unknown-host"
    try:
        parsed = urllib.parse.urlparse(raw)
    except ValueError:
        return "unknown-host"
    host = (parsed.hostname or "").strip().casefold()
    if host.startswith("www."):
        host = host[4:]
    return host or "unknown-host"


def _carrier_for_record(record: dict | None) -> str:
    if not isinstance(record, dict):
        return "unknown-carrier"
    raw = str(record.get("carrier") or "").strip()
    return raw or "unknown-carrier"


def _profile_key(url: str | None, record: dict | None) -> str:
    host = _host_for_url(url)
    carrier = re.sub(r"[^a-z0-9]+", "-", _carrier_for_record(record).casefold()).strip("-")
    return f"{carrier or 'unknown-carrier'}@{host}"


def _load_profiles() -> dict[str, Any]:
    path = _profile_path()
    if not path.is_file():
        return {"version": PROFILE_VERSION, "profiles": {}}
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, UnicodeError, json.JSONDecodeError):
        return {"version": PROFILE_VERSION, "profiles": {}}
    if not isinstance(data, dict):
        return {"version": PROFILE_VERSION, "profiles": {}}
    profiles = data.get("profiles")
    if not isinstance(profiles, dict):
        profiles = {}
    return {"version": PROFILE_VERSION, "profiles": profiles}


def _save_profiles(data: dict[str, Any]) -> None:
    path = _profile_path()
    data["version"] = PROFILE_VERSION
    data["updated_at"] = _utc_now()
    tmp = path.with_suffix(".tmp")
    tmp.write_text(json.dumps(data, indent=2, ensure_ascii=True) + "\n", encoding="utf-8")
    tmp.replace(path)


def _append_event_log(profile_key: str, event: dict[str, Any]) -> None:
    payload = {
        "logged_at": _utc_now(),
        "profile_key": profile_key,
        "event": event,
    }
    try:
        with _event_log_path().open("a", encoding="utf-8", newline="\n") as handle:
            handle.write(json.dumps(payload, ensure_ascii=True, sort_keys=True) + "\n")
    except OSError:
        pass


def _json_value_from_eval(result: dict | None) -> dict[str, Any] | None:
    if not isinstance(result, dict):
        return None
    payload = result.get("result")
    if not isinstance(payload, dict):
        return None
    value = payload.get("value")
    if not isinstance(value, str):
        return None
    try:
        decoded = json.loads(value)
    except json.JSONDecodeError:
        return None
    return decoded if isinstance(decoded, dict) else None


def _eval_json(ws_url: str, expression: str, *, timeout: float = 5.0) -> dict[str, Any] | None:
    try:
        with _CdpSession(ws_url, timeout=timeout) as cdp:
            result = cdp.call(
                "Runtime.evaluate",
                {
                    "expression": expression,
                    "returnByValue": True,
                    "awaitPromise": True,
                },
            )
    except Exception:
        return None
    return _json_value_from_eval(result)


_OBSERVER_SCRIPT = r"""
(() => {
  if (window.__podReadinessObserver && window.__podReadinessObserver.version === 1) {
    return JSON.stringify({ok: true, already: true});
  }
  const startedAt = performance.now();
  const maxEvents = 180;
  const cssEscape = (value) => {
    const s = String(value || "");
    if (window.CSS && CSS.escape) return CSS.escape(s);
    return s.replace(/[^a-zA-Z0-9_-]/g, "\\$&");
  };
  const attrSelector = (tag, attr, value) => {
    const safe = String(value || "").replace(/\\/g, "\\\\").replace(/"/g, '\\"');
    return `${tag}[${attr}="${safe}"]`;
  };
  const cleanToken = (value) => {
    const s = String(value || "").trim();
    if (!s || s.length > 48) return "";
    if (/[0-9a-f]{10,}/i.test(s)) return "";
    if (/\d{5,}/.test(s)) return "";
    return s;
  };
  const visible = (el) => {
    if (!el || el.nodeType !== 1 || !el.getBoundingClientRect) return false;
    const rect = el.getBoundingClientRect();
    if (rect.width < 2 || rect.height < 2) return false;
    const style = getComputedStyle(el);
    return style.display !== "none" && style.visibility !== "hidden" && Number(style.opacity || 1) > 0;
  };
  const shortText = (el) => {
    const text = (el && (el.innerText || el.textContent) || "").replace(/\s+/g, " ").trim();
    return text.slice(0, 180);
  };
  const selectorFor = (el) => {
    if (!el || el.nodeType !== 1) return "";
    const tag = el.tagName.toLowerCase();
    const id = cleanToken(el.id);
    if (id) return `${tag}#${cssEscape(id)}`;
    for (const attr of ["data-testid", "data-test", "data-qa", "aria-label", "name", "role"]) {
      const value = cleanToken(el.getAttribute(attr));
      if (value) return attrSelector(tag, attr, value);
    }
    const parts = [];
    let cur = el;
    for (let depth = 0; cur && cur.nodeType === 1 && depth < 5; depth += 1) {
      const curTag = cur.tagName.toLowerCase();
      let part = curTag;
      const classes = Array.from(cur.classList || []).map(cleanToken).filter(Boolean).slice(0, 3);
      if (classes.length) part += "." + classes.map(cssEscape).join(".");
      const parent = cur.parentElement;
      if (parent) {
        const sameTag = Array.from(parent.children).filter((child) => child.tagName === cur.tagName);
        if (sameTag.length > 1) part += `:nth-of-type(${sameTag.indexOf(cur) + 1})`;
      }
      parts.unshift(part);
      if (cur.id && cleanToken(cur.id)) break;
      cur = parent;
    }
    return parts.join(" > ");
  };
  const describe = (el, kind) => {
    if (!el || el.nodeType !== 1) return null;
    const rect = el.getBoundingClientRect ? el.getBoundingClientRect() : {x: 0, y: 0, width: 0, height: 0};
    const text = shortText(el);
    return {
      kind,
      selector: selectorFor(el),
      tag: el.tagName.toLowerCase(),
      id: cleanToken(el.id),
      classes: Array.from(el.classList || []).map(cleanToken).filter(Boolean).slice(0, 6),
      text,
      visible: visible(el),
      rect: {
        x: Math.round(rect.x || 0),
        y: Math.round(rect.y || 0),
        width: Math.round(rect.width || 0),
        height: Math.round(rect.height || 0)
      }
    };
  };
  const state = {
    version: 1,
    startedAt,
    lastMutationAt: performance.now(),
    sequence: 0,
    sentSequence: 0,
    events: [],
    observer: null
  };
  const push = (event) => {
    if (!event) return;
    state.sequence += 1;
    event.sequence = state.sequence;
    event.at_ms = Math.round(performance.now() - startedAt);
    state.lastMutationAt = performance.now();
    state.events.push(event);
    if (state.events.length > maxEvents) state.events.shift();
    try {
      if (event.visible || event.kind.indexOf("css") >= 0) {
        console.log("[POD readiness]", JSON.stringify(event));
      }
    } catch (_err) {}
  };
  const noteNode = (node, kind) => {
    if (!node || node.nodeType !== 1) return;
    const el = node;
    const tag = el.tagName.toLowerCase();
    const eventKind = (tag === "style" || (tag === "link" && /stylesheet/i.test(el.rel || ""))) ? "css-added" : kind;
    const direct = describe(el, eventKind);
    if (direct && (direct.visible || eventKind === "css-added")) push(direct);
    if (el.querySelectorAll) {
      const candidates = el.querySelectorAll("main,section,article,table,[role='main'],[data-testid],[data-test],[data-qa],h1,h2,h3,p,button,a,img,div");
      const slice = Array.from(candidates).filter(visible).slice(-4);
      for (const child of slice) push(describe(child, "element-added-descendant"));
    }
  };
  const observer = new MutationObserver((mutations) => {
    for (const mutation of mutations) {
      if (mutation.type === "childList") {
        for (const node of mutation.addedNodes || []) noteNode(node, "element-added");
      } else if (mutation.type === "attributes") {
        const desc = describe(mutation.target, `attribute-${mutation.attributeName || "changed"}`);
        if (desc && (desc.visible || mutation.attributeName === "style" || mutation.attributeName === "class")) push(desc);
      }
    }
  });
  observer.observe(document.documentElement || document, {
    childList: true,
    subtree: true,
    attributes: true,
    attributeFilter: ["class", "style", "hidden", "aria-hidden"]
  });
  state.observer = observer;
  window.__podReadinessObserver = state;
  push({kind: "observer-start", selector: "html", tag: "html", visible: true, text: document.title || ""});
  return JSON.stringify({ok: true, already: false});
})()
"""


_SNAPSHOT_SCRIPT = r"""
(() => {
  const state = window.__podReadinessObserver || null;
  const now = performance.now();
  const cssEscape = (value) => {
    const s = String(value || "");
    if (window.CSS && CSS.escape) return CSS.escape(s);
    return s.replace(/[^a-zA-Z0-9_-]/g, "\\$&");
  };
  const cleanToken = (value) => {
    const s = String(value || "").trim();
    if (!s || s.length > 48) return "";
    if (/[0-9a-f]{10,}/i.test(s)) return "";
    if (/\d{5,}/.test(s)) return "";
    return s;
  };
  const visible = (el) => {
    if (!el || el.nodeType !== 1 || !el.getBoundingClientRect) return false;
    const rect = el.getBoundingClientRect();
    if (rect.width < 2 || rect.height < 2) return false;
    const style = getComputedStyle(el);
    return style.display !== "none" && style.visibility !== "hidden" && Number(style.opacity || 1) > 0;
  };
  const shortText = (el) => String((el && (el.innerText || el.textContent)) || "").replace(/\s+/g, " ").trim().slice(0, 180);
  const selectorFor = (el) => {
    if (!el || el.nodeType !== 1) return "";
    const tag = el.tagName.toLowerCase();
    const id = cleanToken(el.id);
    if (id) return `${tag}#${cssEscape(id)}`;
    for (const attr of ["data-testid", "data-test", "data-qa", "aria-label", "name", "role"]) {
      const value = cleanToken(el.getAttribute(attr));
      if (value) {
        const safe = value.replace(/\\/g, "\\\\").replace(/"/g, '\\"');
        return `${tag}[${attr}="${safe}"]`;
      }
    }
    const parts = [];
    let cur = el;
    for (let depth = 0; cur && cur.nodeType === 1 && depth < 5; depth += 1) {
      const curTag = cur.tagName.toLowerCase();
      let part = curTag;
      const classes = Array.from(cur.classList || []).map(cleanToken).filter(Boolean).slice(0, 3);
      if (classes.length) part += "." + classes.map(cssEscape).join(".");
      const parent = cur.parentElement;
      if (parent) {
        const sameTag = Array.from(parent.children).filter((child) => child.tagName === cur.tagName);
        if (sameTag.length > 1) part += `:nth-of-type(${sameTag.indexOf(cur) + 1})`;
      }
      parts.unshift(part);
      cur = parent;
    }
    return parts.join(" > ");
  };
  const describe = (el, source) => {
    if (!el || el.nodeType !== 1) return null;
    const rect = el.getBoundingClientRect ? el.getBoundingClientRect() : {x: 0, y: 0, width: 0, height: 0};
    return {
      source,
      selector: selectorFor(el),
      tag: el.tagName.toLowerCase(),
      text: shortText(el),
      visible: visible(el),
      rect: {
        x: Math.round(rect.x || 0),
        y: Math.round(rect.y || 0),
        width: Math.round(rect.width || 0),
        height: Math.round(rect.height || 0)
      }
    };
  };
  const events = state ? state.events.slice() : [];
  let candidate = null;
  for (let i = events.length - 1; i >= 0; i -= 1) {
    const event = events[i];
    if (event && event.visible && event.selector && event.kind !== "observer-start") {
      candidate = event;
      break;
    }
  }
  if (!candidate) {
    const nodes = Array.from(document.querySelectorAll("main,section,article,table,[role='main'],[data-testid],[data-test],[data-qa],h1,h2,h3,p,button,a,div"));
    let best = null;
    let bestScore = 0;
    for (const node of nodes) {
      if (!visible(node)) continue;
      const text = shortText(node);
      const rect = node.getBoundingClientRect();
      const score = Math.min(text.length, 300) + Math.min(rect.width * rect.height / 20000, 100);
      if (score > bestScore) {
        bestScore = score;
        best = node;
      }
    }
    candidate = best ? describe(best, "snapshot-best") : null;
  }
  const bodyText = document.body ? (document.body.innerText || "") : "";
  const seq = state ? Number(state.sequence || 0) : 0;
  const sent = state ? Number(state.sentSequence || 0) : 0;
  const newEvents = state ? state.events.filter((ev) => Number(ev.sequence || 0) > sent) : [];
  if (state) state.sentSequence = seq;
  return JSON.stringify({
    href: location.href,
    title: document.title || "",
    readyState: document.readyState || "",
    elapsed_ms: state ? Math.round(now - state.startedAt) : 0,
    quiet_ms: state ? Math.round(now - state.lastMutationAt) : 0,
    sequence: seq,
    new_events: newEvents,
    text_length: bodyText.replace(/\s+/g, " ").trim().length,
    element_count: document.getElementsByTagName("*").length,
    stylesheet_count: document.styleSheets ? document.styleSheets.length : 0,
    candidate
  });
})()
"""


def _selector_visible_script(selector: str) -> str:
    selector_json = json.dumps(selector)
    return f"""
(() => {{
  const selector = {selector_json};
  let el = null;
  try {{
    el = document.querySelector(selector);
  }} catch (_err) {{
    return JSON.stringify({{ok: false, selector, visible: false, error: "bad selector"}});
  }}
  if (!el || !el.getBoundingClientRect) {{
    return JSON.stringify({{ok: true, selector, visible: false}});
  }}
  const rect = el.getBoundingClientRect();
  const style = getComputedStyle(el);
  const visible = rect.width >= 2 && rect.height >= 2 && style.display !== "none" && style.visibility !== "hidden" && Number(style.opacity || 1) > 0;
  const text = String(el.innerText || el.textContent || "").replace(/\\s+/g, " ").trim().slice(0, 180);
  return JSON.stringify({{
    ok: true,
    selector,
    visible,
    tag: el.tagName.toLowerCase(),
    text,
    rect: {{
      x: Math.round(rect.x || 0),
      y: Math.round(rect.y || 0),
      width: Math.round(rect.width || 0),
      height: Math.round(rect.height || 0)
    }}
  }});
}})()
"""


def _overlay_script(selector: str | None) -> str:
    selector_json = json.dumps(selector or "")
    return f"""
(() => {{
  const selector = {selector_json};
  const visible = (el) => {{
    if (!el || !el.getBoundingClientRect) return false;
    const rect = el.getBoundingClientRect();
    if (rect.width < 2 || rect.height < 2) return false;
    const style = getComputedStyle(el);
    return style.display !== "none" && style.visibility !== "hidden" && Number(style.opacity || 1) > 0;
  }};
  let el = null;
  if (selector) {{
    try {{ el = document.querySelector(selector); }} catch (_err) {{ el = null; }}
  }}
  if (!visible(el)) {{
    const nodes = Array.from(document.querySelectorAll("main,section,article,table,[role='main'],[data-testid],[data-test],[data-qa],h1,h2,h3,p,button,a,div"));
    let best = null;
    let bestScore = 0;
    for (const node of nodes) {{
      if (!visible(node)) continue;
      const text = String(node.innerText || node.textContent || "").replace(/\\s+/g, " ").trim();
      const rect = node.getBoundingClientRect();
      const score = Math.min(text.length, 300) + Math.min(rect.width * rect.height / 20000, 100);
      if (score > bestScore) {{
        bestScore = score;
        best = node;
      }}
    }}
    el = best;
  }}
  if (!visible(el)) return JSON.stringify({{ok: false, reason: "no visible element"}});
  let box = document.getElementById("__podReadinessBlueBox");
  if (!box) {{
    box = document.createElement("div");
    box.id = "__podReadinessBlueBox";
    box.style.position = "fixed";
    box.style.pointerEvents = "none";
    box.style.zIndex = "2147483647";
    box.style.border = "3px solid rgba(37, 99, 235, 0.95)";
    box.style.background = "rgba(59, 130, 246, 0.16)";
    box.style.boxShadow = "0 0 0 9999px rgba(37, 99, 235, 0.045)";
    box.style.borderRadius = "6px";
    document.documentElement.appendChild(box);
  }}
  const update = () => {{
    if (!visible(el)) return;
    const rect = el.getBoundingClientRect();
    box.style.left = `${{Math.max(0, rect.left - 4)}}px`;
    box.style.top = `${{Math.max(0, rect.top - 4)}}px`;
    box.style.width = `${{Math.max(0, rect.width + 8)}}px`;
    box.style.height = `${{Math.max(0, rect.height + 8)}}px`;
  }};
  update();
  if (window.__podReadinessBlueBoxTimer) clearInterval(window.__podReadinessBlueBoxTimer);
  window.__podReadinessBlueBoxTimer = setInterval(update, 250);
  window.addEventListener("scroll", update, {{passive: true}});
  window.addEventListener("resize", update);
  const rect = el.getBoundingClientRect();
  return JSON.stringify({{
    ok: true,
    selector,
    tag: el.tagName.toLowerCase(),
    text: String(el.innerText || el.textContent || "").replace(/\\s+/g, " ").trim().slice(0, 180),
    rect: {{
      x: Math.round(rect.x || 0),
      y: Math.round(rect.y || 0),
      width: Math.round(rect.width || 0),
      height: Math.round(rect.height || 0)
    }}
  }});
}})()
"""


_REMOVE_OVERLAY_SCRIPT = r"""
(() => {
  try {
    if (window.__podReadinessBlueBoxTimer) {
      clearInterval(window.__podReadinessBlueBoxTimer);
      window.__podReadinessBlueBoxTimer = null;
    }
  } catch (_err) {}
  const box = document.getElementById("__podReadinessBlueBox");
  if (box && box.parentNode) {
    box.parentNode.removeChild(box);
  }
  return JSON.stringify({ok: true, removed: Boolean(box)});
})()
"""


_CANDIDATES_SCRIPT = r"""
(() => {
  const cssEscape = (value) => {
    const s = String(value || "");
    if (window.CSS && CSS.escape) return CSS.escape(s);
    return s.replace(/[^a-zA-Z0-9_-]/g, "\\$&");
  };
  const cleanToken = (value) => {
    const s = String(value || "").trim();
    if (!s || s.length > 48) return "";
    if (/[0-9a-f]{10,}/i.test(s)) return "";
    if (/\d{5,}/.test(s)) return "";
    return s;
  };
  const visible = (el) => {
    if (!el || el.nodeType !== 1 || !el.getBoundingClientRect) return false;
    const rect = el.getBoundingClientRect();
    if (rect.width < 2 || rect.height < 2) return false;
    const style = getComputedStyle(el);
    return style.display !== "none" && style.visibility !== "hidden" && Number(style.opacity || 1) > 0;
  };
  const shortText = (el) => String((el && (el.innerText || el.textContent)) || "").replace(/\s+/g, " ").trim().slice(0, 180);
  const selectorFor = (el) => {
    if (!el || el.nodeType !== 1) return "";
    const tag = el.tagName.toLowerCase();
    const id = cleanToken(el.id);
    if (id) return `${tag}#${cssEscape(id)}`;
    for (const attr of ["data-testid", "data-test", "data-qa", "aria-label", "name", "role"]) {
      const value = cleanToken(el.getAttribute(attr));
      if (value) {
        const safe = value.replace(/\\/g, "\\\\").replace(/"/g, '\\"');
        return `${tag}[${attr}="${safe}"]`;
      }
    }
    const parts = [];
    let cur = el;
    for (let depth = 0; cur && cur.nodeType === 1 && depth < 5; depth += 1) {
      const curTag = cur.tagName.toLowerCase();
      let part = curTag;
      const classes = Array.from(cur.classList || []).map(cleanToken).filter(Boolean).slice(0, 3);
      if (classes.length) part += "." + classes.map(cssEscape).join(".");
      const parent = cur.parentElement;
      if (parent) {
        const sameTag = Array.from(parent.children).filter((child) => child.tagName === cur.tagName);
        if (sameTag.length > 1) part += `:nth-of-type(${sameTag.indexOf(cur) + 1})`;
      }
      parts.unshift(part);
      cur = parent;
    }
    return parts.join(" > ");
  };
  const describe = (el, source) => {
    if (!visible(el)) return null;
    const rect = el.getBoundingClientRect();
    const text = shortText(el);
    const selector = selectorFor(el);
    if (!selector) return null;
    const area = Math.max(0, rect.width * rect.height);
    const textScore = Math.min(text.length, 320);
    const areaScore = Math.min(area / 18000, 140);
    const tag = el.tagName.toLowerCase();
    const role = String(el.getAttribute("role") || "");
    const semanticBoost = (
      ["main", "article", "section", "table"].includes(tag) ||
      ["main", "article", "status", "region"].includes(role) ||
      el.hasAttribute("data-testid") ||
      el.hasAttribute("data-test") ||
      el.hasAttribute("data-qa")
    ) ? 95 : 0;
    const trackingBoost = /tracking|deliver|shipment|status|package|proof|pod|estimated|transit/i.test(text) ? 120 : 0;
    return {
      source,
      selector,
      tag,
      role,
      text,
      score: Math.round(textScore + areaScore + semanticBoost + trackingBoost),
      rect: {
        x: Math.round(rect.x || 0),
        y: Math.round(rect.y || 0),
        width: Math.round(rect.width || 0),
        height: Math.round(rect.height || 0)
      }
    };
  };
  const query = [
    "main",
    "article",
    "section",
    "table",
    "[role='main']",
    "[role='status']",
    "[role='region']",
    "[data-testid]",
    "[data-test]",
    "[data-qa]",
    "h1",
    "h2",
    "h3",
    "p",
    "button",
    "a",
    "img",
    "div"
  ].join(",");
  const seen = new Set();
  const candidates = [];
  const state = window.__podReadinessObserver || null;
  const events = state && Array.isArray(state.events) ? state.events : [];
  const elapsedMs = state ? Math.round(performance.now() - state.startedAt) : null;
  const timingFor = (item) => {
    let firstSeen = null;
    let lastSeen = null;
    let lastKind = "";
    for (const event of events) {
      if (!event || !event.selector) continue;
      const evSelector = String(event.selector || "");
      const selector = String(item.selector || "");
      const matches = evSelector === selector || evSelector.indexOf(selector) >= 0 || selector.indexOf(evSelector) >= 0;
      if (!matches) continue;
      const at = Number(event.at_ms || 0);
      if (!Number.isFinite(at)) continue;
      if (firstSeen === null || at < firstSeen) firstSeen = at;
      if (lastSeen === null || at > lastSeen) {
        lastSeen = at;
        lastKind = String(event.kind || "");
      }
    }
    return {
      first_seen_ms: firstSeen,
      last_seen_ms: lastSeen,
      last_event_kind: lastKind,
      page_elapsed_ms: elapsedMs
    };
  };
  for (const el of Array.from(document.querySelectorAll(query))) {
    const item = describe(el, "candidate");
    if (!item || seen.has(item.selector)) continue;
    seen.add(item.selector);
    candidates.push(Object.assign(item, timingFor(item)));
  }
  candidates.sort((a, b) => {
    if (b.score !== a.score) return b.score - a.score;
    return (a.rect.y || 0) - (b.rect.y || 0);
  });
  return JSON.stringify({candidates: candidates.slice(0, 48)});
})()
"""


def _profile_ready_selector(profile: dict[str, Any] | None) -> str | None:
    if not isinstance(profile, dict):
        return None
    selector = str(profile.get("ready_selector") or "").strip()
    confirmations = int(profile.get("ready_selector_confirmations") or 0)
    if selector and confirmations >= 3:
        return selector
    return None


def _record_observation(
    *,
    key: str,
    url: str | None,
    record: dict | None,
    result: dict[str, Any],
) -> None:
    ready_element = result.get("ready_element")
    selector = ""
    if isinstance(ready_element, dict):
        selector = str(ready_element.get("selector") or "").strip()
    data = _load_profiles()
    profiles = data.setdefault("profiles", {})
    if not isinstance(profiles, dict):
        profiles = {}
        data["profiles"] = profiles
    profile = profiles.setdefault(key, {})
    if not isinstance(profile, dict):
        profile = {}
        profiles[key] = profile

    observations = profile.get("observations")
    if not isinstance(observations, list):
        observations = []
    observation = {
        "observed_at": _utc_now(),
        "url": str(url or "")[:500],
        "host": _host_for_url(url),
        "carrier": _carrier_for_record(record),
        "mode": str(result.get("mode") or ""),
        "elapsed_seconds": result.get("elapsed_seconds"),
        "quiet_seconds": result.get("quiet_seconds"),
        "event_count": result.get("event_count"),
        "ready_element": ready_element if isinstance(ready_element, dict) else None,
    }
    observations.append(observation)
    profile["observations"] = observations[-12:]
    profile["runs"] = int(profile.get("runs") or 0) + 1
    profile["host"] = _host_for_url(url)
    profile["carrier"] = _carrier_for_record(record)
    profile["updated_at"] = _utc_now()
    profile["quiet_seconds"] = DEFAULT_QUIET_SECONDS

    selectors = [
        str((obs.get("ready_element") or {}).get("selector") or "").strip()
        for obs in observations
        if isinstance(obs, dict)
    ]
    selectors = [s for s in selectors if s]
    last_three = selectors[-3:]
    if len(last_three) == 3 and len(set(last_three)) == 1:
        profile["ready_selector"] = last_three[-1]
        profile["ready_selector_confirmations"] = 3
        profile["ready_element"] = ready_element
    elif len(selectors) >= 3:
        selector_counts = Counter(selectors[-8:])
        selected, count = selector_counts.most_common(1)[0]
        if count >= 3:
            profile["ready_selector"] = selected
            profile["ready_selector_confirmations"] = count
            if selector == selected:
                profile["ready_element"] = ready_element

    try:
        _save_profiles(data)
    except OSError:
        pass


def record_user_selected_ready_element(
    *,
    url: str | None,
    record: dict | None,
    selected_element: dict[str, Any],
    elapsed_seconds: float | None = None,
) -> None:
    """Persist the user's chosen DOM selector as the ready marker for this carrier page."""
    key = _profile_key(url, record)
    selector = str(selected_element.get("selector") or "").strip()
    if not selector:
        return
    result = {
        "mode": "user-selected",
        "ready_selector": selector,
        "ready_element": dict(selected_element),
        "elapsed_seconds": elapsed_seconds,
        "quiet_seconds": None,
        "event_count": None,
    }
    _record_observation(key=key, url=url, record=record, result=result)
    _append_event_log(
        key,
        {
            "kind": "user-selected-ready-element",
            "selector": selector,
            "tag": selected_element.get("tag"),
            "text": str(selected_element.get("text") or "")[:180],
        },
    )


def wait_for_pod_dom_ready(
    ws_url: str,
    *,
    url: str | None,
    record: dict | None,
    timeout_seconds: float = 75.0,
    quiet_seconds: float = DEFAULT_QUIET_SECONDS,
    profile_grace_seconds: float = DEFAULT_PROFILE_GRACE_SECONDS,
    notify=None,
) -> dict[str, Any]:
    """Wait until a tracking page is ready, using learned selector profiles when available."""
    key = _profile_key(url, record)
    data = _load_profiles()
    profiles = data.get("profiles") if isinstance(data, dict) else {}
    profile = profiles.get(key) if isinstance(profiles, dict) else None
    ready_selector = _profile_ready_selector(profile if isinstance(profile, dict) else None)
    _eval_json(ws_url, _OBSERVER_SCRIPT)

    start = time.monotonic()
    deadline = start + max(float(timeout_seconds), quiet_seconds + 1.0)
    profile_seen_at: float | None = None
    event_count = 0
    last_snapshot: dict[str, Any] | None = None
    if notify is not None:
        try:
            if ready_selector:
                notify("progress", "Readiness profile found. Waiting for the learned page element...")
            else:
                notify("progress", "Watching the tracking page DOM until it is quiet for 5 seconds...")
        except Exception:
            pass

    while time.monotonic() < deadline:
        snapshot = _eval_json(ws_url, _SNAPSHOT_SCRIPT)
        if snapshot:
            last_snapshot = snapshot
            new_events = snapshot.get("new_events")
            if isinstance(new_events, list):
                for event in new_events:
                    if isinstance(event, dict):
                        event_count += 1
                        _append_event_log(key, event)

        if ready_selector:
            visible = _eval_json(ws_url, _selector_visible_script(ready_selector), timeout=3.0)
            if visible and visible.get("visible") is True:
                if profile_seen_at is None:
                    profile_seen_at = time.monotonic()
                    _append_event_log(
                        key,
                        {
                            "kind": "profile-selector-visible",
                            "selector": ready_selector,
                            "at_ms": int((profile_seen_at - start) * 1000),
                            "text": str(visible.get("text") or "")[:180],
                        },
                    )
                if time.monotonic() - profile_seen_at >= profile_grace_seconds:
                    result = {
                        "mode": "profile-hit",
                        "profile_key": key,
                        "ready_selector": ready_selector,
                        "ready_element": {
                            "selector": ready_selector,
                            "tag": visible.get("tag"),
                            "text": visible.get("text"),
                            "visible": True,
                            "rect": visible.get("rect"),
                        },
                        "elapsed_seconds": round(time.monotonic() - start, 3),
                        "quiet_seconds": round(float((last_snapshot or {}).get("quiet_ms") or 0) / 1000, 3),
                        "event_count": event_count,
                    }
                    _record_observation(key=key, url=url, record=record, result=result)
                    return result

        if last_snapshot:
            ready_state = str(last_snapshot.get("readyState") or "")
            text_length = int(last_snapshot.get("text_length") or 0)
            element_count = int(last_snapshot.get("element_count") or 0)
            quiet_ms = int(last_snapshot.get("quiet_ms") or 0)
            elapsed = time.monotonic() - start
            page_has_content = text_length >= MIN_READY_TEXT_CHARS or element_count >= 35
            state_ready = ready_state in ("interactive", "complete") or elapsed >= quiet_seconds
            if page_has_content and state_ready and quiet_ms >= int(quiet_seconds * 1000):
                candidate = last_snapshot.get("candidate")
                ready_element = candidate if isinstance(candidate, dict) else None
                result = {
                    "mode": "quiet-window",
                    "profile_key": key,
                    "ready_selector": str((ready_element or {}).get("selector") or ""),
                    "ready_element": ready_element,
                    "elapsed_seconds": round(elapsed, 3),
                    "quiet_seconds": round(quiet_ms / 1000, 3),
                    "event_count": event_count,
                }
                _record_observation(key=key, url=url, record=record, result=result)
                return result

        time.sleep(0.45)

    candidate = (last_snapshot or {}).get("candidate") if last_snapshot else None
    ready_element = candidate if isinstance(candidate, dict) else None
    result = {
        "mode": "timeout",
        "profile_key": key,
        "ready_selector": str((ready_element or {}).get("selector") or ""),
        "ready_element": ready_element,
        "elapsed_seconds": round(time.monotonic() - start, 3),
        "quiet_seconds": round(float((last_snapshot or {}).get("quiet_ms") or 0) / 1000, 3),
        "event_count": event_count,
    }
    _record_observation(key=key, url=url, record=record, result=result)
    return result


def highlight_pod_ready_element(
    ws_url: str,
    *,
    url: str | None,
    record: dict | None,
    readiness: dict[str, Any] | None = None,
) -> dict[str, Any] | None:
    """Draw the blue debug box around the learned or current likely-ready element."""
    selector = ""
    if isinstance(readiness, dict):
        ready_element = readiness.get("ready_element")
        if isinstance(ready_element, dict):
            selector = str(ready_element.get("selector") or "").strip()
        if not selector:
            selector = str(readiness.get("ready_selector") or "").strip()
    if not selector:
        data = _load_profiles()
        profiles = data.get("profiles") if isinstance(data, dict) else {}
        profile = profiles.get(_profile_key(url, record)) if isinstance(profiles, dict) else None
        selector = _profile_ready_selector(profile if isinstance(profile, dict) else None) or ""
    result = _eval_json(ws_url, _overlay_script(selector), timeout=4.0)
    key = _profile_key(url, record)
    if result:
        _append_event_log(
            key,
            {
                "kind": "debug-overlay",
                "selector": selector,
                "overlay_result": result,
            },
        )
    return result


def pod_selector_candidates(
    ws_url: str,
    *,
    url: str | None,
    record: dict | None,
    readiness: dict[str, Any] | None = None,
) -> list[dict[str, Any]]:
    """Return visible candidate elements, with the learned/current ready element first when possible."""
    result = _eval_json(ws_url, _CANDIDATES_SCRIPT, timeout=5.0)
    raw_candidates = result.get("candidates") if isinstance(result, dict) else []
    candidates: list[dict[str, Any]] = [
        dict(item) for item in raw_candidates if isinstance(item, dict) and str(item.get("selector") or "").strip()
    ][:MAX_SELECTOR_CANDIDATES]

    preferred_selector = ""
    if isinstance(readiness, dict):
        ready_element = readiness.get("ready_element")
        if isinstance(ready_element, dict):
            preferred_selector = str(ready_element.get("selector") or "").strip()
        if not preferred_selector:
            preferred_selector = str(readiness.get("ready_selector") or "").strip()
    if not preferred_selector:
        data = _load_profiles()
        profiles = data.get("profiles") if isinstance(data, dict) else {}
        profile = profiles.get(_profile_key(url, record)) if isinstance(profiles, dict) else None
        preferred_selector = _profile_ready_selector(profile if isinstance(profile, dict) else None) or ""

    if preferred_selector:
        for index, item in enumerate(candidates):
            if str(item.get("selector") or "") == preferred_selector:
                if index != 0:
                    candidates.insert(0, candidates.pop(index))
                break
        else:
            selected = _eval_json(ws_url, _selector_visible_script(preferred_selector), timeout=3.0)
            if selected and selected.get("visible") is True:
                candidates.insert(
                    0,
                    {
                        "source": "profile",
                        "selector": preferred_selector,
                        "tag": selected.get("tag"),
                        "text": selected.get("text"),
                        "score": 999,
                        "rect": selected.get("rect"),
                    },
                )
    return candidates[:MAX_SELECTOR_CANDIDATES]


def highlight_pod_selector(
    ws_url: str,
    *,
    selector: str | None,
    url: str | None,
    record: dict | None,
) -> dict[str, Any] | None:
    """Draw the blue selection box around *selector* or the best visible fallback."""
    result = _eval_json(ws_url, _overlay_script(selector), timeout=4.0)
    if result:
        _append_event_log(
            _profile_key(url, record),
            {
                "kind": "selector-overlay",
                "selector": selector or "",
                "overlay_result": result,
            },
        )
    return result


def remove_pod_debug_overlay(ws_url: str) -> dict[str, Any] | None:
    """Remove the blue POD selector/debug overlay before a capture is saved."""
    return _eval_json(ws_url, _REMOVE_OVERLAY_SCRIPT, timeout=3.0)
