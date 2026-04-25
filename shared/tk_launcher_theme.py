"""
Tk widgets and colors aligned with ``email_sorter_launcher`` Settings / ``THEME``.

Single source for palette: :data:`launcher_progress_ui.THEME`.
"""

from __future__ import annotations

import tkinter as tk
from tkinter import ttk

from launcher_progress_ui import THEME

_DANGER_BG = THEME["stop_fg"]
_DANGER_ACTIVE_BG = "#da3633"


def theme_font(name: str) -> tuple[str, int] | tuple[str, int, str]:
    fonts = {
        "titlebar": ("Segoe UI", 9),
        "title": ("Segoe UI", 14, "bold"),
        "body": ("Segoe UI", 10),
        "button": ("Segoe UI", 10, "bold"),
        "input": ("Segoe UI", 10),
    }
    return fonts[name]


def blend_hex_color(hex_color: str, target: str = "#ffffff", amount: float = 0.14) -> str:
    try:
        base = hex_color.strip().lstrip("#")
        dest = target.strip().lstrip("#")
        if len(base) != 6 or len(dest) != 6:
            return hex_color
        b = tuple(int(base[i : i + 2], 16) for i in (0, 2, 4))
        d = tuple(int(dest[i : i + 2], 16) for i in (0, 2, 4))
    except ValueError:
        return hex_color
    mixed = tuple(max(0, min(255, round(c + (t - c) * amount))) for c, t in zip(b, d))
    return "#" + "".join(f"{c:02x}" for c in mixed)


def add_button_hover(
    btn: tk.Button, *, normal_bg: str, hover_bg: str | None = None
) -> tk.Button:
    hover = hover_bg or blend_hex_color(normal_bg)

    def on_enter(_event: tk.Event) -> None:
        try:
            if str(btn.cget("state")) != tk.DISABLED:
                btn.configure(bg=hover)
        except tk.TclError:
            pass

    def on_leave(_event: tk.Event) -> None:
        try:
            btn.configure(bg=normal_bg)
        except tk.TclError:
            pass

    btn.bind("<Enter>", on_enter, add="+")
    btn.bind("<Leave>", on_leave, add="+")
    return btn


def make_flat_button(
    parent: tk.Misc,
    *,
    text: str,
    command,
    bg: str,
    fg: str = "#ffffff",
    active_bg: str | None = None,
    active_fg: str | None = None,
    padx: int = 18,
    pady: int = 7,
    width: int | None = None,
) -> tk.Button:
    opts: dict[str, object] = {
        "text": text,
        "command": command,
        "font": theme_font("button"),
        "bg": bg,
        "fg": fg,
        "activebackground": active_bg or bg,
        "activeforeground": active_fg or fg,
        "relief": tk.FLAT,
        "bd": 0,
        "highlightthickness": 0,
        "cursor": "hand2",
        "padx": padx,
        "pady": pady,
    }
    if width is not None:
        opts["width"] = width
    btn = tk.Button(parent, **opts)
    return add_button_hover(btn, normal_bg=bg, hover_bg=active_bg)


def settings_label_opts() -> dict[str, object]:
    return {"font": theme_font("body"), "fg": THEME["fg"], "bg": THEME["bg"]}


def settings_entry_opts() -> dict[str, object]:
    return {
        "font": theme_font("input"),
        "fg": THEME["fg"],
        "bg": THEME["surface"],
        "insertbackground": THEME["fg"],
        "relief": tk.FLAT,
        "highlightthickness": 1,
        "highlightbackground": THEME["border"],
        "highlightcolor": THEME["run_accent"],
    }


def danger_colors() -> tuple[str, str]:
    return _DANGER_BG, _DANGER_ACTIVE_BG


class SettingsStyleSwitch(tk.Frame):
    """On/off pill (IntVar 0/1) matching Settings dialog."""

    def __init__(self, parent: tk.Misc, variable: tk.IntVar) -> None:
        try:
            bg = parent.cget("bg")
        except tk.TclError:
            bg = "SystemButtonFace"
        super().__init__(parent, bg=bg)
        self._var = variable
        self._cv = tk.Canvas(
            self,
            width=48,
            height=26,
            highlightthickness=0,
            bg=bg,
            cursor="hand2",
        )
        self._cv.pack(side=tk.LEFT)
        self._var.trace_add("write", lambda *_: self._draw())
        self._cv.bind("<Button-1>", self._toggle)
        self._draw()

    def _toggle(self, _evt: object | None = None) -> None:
        self._var.set(1 - (1 if self._var.get() else 0))

    def _draw(self) -> None:
        self._cv.delete("all")
        on = bool(self._var.get())
        color = THEME["excel_accent"] if on else THEME["track"]
        self._cv.create_rectangle(2, 6, 46, 20, fill=color, outline="", width=0)
        cx = 33 if on else 15
        self._cv.create_oval(cx - 9, 4, cx + 9, 22, fill="#ffffff", outline=THEME["border"], width=1)


def apply_launcher_theme_root(win: tk.Misc) -> None:
    win.configure(bg=THEME["bg"])


def configure_launcher_ttk_styles(win: tk.Misc) -> ttk.Style:
    style = ttk.Style(win)
    try:
        style.theme_use("clam")
    except tk.TclError:
        pass
    style.configure("Launcher.TFrame", background=THEME["bg"])
    style.configure(
        "Launcher.TLabel",
        background=THEME["bg"],
        foreground=THEME["fg"],
        font=theme_font("body"),
    )
    style.configure(
        "Launcher.TLabelframe",
        background=THEME["bg"],
        foreground=THEME["fg"],
        font=theme_font("body"),
    )
    style.configure("Launcher.TLabelframe.Label", background=THEME["bg"], foreground=THEME["muted"])
    style.configure(
        "Launcher.Treeview",
        background=THEME["surface"],
        fieldbackground=THEME["surface"],
        foreground=THEME["fg"],
        font=theme_font("body"),
    )
    style.configure(
        "Launcher.Treeview.Heading",
        background=THEME["track"],
        foreground=THEME["fg"],
        font=theme_font("button"),
    )
    style.map(
        "Launcher.Treeview",
        background=[("selected", THEME["run_accent_dim"])],
        foreground=[("selected", "#ffffff")],
    )
    return style


def launcher_scrollbar(master: tk.Misc, orient: str, command) -> tk.Scrollbar:
    return tk.Scrollbar(
        master,
        orient=orient,
        command=command,
        troughcolor=THEME["bg"],
        bg=THEME["track"],
        activebackground=THEME["border"],
        highlightthickness=0,
        bd=0,
        width=14,
        elementborderwidth=1,
        jump=0,
    )


def style_scrolled_text(widget: tk.Text) -> None:
    widget.configure(
        bg=THEME["surface"],
        fg=THEME["fg"],
        insertbackground=THEME["fg"],
        relief=tk.FLAT,
        highlightthickness=0,
        borderwidth=0,
    )
