from __future__ import annotations

import tkinter as tk
from tkinter import ttk

UI_BG = "#1e293b"
UI_BG_PANEL = "#0f172a"
UI_TREE_BG = "#334155"
UI_TREE_HEAD = "#60a5fa"
UI_TREE_HEAD_ACTIVE = "#2563eb"
UI_FG = "#f8fafc"
UI_FG_DIM = "#94a3b8"
UI_SEL = "#3b82f6"
UI_BTN = "#3b82f6"
UI_BTN_ACTIVE = "#2563eb"
UI_SCROLL_TROUGH = "#1e293b"
UI_SCROLL_THUMB = "#64748b"
UI_SCROLL_THUMB_ACTIVE = "#94a3b8"


def setup_dark_theme(root: tk.Misc) -> ttk.Style:
    root.configure(bg=UI_BG)
    style = ttk.Style(root)
    try:
        style.theme_use("clam")
    except tk.TclError:
        pass
    style.configure("TFrame", background=UI_BG)
    style.configure("TLabel", background=UI_BG, foreground=UI_FG)
    style.configure(
        "TButton",
        background=UI_BTN,
        foreground=UI_FG,
        borderwidth=0,
        focuscolor="none",
        padding=(14, 6),
    )
    style.map(
        "TButton",
        background=[("active", UI_BTN_ACTIVE), ("disabled", "#475569")],
        foreground=[("disabled", UI_FG_DIM)],
    )
    style.configure(
        "Treeview",
        background=UI_TREE_BG,
        fieldbackground=UI_TREE_BG,
        foreground=UI_FG,
        borderwidth=0,
        rowheight=26,
    )
    style.configure(
        "Treeview.Heading",
        background=UI_TREE_HEAD,
        foreground=UI_FG,
        borderwidth=0,
        relief="flat",
        anchor="w",
    )
    style.map(
        "Treeview.Heading",
        background=[("active", UI_TREE_HEAD_ACTIVE), ("pressed", UI_TREE_HEAD_ACTIVE)],
        foreground=[("active", UI_FG), ("pressed", UI_FG)],
    )
    style.map(
        "Treeview",
        background=[("selected", UI_SEL)],
        foreground=[("selected", UI_FG)],
    )
    return style


def dark_tk_scrollbar(master: tk.Misc, orient: str, command) -> tk.Scrollbar:
    return tk.Scrollbar(
        master,
        orient=orient,
        command=command,
        troughcolor=UI_SCROLL_TROUGH,
        bg=UI_SCROLL_THUMB,
        activebackground=UI_SCROLL_THUMB_ACTIVE,
        highlightthickness=0,
        bd=0,
        width=14,
        elementborderwidth=1,
        jump=0,
    )


def style_text_widget(widget: tk.Text) -> None:
    widget.configure(
        bg=UI_TREE_BG,
        fg=UI_FG,
        insertbackground=UI_FG,
        relief=tk.FLAT,
        highlightthickness=0,
        borderwidth=0,
    )
