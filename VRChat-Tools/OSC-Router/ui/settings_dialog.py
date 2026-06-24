"""
ui/settings_dialog.py
─────────────────────
Settings modal for OSC-Router

Scroll behaviour mirrors BuilderTab: a tk.Canvas + Scrollbar with
mousewheel bindings on every child widget.

The theme list is collapsible and collapsed by default.
"""

import tkinter as tk
from tkinter import messagebox

from ui.circle_toggle import CircleToggle
from ui.theme import BG, PANEL, BORDER, ACCENT, ACCENT2, TEXT, SUBTEXT, FONT, THEMES, THEME_LABELS, colour_mode


def open_settings(root, cfg: dict, save_cb, reset_cb, theme_cb):
    win = tk.Toplevel(root)
    win.title("Settings")
    win.configure(bg=BG)
    win.resizable(True, True)
    root.update_idletasks()
    win.geometry(f"{root.winfo_width()}x{root.winfo_height()}+{root.winfo_x()}+{root.winfo_y()}")

    # ── Header (fixed, outside scroll) ───────────────────────────────────────
    hdr = tk.Frame(win, bg=PANEL, pady=10)
    hdr.pack(fill="x")
    tk.Label(hdr, text="Settings", bg=PANEL, fg=ACCENT2, font=(FONT, 12, "bold")).pack(side="left", padx=16)
    tk.Frame(win, bg=BORDER, height=1).pack(fill="x")

    # ── Scrollable canvas ─────────────────────────────────────────────────────
    scroll_frame = tk.Frame(win, bg=PANEL)
    scroll_frame.pack(fill="both", expand=True)
    scroll_frame.rowconfigure(0, weight=1)
    scroll_frame.columnconfigure(0, weight=1)

    canvas = tk.Canvas(scroll_frame, bg=PANEL, highlightthickness=0)
    canvas.grid(row=0, column=0, sticky="nsew")

    vsb = tk.Scrollbar(scroll_frame, orient="vertical", command=canvas.yview)
    vsb.grid(row=0, column=1, sticky="ns")
    canvas.configure(yscrollcommand=vsb.set)

    inner = tk.Frame(canvas, bg=PANEL)
    canvas_win = canvas.create_window((0, 0), window=inner, anchor="nw")

    def _on_inner_configure(_e):
        canvas.configure(scrollregion=canvas.bbox("all"))

    def _on_canvas_configure(e):
        canvas.itemconfigure(canvas_win, width=e.width)

    inner.bind("<Configure>", _on_inner_configure)
    canvas.bind("<Configure>", _on_canvas_configure)

    # ── Mousewheel ────────────────────────────────────────────────────────────
    def _on_wheel(event):
        if event.num == 4:
            canvas.yview_scroll(-1, "units")
        elif event.num == 5:
            canvas.yview_scroll(1, "units")
        elif event.delta:
            canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

    def _bind_wheel(widget):
        widget.bind("<MouseWheel>", _on_wheel, add="+")
        widget.bind("<Button-4>",   _on_wheel, add="+")
        widget.bind("<Button-5>",   _on_wheel, add="+")
        for child in widget.winfo_children():
            _bind_wheel(child)

    # Bind after all widgets are built
    win.after(100, lambda: _bind_wheel(inner))

    # ── Section helper ────────────────────────────────────────────────────────
    def section(label):
        tk.Label(inner, text=label, bg=PANEL, fg=ACCENT2,
                 font=(FONT, 10, "bold")).pack(pady=(16, 6))

    # ── Theme (collapsible, collapsed by default) ─────────────────────────────
    theme_header = tk.Frame(inner, bg=PANEL, cursor="hand2")
    theme_header.pack(fill="x", padx=20, pady=(16, 0))

    _theme_open = tk.BooleanVar(value=False)

    arrow_lbl = tk.Label(theme_header, text="▶", bg=PANEL, fg=ACCENT2,
                         font=(FONT, 12, "bold"), cursor="hand2")
    arrow_lbl.pack(side="left")

    tk.Label(theme_header, text="  Themes", bg=PANEL, fg=ACCENT2,
             font=(FONT, 12, "bold"), cursor="hand2").pack(side="left")

    _preview_lbl = tk.Label(theme_header, text="",
                            bg=PANEL, fg=SUBTEXT, font=(FONT, 9), cursor="hand2")
    _preview_lbl.pack(side="left")

    restart_lbl = tk.Label(inner, text="Restart required to apply", bg=PANEL,
                           fg=SUBTEXT, font=(FONT, 8))

    theme_body = tk.Frame(inner, bg=PANEL)

    current_theme = cfg.get("theme_mode", colour_mode)
    theme_state   = {"selected": current_theme}
    theme_rows: list[dict] = []

    def _refresh_theme_rows():
        for row_data in theme_rows:
            is_sel = row_data["mode"] == theme_state["selected"]
            row_data["toggle"].set(is_sel)
            row_data["label"].config(fg=ACCENT2 if is_sel else TEXT)

    def _select_theme(mode):
        theme_state["selected"] = mode
        _refresh_theme_rows()
        _preview_lbl.config(text=f"({THEME_LABELS.get(mode, mode)})")
        theme_cb(mode)

    theme_list_frame = tk.Frame(theme_body, bg=PANEL)
    theme_list_frame.pack(anchor="w", padx=20, pady=(4, 0))

    for mode, label_text in THEME_LABELS.items():
        row = tk.Frame(theme_list_frame, bg=PANEL, cursor="hand2")
        row.pack(anchor="w", pady=3, fill="x")

        toggle = CircleToggle(row, enabled=(mode == current_theme), bg=PANEL, color=ACCENT)
        toggle.pack(side="left", padx=(0, 4))

        lbl = tk.Label(row, text=label_text, bg=PANEL, font=(FONT, 9), cursor="hand2")
        lbl.pack(side="left", padx=(4, 8))

        swatch = tk.Frame(row, bg=PANEL, cursor="hand2")
        swatch.pack(side="left")
        for colour_key in ("BG", "PANEL", "ACCENT", "ACCENT2"):
            tk.Frame(swatch, bg=THEMES[mode][colour_key], width=14, height=14,
                     highlightthickness=1, highlightbackground=BORDER).pack(side="left", padx=1)

        row_data = {"mode": mode, "toggle": toggle, "label": lbl}
        theme_rows.append(row_data)

        for widget in (row, lbl, swatch, toggle):
            widget.bind("<Button-1>", lambda e, m=mode: _select_theme(m))

    _refresh_theme_rows()

    def _toggle_theme_body(*_):
        if _theme_open.get():
            _theme_open.set(False)
            arrow_lbl.config(text="▶")
            restart_lbl.pack_forget()
            theme_body.pack_forget()
        else:
            _theme_open.set(True)
            arrow_lbl.config(text="▼")
            restart_lbl.pack(after=theme_header)
            theme_body.pack(after=restart_lbl, fill="x")
        win.after(50, lambda: _bind_wheel(inner))

    for w in (theme_header, arrow_lbl, _preview_lbl):
        w.bind("<Button-1>", _toggle_theme_body)
    for child in theme_header.winfo_children():
        child.bind("<Button-1>", _toggle_theme_body)

    # ── Config reset ──────────────────────────────────────────────────────────
    section("Config")
    tk.Button(
        inner, text="Reset to Defaults",
        bg=PANEL, fg=SUBTEXT, relief="flat",
        activebackground=BORDER, activeforeground=TEXT,
        cursor="hand2", font=(FONT, 9, "bold"),
        command=lambda: messagebox.askyesno("Reset", "Reset all settings to defaults?") and reset_cb(),
    ).pack(pady=6)

    # ── Action Buttons ────────────────────────────────────────────────────────
    section("Actions")

    def _trigger_reset():
        if messagebox.askyesno("Reset", "Are you sure you want to restore default values?", parent=win):
            reset_cb()
            win.destroy()

    btn_frame = tk.Frame(inner, bg=PANEL, pady=12)
    btn_frame.pack(fill="x", padx=16)

    tk.Button(
        btn_frame, text="Restore Defaults", bg=BG, fg=TEXT, relief="flat",
        font=(FONT, 9), activebackground=BORDER, activeforeground=TEXT,
        padx=12, pady=4, cursor="hand2", command=_trigger_reset,
    ).pack(side="left")

    tk.Button(
        btn_frame, text="Close Settings", bg=ACCENT, fg=BG, relief="flat",
        font=(FONT, 9, "bold"), activebackground=ACCENT2, activeforeground=BG,
        padx=16, pady=4, cursor="hand2", command=win.destroy,
    ).pack(side="right")

    # bottom padding
    tk.Frame(inner, bg=PANEL, height=20).pack()