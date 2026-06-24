"""
ui/settings_dialog.py
─────────────────────
Settings modal: theme, progress bar chars, feature flags.

Scroll behaviour mirrors BuilderTab: a tk.Canvas + Scrollbar with
mousewheel bindings on every child widget.

The theme list is collapsible and collapsed by default.
"""

import tkinter as tk
from tkinter import messagebox

from config import normalize_char
from state import AppState, DEFAULT_SLEEP, SLOW_SLEEP, SPEED_SLEEP
from state import DEFAULT_PROGRESS_FILLED, DEFAULT_PROGRESS_BORDER, DEFAULT_PROGRESS_EMPTY
from ui.circle_toggle import CircleToggle
from ui.dev_menu import open_dev_menu
from ui.theme import BG, PANEL, BORDER, ACCENT, ACCENT2, TEXT, SUBTEXT, FONT, THEMES, THEME_LABELS, colour_mode


def open_settings(root, state: AppState, cfg: dict, save_cb, reset_cb, theme_cb):
    win = tk.Toplevel(root)
    win.title("Settings")
    win.configure(bg=BG)
    win.resizable(True, True)
    root.update_idletasks()
    win.geometry(f"{root.winfo_width()}x{root.winfo_height()}+{root.winfo_x()}+{root.winfo_y()}")

    # ── Header (fixed, outside scroll) ───────────────────────────────────────
    hdr = tk.Frame(win, bg=PANEL, pady=10)
    hdr.pack(fill="x")
    tk.Label(hdr, text="Settings", bg=PANEL, fg=ACCENT2,
             font=(FONT, 12, "bold")).pack(side="left", padx=16)
    tk.Frame(win, bg=BORDER, height=1).pack(fill="x")

    # ── Scrollable canvas (builder-style) ────────────────────────────────────
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

    # ── Mousewheel (mirrors BuilderTab) ──────────────────────────────────────
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

    current_theme_name = THEME_LABELS.get(cfg.get("theme_mode", colour_mode), "")
    _preview_lbl = tk.Label(theme_header, text=f"",
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
            # collapse
            _theme_open.set(False)
            arrow_lbl.config(text="▶")
            restart_lbl.pack_forget()
            theme_body.pack_forget()
        else:
            # expand
            _theme_open.set(True)
            arrow_lbl.config(text="▼")
            restart_lbl.pack(after=theme_header)
            theme_body.pack(after=restart_lbl, fill="x")
        win.after(50, lambda: _bind_wheel(inner))

    for w in (theme_header, arrow_lbl, _preview_lbl):
        w.bind("<Button-1>", _toggle_theme_body)
    # also bind every child of theme_header
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

    # ── Progress bar chars ────────────────────────────────────────────────────
    section("Progress Bar Characters")
    chars_frame = tk.Frame(inner, bg=PANEL)
    chars_frame.pack()

    entries = []
    for col, (lbl, val) in enumerate((
            ("Filled", state.progress_filled),
            ("Border", state.progress_border),
            ("Empty",  state.progress_empty),
    )):
        tk.Label(chars_frame, text=lbl, bg=PANEL, fg=SUBTEXT,
                 font=(FONT, 8)).grid(row=0, column=col, padx=6)
        e = tk.Entry(
            chars_frame, width=4, justify="center",
            bg=PANEL, fg=TEXT, insertbackground=ACCENT, relief="flat",
            font=(FONT, 9), highlightthickness=1,
            highlightbackground=BORDER, highlightcolor=ACCENT,
        )
        e.insert(0, val)
        e.grid(row=1, column=col, padx=6, pady=2)
        entries.append(e)

    preview_frame = tk.Frame(inner, bg=PANEL)
    preview_frame.pack(pady=(4, 8))
    tk.Label(preview_frame, text="Preview", bg=PANEL, fg=SUBTEXT, font=(FONT, 8)).pack()

    previews = []
    for ch, color in (
            (state.progress_filled, TEXT),
            (state.progress_border, TEXT),
            (state.progress_empty,  ACCENT2),
    ):
        lbl = tk.Label(preview_frame, text=ch * 8, bg=BORDER, fg=color,
                       font=(FONT, 10), padx=4, pady=2)
        lbl.pack(side="left", padx=4)
        previews.append(lbl)

    def _apply_chars(_=None):
        state.progress_filled = normalize_char(entries[0].get(), DEFAULT_PROGRESS_FILLED)
        state.progress_border = normalize_char(entries[1].get(), DEFAULT_PROGRESS_BORDER)
        state.progress_empty  = normalize_char(entries[2].get(), DEFAULT_PROGRESS_EMPTY)
        for entry, ch in zip(entries, (state.progress_filled, state.progress_border, state.progress_empty)):
            if entry.get() != ch:
                entry.delete(0, tk.END)
                entry.insert(0, ch)
        for prev, ch in zip(previews, (state.progress_filled, state.progress_border, state.progress_empty)):
            prev.config(text=ch * 8)
        save_cb()

    for e in entries:
        e.bind("<KeyRelease>", _apply_chars)
        e.bind("<FocusOut>",   _apply_chars)

    # ── Feature flags ─────────────────────────────────────────────────────────
    section("Features")

    flags = [
        ("Trim Media Titles", "media_title_trim", "Removes words like official, lyrics, video"),
        ("Slow Mode",         "slow_mode",        f"Sets update interval to {SLOW_SLEEP:.0f}s"),
        ("Speed Mode",        "speed_mode",       f"Sets update interval to {SPEED_SLEEP:.1f}s"),
        ("Testing Mode",      "testing",          "Enables dev testing"),
    ]

    for label, attr, hint in flags:
        var = tk.BooleanVar(value=getattr(state, attr, False))

        def _changed(*_, a=attr, v=var):
            setattr(state, a, v.get())
            save_cb()
            _refresh_dev_btn()

        tk.Checkbutton(
            inner, text=label, bg=PANEL, fg=TEXT,
            variable=var, onvalue=True, offvalue=False,
            selectcolor=PANEL, activebackground=PANEL,
            font=(FONT, 9), command=_changed,
        ).pack(anchor="w", padx=20, pady=(8, 0))
        tk.Label(inner, text=hint, bg=PANEL, fg=SUBTEXT,
                 font=(FONT, 8)).pack(anchor="w", padx=40)

    # ── LHM startup preference ────────────────────────────────────────────────
    section("Libre Hardware Monitor")

    lhm_options = [
        ("always",  "Always start LHM on launch"),
        ("ask",     "Ask every time"),
        ("never",   "Never start / don't ask"),
    ]

    lhm_var = tk.StringVar(value=cfg.get("lhm_prompt", "ask"))

    lhm_frame = tk.Frame(inner, bg=PANEL)
    lhm_frame.pack(anchor="w", padx=20, pady=(0, 4))

    def _lhm_changed(*_):
        cfg["lhm_prompt"] = lhm_var.get()
        save_cb()

    for value, label_text in lhm_options:
        row = tk.Frame(lhm_frame, bg=PANEL)
        row.pack(anchor="w", pady=3)
        rb = tk.Radiobutton(
            row,
            text=label_text,
            variable=lhm_var,
            value=value,
            bg=PANEL, fg=TEXT,
            selectcolor=PANEL,
            activebackground=PANEL,
            activeforeground=ACCENT2,
            font=(FONT, 9),
            cursor="hand2",
            command=_lhm_changed,
        )
        rb.pack(side="left")

    # ── Action buttons ────────────────────────────────────────────────────────
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

    # Dev Menu button — shown only when Testing Mode is on
    dev_btn = tk.Button(
        btn_frame, text="Dev Menu", bg=PANEL, fg=ACCENT2, relief="flat",
        font=(FONT, 9, "bold"), activebackground=BORDER, activeforeground=ACCENT2,
        padx=12, pady=4, cursor="hand2",
        command=lambda: open_dev_menu(win, state, cfg, save_cb),
    )

    def _refresh_dev_btn():
        if getattr(state, "testing", False):
            dev_btn.pack(side="left", padx=(8, 0))
        else:
            dev_btn.pack_forget()

    _refresh_dev_btn()  # set initial visibility

    tk.Button(
        btn_frame, text="Close Settings", bg=ACCENT, fg=BG, relief="flat",
        font=(FONT, 9, "bold"), activebackground=ACCENT2, activeforeground=BG,
        padx=16, pady=4, cursor="hand2", command=win.destroy,
    ).pack(side="right")

    # bottom padding
    tk.Frame(inner, bg=PANEL, height=20).pack()