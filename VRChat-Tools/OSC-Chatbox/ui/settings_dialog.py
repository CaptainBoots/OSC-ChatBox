"""
ui/settings_dialog.py
─────────────────────
Settings modal: theme, progress bar chars, feature flags.
"""

import tkinter as tk
from tkinter import messagebox

from config import normalize_char
from state import AppState, DEFAULT_SLEEP, SLOW_SLEEP, SPEED_SLEEP
from state import DEFAULT_PROGRESS_FILLED, DEFAULT_PROGRESS_BORDER, DEFAULT_PROGRESS_EMPTY
from sympy import true
from ui.circle_toggle import CircleToggle
from ui.theme import BG, PANEL, BORDER, ACCENT, ACCENT2, TEXT, SUBTEXT, FONT, THEMES, THEME_LABELS, colour_mode


def open_settings(root, state: AppState, cfg: dict, save_cb, reset_cb, theme_cb):
    win = tk.Toplevel(root)
    win.title("Settings")
    win.configure(bg=BG)
    win.resizable(True, True)
    root.update_idletasks()
    win.geometry(f"{root.winfo_width()}x{root.winfo_height()}+{root.winfo_x()}+{root.winfo_y()}")

    # Header
    hdr = tk.Frame(win, bg=PANEL, pady=10)
    hdr.pack(fill="x")
    tk.Label(hdr, text="Settings", bg=PANEL, fg=ACCENT2, font=(FONT, 12, "bold")).pack(side="left", padx=16)
    tk.Frame(win, bg=BORDER, height=1).pack(fill="x")

    # Scrollable content
    canvas = tk.Canvas(win, bg=PANEL, highlightthickness=0)
    vsb    = tk.Scrollbar(win, orient="vertical", command=canvas.yview)
    canvas.configure(yscrollcommand=vsb.set)
    vsb.pack(side="right", fill="y")
    canvas.pack(fill="both", expand=True, padx=20, pady=14)

    inner = tk.Frame(canvas, bg=PANEL)
    inner.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
    canvas.create_window((0, 0), window=inner, anchor="nw")

    # The helper function now only takes care of creating the section heading
    def section(label):
        tk.Label(inner, text=label, bg=PANEL, fg=ACCENT2, font=(FONT, 10, "bold")).pack(pady=(16, 6))

    # ── Theme ─────────────────────────────────────────────────────────────────
    # This is now properly placed outside of the def section block
    section("Theme")
    tk.Label(inner, text="Restart required to apply", bg=PANEL, fg=SUBTEXT,
             font=(FONT, 8)).pack()

    current_theme = cfg.get("theme_mode", colour_mode)
    theme_state = {"selected": current_theme}
    theme_rows: list[dict] = []

    def _refresh_theme_rows():
        for row_data in theme_rows:
            is_selected = row_data["mode"] == theme_state["selected"]
            row_data["toggle"].set(is_selected)
            row_data["label"].config(fg=ACCENT2 if is_selected else TEXT)

    def _select_theme(mode):
        theme_state["selected"] = mode
        _refresh_theme_rows()
        theme_cb(mode)

    theme_list_frame = tk.Frame(inner, bg=PANEL)
    theme_list_frame.pack(anchor="w", padx=20, pady=(6, 0))

    for mode, label_text in THEME_LABELS.items():
        row = tk.Frame(theme_list_frame, bg=PANEL, cursor="hand2")
        row.pack(anchor="w", pady=3, fill="x")

        toggle = CircleToggle(
            row,
            enabled=(mode == current_theme),
            bg=PANEL,
            color=ACCENT,
        )
        toggle.pack(side="left", padx=(0, 4))

        lbl = tk.Label(row, text=label_text, bg=PANEL, font=(FONT, 9), cursor="hand2")
        lbl.pack(side="left", padx=(4, 8))

        # Swatch preview of the palette's accent colours
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
        tk.Label(chars_frame, text=lbl, bg=PANEL, fg=SUBTEXT, font=(FONT, 8)).grid(row=0, column=col, padx=6)
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
    for ch, color in ((state.progress_filled, TEXT), (state.progress_border, TEXT), (state.progress_empty, ACCENT2)):
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
        ("Testing Mode",        "testing",       "Enables dev testing"),
    ]

    for label, attr, hint in flags:
        var = tk.BooleanVar(value=getattr(state, attr, False))

        def _changed(*_, a=attr, v=var):
            setattr(state, a, v.get())
            save_cb()

        tk.Checkbutton(
            inner, text=label, bg=PANEL, fg=TEXT,
            variable=var, onvalue=True, offvalue=False,
            selectcolor=PANEL, activebackground=PANEL,
            font=(FONT, 9), command=_changed,
        ).pack(anchor="w", padx=20, pady=(8, 0))
        tk.Label(inner, text=hint, bg=PANEL, fg=SUBTEXT, font=(FONT, 8)).pack(anchor="w", padx=40)

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
        padx=12, pady=4, cursor="hand2", command=_trigger_reset
    ).pack(side="left")

    tk.Button(
        btn_frame, text="Close Settings", bg=ACCENT, fg=BG, relief="flat",
        font=(FONT, 9, "bold"), activebackground=ACCENT2, activeforeground=BG,
        padx=16, pady=4, cursor="hand2", command=win.destroy
    ).pack(side="right")