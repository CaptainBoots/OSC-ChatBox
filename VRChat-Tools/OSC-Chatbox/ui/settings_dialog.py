"""
ui/settings_dialog.py
─────────────────────
Settings modal: UI scale, progress bar chars, feature flags.
"""

import tkinter as tk
from tkinter import messagebox

from config import normalize_char
from state import AppState, DEFAULT_SLEEP, SLOW_SLEEP, SPEED_SLEEP
from state import DEFAULT_PROGRESS_FILLED, DEFAULT_PROGRESS_BORDER, DEFAULT_PROGRESS_EMPTY
from ui.theme import BG, PANEL, BORDER, ACCENT, ACCENT2, TEXT, SUBTEXT, FONT


def open_settings(root, state: AppState, save_cb, reset_cb, apply_scale_fn, get_scale_fn):
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

    def section(label):
        tk.Label(inner, text=label, bg=PANEL, fg=ACCENT2, font=(FONT, 10, "bold")).pack(pady=(16, 6))

    # ── UI Scale ──────────────────────────────────────────────────────────────
    section("UI Scale")
    scale_var = tk.DoubleVar(value=get_scale_fn())
    pct_lbl   = tk.Label(inner, text="", bg=PANEL, fg=SUBTEXT, font=(FONT, 9))

    def _scale_changed(v):
        apply_scale_fn(float(v))
        pct_lbl.config(text=f"{int(float(v) * 100)}%")

    tk.Scale(
        inner, from_=0.7, to=2.0, resolution=0.05,
        orient="horizontal", variable=scale_var,
        bg=PANEL, fg=TEXT, troughcolor=BORDER, activebackground=ACCENT2,
        highlightthickness=0, sliderrelief="flat", length=300,
        command=_scale_changed,
    ).pack(pady=4)
    pct_lbl.pack()
    _scale_changed(scale_var.get())

    # ── Config reset ──────────────────────────────────────────────────────────
    section("Config")
    tk.Button(
        inner, text="Reset to Defaults",
        bg=PANEL, fg=SUBTEXT, relief="flat",
        activebackground=BORDER, activeforeground=TEXT,
        cursor="hand2", font=(FONT, 9, "bold"),
        command=lambda: messagebox.askyesno("Reset", "Reset all settings?") and reset_cb(),
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

    # ── Close ─────────────────────────────────────────────────────────────────
    tk.Button(
        win, text="Close", bg=ACCENT, fg=BG, relief="flat",
        cursor="hand2", font=(FONT, 10, "bold"),
        activebackground=ACCENT2, activeforeground=BG,
        command=win.destroy,
    ).pack(pady=12)