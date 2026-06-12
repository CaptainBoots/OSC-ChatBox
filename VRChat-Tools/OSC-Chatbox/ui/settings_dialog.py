"""
ui/settings_dialog.py
─────────────────────
Settings modal: Progress bar chars and feature flags.
"""

import tkinter as tk
from tkinter import messagebox

from config import normalize_char
from state import AppState, DEFAULT_SLEEP, SLOW_SLEEP, SPEED_SLEEP
from state import DEFAULT_PROGRESS_FILLED, DEFAULT_PROGRESS_BORDER, DEFAULT_PROGRESS_EMPTY
from ui.theme import BG, PANEL, BORDER, ACCENT, ACCENT2, TEXT, SUBTEXT, FONT


def open_settings(root, state: AppState, save_cb, reset_cb):
    win = tk.Toplevel(root)
    win.title("Settings")
    win.configure(bg=BG)
    win.resizable(True, True)
    root.update_idletasks()
    win.geometry(f"{root.winfo_width()}x{root.winfo_height()}+{root.winfo_x()}+{root.winfo_y()}")

    # Header section
    hdr = tk.Frame(win, bg=PANEL, pady=10)
    hdr.pack(fill="x")
    tk.Label(hdr, text="Settings", bg=PANEL, fg=ACCENT2, font=(FONT, 12, "bold")).pack(side="left", padx=16)
    tk.Frame(win, bg=BORDER, height=1).pack(fill="x")

    # Scrollable container frames
    canvas = tk.Canvas(win, bg=PANEL, highlightthickness=0)
    vsb    = tk.Scrollbar(win, orient="vertical", command=canvas.yview)
    canvas.configure(yscrollcommand=vsb.set)

    vsb.pack(side="right", fill="y")
    canvas.pack(side="left", fill="both", expand=True)

    inner = tk.Frame(canvas, bg=PANEL)
    canvas_win = canvas.create_window((0, 0), window=inner, anchor="nw")

    def _on_canvas_configure(e):
        canvas.itemconfigure(canvas_win, width=e.width)

    def _on_inner_configure(_):
        canvas.configure(scrollregion=canvas.bbox("all"))

    canvas.bind("<Configure>", _on_canvas_configure)
    inner.bind("<Configure>", _on_inner_configure)

    # Mousewheel scrolling hook
    def _on_mousewheel(e):
        canvas.yview_scroll(int(-1 * (e.delta / 120)), "units")
    win.bind("<MouseWheel>", _on_mousewheel)

    # Section generator helper
    def section(title: str):
        f = tk.Frame(inner, bg=PANEL, pady=8)
        f.pack(fill="x", padx=16, pady=(16, 0))
        tk.Label(f, text=title.upper(), bg=PANEL, fg=ACCENT2, font=(FONT, 9, "bold")).pack(side="left")
        div = tk.Frame(inner, bg=BORDER, height=1)
        div.pack(fill="x", padx=16, pady=(0, 12))

    # ── Custom Characters ─────────────────────────────────────────────────────
    section("Custom Characters")

    char_frame = tk.Frame(inner, bg=PANEL)
    char_frame.pack(fill="x", padx=16)
    char_frame.columnconfigure(1, weight=1)

    fields = [
        ("Progress Bar Filled", "progress_filled", state.progress_filled),
        ("Progress Bar Border", "progress_border", state.progress_border),
        ("Progress Bar Empty",  "progress_empty",  state.progress_empty),
    ]

    entries = []
    for i, (label, attr, current_val) in enumerate(fields):
        tk.Label(char_frame, text=label, bg=PANEL, fg=TEXT, font=(FONT, 9)).grid(row=i, column=0, sticky="w", pady=6)

        var = tk.StringVar(value=current_val)
        e = tk.Entry(
            char_frame, textvariable=var, width=6, bg=BG, fg=TEXT,
            insertbackground=ACCENT, font=(FONT, 9, "bold"), relief="flat",
            highlightthickness=1, highlightbackground=BORDER, highlightcolor=ACCENT, justify="center"
        )
        e.grid(row=i, column=1, sticky="w", padx=12, pady=6)
        entries.append((var, attr))

    # Live preview layout for configuration characters
    prev_frame = tk.Frame(inner, bg=PANEL, pady=8)
    prev_frame.pack(fill="x", padx=16)
    tk.Label(prev_frame, text="Live Preview:", bg=PANEL, fg=SUBTEXT, font=(FONT, 8)).pack(side="left")

    b_l = state.progress_border if state.progress_border else ""
    b_r = state.progress_border if state.progress_border else ""
    init_str = f"{b_l}{state.progress_filled * 5}{state.progress_empty * 3}{b_r}"

    prev = tk.Label(prev_frame, text=init_str, bg=PANEL, fg=TEXT, font=(FONT, 9, "bold"))
    prev.pack(side="left", padx=8)

    def _apply_chars(*_):
        for var, attr in entries:
            raw = var.get()
            val = normalize_char(raw) if attr != "progress_border" else raw
            setattr(state, attr, val)

        ch = state.progress_filled if state.progress_filled else "■"
        border = state.progress_border if state.progress_border else ""
        empty = state.progress_empty if state.progress_empty else " "

        prev.config(text=f"{border}{ch * 5}{empty * 3}{border}")
        save_cb()

    for var, _ in entries:
        var.trace_add("write", _apply_chars)

    # ── Feature Flags ─────────────────────────────────────────────────────────
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

        chk = tk.Checkbutton(
            inner, text=label, bg=PANEL, fg=TEXT,
            variable=var, onvalue=True, offvalue=False,
            selectcolor=BG, activebackground=PANEL, activeforeground=TEXT,
            font=(FONT, 9), command=_changed, relief="flat", bd=0, highlightthickness=0
        )
        chk.pack(anchor="w", padx=20, pady=(8, 0))
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