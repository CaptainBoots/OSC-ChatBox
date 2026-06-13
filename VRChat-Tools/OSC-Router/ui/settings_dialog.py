"""
ui/settings_dialog.py
─────────────────────
Settings modal for OSC-Router: UI scale and config reset.
"""

import tkinter as tk
from tkinter import messagebox
from ui.theme import BG, PANEL, BORDER, ACCENT, ACCENT2, TEXT, SUBTEXT, FONT


def open_settings(root, save_cb, reset_cb, apply_scale_fn, get_scale_fn):
    win = tk.Toplevel(root)
    win.title("Settings")
    win.configure(bg=BG)
    win.resizable(True, True)
    root.update_idletasks()
    win.geometry(f"{root.winfo_width()}x{root.winfo_height()}+{root.winfo_x()}+{root.winfo_y()}")

    hdr = tk.Frame(win, bg=PANEL, pady=10)
    hdr.pack(fill="x")
    tk.Label(hdr, text="Settings", bg=PANEL, fg=ACCENT2, font=(FONT, 12, "bold")).pack(side="left", padx=16)
    tk.Frame(win, bg=BORDER, height=1).pack(fill="x")

    inner = tk.Frame(win, bg=PANEL)
    inner.pack(fill="both", expand=True, padx=20, pady=14)

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
        command=lambda: messagebox.askyesno("Reset", "Reset all settings to defaults?") and reset_cb(),
    ).pack(pady=6)

    tk.Button(
        win, text="Close", bg=ACCENT, fg=BG, relief="flat",
        cursor="hand2", font=(FONT, 10, "bold"),
        activebackground=ACCENT2, activeforeground=BG,
        command=win.destroy,
    ).pack(pady=12)
