"""
ui/settings_dialog.py
─────────────────────
Settings modal: Progress bar chars and feature flags.
"""

import tkinter as tk
from tkinter import messagebox
from ui.theme import BG, PANEL, BORDER, ACCENT, ACCENT2, TEXT, SUBTEXT, GREEN, RED, YELLOW, FONT


def open_settings(root, save_cb, reset_cb, apply_scale_fn=None, get_scale_fn=None):
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

    # ── Config reset ──────────────────────────────────────────────────────────
    section("Config")
    tk.Button(
        inner, text="Reset to Defaults",
        bg=RED, fg=BG, relief="flat",
        activebackground=BORDER, activeforeground=TEXT,
        cursor="hand2", font=(FONT, 9, "bold"),
        command=lambda: messagebox.askyesno("Config Reset", "Are you sure?") and reset_cb(),
    ).pack(pady=6)

    tk.Button(
        win, text="Close", bg=ACCENT, fg=BG, relief="flat",
        cursor="hand2", font=(FONT, 10, "bold"),
        activebackground=ACCENT2, activeforeground=BG,
        command=win.destroy,
    ).pack(pady=12)


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
