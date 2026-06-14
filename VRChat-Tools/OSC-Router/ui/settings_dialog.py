"""
ui/settings_dialog.py
─────────────────────
Settings modal: Colour Mode and Defaults.
"""

import tkinter as tk
from tkinter import messagebox
from ui.theme import BG, PANEL, BORDER, ACCENT, ACCENT2, TEXT, SUBTEXT, FONT


def open_settings(root, apply_color_fn, get_color_fn, save_cb, reset_cb):
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

    # Main Container
    inner = tk.Frame(win, bg=PANEL)
    inner.pack(fill="both", expand=True, padx=20, pady=14)

    # Section generator helper
    def section(title: str):
        f = tk.Frame(inner, bg=PANEL, pady=8)
        f.pack(fill="x", pady=(16, 0))
        tk.Label(f, text=title.upper(), bg=PANEL, fg=ACCENT2, font=(FONT, 9, "bold")).pack(side="left")
        div = tk.Frame(inner, bg=BORDER, height=1)
        div.pack(fill="x", pady=(0, 12))

    # ── Colour Mode Component ─────────────────────────────────────────────────
    section("Appearance")
    row_color = tk.Frame(inner, bg=PANEL, pady=10)
    row_color.pack(fill="x")
    tk.Label(row_color, text="Colour Mode:", bg=PANEL, fg=TEXT, font=(FONT, 10)).pack(side="left")

    current_mode = get_color_fn()
    mode_var = tk.StringVar(value=current_mode)

    def on_mode_change(*args):
        apply_color_fn(mode_var.get())

    mode_menu = tk.OptionMenu(row_color, mode_var, "light", "old", "new", command=on_mode_change)
    mode_menu.config(
        bg=BORDER, fg=TEXT, activebackground=ACCENT, activeforeground=TEXT,
        highlightthickness=0, font=(FONT, 10), bd=0, relief="flat"
    )
    mode_menu["menu"].config(
        bg=PANEL, fg=TEXT, activebackground=ACCENT, activeforeground=TEXT,
        font=(FONT, 10), bd=0, relief="flat"
    )
    mode_menu.pack(side="right")

    # Restart Notice
    lbl_notice = tk.Label(
        inner, text="* Changing color modes requires an application restart.",
        bg=PANEL, fg=SUBTEXT, font=(FONT, 8, "italic"), anchor="w"
    )
    lbl_notice.pack(fill="x", pady=(10, 0))

    # ── Actions Footer ────────────────────────────────────────────────────────
    def _trigger_reset():
        if messagebox.askyesno("Reset", "Are you sure you want to restore default values?", parent=win):
            reset_cb()
            win.destroy()

    def _close_and_save():
        save_cb()
        win.destroy()

    btn_frame = tk.Frame(win, bg=PANEL, pady=12)
    btn_frame.pack(fill="x", side="bottom")
    tk.Frame(win, bg=BORDER, height=1).pack(fill="x", side="bottom")

    tk.Button(
        btn_frame, text="Restore Defaults", bg=PANEL, fg=TEXT, relief="flat",
        font=(FONT, 9), activebackground=BORDER, activeforeground=TEXT,
        padx=12, pady=4, cursor="hand2", command=_trigger_reset
    ).pack(side="left", padx=16)

    tk.Button(
        btn_frame, text="Save & Close", bg=ACCENT, fg=BG, relief="flat",
        font=(FONT, 9, "bold"), activebackground=ACCENT2, activeforeground=BG,
        padx=16, pady=4, cursor="hand2", command=_close_and_save
    ).pack(side="right", padx=16)