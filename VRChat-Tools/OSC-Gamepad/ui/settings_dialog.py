"""
ui/settings_dialog.py
─────────────────────
Settings modal for OSC-Gamepad: Colour Mode selection.
"""

import tkinter as tk
from ui.theme import BG, PANEL, BORDER, ACCENT, ACCENT2, TEXT, SUBTEXT, FONT


def open_settings(root, apply_color_fn, get_color_fn, save_cb):
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

    # Main Container
    inner = tk.Frame(win, bg=PANEL)
    inner.pack(fill="both", expand=True, padx=20, pady=14)

    # ─────────────────────────────────────────────────────────────────
    # Colour Mode Component
    # ─────────────────────────────────────────────────────────────────
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

    # Restart Notice Notice
    lbl_notice = tk.Label(
        inner, text="* Changing color modes requires an application restart.",
        bg=PANEL, fg=SUBTEXT, font=(FONT, 8, "italic"), anchor="w"
    )
    lbl_notice.pack(fill="x", pady=(10, 0))

    # ─────────────────────────────────────────────────────────────────
    # Footer & Save Actions
    # ─────────────────────────────────────────────────────────────────
    footer = tk.Frame(win, bg=PANEL, pady=10)
    footer.pack(fill="x", side="bottom")
    tk.Frame(win, bg=BORDER, height=1).pack(fill="x", side="bottom")

    def close_and_save():
        save_cb()
        win.destroy()

    btn_save = tk.Button(
        footer, text="Save & Close", command=close_and_save,
        bg=ACCENT, fg=TEXT, activebackground=ACCENT2, activeforeground=TEXT,
        font=(FONT, 10, "bold"), bd=0, padx=15, pady=5, cursor="hand2"
    )
    btn_save.pack(side="right", padx=16)