"""
ui/settings_dialog.py
─────────────────────
Settings modal for OSC-Router: theme and config reset.
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

    hdr = tk.Frame(win, bg=PANEL, pady=10)
    hdr.pack(fill="x")
    tk.Label(hdr, text="Settings", bg=PANEL, fg=ACCENT2, font=(FONT, 12, "bold")).pack(side="left", padx=16)
    tk.Frame(win, bg=BORDER, height=1).pack(fill="x")

    inner = tk.Frame(win, bg=PANEL)
    inner.pack(fill="both", expand=True, padx=20, pady=14)

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