"""
ui/dev_menu.py
──────────────
Developer menu modal: testing tools and internal diagnostics.

Structure mirrors settings_dialog.py — scrollable canvas with a
fixed header. Opened from the Settings dialog when Testing Mode
is enabled.
"""

import tkinter as tk
from tkinter import messagebox

from state import AppState
from ui.theme import BG, PANEL, BORDER, ACCENT, ACCENT2, TEXT, SUBTEXT, FONT


def open_dev_menu(root, state: AppState, cfg: dict, save_cb):
    win = tk.Toplevel(root)
    win.title("Dev Menu")
    win.configure(bg=BG)
    win.resizable(True, True)
    root.update_idletasks()
    win.geometry(f"{root.winfo_width()}x{root.winfo_height()}+{root.winfo_x()}+{root.winfo_y()}")

    # ── Header (fixed, outside scroll) ───────────────────────────────────────
    hdr = tk.Frame(win, bg=PANEL, pady=10)
    hdr.pack(fill="x")
    tk.Label(hdr, text="Dev Menu", bg=PANEL, fg=ACCENT2,
             font=(FONT, 12, "bold")).pack(side="left", padx=16)
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

    win.after(100, lambda: _bind_wheel(inner))

    # ── Section helper ────────────────────────────────────────────────────────
    def section(label):
        tk.Label(inner, text=label, bg=PANEL, fg=ACCENT2,
                 font=(FONT, 10, "bold")).pack(pady=(16, 6))

    # ══════════════════════════════════════════════════════════════════════════
    # Dev sections go here
    # ══════════════════════════════════════════════════════════════════════════

    section("Developer Tools")

    tk.Label(
        inner,
        text="Dev tools and diagnostics will appear here.",
        bg=PANEL, fg=SUBTEXT, font=(FONT, 9),
    ).pack(anchor="w", padx=24, pady=(0, 8))

    # ── Action buttons ────────────────────────────────────────────────────────
    section("Actions")

    btn_frame = tk.Frame(inner, bg=PANEL, pady=12)
    btn_frame.pack(fill="x", padx=16)

    tk.Button(
        btn_frame, text="Close", bg=ACCENT, fg=BG, relief="flat",
        font=(FONT, 9, "bold"), activebackground=ACCENT2, activeforeground=BG,
        padx=16, pady=4, cursor="hand2", command=win.destroy,
    ).pack(side="right")

    # bottom padding
    tk.Frame(inner, bg=PANEL, height=20).pack()