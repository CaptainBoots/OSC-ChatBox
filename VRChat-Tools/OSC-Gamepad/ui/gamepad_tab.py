"""
ui/gamepad_tab.py
─────────────────
Pads tab: toolbar with + Add Pad / ? Help / ⚙ Settings, then a
scrollable list of PadCards. Mirrors the layout style of
RouterTab / ChatboxTab (status-row + button-row + scrollable body).
"""

import tkinter as tk
from tkinter import ttk

from ui.pad_card import PadCard
from ui.theme import BG, PANEL, BORDER, ACCENT, ACCENT2, TEXT, SUBTEXT, GREEN, FONT, STRIPE_COLOURS, draw_stripes


class GamepadTab(tk.Frame):
    def __init__(self, parent, cfg: dict, save_cb, help_cb, settings_cb):
        super().__init__(parent, bg=BG)
        self._cfg         = cfg
        self._save_cb     = save_cb
        self._help_cb     = help_cb
        self._settings_cb = settings_cb

        self.cards: list[PadCard] = []
        self._pad_counter = 0

        self.columnconfigure(0, weight=1)
        self.rowconfigure(2, weight=1)

        if STRIPE_COLOURS:
            self._stripe_canvas = tk.Canvas(self, bg=BG, highlightthickness=0, bd=0)
            self._stripe_canvas.place(x=0, y=0, relwidth=1, relheight=1)
            self.bind("<Configure>", self._on_resize)

        self._build()

    def _on_resize(self, event):
        draw_stripes(self._stripe_canvas, event.width, event.height, STRIPE_COLOURS)
        self._stripe_canvas.tk.call("lower", self._stripe_canvas._w)

    # ── Layout ────────────────────────────────────────────────────────────────

    def _build(self):
        # ── Status bar ────────────────────────────────────────────────────────
        sf = tk.Frame(self, bg=PANEL, pady=6)
        sf.grid(row=0, column=0, sticky="ew", padx=8, pady=(8, 4))
        sf.columnconfigure(1, weight=1)

        tk.Label(sf, text="Status:", bg=PANEL, fg=SUBTEXT, font=(FONT, 9)).grid(row=0, column=0, padx=(10, 4))
        self._pad_count_lbl = tk.Label(sf, text="0 pads", bg=PANEL, fg=GREEN, font=(FONT, 9, "bold"))
        self._pad_count_lbl.grid(row=0, column=1, sticky="w")

        # ── Button row ────────────────────────────────────────────────────────
        bf = tk.Frame(self, bg=BG)
        bf.grid(row=1, column=0, sticky="ew", padx=8, pady=4)

        tk.Button(
            bf, text="＋  Add Pad", bg=PANEL, fg=ACCENT, relief="flat", cursor="hand2",
            font=(FONT, 10, "bold"), activebackground=BORDER, activeforeground=ACCENT2,
            width=12, pady=6, command=self._add_pad,
        ).pack(side="left", padx=4)

        tk.Button(
            bf, text="⚙ Settings", bg=PANEL, fg=SUBTEXT, relief="flat", cursor="hand2",
            font=(FONT, 9), activebackground=BORDER, activeforeground=TEXT,
            command=self._settings_cb,
        ).pack(side="right", padx=4)

        tk.Button(
            bf, text="? Help", bg=PANEL, fg=SUBTEXT, relief="flat", cursor="hand2",
            font=(FONT, 9), activebackground=BORDER, activeforeground=TEXT,
            command=self._help_cb,
        ).pack(side="right", padx=4)

        # ── Scrollable pad list ───────────────────────────────────────────────
        outer = tk.Frame(self, bg=BG)
        outer.grid(row=2, column=0, sticky="nsew", padx=8, pady=(0, 8))
        outer.columnconfigure(0, weight=1)
        outer.rowconfigure(0, weight=1)

        style = ttk.Style()
        style.configure(
            "Dark.Vertical.TScrollbar",
            background=PANEL, troughcolor=BG,
            arrowcolor=SUBTEXT, bordercolor=BG,
            lightcolor=PANEL, darkcolor=PANEL,
        )

        self._canvas = tk.Canvas(outer, bg=BG, highlightthickness=0)
        scrollbar = ttk.Scrollbar(outer, orient="vertical",
                                  command=self._canvas.yview, style="Dark.Vertical.TScrollbar")
        self._canvas.configure(yscrollcommand=scrollbar.set)
        scrollbar.grid(row=0, column=1, sticky="ns")
        self._canvas.grid(row=0, column=0, sticky="nsew")

        self._inner = tk.Frame(self._canvas, bg=BG)
        self._win_id = self._canvas.create_window((0, 0), window=self._inner, anchor="nw")

        self._inner.bind("<Configure>", lambda e: self._canvas.configure(scrollregion=self._canvas.bbox("all")))
        self._canvas.bind("<Configure>", lambda e: self._canvas.itemconfig(self._win_id, width=e.width))
        self._canvas.bind("<MouseWheel>", lambda e: self._canvas.yview_scroll(-1 * (e.delta // 120), "units"))

    # ── Pads ──────────────────────────────────────────────────────────────────

    def load_pads(self):
        saved = self._cfg.get("pads", [])
        if saved:
            for pad_cfg in saved:
                self.add_pad(
                    host=pad_cfg.get("host", "127.0.0.1"),
                    port=pad_cfg.get("port", "9000"),
                    style=pad_cfg.get("style", "nes"),
                    name=pad_cfg.get("name", ""),
                )
        else:
            self.add_pad()

    def add_pad(self, host="127.0.0.1", port="9000", style="nes", name=""):
        self._pad_counter += 1
        card = PadCard(self._inner, self._pad_counter, self._remove_pad,
                       host=host, port=str(port), style=style, name=name)
        card.pack(fill=tk.X, padx=4, pady=6)
        self.cards.append(card)
        self._update_count()

    def _add_pad(self):
        self.add_pad()
        self._save_cb()

    def _remove_pad(self, card: PadCard):
        card.destroy_state()
        card.destroy()
        self.cards.remove(card)
        self._update_count()
        self._save_cb()

    def _update_count(self):
        n = len(self.cards)
        self._pad_count_lbl.config(text=f"{n} pad{'s' if n != 1 else ''}")

    # ── Config I/O ────────────────────────────────────────────────────────────

    def collect_pads(self) -> list[dict]:
        return [c.get_config() for c in self.cards]

    def destroy_all(self):
        for c in self.cards:
            c.destroy_state()