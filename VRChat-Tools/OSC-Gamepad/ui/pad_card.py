"""
ui/pad_card.py
──────────────
PadCard: one pad's config row (name, host, port, style, connect)
plus its active pad area (NES / Joystick).
"""

import tkinter as tk

from core.pad_state import PadState
from ui.theme import BG, PANEL, BORDER, ACCENT, ACCENT2, TEXT, SUBTEXT, GREEN, RED, FONT
from ui.widgets import NESPad, JoystickPad


class PadCard(tk.Frame):
    def __init__(self, parent, index: int, on_remove, host="127.0.0.1", port="9000",
                 style="nes", name="", **kwargs):
        super().__init__(parent, bg=PANEL, highlightthickness=1, highlightbackground=BORDER, **kwargs)
        self.index     = index
        self.on_remove = on_remove
        self.state: PadState | None = None

        self._default_host  = host
        self._default_port  = str(port)
        self._default_style = style
        self._default_name  = name or f"Pad {index}"
        self._connected     = False

        self._build()

    def _build(self):
        # ── Header ────────────────────────────────────────────────────────────
        hdr = tk.Frame(self, bg=PANEL)
        hdr.pack(fill=tk.X, padx=8, pady=(8, 4))

        tk.Label(hdr, text="◈", font=(FONT, 10, "bold"), fg=ACCENT2, bg=PANEL).pack(side=tk.LEFT)

        self._name_var = tk.StringVar(value=self._default_name)
        tk.Entry(
            hdr, textvariable=self._name_var,
            font=(FONT, 10, "bold"), fg=ACCENT2, bg=PANEL,
            insertbackground=ACCENT2, relief="flat",
            highlightthickness=0, width=18,
        ).pack(side=tk.LEFT, padx=(4, 0))

        rm_btn = tk.Label(hdr, text="✕", font=(FONT, 10), fg=RED, bg=PANEL, cursor="hand2")
        rm_btn.pack(side=tk.RIGHT, padx=4)
        rm_btn.bind("<Button-1>", lambda e: self.on_remove(self))

        tk.Frame(self, bg=BORDER, height=1).pack(fill=tk.X, padx=8)

        # ── Config row ────────────────────────────────────────────────────────
        cfg_row = tk.Frame(self, bg=PANEL)
        cfg_row.pack(fill=tk.X, padx=8, pady=6)

        def lbl(text):
            tk.Label(cfg_row, text=text, font=(FONT, 8), fg=SUBTEXT, bg=PANEL).pack(side=tk.LEFT)

        def entry(default, width):
            e = tk.Entry(
                cfg_row, font=(FONT, 9), width=width,
                bg=PANEL, fg=TEXT, insertbackground=ACCENT,
                relief="flat", highlightthickness=1,
                highlightbackground=BORDER, highlightcolor=ACCENT,
            )
            e.insert(0, default)
            e.pack(side=tk.LEFT, padx=(2, 8))
            return e

        lbl("Host:")
        self._host = entry(self._default_host, 13)
        lbl("Port:")
        self._port = entry(self._default_port, 6)

        self._style_var = tk.StringVar(value=self._default_style)
        for val, txt in (("nes", "NES"), ("joy", "Joystick")):
            tk.Radiobutton(
                cfg_row, text=txt, variable=self._style_var, value=val,
                font=(FONT, 8), fg=TEXT, bg=PANEL,
                selectcolor=PANEL, activebackground=PANEL,
                command=self._rebuild_pad,
            ).pack(side=tk.LEFT, padx=2)

        self._conn_btn = tk.Button(
            cfg_row, text="Connect", font=(FONT, 8, "bold"),
            fg=BG, bg=ACCENT, relief="flat",
            activebackground=ACCENT2, activeforeground=BG,
            cursor="hand2", padx=6, command=self._toggle_connect,
        )
        self._conn_btn.pack(side=tk.LEFT, padx=(8, 0))

        self._status_dot = tk.Label(cfg_row, text="●", font=(FONT, 10), fg=SUBTEXT, bg=PANEL)
        self._status_dot.pack(side=tk.LEFT, padx=4)

        # ── Pad area ──────────────────────────────────────────────────────────
        self._pad_area = tk.Frame(self, bg=PANEL)
        self._pad_area.pack(padx=4, pady=(0, 8))
        self._show_placeholder()

    # ── Connection ────────────────────────────────────────────────────────────

    def _toggle_connect(self):
        self._disconnect() if self._connected else self._connect()

    def _connect(self):
        host = self._host.get().strip()
        try:
            port = int(self._port.get().strip())
        except ValueError:
            self._status_dot.config(fg=RED)
            return

        if self.state:
            self.state.stop()

        self.state = PadState(host, port)
        self._connected = True
        self._status_dot.config(fg=GREEN)
        self._conn_btn.config(
            text="Disconnect", fg=BG, bg=RED,
            activebackground=RED, activeforeground=BG,
        )
        self._rebuild_pad()

    def _disconnect(self):
        if self.state:
            self.state.stop()
            self.state = None
        self._connected = False
        self._status_dot.config(fg=SUBTEXT)
        self._conn_btn.config(
            text="Connect", fg=BG, bg=ACCENT,
            activebackground=ACCENT2, activeforeground=BG,
        )
        for w in self._pad_area.winfo_children():
            w.destroy()
        self._show_placeholder()

    def _show_placeholder(self):
        tk.Label(
            self._pad_area, text="Press Connect to activate",
            font=(FONT, 8), fg=SUBTEXT, bg=PANEL,
        ).pack(pady=16)

    def _rebuild_pad(self):
        if not self.state:
            return
        for w in self._pad_area.winfo_children():
            w.destroy()
        cls = NESPad if self._style_var.get() == "nes" else JoystickPad
        cls(self._pad_area, self.state).pack()

    # ── Config I/O ────────────────────────────────────────────────────────────

    def get_config(self) -> dict:
        return {
            "host":  self._host.get().strip(),
            "port":  self._port.get().strip(),
            "style": self._style_var.get(),
            "name":  self._name_var.get().strip(),
        }

    def destroy_state(self):
        if self.state:
            self.state.stop()
