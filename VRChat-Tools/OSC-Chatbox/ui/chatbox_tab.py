"""
ui/chatbox_tab.py
─────────────────
Chatbox tab: live preview, start/stop/restart, config fields, forced text.

For flag themes, STRIPE_COLOURS drives repeating diagonal stripes drawn on a
canvas that fills the root window. Widgets sit in a normal Frame on top with
their own PANEL backgrounds so they remain fully readable. The stripe canvas
is managed by the App root window, not by this tab — this tab just uses BG
as its own background and the stripes show in the gaps between panels.
"""

import tkinter as tk

from ui.theme import BG, PANEL, BORDER, ACCENT, ACCENT2, TEXT, SUBTEXT, GREEN, RED, FONT, STRIPE_COLOURS


def draw_stripes(canvas: tk.Canvas, width: int, height: int, colours: list):
    """
    Fill *canvas* with repeating ~45° diagonal stripes that tile across the
    full width × height, like the reference image.
    """
    canvas.delete("stripe")
    if not colours or width <= 0 or height <= 0:
        return

    stripe_w = 28
    cycle    = stripe_w * len(colours)
    extent   = width + height + cycle * 2

    for start in range(-cycle, extent, cycle):
        for i, colour in enumerate(colours):
            x0 = start + i * stripe_w
            points = [
                x0,                0,
                x0 + stripe_w,     0,
                x0 + stripe_w + height, height,
                x0            + height, height,
                ]
            canvas.create_polygon(points, fill=colour, outline="", tags="stripe")

    canvas.tag_lower("stripe")


class ChatboxTab(tk.Frame):
    def __init__(self, parent, cfg: dict, state, save_cb, start_cb, stop_cb,
                 restart_cb, settings_cb, help_cb):
        super().__init__(parent, bg=BG)
        self._cfg         = cfg
        self._state       = state
        self._save_cb     = save_cb
        self._start_cb    = start_cb
        self._stop_cb     = stop_cb
        self._restart_cb  = restart_cb
        self._settings_cb = settings_cb
        self._help_cb     = help_cb

        self.columnconfigure(0, weight=1)

        if STRIPE_COLOURS:
            # Draw stripes directly on this frame's background via a canvas
            # that sits at z-order bottom; all child widgets grid on top normally.
            self._stripe_canvas = tk.Canvas(self, bg=BG, highlightthickness=0, bd=0)
            self._stripe_canvas.place(x=0, y=0, relwidth=1, relheight=1)
            self.bind("<Configure>", self._on_resize)

        self._build()

    def _on_resize(self, event):
        w, h = event.width, event.height
        draw_stripes(self._stripe_canvas, w, h, STRIPE_COLOURS)
        self._stripe_canvas.tk.call("lower", self._stripe_canvas._w)

    def _build(self):
        row = 0

        # ── Status bar ────────────────────────────────────────────────────────
        status_frame = tk.Frame(self, bg=PANEL, pady=6)
        status_frame.grid(row=row, column=0, sticky="ew", padx=8, pady=(8, 4))
        status_frame.columnconfigure(1, weight=1)
        row += 1

        tk.Label(status_frame, text="Status:", bg=PANEL, fg=SUBTEXT,
                 font=(FONT, 9)).grid(row=0, column=0, padx=(10, 4))
        self._status_lbl = tk.Label(status_frame, text="Stopped", bg=PANEL, fg=RED,
                                    font=(FONT, 9, "bold"))
        self._status_lbl.grid(row=0, column=1, sticky="w")

        # ── Control buttons ───────────────────────────────────────────────────
        btn_frame = tk.Frame(self, bg=BG)
        btn_frame.grid(row=row, column=0, sticky="ew", padx=8, pady=4)
        row += 1

        for text, cmd, fg in (
                ("▶  Start",   self._start_cb,   ACCENT),
                ("■  Stop",    self._stop_cb,    ACCENT),
                ("↺  Restart", self._restart_cb, ACCENT),
        ):
            tk.Button(
                btn_frame, text=text, bg=PANEL, fg=fg,
                relief="flat", cursor="hand2", font=(FONT, 10, "bold"),
                activebackground=BORDER, activeforeground=TEXT,
                width=12, pady=6, command=cmd,
            ).pack(side="left", padx=4)

        tk.Button(
            btn_frame, text="⚙ Settings", bg=PANEL, fg=SUBTEXT,
            relief="flat", cursor="hand2", font=(FONT, 9),
            activebackground=BORDER, activeforeground=TEXT,
            command=self._settings_cb,
        ).pack(side="right", padx=4)

        tk.Button(
            btn_frame, text="? Help", bg=PANEL, fg=SUBTEXT,
            relief="flat", cursor="hand2", font=(FONT, 9),
            activebackground=BORDER, activeforeground=TEXT,
            command=self._help_cb,
        ).pack(side="right", padx=4)

        tk.Frame(self, bg=BORDER, height=1).grid(row=row, column=0, sticky="ew", padx=8, pady=4)
        row += 1

        # ── Live preview ──────────────────────────────────────────────────────
        tk.Label(self, text="Live Chatbox Preview", bg=BG, fg=ACCENT2,
                 font=(FONT, 9, "bold")).grid(row=row, column=0, sticky="w", padx=12)
        row += 1

        preview_frame = tk.Frame(self, bg=PANEL, highlightthickness=1, highlightbackground=BORDER)
        preview_frame.grid(row=row, column=0, sticky="ew", padx=8, pady=(0, 4))
        preview_frame.columnconfigure(0, weight=1)
        row += 1

        self._preview = tk.Text(
            preview_frame, bg=PANEL, fg=TEXT,
            font=(FONT, 10), relief="flat",
            height=8, wrap="word",
            state="disabled", padx=10, pady=8,
        )
        self._preview.pack(fill="x")

        self._page_lbl = tk.Label(preview_frame, text="", bg=PANEL, fg=SUBTEXT, font=(FONT, 8))
        self._page_lbl.pack(anchor="e", padx=8, pady=(0, 4))

        tk.Frame(self, bg=BORDER, height=1).grid(row=row, column=0, sticky="ew", padx=8, pady=4)
        row += 1

        # ── Config fields ─────────────────────────────────────────────────────
        tk.Label(self, text="Configuration", bg=BG, fg=ACCENT2,
                 font=(FONT, 9, "bold")).grid(row=row, column=0, sticky="w", padx=12)
        row += 1

        cfg_frame = tk.Frame(self, bg=BG)
        cfg_frame.grid(row=row, column=0, sticky="ew", padx=12, pady=4)
        cfg_frame.columnconfigure(1, weight=1)
        cfg_frame.columnconfigure(3, weight=1)
        row += 1

        self._entries = {}

        fields = [
            ("OSC IP",       "osc_ip",         0, 0, 1),
            ("OSC Port",     "osc_port",        0, 2, 3),
            ("Interface",    "interface",       1, 0, 1),
            ("Interval (s)", "switch_interval", 1, 2, 3),
            ("LHM URL",      "lhm_api",         2, 0, 1),
            ("Location",     "location",        2, 2, 3),
        ]

        for label, key, r, cl, ce in fields:
            tk.Label(cfg_frame, text=label, bg=BG, fg=SUBTEXT,
                     font=(FONT, 9), anchor="e").grid(row=r, column=cl, sticky="e",
                                                      padx=(8, 4), pady=3)
            e = tk.Entry(
                cfg_frame, bg=PANEL, fg=TEXT, insertbackground=ACCENT,
                relief="flat", font=(FONT, 9),
                highlightthickness=1, highlightbackground=BORDER, highlightcolor=ACCENT,
            )
            e.insert(0, str(self._cfg.get(key, "")))
            e.grid(row=r, column=ce, sticky="ew", pady=3)
            self._entries[key] = e

            def _on_change(event, k=key, entry=e):
                self._cfg[k] = entry.get()
                self._save_cb()

            e.bind("<FocusOut>", _on_change)
            e.bind("<Return>",   _on_change)

        tk.Frame(self, bg=BORDER, height=1).grid(row=row, column=0, sticky="ew", padx=8, pady=4)
        row += 1

        # ── Forced text ───────────────────────────────────────────────────────
        tk.Label(self, text="Forced Text (overrides all pages)",
                 bg=BG, fg=ACCENT2, font=(FONT, 9, "bold")).grid(
            row=row, column=0, sticky="w", padx=12)
        row += 1

        forced_frame = tk.Frame(self, bg=BG)
        forced_frame.grid(row=row, column=0, sticky="ew", padx=12, pady=(0, 8))
        forced_frame.columnconfigure(0, weight=1)

        self._forced_var = tk.StringVar()
        forced_entry = tk.Entry(
            forced_frame, textvariable=self._forced_var,
            bg=PANEL, fg=TEXT, insertbackground=ACCENT,
            relief="flat", font=(FONT, 10),
            highlightthickness=1, highlightbackground=BORDER, highlightcolor=ACCENT,
        )
        forced_entry.grid(row=0, column=0, sticky="ew")
        tk.Label(forced_frame, text="Leave blank to use pages",
                 bg=BG, fg=SUBTEXT, font=(FONT, 8)).grid(row=1, column=0, sticky="w")

        def _forced_changed(*_):
            self._state.forced_text = self._forced_var.get()

        self._forced_var.trace_add("write", _forced_changed)

    # ── Public update methods ─────────────────────────────────────────────────

    def set_status(self, text: str):
        colour = GREEN if "running" in text.lower() else RED
        self._status_lbl.config(text=text, fg=colour)

    def set_preview(self, text: str):
        self._preview.config(state="normal")
        self._preview.delete("1.0", tk.END)
        self._preview.insert("1.0", text)
        self._preview.config(state="disabled")

    def set_page_label(self, text: str):
        self._page_lbl.config(text=text)

    def get_forced_text(self) -> str:
        return self._forced_var.get()