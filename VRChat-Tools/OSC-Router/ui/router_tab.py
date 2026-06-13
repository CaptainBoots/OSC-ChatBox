"""
ui/router_tab.py
────────────────
Router tab with two inner subtabs:
  • Sources  — input listeners (name + port)
  • Outputs  — output targets (name + ip + port + source checkboxes)

Plus status bar, Start/Stop/Restart, and live stats at the top.
"""

import tkinter as tk
from tkinter import ttk
from ui.theme import BG, PANEL, BORDER, ACCENT, ACCENT2, TEXT, SUBTEXT, GREEN, RED, YELLOW, FONT


class RouterTab(tk.Frame):
    def __init__(self, parent, cfg, router, save_cb, start_cb, stop_cb, restart_cb, settings_cb, help_cb):
        super().__init__(parent, bg=BG)
        self._cfg         = cfg
        self._router      = router
        self._save_cb     = save_cb
        self._start_cb    = start_cb
        self._stop_cb     = stop_cb
        self._restart_cb  = restart_cb
        self._settings_cb = settings_cb
        self._help_cb     = help_cb

        # Internal row state
        self._src_rows: list[dict] = []   # {frame, name_entry, port_entry, stats_label}
        self._out_rows: list[dict] = []   # {frame, name_entry, ip_entry, port_entry, src_vars, stats_label}

        self.columnconfigure(0, weight=1)
        self.rowconfigure(3, weight=1)
        self._build()

    # ── Top section (always visible) ──────────────────────────────────────────

    def _build(self):
        # Status bar
        sf = tk.Frame(self, bg=PANEL, pady=6)
        sf.grid(row=0, column=0, sticky="ew", padx=8, pady=(8, 4))
        sf.columnconfigure(1, weight=1)
        tk.Label(sf, text="Status:", bg=PANEL, fg=SUBTEXT, font=(FONT, 9)).grid(row=0, column=0, padx=(10, 4))
        self._status_lbl = tk.Label(sf, text="Stopped", bg=PANEL, fg=RED, font=(FONT, 9, "bold"))
        self._status_lbl.grid(row=0, column=1, sticky="w")

        # Control buttons
        bf = tk.Frame(self, bg=BG)
        bf.grid(row=1, column=0, sticky="ew", padx=8, pady=4)
        for text, cmd, fg in (
            ("▶  Start",   self._start_cb,   GREEN),
            ("■  Stop",    self._stop_cb,    RED),
            ("↺  Restart", self._restart_cb, ACCENT2),
        ):
            tk.Button(bf, text=text, bg=PANEL, fg=fg, relief="flat", cursor="hand2",
                      font=(FONT, 10, "bold"), activebackground=BORDER, activeforeground=TEXT,
                      width=12, pady=6, command=cmd).pack(side="left", padx=4)
        tk.Button(bf, text="⚙ Settings", bg=PANEL, fg=SUBTEXT, relief="flat", cursor="hand2",
                  font=(FONT, 9), activebackground=BORDER, activeforeground=TEXT,
                  command=self._settings_cb).pack(side="right", padx=4)
        tk.Button(bf, text="? Help", bg=PANEL, fg=SUBTEXT, relief="flat", cursor="hand2",
                  font=(FONT, 9), activebackground=BORDER, activeforeground=TEXT,
                  command=self._help_cb).pack(side="right", padx=4)

        # Live stats
        lf = tk.Frame(self, bg=PANEL, highlightthickness=1, highlightbackground=BORDER)
        lf.grid(row=2, column=0, sticky="ew", padx=8, pady=(0, 4))
        lf.columnconfigure(1, weight=1)
        tk.Label(lf, text="Live Stats", bg=PANEL, fg=ACCENT2,
                 font=(FONT, 9, "bold")).grid(row=0, column=0, columnspan=3, sticky="w", padx=10, pady=(6, 2))
        self._lbl_fwd      = tk.Label(lf, text="Forwarded: —",   bg=PANEL, fg=TEXT, font=(FONT, 9))
        self._lbl_conflict = tk.Label(lf, text="Conflicts: —",    bg=PANEL, fg=TEXT, font=(FONT, 9))
        self._lbl_sources  = tk.Label(lf, text="Sources: 0 / 0",  bg=PANEL, fg=TEXT, font=(FONT, 9))
        self._lbl_outputs  = tk.Label(lf, text="Outputs: 0 / 0",  bg=PANEL, fg=TEXT, font=(FONT, 9))
        self._lbl_fwd.grid(row=1, column=0, padx=10, pady=(0, 6), sticky="w")
        self._lbl_conflict.grid(row=1, column=1, padx=10, pady=(0, 6), sticky="w")
        self._lbl_sources.grid(row=1, column=2, padx=10, pady=(0, 6), sticky="w")
        self._lbl_outputs.grid(row=1, column=3, padx=10, pady=(0, 6), sticky="w")

        # Inner notebook (Sources / Outputs)
        style = ttk.Style()
        style.configure("Inner.TNotebook", background=BG, borderwidth=0, tabmargins=0)
        style.configure("Inner.TNotebook.Tab", background=PANEL, foreground=SUBTEXT,
                        font=(FONT, 9), padding=(12, 4), borderwidth=0)
        style.map("Inner.TNotebook.Tab",
                  background=[("selected", BORDER)],
                  foreground=[("selected", ACCENT2)])

        nb = ttk.Notebook(self, style="Inner.TNotebook")
        nb.grid(row=3, column=0, sticky="nsew", padx=8, pady=(4, 8))

        self._src_tab = self._make_sources_tab(nb)
        self._out_tab = self._make_outputs_tab(nb)
        nb.add(self._src_tab, text="  Sources  ")
        nb.add(self._out_tab, text="  Outputs  ")

    # ── Sources subtab ────────────────────────────────────────────────────────

    def _make_sources_tab(self, parent) -> tk.Frame:
        frame = tk.Frame(parent, bg=BG)
        frame.columnconfigure(0, weight=1)
        frame.rowconfigure(1, weight=1)

        toolbar = tk.Frame(frame, bg=BG)
        toolbar.grid(row=0, column=0, sticky="ew", padx=4, pady=(6, 2))
        tk.Label(toolbar, text="Input Sources", bg=BG, fg=ACCENT2,
                 font=(FONT, 9, "bold")).pack(side="left", padx=4)
        tk.Button(toolbar, text="+ Add Source", bg=PANEL, fg=ACCENT2, relief="flat",
                  cursor="hand2", font=(FONT, 9), activebackground=BORDER, activeforeground=TEXT,
                  command=self._add_source).pack(side="left", padx=4)

        self._src_canvas, self._src_inner = self._scrollable(frame, row=1)
        return frame

    # ── Outputs subtab ────────────────────────────────────────────────────────

    def _make_outputs_tab(self, parent) -> tk.Frame:
        frame = tk.Frame(parent, bg=BG)
        frame.columnconfigure(0, weight=1)
        frame.rowconfigure(1, weight=1)

        toolbar = tk.Frame(frame, bg=BG)
        toolbar.grid(row=0, column=0, sticky="ew", padx=4, pady=(6, 2))
        tk.Label(toolbar, text="Output Targets", bg=BG, fg=ACCENT2,
                 font=(FONT, 9, "bold")).pack(side="left", padx=4)
        tk.Button(toolbar, text="+ Add Output", bg=PANEL, fg=ACCENT2, relief="flat",
                  cursor="hand2", font=(FONT, 9), activebackground=BORDER, activeforeground=TEXT,
                  command=self._add_output).pack(side="left", padx=4)

        self._out_canvas, self._out_inner = self._scrollable(frame, row=1)
        return frame

    # ── Source rows ───────────────────────────────────────────────────────────

    def add_source_row(self, name: str = "Source", port: int = 9001):
        idx = len(self._src_rows)
        card = tk.Frame(self._src_inner, bg=PANEL, highlightthickness=1, highlightbackground=BORDER)
        card.grid(row=idx, column=0, sticky="ew", padx=4, pady=3)
        card.columnconfigure(1, weight=1)
        card.columnconfigure(3, weight=1)

        tk.Label(card, text=f"#{idx+1}", bg=PANEL, fg=SUBTEXT,
                 font=(FONT, 8), width=3, anchor="e").grid(row=0, column=0, padx=(8, 4), pady=6)

        name_e = self._entry(card, name, width=18)
        name_e.grid(row=0, column=1, sticky="ew", padx=4, pady=6)

        tk.Label(card, text="Port:", bg=PANEL, fg=SUBTEXT, font=(FONT, 9)).grid(row=0, column=2, padx=(8, 4))

        port_e = self._entry(card, str(port), width=7)
        port_e.grid(row=0, column=3, sticky="ew", padx=(0, 4), pady=6)

        stats = tk.Label(card, text="●", bg=PANEL, fg=SUBTEXT, font=(FONT, 9))
        stats.grid(row=0, column=4, padx=6)

        tk.Button(card, text="✕", bg=PANEL, fg=RED, relief="flat", cursor="hand2",
                  font=(FONT, 9), activebackground=BORDER, activeforeground=RED,
                  command=lambda i=idx: self._remove_source(i)).grid(row=0, column=5, padx=(0, 8))

        row = {"frame": card, "name_entry": name_e, "port_entry": port_e, "stats_label": stats}
        self._src_rows.append(row)

        for e in (name_e, port_e):
            e.bind("<FocusOut>", self._on_source_change)
            e.bind("<Return>",   self._on_source_change)

        self._src_canvas.configure(scrollregion=self._src_canvas.bbox("all"))

    def _add_source(self):
        self.add_source_row("Source", 9010 + len(self._src_rows) + 1)
        self._on_source_change()
        # Refresh all output rows so they pick up the new source
        self._rebuild_output_source_checks()

    def _remove_source(self, idx: int):
        if idx >= len(self._src_rows):
            return
        self._src_rows[idx]["frame"].destroy()
        self._src_rows.pop(idx)
        for i, r in enumerate(self._src_rows):
            r["frame"].grid(row=i)
        self._on_source_change()
        self._rebuild_output_source_checks()

    def _on_source_change(self, _=None):
        self._cfg["sources"] = self._collect_sources()
        self._save_cb()

    def _collect_sources(self) -> list[dict]:
        out = []
        for r in self._src_rows:
            name = r["name_entry"].get().strip() or "Source"
            try:
                port = int(r["port_entry"].get())
            except ValueError:
                port = 9001
            out.append({"name": name, "port": port})
        return out

    def source_names(self) -> list[str]:
        return [r["name_entry"].get().strip() or "Source" for r in self._src_rows]

    # ── Output rows ───────────────────────────────────────────────────────────

    def add_output_row(self, name: str = "Output", ip: str = "127.0.0.1",
                       port: int = 9000, subscribed: list[str] | None = None):
        if subscribed is None:
            subscribed = self.source_names()

        idx = len(self._out_rows)
        card = tk.Frame(self._out_inner, bg=PANEL, highlightthickness=1, highlightbackground=BORDER)
        card.grid(row=idx, column=0, sticky="ew", padx=4, pady=3)
        card.columnconfigure(1, weight=1)
        card.columnconfigure(3, weight=1)

        # ── Header row ────────────────────────────────────────────────────────
        hdr = tk.Frame(card, bg=PANEL)
        hdr.grid(row=0, column=0, columnspan=7, sticky="ew")
        hdr.columnconfigure(1, weight=1)
        hdr.columnconfigure(3, weight=1)

        tk.Label(hdr, text=f"#{idx+1}", bg=PANEL, fg=SUBTEXT,
                 font=(FONT, 8), width=3, anchor="e").grid(row=0, column=0, padx=(8, 4), pady=(8, 4))

        name_e = self._entry(hdr, name, width=14)
        name_e.grid(row=0, column=1, sticky="ew", padx=4, pady=(8, 4))

        tk.Label(hdr, text="IP:", bg=PANEL, fg=SUBTEXT, font=(FONT, 9)).grid(row=0, column=2, padx=(8, 4))
        ip_e = self._entry(hdr, ip, width=14)
        ip_e.grid(row=0, column=3, sticky="ew", padx=(0, 4), pady=(8, 4))

        tk.Label(hdr, text="Port:", bg=PANEL, fg=SUBTEXT, font=(FONT, 9)).grid(row=0, column=4, padx=(8, 4))
        port_e = self._entry(hdr, str(port), width=6)
        port_e.grid(row=0, column=5, sticky="ew", padx=(0, 4), pady=(8, 4))

        stats = tk.Label(hdr, text="●", bg=PANEL, fg=SUBTEXT, font=(FONT, 9))
        stats.grid(row=0, column=6, padx=4)

        tk.Button(hdr, text="✕", bg=PANEL, fg=RED, relief="flat", cursor="hand2",
                  font=(FONT, 9), activebackground=BORDER, activeforeground=RED,
                  command=lambda i=idx: self._remove_output(i)).grid(row=0, column=7, padx=(2, 8))

        # ── Source checkboxes ─────────────────────────────────────────────────
        check_frame = tk.Frame(card, bg=PANEL)
        check_frame.grid(row=1, column=0, columnspan=7, sticky="ew", padx=12, pady=(0, 8))

        tk.Label(check_frame, text="Receives from:", bg=PANEL, fg=SUBTEXT,
                 font=(FONT, 8)).grid(row=0, column=0, sticky="w", padx=(0, 8))

        src_vars: dict[str, tk.BooleanVar] = {}
        check_widgets: dict[str, tk.Checkbutton] = {}

        for col, src_name in enumerate(self.source_names()):
            var = tk.BooleanVar(value=(src_name in subscribed))
            cb  = tk.Checkbutton(
                check_frame, text=src_name, variable=var,
                bg=PANEL, fg=TEXT, selectcolor=BORDER,
                activebackground=PANEL, activeforeground=ACCENT2,
                font=(FONT, 9), cursor="hand2",
                command=self._on_output_change,
            )
            cb.grid(row=0, column=col + 1, padx=6, sticky="w")
            src_vars[src_name]    = var
            check_widgets[src_name] = cb

        row_data = {
            "frame":          card,
            "name_entry":     name_e,
            "ip_entry":       ip_e,
            "port_entry":     port_e,
            "stats_label":    stats,
            "src_vars":       src_vars,
            "check_frame":    check_frame,
            "check_widgets":  check_widgets,
        }
        self._out_rows.append(row_data)

        for e in (name_e, ip_e, port_e):
            e.bind("<FocusOut>", self._on_output_change)
            e.bind("<Return>",   self._on_output_change)

        self._out_canvas.configure(scrollregion=self._out_canvas.bbox("all"))

    def _add_output(self):
        self.add_output_row(
            name="Output",
            ip="127.0.0.1",
            port=9000 + len(self._out_rows),
            subscribed=self.source_names(),
        )
        self._on_output_change()

    def _remove_output(self, idx: int):
        if idx >= len(self._out_rows):
            return
        self._out_rows[idx]["frame"].destroy()
        self._out_rows.pop(idx)
        for i, r in enumerate(self._out_rows):
            r["frame"].grid(row=i)
        self._on_output_change()

    def _on_output_change(self, _=None):
        self._cfg["outputs"] = self._collect_outputs()
        self._save_cb()

    def _collect_outputs(self) -> list[dict]:
        out = []
        for r in self._out_rows:
            name = r["name_entry"].get().strip() or "Output"
            ip   = r["ip_entry"].get().strip()   or "127.0.0.1"
            try:
                port = int(r["port_entry"].get())
            except ValueError:
                port = 9000
            sources = [n for n, var in r["src_vars"].items() if var.get()]
            out.append({"name": name, "ip": ip, "port": port, "sources": sources})
        return out

    def _rebuild_output_source_checks(self):
        """
        After sources change, rebuild every output row's checkbox list
        so new sources appear and removed ones disappear.
        Preserves existing checked state by name where possible.
        """
        current_outputs = self._collect_outputs()

        # Destroy all output rows
        for r in self._out_rows:
            r["frame"].destroy()
        self._out_rows.clear()

        # Re-add with updated source list
        for o in current_outputs:
            self.add_output_row(o["name"], o["ip"], o["port"], o["sources"])

    # ── Stats tick ────────────────────────────────────────────────────────────

    def tick(self):
        if self._router.running:
            active_src = sum(1 for s in self._router.sources if s.running)
            total_src  = len(self._router.sources)
            active_out = sum(1 for o in self._router.outputs if not o.failed)
            total_out  = len(self._router.outputs)

            self._lbl_fwd.config(text=f"Forwarded: {self._router.total_forwarded:,}")
            self._lbl_conflict.config(text=f"Conflicts: {self._router.live_conflicts} live")
            self._lbl_sources.config(text=f"Sources: {active_src} / {total_src}")
            self._lbl_outputs.config(text=f"Outputs: {active_out} / {total_out}")

            # Per source stats
            src_by_name = {s.name: s for s in self._router.sources}
            for r in self._src_rows:
                name = r["name_entry"].get().strip()
                src  = src_by_name.get(name)
                if src:
                    if src.running:
                        r["stats_label"].config(text=f"● {src.rx_count:,} rx", fg=GREEN)
                    else:
                        r["stats_label"].config(text="✗ failed", fg=RED)
                else:
                    r["stats_label"].config(text="●", fg=SUBTEXT)

            # Per output stats
            out_by_name = {o.name: o for o in self._router.outputs}
            for r in self._out_rows:
                name = r["name_entry"].get().strip()
                out  = out_by_name.get(name)
                if out:
                    if out.failed:
                        r["stats_label"].config(text="✗ failed", fg=RED)
                    else:
                        r["stats_label"].config(text=f"▶ {out.fwd_total:,} sent", fg=GREEN)
                else:
                    r["stats_label"].config(text="●", fg=SUBTEXT)
        else:
            self._lbl_fwd.config(text="Forwarded: —")
            self._lbl_conflict.config(text="Conflicts: —")
            self._lbl_sources.config(text=f"Sources: 0 / {len(self._src_rows)}")
            self._lbl_outputs.config(text=f"Outputs: 0 / {len(self._out_rows)}")
            for r in self._src_rows + self._out_rows:
                r["stats_label"].config(text="●", fg=SUBTEXT)

    def set_status(self, text: str):
        colour = GREEN if "running" in text.lower() else RED
        self._status_lbl.config(text=text, fg=colour)

    # ── collect_config (for app.py to use at start time) ──────────────────────

    def collect_config(self) -> dict:
        return {
            "sources": self._collect_sources(),
            "outputs": self._collect_outputs(),
        }

    # ── Helpers ───────────────────────────────────────────────────────────────

    def _entry(self, parent, value: str = "", width: int = 14) -> tk.Entry:
        e = tk.Entry(parent, bg=PANEL, fg=TEXT, insertbackground=ACCENT,
                     relief="flat", font=(FONT, 9), width=width,
                     highlightthickness=1, highlightbackground=BORDER, highlightcolor=ACCENT)
        e.insert(0, value)
        return e

    def _scrollable(self, parent, row: int):
        outer = tk.Frame(parent, bg=BG)
        outer.grid(row=row, column=0, sticky="nsew", padx=4)
        outer.columnconfigure(0, weight=1)
        outer.rowconfigure(0, weight=1)
        parent.rowconfigure(row, weight=1)

        canvas = tk.Canvas(outer, bg=BG, highlightthickness=0)
        vsb    = tk.Scrollbar(outer, orient="vertical", command=canvas.yview)
        canvas.configure(yscrollcommand=vsb.set)
        vsb.grid(row=0, column=1, sticky="ns")
        canvas.grid(row=0, column=0, sticky="nsew")
        canvas.bind("<MouseWheel>", lambda e: canvas.yview_scroll(-1*(e.delta//120), "units"))

        inner = tk.Frame(canvas, bg=BG)
        inner.columnconfigure(0, weight=1)
        inner.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=inner, anchor="nw")

        return canvas, inner
