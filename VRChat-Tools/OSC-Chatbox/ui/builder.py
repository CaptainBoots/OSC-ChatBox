"""
ui/builder.py
─────────────
Builder tab: palette on the left, pages + slots on the right.

Users can:
  • Click a module in the palette to add it to the selected page
  • Drag slots up/down to reorder
  • Click ✕ to remove a slot
  • Edit custom_text inline
  • Add / remove pages
  • Toggle page enabled/disabled
  • Set per-page duration with a spinbox

The builder reads/writes cfg["pages"] directly and calls save_cb() after
every change so config is always in sync.
"""

import tkinter as tk
from tkinter import ttk

from modules.registry import CATEGORIES, MODULE_BY_ID
from ui.theme import BG, PANEL, BORDER, ACCENT, ACCENT2, TEXT, SUBTEXT, FONT, RED


class BuilderTab(tk.Frame):
    def __init__(self, parent, cfg: dict, save_cb):
        super().__init__(parent, bg=BG)
        self._cfg     = cfg
        self._save_cb = save_cb
        self._sel_page = 0           # currently selected page index
        self._drag     = {}          # drag state

        self._build_ui()
        self._refresh_pages()

    # ── Layout ────────────────────────────────────────────────────────────────

    def _build_ui(self):
        self.columnconfigure(0, minsize=200)
        self.columnconfigure(1, weight=1)
        self.rowconfigure(0, weight=1)

        # Left panel: module palette
        left = tk.Frame(self, bg=PANEL, highlightthickness=1, highlightbackground=BORDER)
        left.grid(row=0, column=0, sticky="nswe", padx=(8, 4), pady=8)
        left.rowconfigure(1, weight=1)
        left.columnconfigure(0, weight=1)

        tk.Label(left, text="Modules", bg=PANEL, fg=ACCENT2,
                 font=(FONT, 10, "bold")).grid(row=0, column=0, pady=(10, 4))

        palette_scroll = tk.Frame(left, bg=PANEL)
        palette_scroll.grid(row=1, column=0, sticky="nswe")
        palette_scroll.columnconfigure(0, weight=1)

        canvas = tk.Canvas(palette_scroll, bg=PANEL, highlightthickness=0)
        vsb    = tk.Scrollbar(palette_scroll, orient="vertical", command=canvas.yview)
        canvas.configure(yscrollcommand=vsb.set)
        vsb.pack(side="right", fill="y")
        canvas.pack(side="left", fill="both", expand=True)

        self._palette_inner = tk.Frame(canvas, bg=PANEL)
        self._palette_inner.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        canvas.create_window((0, 0), window=self._palette_inner, anchor="nw")
        canvas.bind("<MouseWheel>", lambda e: canvas.yview_scroll(-1 * (e.delta // 120), "units"))

        self._build_palette()

        # Right panel: pages
        right = tk.Frame(self, bg=BG)
        right.grid(row=0, column=1, sticky="nswe", padx=(4, 8), pady=8)
        right.columnconfigure(0, weight=1)
        right.rowconfigure(1, weight=1)

        # Page list toolbar
        toolbar = tk.Frame(right, bg=BG)
        toolbar.grid(row=0, column=0, sticky="ew", pady=(0, 4))

        tk.Label(toolbar, text="Pages", bg=BG, fg=ACCENT2,
                 font=(FONT, 10, "bold")).pack(side="left", padx=4)

        tk.Button(
            toolbar, text="+ Add Page", bg=PANEL, fg=ACCENT2,
            relief="flat", cursor="hand2", font=(FONT, 9),
            activebackground=BORDER, activeforeground=TEXT,
            command=self._add_page,
        ).pack(side="left", padx=4)

        # Scrollable page list
        page_outer = tk.Frame(right, bg=BG)
        page_outer.grid(row=1, column=0, sticky="nswe")
        page_outer.columnconfigure(0, weight=1)
        page_outer.rowconfigure(0, weight=1)

        self._page_canvas = tk.Canvas(page_outer, bg=BG, highlightthickness=0)
        vsb2 = tk.Scrollbar(page_outer, orient="vertical", command=self._page_canvas.yview)
        self._page_canvas.configure(yscrollcommand=vsb2.set)
        vsb2.grid(row=0, column=1, sticky="ns")
        self._page_canvas.grid(row=0, column=0, sticky="nswe")
        self._page_canvas.bind(
            "<Configure>",
            lambda e: self._page_canvas.configure(scrollregion=self._page_canvas.bbox("all"))
        )
        self._page_canvas.bind(
            "<MouseWheel>",
            lambda e: self._page_canvas.yview_scroll(-1 * (e.delta // 120), "units"),
        )

        self._pages_frame = tk.Frame(self._page_canvas, bg=BG)
        self._pages_frame.columnconfigure(0, weight=1)
        self._page_canvas.create_window((0, 0), window=self._pages_frame, anchor="nw")
        self._pages_frame.bind(
            "<Configure>",
            lambda e: self._page_canvas.configure(scrollregion=self._page_canvas.bbox("all"))
        )

    # ── Palette ───────────────────────────────────────────────────────────────

    def _build_palette(self):
        for cat, mods in CATEGORIES.items():
            tk.Label(
                self._palette_inner, text=cat,
                bg=PANEL, fg=SUBTEXT, font=(FONT, 8, "bold"),
                anchor="w",
            ).pack(fill="x", padx=8, pady=(8, 2))

            for mod in mods:
                btn = tk.Button(
                    self._palette_inner,
                    text=f"  {mod['label']}",
                    bg=PANEL, fg=TEXT,
                    relief="flat", anchor="w", cursor="hand2",
                    font=(FONT, 9),
                    activebackground=BORDER, activeforeground=ACCENT2,
                    command=lambda m=mod: self._add_slot(m),
                )
                btn.pack(fill="x", padx=4, pady=1)

    # ── Page list ─────────────────────────────────────────────────────────────

    def _refresh_pages(self):
        for w in self._pages_frame.winfo_children():
            w.destroy()

        pages = self._cfg.get("pages", [])
        for i, page in enumerate(pages):
            self._build_page_card(i, page)

        self._page_canvas.configure(scrollregion=self._page_canvas.bbox("all"))

    def _build_page_card(self, page_idx: int, page: dict):
        card = tk.Frame(
            self._pages_frame, bg=PANEL,
            highlightthickness=1,
            highlightbackground=ACCENT if page_idx == self._sel_page else BORDER,
        )
        card.pack(fill="x", padx=4, pady=4)
        card.columnconfigure(0, weight=1)

        # ── Page header ───────────────────────────────────────────────────────
        header = tk.Frame(card, bg=PANEL)
        header.grid(row=0, column=0, sticky="ew", padx=6, pady=(6, 0))
        header.columnconfigure(1, weight=1)

        # Enabled toggle
        enabled_var = tk.BooleanVar(value=page.get("enabled", True))

        def _toggle_enabled(var=enabled_var, idx=page_idx):
            page["enabled"] = var.get()
            self._save()

        tk.Checkbutton(
            header, variable=enabled_var, bg=PANEL,
            selectcolor=PANEL, activebackground=PANEL,
            command=_toggle_enabled,
        ).grid(row=0, column=0, padx=(0, 4))

        # Page label — click to select
        label_text = f"Page {page_idx + 1}"
        lbl = tk.Label(
            header, text=label_text, bg=PANEL, fg=ACCENT2,
            font=(FONT, 10, "bold"), cursor="hand2",
        )
        lbl.grid(row=0, column=1, sticky="w")
        lbl.bind("<Button-1>", lambda e, i=page_idx: self._select_page(i))

        # Duration spinbox
        tk.Label(header, text="Duration:", bg=PANEL, fg=SUBTEXT, font=(FONT, 8)).grid(
            row=0, column=2, padx=(8, 2)
        )
        dur_var = tk.IntVar(value=int(page.get("duration", 20)))
        dur_spin = tk.Spinbox(
            header, from_=1, to=3600, textvariable=dur_var, width=5,
            bg=PANEL, fg=TEXT, insertbackground=ACCENT,
            relief="flat", highlightthickness=1,
            highlightbackground=BORDER, highlightcolor=ACCENT,
            font=(FONT, 9), buttonbackground=BORDER,
        )
        dur_spin.grid(row=0, column=3, padx=2)
        tk.Label(header, text="s", bg=PANEL, fg=SUBTEXT, font=(FONT, 8)).grid(row=0, column=4)

        def _dur_changed(*_, var=dur_var, idx=page_idx):
            try:
                self._cfg["pages"][idx]["duration"] = int(var.get())
                self._save()
            except (ValueError, IndexError):
                pass

        dur_var.trace_add("write", _dur_changed)

        # Remove page button
        tk.Button(
            header, text="✕", bg=PANEL, fg=RED, relief="flat",
            cursor="hand2", font=(FONT, 9),
            activebackground=BORDER, activeforeground=RED,
            command=lambda i=page_idx: self._remove_page(i),
        ).grid(row=0, column=5, padx=(4, 0))

        # ── Slot list ─────────────────────────────────────────────────────────
        slots_frame = tk.Frame(card, bg=PANEL)
        slots_frame.grid(row=1, column=0, sticky="ew", padx=10, pady=(4, 6))
        slots_frame.columnconfigure(1, weight=1)

        slots = page.get("slots", [])
        for slot_idx, slot in enumerate(slots):
            self._build_slot_row(slots_frame, page_idx, slot_idx, slot, slots)

        if not slots:
            tk.Label(
                slots_frame, text="No modules — click a module on the left to add",
                bg=PANEL, fg=SUBTEXT, font=(FONT, 8),
            ).grid(row=0, column=0, columnspan=5, pady=4, sticky="w")

    def _build_slot_row(self, parent, page_idx, slot_idx, slot, slots):
        mod = MODULE_BY_ID.get(slot.get("module", ""))
        label = mod["label"] if mod else slot.get("module", "unknown")

        # Fixed column layout:
        #  0: drag handle  1: ▲  2: ▼  3: label  4: text entry (optional)  5: ✕
        row_frame = tk.Frame(parent, bg=PANEL)
        row_frame.grid(row=slot_idx, column=0, columnspan=6, sticky="ew", pady=1)
        row_frame.columnconfigure(3, weight=1)  # label stretches
        row_frame.columnconfigure(4, weight=2)  # text entry stretches more

        # ── Drag handle ───────────────────────────────────────────────────────
        handle = tk.Label(row_frame, text="⠿", bg=PANEL, fg=SUBTEXT,
                          font=(FONT, 10), cursor="fleur")
        handle.grid(row=0, column=0, padx=(0, 2))
        handle.bind("<ButtonPress-1>",   lambda e, pi=page_idx, si=slot_idx: self._drag_start(e, pi, si))
        handle.bind("<B1-Motion>",       lambda e, pi=page_idx: self._drag_motion(e, pi))
        handle.bind("<ButtonRelease-1>", lambda e, pi=page_idx: self._drag_end(e, pi))

        # ▲ up
        tk.Button(
            row_frame, text="▲", bg=PANEL, fg=SUBTEXT, relief="flat",
            font=(FONT, 7), cursor="hand2", width=2,
            activebackground=BORDER, activeforeground=TEXT,
            command=lambda pi=page_idx, si=slot_idx: self._move_slot(pi, si, -1),
        ).grid(row=0, column=1, padx=1)

        # ▼ down
        tk.Button(
            row_frame, text="▼", bg=PANEL, fg=SUBTEXT, relief="flat",
            font=(FONT, 7), cursor="hand2", width=2,
            activebackground=BORDER, activeforeground=TEXT,
            command=lambda pi=page_idx, si=slot_idx: self._move_slot(pi, si, 1),
        ).grid(row=0, column=2, padx=(1, 4))

        # Module label
        tk.Label(row_frame, text=label, bg=PANEL, fg=TEXT,
                 font=(FONT, 9), anchor="w").grid(row=0, column=3, sticky="w", padx=(0, 4))

        # Inline text entry for custom_text (column 4)
        if mod and mod.get("has_text"):
            txt_var = tk.StringVar(value=slot.get("text", ""))
            entry = tk.Entry(
                row_frame, textvariable=txt_var, width=18,
                bg=BG, fg=TEXT, insertbackground=ACCENT,
                relief="flat", font=(FONT, 9),
                highlightthickness=1, highlightbackground=BORDER, highlightcolor=ACCENT,
            )
            entry.grid(row=0, column=4, padx=4, sticky="ew")

            def _text_changed(*_, var=txt_var, pi=page_idx, si=slot_idx):
                try:
                    self._cfg["pages"][pi]["slots"][si]["text"] = var.get()
                    self._save()
                except IndexError:
                    pass

            txt_var.trace_add("write", _text_changed)

        # ✕ remove — always in column 5
        tk.Button(
            row_frame, text="✕", bg=PANEL, fg=RED, relief="flat",
            font=(FONT, 9), cursor="hand2",
            activebackground=BORDER, activeforeground=RED,
            command=lambda pi=page_idx, si=slot_idx: self._remove_slot(pi, si),
        ).grid(row=0, column=5, padx=(2, 4))

    # ── Actions ───────────────────────────────────────────────────────────────

    def _select_page(self, idx: int):
        self._sel_page = idx
        self._refresh_pages()

    def _add_page(self):
        default_dur = int(self._cfg.get("switch_interval", 20))
        self._cfg.setdefault("pages", []).append({
            "enabled":  True,
            "duration": default_dur,
            "slots":    [],
        })
        self._sel_page = len(self._cfg["pages"]) - 1
        self._save()
        self._refresh_pages()

    def _remove_page(self, idx: int):
        pages = self._cfg.get("pages", [])
        if len(pages) <= 1:
            return
        pages.pop(idx)
        self._sel_page = max(0, self._sel_page - 1)
        self._save()
        self._refresh_pages()

    def _add_slot(self, mod: dict):
        pages = self._cfg.get("pages", [])
        if not pages:
            return
        idx = min(self._sel_page, len(pages) - 1)
        slot = {"module": mod["id"]}
        if mod.get("has_text"):
            slot["text"] = ""
        pages[idx].setdefault("slots", []).append(slot)
        self._save()
        self._refresh_pages()

    def _remove_slot(self, page_idx: int, slot_idx: int):
        try:
            self._cfg["pages"][page_idx]["slots"].pop(slot_idx)
            self._save()
            self._refresh_pages()
        except IndexError:
            pass

    def _move_slot(self, page_idx: int, slot_idx: int, direction: int):
        try:
            slots = self._cfg["pages"][page_idx]["slots"]
            new_idx = slot_idx + direction
            if 0 <= new_idx < len(slots):
                slots[slot_idx], slots[new_idx] = slots[new_idx], slots[slot_idx]
                self._save()
                self._refresh_pages()
        except IndexError:
            pass

    # ── Drag-and-drop slot reorder ────────────────────────────────────────────

    def _drag_start(self, event, page_idx: int, slot_idx: int):
        self._drag = {
            "page":    page_idx,
            "src":     slot_idx,
            "y_start": event.y_root,
        }

    def _drag_motion(self, event, page_idx: int):
        if not self._drag:
            return

    def _drag_end(self, event, page_idx: int):
        if not self._drag:
            return
        dy     = event.y_root - self._drag.get("y_start", event.y_root)
        steps  = dy // 28   # approximate row height
        if steps != 0:
            self._move_slot(self._drag["page"], self._drag["src"], steps)
        self._drag = {}

    # ── Save ──────────────────────────────────────────────────────────────────

    def _save(self):
        self._save_cb()

    def refresh(self):
        """Called externally when cfg changes outside the builder."""
        self._refresh_pages()