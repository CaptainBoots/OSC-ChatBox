"""
ui/builder.py
─────────────
Builder tab: Single-panel layout where pages and slots fill the view.

Users can:
  • Scroll the entire page setup layout with the mouse wheel smoothly
  • Drag slots up/down to reorder
  • Click ✕ to remove a slot or sub-module
  • Edit custom_text inline
  • Add / remove pages
  • Toggle page enabled/disabled
  • Set per-page duration with a custom pure-dark numeric stepper
  • Add new lines and modules contextually using the "+ Add New Line" button
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
        self._canvas_win = None      # Tracks the canvas window item ID
        self._cards_frame = None     # Tracks the active layout container frame

        self._build_ui()
        self._refresh_pages()

    # ── Layout ────────────────────────────────────────────────────────────────

    def _build_ui(self):
        self.columnconfigure(0, weight=1)
        self.rowconfigure(0, weight=1)

        # Main panel: pages and items fill the entire frame width
        right = tk.Frame(self, bg=BG)
        right.grid(row=0, column=0, sticky="nsew", padx=4)
        right.rowconfigure(1, weight=1)
        right.columnconfigure(0, weight=1)

        # Header for main panel
        r_hdr = tk.Frame(right, bg=BG)
        r_hdr.grid(row=0, column=0, sticky="ew", pady=(8, 8))

        tk.Label(r_hdr, text="Output Pages Setup", bg=BG, fg=TEXT, font=(FONT, 11, "bold")).pack(side="left")

        tk.Button(
            r_hdr, text="+ Create Page", bg=PANEL, fg=ACCENT, relief="flat",
            font=(FONT, 9, "bold"), activebackground=BORDER, activeforeground=TEXT,
            padx=12, pady=4, cursor="hand2", command=self._add_page
        ).pack(side="right")

        # Scrollable canvas explicitly locked to BG color with zero-width borders
        self._canvas = tk.Canvas(right, bg=BG, highlightthickness=0, borderwidth=0)
        self._canvas.grid(row=1, column=0, sticky="nsew")

        self._vsb = tk.Scrollbar(right, orient="vertical", command=self._canvas.yview)
        self._vsb.grid(row=1, column=1, sticky="ns")
        self._canvas.configure(yscrollcommand=self._vsb.set)

        # Bind mouse wheel directly to canvas container
        self._bind_mouse_wheel(self._canvas)

        # Handle responsive scaling safely on the true internal window reference
        def _on_canvas_configure(e):
            if self._canvas_win is not None:
                self._canvas.itemconfigure(self._canvas_win, width=e.width)
        self._canvas.bind("<Configure>", _on_canvas_configure)

    # ── Mouse Wheel Event Binding Hook ────────────────────────────────────────

    def _bind_mouse_wheel(self, widget):
        """Recursively registers platforms-agnostic scroll events to a widget and its children."""
        widget.bind("<MouseWheel>", self._on_mouse_wheel, add="+")
        widget.bind("<Button-4>", self._on_mouse_wheel, add="+")
        widget.bind("<Button-5>", self._on_mouse_wheel, add="+")

    def _on_mouse_wheel(self, event):
        """Intercepts hardware scroll delta codes across Windows, macOS, and Linux."""
        if event.num == 4:    # Linux scroll up
            self._canvas.yview_scroll(-1, "units")
        elif event.num == 5:  # Linux scroll down
            self._canvas.yview_scroll(1, "units")
        elif event.delta:     # Windows/macOS scroll delta handling
            self._canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

    # ── Refresh / Card Generator ──────────────────────────────────────────────

    def _refresh_pages(self):
        """Uses twin-frame swap buffering to keep the canvas filled, completely eliminating flashes."""
        old_frame = self._cards_frame

        # Build the next layout fully in memory first
        new_frame = tk.Frame(self._canvas, bg=BG)

        pages = self._cfg.get("pages", [])
        for page_idx, page in enumerate(pages):
            self._build_page_card(new_frame, page_idx, page)

        def _on_frame_configure(_):
            self._canvas.configure(scrollregion=self._canvas.bbox("all"))
        new_frame.bind("<Configure>", _on_frame_configure)

        # Drop the ready frame instantly onto the active window coordinate
        self._canvas_win = self._canvas.create_window((0, 0), window=new_frame, anchor="nw")
        self._canvas.itemconfigure(self._canvas_win, width=self._canvas.winfo_width())
        self._cards_frame = new_frame

        # Recursively apply scroll binds throughout all freshly compiled UI objects
        self._bind_mouse_wheel(new_frame)

        # Safely sweep out the old container frame without revealing an empty workspace background
        if old_frame:
            old_frame.destroy()

    def _build_page_card(self, parent_frame, page_idx, page):
        is_sel = (page_idx == self._sel_page)
        border_clr = ACCENT if is_sel else BORDER

        card = tk.Frame(parent_frame, bg=PANEL, highlightthickness=1, highlightbackground=border_clr, pady=8)
        card.pack(fill="x", pady=(0, 12), padx=(0, 8))
        self._bind_mouse_wheel(card)

        def _select(_e, idx=page_idx):
            if self._sel_page != idx:
                self._sel_page = idx
                self._refresh_pages()

        card.bind("<Button-1>", _select)

        # Title bar frame
        header = tk.Frame(card, bg=PANEL)
        header.pack(fill="x", padx=12, pady=(0, 8))
        header.columnconfigure(1, weight=1)
        self._bind_mouse_wheel(header)

        # Primitive tk Checkbutton. Locked background, no OS flashes.
        en_var = tk.BooleanVar(value=page.get("enabled", True))
        chk = tk.Checkbutton(
            header, variable=en_var, bg=PANEL, fg=TEXT,
            selectcolor=BG, activebackground=PANEL, activeforeground=TEXT,
            relief="flat", bd=0, highlightthickness=0
        )
        chk.grid(row=0, column=0, sticky="w", padx=(2, 4))
        self._bind_mouse_wheel(chk)

        def _toggle_enabled(*_):
            page["enabled"] = en_var.get()
            self._save()
        en_var.trace_add("write", _toggle_enabled)

        # Title Label
        lbl_title = tk.Label(header, text=f"Page {page_idx + 1}", bg=PANEL, fg=TEXT, font=(FONT, 10, "bold"))
        lbl_title.grid(row=0, column=1, sticky="w", padx=4)
        self._bind_mouse_wheel(lbl_title)

        # Duration section label
        dur_lbl = tk.Label(header, text="Duration:", bg=PANEL, fg=SUBTEXT, font=(FONT, 8))
        dur_lbl.grid(row=0, column=2, padx=(8, 2))
        self._bind_mouse_wheel(dur_lbl)

        # Pure Tkinter flat stepper construct. No native layout white frames, fits numbers perfectly.
        counter_frame = tk.Frame(header, bg=BORDER, padx=1, pady=1)
        counter_frame.grid(row=0, column=3, padx=2)
        self._bind_mouse_wheel(counter_frame)

        dur_var = tk.StringVar(value=str(page.get("duration", 20)))

        def _decrement():
            try:
                val = int(dur_var.get().strip())
                if val > 1:
                    dur_var.set(str(val - 1))
            except ValueError:
                dur_var.set("1")

        def _increment():
            try:
                val = int(dur_var.get().strip())
                if val < 3600:
                    dur_var.set(str(val + 1))
            except ValueError:
                dur_var.set("20")

        btn_down = tk.Button(
            counter_frame, text="-", bg=PANEL, fg=SUBTEXT, font=(FONT, 8, "bold"),
            relief="flat", bd=0, width=2, height=1, cursor="hand2",
            activebackground=BORDER, activeforeground=TEXT, command=_decrement
        )
        btn_down.pack(side="left")
        self._bind_mouse_wheel(btn_down)

        dur_entry = tk.Entry(
            counter_frame, textvariable=dur_var, width=3, bg=BG, fg=TEXT,
            insertbackground=ACCENT, font=(FONT, 9, "bold"), relief="flat", justify="center"
        )
        dur_entry.pack(side="left", padx=1)
        self._bind_mouse_wheel(dur_entry)

        btn_up = tk.Button(
            counter_frame, text="+", bg=PANEL, fg=SUBTEXT, font=(FONT, 8, "bold"),
            relief="flat", bd=0, width=2, height=1, cursor="hand2",
            activebackground=BORDER, activeforeground=TEXT, command=_increment
        )
        btn_up.pack(side="left")
        self._bind_mouse_wheel(btn_up)

        def _dur_changed(*_, p_idx=page_idx, v=dur_var):
            val = v.get().strip()
            if not val:
                return
            try:
                self._cfg["pages"][p_idx]["duration"] = int(val)
                self._save()
            except (ValueError, KeyError, IndexError):
                pass

        dur_var.trace_add("write", _dur_changed)

        # Delete Page Button
        btn_del = tk.Button(
            header, text="✕", bg=PANEL, fg=RED, relief="flat", cursor="hand2", font=(FONT, 9, "bold"),
            activebackground=BORDER, activeforeground=RED, command=lambda: self._delete_page(page_idx)
        )
        btn_del.grid(row=0, column=4, padx=(8, 2))
        self._bind_mouse_wheel(btn_del)

        # Divider line inside card
        div = tk.Frame(card, bg=BORDER, height=1)
        div.pack(fill="x", padx=12, pady=(0, 8))
        self._bind_mouse_wheel(div)

        # Container for lines
        slots_frame = tk.Frame(card, bg=PANEL)
        slots_frame.pack(fill="x", padx=12)
        slots_frame.columnconfigure(0, weight=1)
        self._bind_mouse_wheel(slots_frame)

        slots = page.get("slots", [])
        for slot_idx, slot in enumerate(slots):
            self._build_slot_row(slots_frame, page_idx, slot_idx, slot, slots)

        # Bottom Row: Dedicated 'Add New Line' area panel
        add_row_idx = len(slots) if slots else 1
        bottom_action_frame = tk.Frame(slots_frame, bg=PANEL)
        bottom_action_frame.grid(row=add_row_idx, column=0, columnspan=6, sticky="ew", pady=(6, 2))
        self._bind_mouse_wheel(bottom_action_frame)

        if not slots:
            lbl_none = tk.Label(
                slots_frame, text="No modules on this page yet.",
                bg=PANEL, fg=SUBTEXT, font=(FONT, 8),
            )
            lbl_none.grid(row=0, column=0, columnspan=6, pady=4, sticky="w")
            self._bind_mouse_wheel(lbl_none)

        btn_add_line = tk.Button(
            bottom_action_frame, text="+ Add New Line", bg=BG, fg=ACCENT2, relief="flat",
            font=(FONT, 9, "bold"), cursor="hand2", activebackground=BORDER, activeforeground=TEXT,
            padx=10, pady=3
        )
        btn_add_line.pack(anchor="w", padx=2)
        btn_add_line.bind("<Button-1>", lambda e, pi=page_idx: self._prompt_add_new_line_row(pi))
        self._bind_mouse_wheel(btn_add_line)

        # Bind click triggers safely on child items to select the page without breaking backgrounds
        for w in [header, lbl_title, slots_frame, bottom_action_frame]:
            w.bind("<Button-1>", _select)

    def _build_slot_row(self, parent, page_idx, slot_idx, slot, slots):
        # Normalize original configurations into horizontal multi-module format
        if "modules" not in slot:
            slot["modules"] = [{"module": slot.get("module", ""), "text": slot.get("text", "")}]

        # Master line row wrapper block
        row_frame = tk.Frame(parent, bg=PANEL)
        row_frame.grid(row=slot_idx, column=0, columnspan=6, sticky="ew", pady=2)
        self._bind_mouse_wheel(row_frame)

        # Controls reordering tools layout
        handle = tk.Label(row_frame, text="⠿", bg=PANEL, fg=SUBTEXT, font=(FONT, 10), cursor="fleur")
        handle.pack(side="left", padx=(0, 2))
        handle.bind("<ButtonPress-1>",   lambda e, pi=page_idx, si=slot_idx: self._drag_start(e, pi, si))
        handle.bind("<B1-Motion>",       lambda e, pi=page_idx: self._drag_motion(e, pi))
        handle.bind("<ButtonRelease-1>", lambda e, pi=page_idx: self._drag_end(e, pi))
        self._bind_mouse_wheel(handle)

        btn_up = tk.Button(
            row_frame, text="▲", bg=PANEL, fg=SUBTEXT, relief="flat",
            font=(FONT, 7), cursor="hand2", width=2, activebackground=BORDER, activeforeground=TEXT,
            command=lambda pi=page_idx, si=slot_idx: self._move_slot(pi, si, -1),
        )
        btn_up.pack(side="left", padx=1)
        self._bind_mouse_wheel(btn_up)

        btn_down = tk.Button(
            row_frame, text="▼", bg=PANEL, fg=SUBTEXT, relief="flat",
            font=(FONT, 7), cursor="hand2", width=2, activebackground=BORDER, activeforeground=TEXT,
            command=lambda pi=page_idx, si=slot_idx: self._move_slot(pi, si, 1),
        )
        btn_down.pack(side="left", padx=(1, 6))
        self._bind_mouse_wheel(btn_down)

        # Container capsule holding side-by-side modules
        capsule = tk.Frame(row_frame, bg=PANEL)
        capsule.pack(side="left", fill="x", expand=True)
        self._bind_mouse_wheel(capsule)

        for m_idx, sub_slot in enumerate(slot["modules"]):
            mod = MODULE_BY_ID.get(sub_slot.get("module", ""))
            label = mod["label"] if mod else sub_slot.get("module", "unknown")

            mod_block = tk.Frame(capsule, bg=BORDER, padx=4, pady=2)
            mod_block.pack(side="left", padx=2)
            self._bind_mouse_wheel(mod_block)

            lbl_mod = tk.Label(mod_block, text=label, bg=BORDER, fg=TEXT, font=(FONT, 9))
            lbl_mod.pack(side="left", padx=4)
            self._bind_mouse_wheel(lbl_mod)

            # Inline customizable text entry if module dictates it
            if mod and mod.get("has_text"):
                txt_var = tk.StringVar(value=sub_slot.get("text", ""))
                entry = tk.Entry(
                    mod_block, textvariable=txt_var, width=12,
                    bg=BG, fg=TEXT, insertbackground=ACCENT, relief="flat", font=(FONT, 9),
                    highlightthickness=1, highlightbackground=BORDER, highlightcolor=ACCENT,
                )
                entry.pack(side="left", padx=4)
                self._bind_mouse_wheel(entry)

                def _text_changed(*_, var=txt_var, pi=page_idx, si=slot_idx, mi=m_idx):
                    try:
                        self._cfg["pages"][pi]["slots"][si]["modules"][mi]["text"] = var.get()
                        self._save()
                    except IndexError:
                        pass
                txt_var.trace_add("write", _text_changed)

            # Inline element delete tool button
            if len(slot["modules"]) > 1:
                btn_rem_sub = tk.Button(
                    mod_block, text="✕", bg=BORDER, fg=SUBTEXT, relief="flat", font=(FONT, 8),
                    activebackground=BORDER, activeforeground=RED, cursor="hand2",
                    command=lambda pi=page_idx, si=slot_idx, mi=m_idx: self._remove_sub_module(pi, si, mi),
                )
                btn_rem_sub.pack(side="left", padx=(2, 2))
                self._bind_mouse_wheel(btn_rem_sub)

        # Right Actions setup
        right_controls = tk.Frame(row_frame, bg=PANEL)
        right_controls.pack(side="right", padx=(2, 4))
        self._bind_mouse_wheel(right_controls)

        btn_app_mod = tk.Button(
            right_controls, text="+ Add Next To", bg=PANEL, fg=ACCENT2, relief="flat",
            font=(FONT, 8, "bold"), cursor="hand2", activebackground=BORDER, activeforeground=TEXT,
            command=lambda pi=page_idx, si=slot_idx: self._prompt_append_module(pi, si),
        )
        btn_app_mod.pack(side="left", padx=4)
        self._bind_mouse_wheel(btn_app_mod)

        btn_rem_slot = tk.Button(
            right_controls, text="✕", bg=PANEL, fg=RED, relief="flat",
            font=(FONT, 9), cursor="hand2", activebackground=BORDER, activeforeground=RED,
            command=lambda pi=page_idx, si=slot_idx: self._remove_slot(pi, si),
        )
        btn_rem_slot.pack(side="left", padx=2)
        self._bind_mouse_wheel(btn_rem_slot)

    # ── Actions / Mutators ───────────────────────────────────────────────────

    def _add_page(self):
        self._cfg.setdefault("pages", []).append({
            "enabled":  True,
            "duration": 6,
            "slots":    []
        })
        self._sel_page = len(self._cfg["pages"]) - 1
        self._save()
        self._refresh_pages()

    def _delete_page(self, page_idx):
        pages = self._cfg.get("pages", [])
        if 0 <= page_idx < len(pages):
            pages.pop(page_idx)
            if self._sel_page >= len(pages):
                self._sel_page = len(pages) - 1
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

    # ── Horizontal & Context Row Actions ──────────────────────────────────────

    def _prompt_append_module(self, page_idx: int, slot_idx: int):
        menu = tk.Menu(self, tearoff=0, bg=PANEL, fg=TEXT, activebackground=ACCENT, activeforeground=BG, font=(FONT, 9))

        for cat, mods in CATEGORIES.items():
            sub_menu = tk.Menu(menu, tearoff=0, bg=PANEL, fg=TEXT, activebackground=ACCENT, activeforeground=BG)
            for m in mods:
                sub_menu.add_command(
                    label=m["label"],
                    command=lambda m_id=m["id"], pi=page_idx, si=slot_idx: self._append_module_to_slot(pi, si, m_id)
                )
            menu.add_cascade(label=cat, menu=sub_menu)

        menu.post(self.winfo_pointerx(), self.winfo_pointery())

    def _append_module_to_slot(self, page_idx: int, slot_idx: int, module_id: str):
        try:
            slots = self._cfg["pages"][page_idx]["slots"]
            slots[slot_idx]["modules"].append({"module": module_id, "text": ""})
            self._save()
            self._refresh_pages()
        except IndexError:
            pass

    def _remove_sub_module(self, page_idx: int, slot_idx: int, module_idx: int):
        try:
            slots = self._cfg["pages"][page_idx]["slots"]
            slots[slot_idx]["modules"].pop(module_idx)
            self._save()
            self._refresh_pages()
        except IndexError:
            pass

    def _prompt_add_new_line_row(self, page_idx: int):
        menu = tk.Menu(self, tearoff=0, bg=PANEL, fg=TEXT, activebackground=ACCENT, activeforeground=BG, font=(FONT, 9))

        for cat, mods in CATEGORIES.items():
            sub_menu = tk.Menu(menu, tearoff=0, bg=PANEL, fg=TEXT, activebackground=ACCENT, activeforeground=BG)
            for m in mods:
                sub_menu.add_command(
                    label=m["label"],
                    command=lambda m_dict=m, pi=page_idx: self._add_slot_for_page(pi, m_dict)
                )
            menu.add_cascade(label=cat, menu=sub_menu)

        menu.post(self.winfo_pointerx(), self.winfo_pointery())

    def _add_slot_for_page(self, page_idx: int, mod: dict):
        try:
            pages = self._cfg.get("pages", [])
            slot = {"module": mod["id"]}
            if mod.get("has_text"):
                slot["text"] = ""
            pages[page_idx].setdefault("slots", []).append(slot)
            self._save()
            self._refresh_pages()
        except IndexError:
            pass

    def refresh(self):
        """Called externally when cfg changes outside the builder."""
        self._refresh_pages()