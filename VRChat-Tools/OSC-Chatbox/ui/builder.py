import tkinter as tk

from modules.registry import CATEGORIES, MODULE_BY_ID
from ui.circle_toggle import CircleToggle
from ui.theme import BG, PANEL, BORDER, ACCENT, ACCENT2, TEXT, SUBTEXT, FONT, RED


class BuilderTab(tk.Frame):
    def __init__(self, parent, cfg: dict, save_cb):
        super().__init__(parent, bg=BG)
        self._cfg        = cfg
        self._save_cb    = save_cb
        self._sel_page   = 0
        self._drag       = {}
        self._canvas_win = None
        self._cards_frame = None

        self._build_ui()
        self._refresh_pages()

    # ── Label-button helper ───────────────────────────────────────────────────

    def _label_btn(self, parent, text, fg, command, *,
                   bg=PANEL, font_size=9, bold=True, padx=6, pady=2, width=None):

        kw = dict(
            bg=bg, fg=fg, cursor="hand2", padx=padx, pady=pady, relief="flat",
            font=(FONT, font_size, "bold" if bold else "normal"),
        )
        if width is not None:
            kw["width"] = width
        w = tk.Label(parent, text=text, **kw)
        w.bind("<Button-1>", lambda _e: command())
        w.bind("<Enter>",    lambda _e: w.config(bg=BORDER))
        w.bind("<Leave>",    lambda _e: w.config(bg=bg))
        return w

    # ── Layout ────────────────────────────────────────────────────────────────

    def _build_ui(self):
        self.columnconfigure(0, weight=1)
        self.rowconfigure(0, weight=1)

        right = tk.Frame(self, bg=BG)
        right.grid(row=0, column=0, sticky="nsew", padx=4)
        right.rowconfigure(1, weight=1)
        right.columnconfigure(0, weight=1)

        # Header
        r_hdr = tk.Frame(right, bg=BG)
        r_hdr.grid(row=0, column=0, sticky="ew", pady=(8, 8))

        tk.Label(
            r_hdr, text="Output Pages Setup",
            bg=BG, fg=TEXT, font=(FONT, 11, "bold"),
        ).pack(side="left")

        self._label_btn(
            r_hdr, "+ Create Page", ACCENT, self._add_page,
            bg=PANEL, padx=12, pady=4,
        ).pack(side="right")

        # Scrollable canvas
        self._canvas = tk.Canvas(right, bg=BG, highlightthickness=0, borderwidth=0)
        self._canvas.grid(row=1, column=0, sticky="nsew")

        self._vsb = tk.Scrollbar(right, orient="vertical", command=self._canvas.yview)
        self._vsb.grid(row=1, column=1, sticky="ns")
        self._canvas.configure(yscrollcommand=self._vsb.set)

        self._bind_mouse_wheel(self._canvas)

        def _on_canvas_configure(e):
            if self._canvas_win is not None:
                self._canvas.itemconfigure(self._canvas_win, width=e.width)

        self._canvas.bind("<Configure>", _on_canvas_configure)

    # ── Mouse wheel ───────────────────────────────────────────────────────────

    def _bind_mouse_wheel(self, widget):
        widget.bind("<MouseWheel>", self._on_mouse_wheel, add="+")
        widget.bind("<Button-4>",   self._on_mouse_wheel, add="+")
        widget.bind("<Button-5>",   self._on_mouse_wheel, add="+")

    def _on_mouse_wheel(self, event):
        if event.num == 4:
            self._canvas.yview_scroll(-1, "units")
        elif event.num == 5:
            self._canvas.yview_scroll(1, "units")
        elif event.delta:
            self._canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

    # ── Refresh ───────────────────────────────────────────────────────────────

    def _refresh_pages(self):
        old_frame = self._cards_frame

        new_frame = tk.Frame(self._canvas, bg=BG)

        pages = self._cfg.get("pages", [])
        for page_idx, page in enumerate(pages):
            self._build_page_card(new_frame, page_idx, page)

        def _on_frame_configure(_):
            self._canvas.configure(scrollregion=self._canvas.bbox("all"))
        new_frame.bind("<Configure>", _on_frame_configure)

        # Swap in-place: reuse the existing canvas window item so there is
        # never a moment where the canvas background is exposed.
        if self._canvas_win is not None:
            self._canvas.itemconfigure(self._canvas_win, window=new_frame)
        else:
            self._canvas_win = self._canvas.create_window((0, 0), window=new_frame, anchor="nw")

        self._canvas.itemconfigure(self._canvas_win, width=self._canvas.winfo_width())
        self._cards_frame = new_frame

        self._bind_mouse_wheel(new_frame)

        if old_frame:
            old_frame.destroy()

    # ── Page card ─────────────────────────────────────────────────────────────

    def _build_page_card(self, parent_frame, page_idx, page):
        is_sel     = (page_idx == self._sel_page)
        border_clr = ACCENT if is_sel else BORDER

        card = tk.Frame(
            parent_frame, bg=PANEL,
            highlightthickness=1, highlightbackground=border_clr, pady=8,
        )
        card.pack(fill="x", pady=(0, 12), padx=(0, 8))
        self._bind_mouse_wheel(card)

        def _select(_e=None, idx=page_idx):
            if self._sel_page != idx:
                self._sel_page = idx
                self._refresh_pages()

        card.bind("<Button-1>", _select)

        # ── Header row ────────────────────────────────────────────────────────
        header = tk.Frame(card, bg=PANEL)
        header.pack(fill="x", padx=12, pady=(0, 8))
        header.columnconfigure(1, weight=1)
        self._bind_mouse_wheel(header)

        # Flash-free Circle Toggle Widget replaces old text label checkbox
        def _toggle_enabled(is_enabled):
            page["enabled"] = is_enabled
            self._save()

        chk_toggle = CircleToggle(
            header,
            enabled=page.get("enabled", True),
            command=_toggle_enabled,
            bg=PANEL,
        )
        chk_toggle.grid(row=0, column=0, sticky="w", padx=(2, 4))
        self._bind_mouse_wheel(chk_toggle)

        # Page title
        lbl_title = tk.Label(
            header, text=f"Page {page_idx + 1}",
            bg=PANEL, fg=TEXT, font=(FONT, 10, "bold"),
        )
        lbl_title.grid(row=0, column=1, sticky="w", padx=4)
        self._bind_mouse_wheel(lbl_title)

        # Duration label
        dur_lbl = tk.Label(header, text="Duration:", bg=PANEL, fg=SUBTEXT, font=(FONT, 8))
        dur_lbl.grid(row=0, column=2, padx=(8, 2))
        self._bind_mouse_wheel(dur_lbl)

        # Duration stepper
        counter_frame = tk.Frame(header, bg=BORDER, padx=1, pady=1)
        counter_frame.grid(row=0, column=3, padx=2)
        self._bind_mouse_wheel(counter_frame)

        dur_var = tk.StringVar(value=str(page.get("duration", 20)))

        def _decrement():
            try:
                v = int(dur_var.get().strip())
                dur_var.set(str(max(1, v - 1)))
            except ValueError:
                dur_var.set("1")

        def _increment():
            try:
                v = int(dur_var.get().strip())
                dur_var.set(str(min(3600, v + 1)))
            except ValueError:
                dur_var.set("20")

        self._label_btn(
            counter_frame, "-", SUBTEXT, _decrement,
            bg=PANEL, font_size=8, padx=4, pady=1, width=2,
        ).pack(side="left")

        dur_entry = tk.Entry(
            counter_frame, textvariable=dur_var, width=3,
            bg=BG, fg=TEXT, insertbackground=ACCENT,
            font=(FONT, 9, "bold"), relief="flat", justify="center",
            highlightthickness=0,
        )
        dur_entry.pack(side="left", padx=1)
        self._bind_mouse_wheel(dur_entry)

        self._label_btn(
            counter_frame, "+", SUBTEXT, _increment,
            bg=PANEL, font_size=8, padx=4, pady=1, width=2,
        ).pack(side="left")

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

        # Delete page button
        self._label_btn(
            header, "✕", RED, lambda: self._delete_page(page_idx),
            bg=PANEL, font_size=9, padx=6, pady=2,
        ).grid(row=0, column=4, padx=(8, 2))

        # ── Divider ───────────────────────────────────────────────────────────
        div = tk.Frame(card, bg=BORDER, height=1)
        div.pack(fill="x", padx=12, pady=(0, 8))
        self._bind_mouse_wheel(div)

        # ── Slots ─────────────────────────────────────────────────────────────
        slots_frame = tk.Frame(card, bg=PANEL)
        slots_frame.pack(fill="x", padx=12)
        slots_frame.columnconfigure(0, weight=1)
        self._bind_mouse_wheel(slots_frame)

        slots = page.get("slots", [])
        for slot_idx, slot in enumerate(slots):
            self._build_slot_row(slots_frame, page_idx, slot_idx, slot, slots)

        # Empty state label
        if not slots:
            lbl_none = tk.Label(
                slots_frame, text="No modules on this page yet.",
                bg=PANEL, fg=SUBTEXT, font=(FONT, 8),
            )
            lbl_none.grid(row=0, column=0, columnspan=6, pady=4, sticky="w")
            self._bind_mouse_wheel(lbl_none)

        # Add new line button
        add_row_idx = len(slots) if slots else 1
        bottom_action_frame = tk.Frame(slots_frame, bg=PANEL)
        bottom_action_frame.grid(row=add_row_idx, column=0, columnspan=6, sticky="ew", pady=(6, 2))
        self._bind_mouse_wheel(bottom_action_frame)

        self._label_btn(
            bottom_action_frame, "+ Add New Line", ACCENT2,
            lambda pi=page_idx: self._prompt_add_new_line_row(pi),
            bg=BG, font_size=9, padx=10, pady=3,
        ).pack(anchor="w", padx=2)

        # Propagate page-select click to background areas
        for w in [header, lbl_title, slots_frame, bottom_action_frame]:
            w.bind("<Button-1>", _select)

    # ── Slot row ──────────────────────────────────────────────────────────────

    def _build_slot_row(self, parent, page_idx, slot_idx, slot, slots):
        # Normalise legacy single-module format
        if "modules" not in slot:
            slot["modules"] = [{"module": slot.get("module", ""), "text": slot.get("text", "")}]

        row_frame = tk.Frame(parent, bg=PANEL)
        row_frame.grid(row=slot_idx, column=0, columnspan=6, sticky="ew", pady=2)
        self._bind_mouse_wheel(row_frame)

        # Drag handle
        handle = tk.Label(row_frame, text="⠿", bg=PANEL, fg=SUBTEXT, font=(FONT, 10), cursor="fleur")
        handle.pack(side="left", padx=(0, 2))
        handle.bind("<ButtonPress-1>",   lambda e, pi=page_idx, si=slot_idx: self._drag_start(e, pi, si))
        handle.bind("<B1-Motion>",       lambda e, pi=page_idx: self._drag_motion(e, pi))
        handle.bind("<ButtonRelease-1>", lambda e, pi=page_idx: self._drag_end(e, pi))
        self._bind_mouse_wheel(handle)

        # Move up / down
        self._label_btn(
            row_frame, "▲", SUBTEXT,
            lambda pi=page_idx, si=slot_idx: self._move_slot(pi, si, -1),
            font_size=7, padx=3, pady=1, width=2,
        ).pack(side="left", padx=1)

        self._label_btn(
            row_frame, "▼", SUBTEXT,
            lambda pi=page_idx, si=slot_idx: self._move_slot(pi, si, 1),
            font_size=7, padx=3, pady=1, width=2,
        ).pack(side="left", padx=(1, 6))

        # Module capsule(s)
        capsule = tk.Frame(row_frame, bg=PANEL)
        capsule.pack(side="left", fill="x", expand=True)
        self._bind_mouse_wheel(capsule)

        for m_idx, sub_slot in enumerate(slot["modules"]):
            mod   = MODULE_BY_ID.get(sub_slot.get("module", ""))
            label = mod["label"] if mod else sub_slot.get("module", "unknown")

            mod_block = tk.Frame(capsule, bg=BORDER, padx=4, pady=2)
            mod_block.pack(side="left", padx=2)
            self._bind_mouse_wheel(mod_block)

            tk.Label(mod_block, text=label, bg=BORDER, fg=TEXT, font=(FONT, 9)).pack(side="left", padx=4)

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

            if len(slot["modules"]) > 1:
                self._label_btn(
                    mod_block, "✕", SUBTEXT,
                    lambda pi=page_idx, si=slot_idx, mi=m_idx: self._remove_sub_module(pi, si, mi),
                    bg=BORDER, font_size=8, padx=3, pady=1,
                ).pack(side="left", padx=(2, 2))

        # Right-side controls
        right_controls = tk.Frame(row_frame, bg=PANEL)
        right_controls.pack(side="right", padx=(2, 4))
        self._bind_mouse_wheel(right_controls)

        self._label_btn(
            right_controls, "+", ACCENT2,
            lambda pi=page_idx, si=slot_idx: self._prompt_append_module(pi, si),
            font_size=13, bold=True, padx=4, pady=1,
        ).pack(side="left", padx=4)

        self._label_btn(
            right_controls, "x", RED,
            lambda pi=page_idx, si=slot_idx: self._remove_slot(pi, si),
            font_size=13, bold=True, padx=4, pady=1,
        ).pack(side="left", padx=4)

    # ── Mutators ──────────────────────────────────────────────────────────────

    def _add_page(self):
        self._cfg.setdefault("pages", []).append({
            "enabled":  True,
            "duration": 6,
            "slots":    [],
        })
        self._sel_page = len(self._cfg["pages"]) - 1
        self._save()
        self._refresh_pages()

    def _delete_page(self, page_idx):
        pages = self._cfg.get("pages", [])
        if 0 <= page_idx < len(pages):
            pages.pop(page_idx)
            if self._sel_page >= len(pages):
                self._sel_page = max(0, len(pages) - 1)
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
            slots   = self._cfg["pages"][page_idx]["slots"]
            new_idx = slot_idx + direction
            if 0 <= new_idx < len(slots):
                slots[slot_idx], slots[new_idx] = slots[new_idx], slots[slot_idx]
                self._save()
                self._refresh_pages()
        except IndexError:
            pass

    # ── Drag-and-drop ─────────────────────────────────────────────────────────

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
        dy    = event.y_root - self._drag.get("y_start", event.y_root)
        steps = dy // 28
        if steps != 0:
            self._move_slot(self._drag["page"], self._drag["src"], steps)
        self._drag = {}

    # ── Save ──────────────────────────────────────────────────────────────────

    def _save(self):
        self._save_cb()

    # ── Context menus ─────────────────────────────────────────────────────────

    def _prompt_append_module(self, page_idx: int, slot_idx: int):
        menu = tk.Menu(
            self, tearoff=0,
            bg=PANEL, fg=TEXT, activebackground=ACCENT, activeforeground=BG, font=(FONT, 9),
        )
        for cat, mods in CATEGORIES.items():
            sub = tk.Menu(menu, tearoff=0, bg=PANEL, fg=TEXT, activebackground=ACCENT, activeforeground=BG)
            for m in mods:
                sub.add_command(
                    label=m["label"],
                    command=lambda m_id=m["id"], pi=page_idx, si=slot_idx:
                        self._append_module_to_slot(pi, si, m_id),
                )
            menu.add_cascade(label=cat, menu=sub)
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
        menu = tk.Menu(
            self, tearoff=0,
            bg=PANEL, fg=TEXT, activebackground=ACCENT, activeforeground=BG, font=(FONT, 9),
        )
        for cat, mods in CATEGORIES.items():
            sub = tk.Menu(menu, tearoff=0, bg=PANEL, fg=TEXT, activebackground=ACCENT, activeforeground=BG)
            for m in mods:
                sub.add_command(
                    label=m["label"],
                    command=lambda m_dict=m, pi=page_idx: self._add_slot_for_page(pi, m_dict),
                )
            menu.add_cascade(label=cat, menu=sub)
        menu.post(self.winfo_pointerx(), self.winfo_pointery())

    def _add_slot_for_page(self, page_idx: int, mod: dict):
        try:
            pages = self._cfg.get("pages", [])
            slot  = {"modules": [{"module": mod["id"], "text": ""}]}
            pages[page_idx].setdefault("slots", []).append(slot)
            self._save()
            self._refresh_pages()
        except IndexError:
            pass

    # ── External refresh ──────────────────────────────────────────────────────

    def refresh(self):
        self._refresh_pages()