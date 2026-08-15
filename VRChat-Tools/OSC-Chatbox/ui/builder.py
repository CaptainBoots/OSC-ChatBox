"""
ui/builder.py
──────────────────────
Qt replacement for ui/builder.py. Same structure: a scrollable list of
page cards, each with a header (enable toggle, title, duration stepper,
delete) and a list of slot rows (drag handle, reorder arrows, module
capsules with optional inline text, add/remove controls).
"""

from PySide6.QtCore import Qt, QMimeData, QPoint
from PySide6.QtGui import QDrag, QCursor
from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QGridLayout, QLabel, QLineEdit,
    QScrollArea, QFrame, QMenu, QSizePolicy,
)

from modules.registry import CATEGORIES, MODULE_BY_ID
from ui.circle_toggle import CircleToggle
from ui import theme
from ui.theme import StripeBackground


def _label_btn(parent_layout, text, fg, command, *, bg=None, font_size=9,
               bold=True, padding="2px 6px"):
    bg = bg or theme.PANEL
    lbl = QLabel(text)
    lbl.setCursor(Qt.PointingHandCursor)
    lbl.setStyleSheet(f"color: {fg}; background-color: {bg}; padding: {padding}; border: none;")
    lbl.setFont(theme.qt_font(font_size, bold=bold))

    def _enter(_e):
        lbl.setStyleSheet(f"color: {fg}; background-color: {theme.BORDER}; padding: {padding}; border: none;")

    def _leave(_e):
        lbl.setStyleSheet(f"color: {fg}; background-color: {bg}; padding: {padding}; border: none;")

    lbl.enterEvent = _enter
    lbl.leaveEvent = _leave
    lbl.mousePressEvent = lambda _e: command()
    parent_layout.addWidget(lbl)
    return lbl


def _hline(parent_layout):
    line = QFrame()
    line.setFixedHeight(1)
    line.setStyleSheet(f"background-color: {theme.BORDER}; border: none;")
    parent_layout.addWidget(line)


class BuilderTab(StripeBackground):
    def __init__(self, cfg: dict, save_cb):
        super().__init__()
        self._cfg = cfg
        self._save_cb = save_cb
        self._sel_page = 0
        self._build_ui()

    # ── Layout ────────────────────────────────────────────────────────────

    def _build_ui(self):
        outer = QVBoxLayout(self)
        outer.setContentsMargins(4, 8, 4, 8)

        header_row = QHBoxLayout()
        title = QLabel("Output Pages Setup")
        title.setStyleSheet(f"color: {theme.TEXT}; background: {theme.BG}; border: none;")
        title.setFont(theme.qt_font(11, bold=True))
        header_row.addWidget(title)
        header_row.addStretch(1)
        _label_btn(header_row, "+ Create Page", theme.ACCENT, self._add_page,
                   font_size=9, padding="4px 12px")
        outer.addLayout(header_row)

        self._scroll = QScrollArea()
        self._scroll.setWidgetResizable(True)
        self._scroll.setStyleSheet("background: transparent; border: none;")
        outer.addWidget(self._scroll, 1)

        self._refresh_pages()

    # ── Refresh ───────────────────────────────────────────────────────────

    def _refresh_pages(self):
        container = QWidget()
        container.setStyleSheet("background: transparent;")
        layout = QVBoxLayout(container)
        layout.setContentsMargins(0, 0, 8, 0)
        layout.setSpacing(12)

        pages = self._cfg.get("pages", [])
        for page_idx, page in enumerate(pages):
            layout.addWidget(self._build_page_card(page_idx, page))
        layout.addStretch(1)

        self._scroll.setWidget(container)

    # ── Page card ─────────────────────────────────────────────────────────

    def _build_page_card(self, page_idx, page):
        is_sel = (page_idx == self._sel_page)
        border_clr = theme.ACCENT if is_sel else theme.BORDER

        card = QFrame()
        card.setStyleSheet(
            f"background-color: {theme.PANEL}; border: 1px solid {border_clr};"
        )
        card_layout = QVBoxLayout(card)
        card_layout.setContentsMargins(12, 8, 12, 8)

        def _select(_evt=None, idx=page_idx):
            if self._sel_page != idx:
                self._sel_page = idx
                self._refresh_pages()

        card.mousePressEvent = _select

        # ── Header row ────────────────────────────────────────────────────
        header = QHBoxLayout()

        def _toggle_enabled(is_enabled):
            page["enabled"] = is_enabled
            self._save()

        chk_toggle = CircleToggle(enabled=page.get("enabled", True), color=theme.ACCENT)
        chk_toggle.toggled.connect(_toggle_enabled)
        header.addWidget(chk_toggle)

        lbl_title = QLabel(f"Page {page_idx + 1}")
        lbl_title.setStyleSheet(f"color: {theme.TEXT}; background: transparent; border: none;")
        lbl_title.setFont(theme.qt_font(10, bold=True))
        header.addWidget(lbl_title)

        header.addSpacing(8)
        dur_lbl = QLabel("Duration:")
        dur_lbl.setStyleSheet(f"color: {theme.SUBTEXT}; background: transparent; border: none;")
        dur_lbl.setFont(theme.qt_font(8))
        header.addWidget(dur_lbl)

        counter = QHBoxLayout()
        counter.setSpacing(1)

        dur_entry = QLineEdit(str(page.get("duration", 20)))
        dur_entry.setFixedWidth(34)
        dur_entry.setAlignment(Qt.AlignCenter)
        dur_entry.setFont(theme.qt_font(9, bold=True))
        dur_entry.setStyleSheet(theme.line_edit_qss())

        def _decrement():
            try:
                v = int(dur_entry.text().strip())
            except ValueError:
                v = 1
            dur_entry.setText(str(max(1, v - 1)))
            _dur_changed()

        def _increment():
            try:
                v = int(dur_entry.text().strip())
            except ValueError:
                v = 20
            dur_entry.setText(str(min(3600, v + 1)))
            _dur_changed()

        minus_lbl = QLabel("-")
        minus_lbl.setCursor(Qt.PointingHandCursor)
        minus_lbl.setStyleSheet(f"color: {theme.SUBTEXT}; background-color: {theme.PANEL}; padding: 1px 4px; border: none;")
        minus_lbl.setFont(theme.qt_font(8))
        minus_lbl.mousePressEvent = lambda _e: _decrement()
        counter.addWidget(minus_lbl)
        counter.addWidget(dur_entry)

        plus_lbl = QLabel("+")
        plus_lbl.setCursor(Qt.PointingHandCursor)
        plus_lbl.setStyleSheet(f"color: {theme.SUBTEXT}; background-color: {theme.PANEL}; padding: 1px 4px; border: none;")
        plus_lbl.setFont(theme.qt_font(8))
        plus_lbl.mousePressEvent = lambda _e: _increment()
        counter.addWidget(plus_lbl)

        counter_wrap = QWidget()
        counter_wrap.setLayout(counter)
        counter_wrap.setStyleSheet(f"background-color: {theme.BG}; border: {theme.BORDER};")
        header.addWidget(counter_wrap)

        def _dur_changed():
            val = dur_entry.text().strip()
            if not val:
                return
            try:
                self._cfg["pages"][page_idx]["duration"] = int(val)
                self._save()
            except (ValueError, KeyError, IndexError):
                pass

        dur_entry.editingFinished.connect(_dur_changed)

        header.addStretch(1)
        _label_btn(header, "✕", theme.RED, lambda: self._delete_page(page_idx), font_size=9)

        card_layout.addLayout(header)
        _hline(card_layout)

        # ── Slots ─────────────────────────────────────────────────────────
        slots_layout = QVBoxLayout()
        slots_layout.setSpacing(3)
        slots = page.get("slots", [])
        for slot_idx, slot in enumerate(slots):
            slots_layout.addWidget(self._build_slot_row(page_idx, slot_idx, slot))

        if not slots:
            none_lbl = QLabel("No modules on this page yet.")
            none_lbl.setStyleSheet(f"color: {theme.SUBTEXT}; background: transparent; border: none;")
            none_lbl.setFont(theme.qt_font(8))
            slots_layout.addWidget(none_lbl)

        card_layout.addLayout(slots_layout)

        add_row = QHBoxLayout()
        _label_btn(add_row, "+ Add New Line", theme.ACCENT2,
                   lambda pi=page_idx: self._prompt_add_new_line_row(pi),
                   bg=theme.BG, font_size=9, padding="3px 10px")
        add_row.addStretch(1)
        card_layout.addLayout(add_row)

        return card

    # ── Slot row ──────────────────────────────────────────────────────────

    def _build_slot_row(self, page_idx, slot_idx, slot):
        if "modules" not in slot:
            slot["modules"] = [{"module": slot.get("module", ""), "text": slot.get("text", "")}]

        row = QFrame()
        row.setStyleSheet(f"background-color: {theme.PANEL}; border: none;")
        row_layout = QHBoxLayout(row)
        row_layout.setContentsMargins(0, 2, 0, 2)

        handle = QLabel("⠿")
        handle.setCursor(Qt.SizeAllCursor)
        handle.setStyleSheet(f"color: {theme.SUBTEXT}; background: transparent; border: none;")
        handle.setFont(theme.qt_font(10))
        self._wire_drag_handle(handle, page_idx, slot_idx)
        row_layout.addWidget(handle)

        up_lbl = QLabel("▲")
        up_lbl.setCursor(Qt.PointingHandCursor)
        up_lbl.setStyleSheet(f"color: {theme.SUBTEXT}; background-color: {theme.PANEL}; padding: 1px 3px; border: none;")
        up_lbl.setFont(theme.qt_font(7))
        up_lbl.mousePressEvent = lambda _e, pi=page_idx, si=slot_idx: self._move_slot(pi, si, -1)
        row_layout.addWidget(up_lbl)

        down_lbl = QLabel("▼")
        down_lbl.setCursor(Qt.PointingHandCursor)
        down_lbl.setStyleSheet(f"color: {theme.SUBTEXT}; background-color: {theme.PANEL}; padding: 1px 3px; border: none;")
        down_lbl.setFont(theme.qt_font(7))
        down_lbl.mousePressEvent = lambda _e, pi=page_idx, si=slot_idx: self._move_slot(pi, si, 1)
        row_layout.addWidget(down_lbl)

        capsule = QHBoxLayout()
        capsule.setSpacing(4)

        for m_idx, sub_slot in enumerate(slot["modules"]):
            mod = MODULE_BY_ID.get(sub_slot.get("module", ""))
            label = mod["label"] if mod else sub_slot.get("module", "unknown")

            mod_block = QFrame()
            mod_block.setStyleSheet(f"background-color: {theme.BORDER}; border: none;")
            mod_block_layout = QHBoxLayout(mod_block)
            mod_block_layout.setContentsMargins(4, 2, 4, 2)

            mlbl = QLabel(label)
            mlbl.setStyleSheet(f"color: {theme.TEXT}; background: transparent; border: none;")
            mlbl.setFont(theme.qt_font(9))
            mod_block_layout.addWidget(mlbl)

            if mod and mod.get("has_text"):
                entry = QLineEdit(sub_slot.get("text", ""))
                entry.setFixedWidth(90)
                entry.setFont(theme.qt_font(9))
                entry.setStyleSheet(theme.line_edit_qss())

                def _text_changed(text, pi=page_idx, si=slot_idx, mi=m_idx):
                    try:
                        self._cfg["pages"][pi]["slots"][si]["modules"][mi]["text"] = text
                        self._save()
                    except IndexError:
                        pass

                entry.textChanged.connect(_text_changed)
                mod_block_layout.addWidget(entry)

            if len(slot["modules"]) > 1:
                x_lbl = QLabel("✕")
                x_lbl.setCursor(Qt.PointingHandCursor)
                x_lbl.setStyleSheet(f"color: {theme.SUBTEXT}; background-color: {theme.BORDER}; padding: 1px 3px; border: none;")
                x_lbl.setFont(theme.qt_font(8))
                x_lbl.mousePressEvent = (
                    lambda _e, pi=page_idx, si=slot_idx, mi=m_idx: self._remove_sub_module(pi, si, mi)
                )
                mod_block_layout.addWidget(x_lbl)

            capsule.addWidget(mod_block)

        capsule.addStretch(1)
        row_layout.addLayout(capsule, 1)

        _label_btn(row_layout, "+", theme.ACCENT2,
                   lambda pi=page_idx, si=slot_idx: self._prompt_append_module(pi, si),
                   font_size=13, padding="1px 6px")
        _label_btn(row_layout, "x", theme.RED,
                   lambda pi=page_idx, si=slot_idx: self._remove_slot(pi, si),
                   font_size=13, padding="1px 6px")

        return row

    # ── Drag-and-drop (mirrors the Tk version: measure dy on release, no
    #    live animation — same behaviour as the original handle) ────────────

    def _wire_drag_handle(self, handle, page_idx, slot_idx):
        drag_state = {}

        def _press(event):
            drag_state["page"] = page_idx
            drag_state["src"] = slot_idx
            drag_state["y_start"] = event.globalPosition().y()

        def _release(event):
            if not drag_state:
                return
            dy = event.globalPosition().y() - drag_state.get("y_start", event.globalPosition().y())
            steps = int(dy // 28)
            if steps != 0:
                self._move_slot(drag_state["page"], drag_state["src"], steps)
            drag_state.clear()

        handle.mousePressEvent = _press
        handle.mouseReleaseEvent = _release

    # ── Mutators ──────────────────────────────────────────────────────────

    def _add_page(self):
        self._cfg.setdefault("pages", []).append({"enabled": True, "duration": 6, "slots": []})
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

    def _remove_slot(self, page_idx, slot_idx):
        try:
            self._cfg["pages"][page_idx]["slots"].pop(slot_idx)
            self._save()
            self._refresh_pages()
        except IndexError:
            pass

    def _move_slot(self, page_idx, slot_idx, direction):
        try:
            slots = self._cfg["pages"][page_idx]["slots"]
            new_idx = slot_idx + direction
            if 0 <= new_idx < len(slots):
                slots[slot_idx], slots[new_idx] = slots[new_idx], slots[slot_idx]
                self._save()
                self._refresh_pages()
        except IndexError:
            pass

    def _save(self):
        self._save_cb()

    # ── Module picker menus ─────────────────────────────────────────────────

    def _build_category_menu(self, on_pick):
        menu = QMenu(self)
        menu.setStyleSheet(
            f"QMenu {{ background-color: {theme.PANEL}; color: {theme.TEXT}; "
            f"border: 1px solid {theme.BORDER}; }} "
            f"QMenu::item:selected {{ background-color: {theme.ACCENT}; color: {theme.BG}; }}"
        )
        for cat, mods in CATEGORIES.items():
            sub = menu.addMenu(cat)
            sub.setStyleSheet(menu.styleSheet())
            for m in mods:
                action = sub.addAction(m["label"])
                action.triggered.connect(lambda _checked=False, m_id=m["id"]: on_pick(m_id))
        return menu

    def _prompt_append_module(self, page_idx, slot_idx):
        menu = self._build_category_menu(
            lambda m_id: self._append_module_to_slot(page_idx, slot_idx, m_id)
        )
        menu.exec(QCursor.pos())

    def _append_module_to_slot(self, page_idx, slot_idx, module_id):
        try:
            slots = self._cfg["pages"][page_idx]["slots"]
            slots[slot_idx]["modules"].append({"module": module_id, "text": ""})
            self._save()
            self._refresh_pages()
        except IndexError:
            pass

    def _remove_sub_module(self, page_idx, slot_idx, module_idx):
        try:
            slots = self._cfg["pages"][page_idx]["slots"]
            slots[slot_idx]["modules"].pop(module_idx)
            self._save()
            self._refresh_pages()
        except IndexError:
            pass

    def _prompt_add_new_line_row(self, page_idx):
        menu = self._build_category_menu(lambda m_id: self._add_slot_for_page(page_idx, m_id))
        menu.exec(QCursor.pos())

    def _add_slot_for_page(self, page_idx, module_id):
        try:
            pages = self._cfg.get("pages", [])
            slot = {"modules": [{"module": module_id, "text": ""}]}
            pages[page_idx].setdefault("slots", []).append(slot)
            self._save()
            self._refresh_pages()
        except IndexError:
            pass

    # ── External refresh ─────────────────────────────────────────────────

    def refresh(self):
        self._refresh_pages()