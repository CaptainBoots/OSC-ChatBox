"""
ui/scriptmaker_tab.py
────────────────────────
Scripts tab: connection bar (listen host/port, default out host/port,
Connect/Disconnect), a live fire-log, a toolbar (+ Add Script / Help /
Settings), then a scrollable list of ScriptCards.
"""

from PySide6.QtCore import Qt, QObject, Signal
from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QLabel, QPushButton, QScrollArea,
    QLineEdit, QPlainTextEdit,
)

from core.models import Script, default_script
from core.script_engine import ScriptEngine
from ui import theme
from ui.theme import StripeBackground
from ui.script_card import ScriptCard

MAX_LOG_LINES = 200


class _Bridge(QObject):
    """Background threads (the OSC listener / timer loop / firing script
    threads) never touch widgets directly — they emit here, and these
    signals are connected to main-thread slots, per the porting guide's
    thread-safety rule (Qt auto-queues cross-thread signal emits)."""
    log_line = Signal(str)


class ScriptMakerTab(StripeBackground):
    def __init__(self, cfg: dict, save_cb, help_cb, settings_cb, parent=None):
        super().__init__(parent)
        self._cfg = cfg
        self._save_cb = save_cb
        self._help_cb = help_cb
        self._settings_cb = settings_cb

        self.cards: list[ScriptCard] = []
        self._uid_counter = 0

        self._bridge = _Bridge()
        self._bridge.log_line.connect(self._append_log)

        self.engine = ScriptEngine(
            get_scripts_cb=self._collect_scripts_for_engine,
            default_out_host=cfg.get("out_host", "127.0.0.1"),
            default_out_port=cfg.get("out_port", "9000"),
            log_cb=lambda msg: self._bridge.log_line.emit(msg),
        )

        self._build()

    # ── Layout ────────────────────────────────────────────────────────────

    def _build(self):
        outer = QVBoxLayout(self)
        outer.setContentsMargins(8, 8, 8, 8)
        outer.setSpacing(4)

        self._build_connection_bar(outer)
        self._build_toolbar(outer)
        self._build_log(outer)

        self._scroll = QScrollArea()
        self._scroll.setWidgetResizable(True)
        self._scroll.setStyleSheet("background: transparent; border: none;")

        self._inner = QWidget()
        self._inner.setStyleSheet("background: transparent;")
        self._inner_layout = QVBoxLayout(self._inner)
        self._inner_layout.setContentsMargins(0, 0, 4, 0)
        self._inner_layout.setSpacing(6)
        self._inner_layout.addStretch(1)

        self._scroll.setWidget(self._inner)
        outer.addWidget(self._scroll, 1)

    def _build_connection_bar(self, outer):
        frame = QWidget()
        frame.setStyleSheet(f"background-color: {theme.PANEL}; border: none;")
        row = QHBoxLayout(frame)
        row.setContentsMargins(10, 6, 10, 6)

        def field(label, cfg_key, width):
            row.addWidget(self._label(label))
            e = QLineEdit(str(self._cfg.get(cfg_key, "")))
            e.setFixedWidth(width)
            e.setFont(theme.qt_font(9))
            e.setStyleSheet(theme.line_edit_qss())
            e.textChanged.connect(lambda t, k=cfg_key: self._cfg.__setitem__(k, t))
            row.addWidget(e)
            return e

        row.addWidget(self._label("Listen:"))
        self._listen_host = field("Host", "listen_host", 90)
        self._listen_port = field("Port", "listen_port", 55)
        row.addSpacing(12)
        row.addWidget(self._label("Send:"))
        self._out_host = field("Host", "out_host", 90)
        self._out_port = field("Port", "out_port", 55)
        row.addSpacing(12)

        def _on_out_changed(_t=None):
            self.engine.default_out_host = self._out_host.text().strip() or "127.0.0.1"
            self.engine.default_out_port = self._out_port.text().strip() or "9000"

        self._out_host.textChanged.connect(_on_out_changed)
        self._out_port.textChanged.connect(_on_out_changed)

        self._status_dot = QLabel("●")
        self._status_dot.setStyleSheet(f"color: {theme.SUBTEXT}; background: transparent; border: none;")
        self._status_dot.setFont(theme.qt_font(11))
        row.addWidget(self._status_dot)

        self._conn_btn = QPushButton("Connect")
        self._conn_btn.setCursor(Qt.PointingHandCursor)
        self._conn_btn.setFont(theme.qt_font(9, bold=True))
        self._set_connect_style(False)
        self._conn_btn.clicked.connect(self._toggle_connect)
        row.addWidget(self._conn_btn)
        row.addStretch(1)

        outer.addWidget(frame)

    def _build_toolbar(self, outer):
        btn_row = QHBoxLayout()
        btn_row.setContentsMargins(0, 4, 0, 4)

        add_btn = QPushButton("＋  Add Script")
        add_btn.setFont(theme.qt_font(10, bold=True))
        add_btn.setCursor(Qt.PointingHandCursor)
        add_btn.setMinimumWidth(130)
        add_btn.clicked.connect(self._add_script)
        btn_row.addWidget(add_btn)
        btn_row.addStretch(1)

        help_btn = QPushButton("? Help")
        help_btn.setStyleSheet(theme.subtle_button_qss())
        help_btn.setFont(theme.qt_font(9))
        help_btn.setCursor(Qt.PointingHandCursor)
        help_btn.clicked.connect(self._help_cb)
        btn_row.addWidget(help_btn)

        settings_btn = QPushButton("⚙ Settings")
        settings_btn.setStyleSheet(theme.subtle_button_qss())
        settings_btn.setFont(theme.qt_font(9))
        settings_btn.setCursor(Qt.PointingHandCursor)
        settings_btn.clicked.connect(self._settings_cb)
        btn_row.addWidget(settings_btn)

        outer.addLayout(btn_row)

    def _build_log(self, outer):
        cap = QLabel("ACTIVITY")
        cap.setStyleSheet(theme.section_caption_qss())
        cap.setFont(theme.qt_font(8, bold=True))
        outer.addWidget(cap)

        self._log = QPlainTextEdit()
        self._log.setReadOnly(True)
        self._log.setFixedHeight(70)
        self._log.setFont(theme.qt_font(8))
        self._log.setStyleSheet(
            f"QPlainTextEdit {{ background-color: {theme.PANEL}; color: {theme.SUBTEXT}; "
            f"border: 1px solid {theme.BORDER}; }}"
        )
        outer.addWidget(self._log)

    @staticmethod
    def _label(text):
        l = QLabel(text)
        l.setStyleSheet(f"color: {theme.SUBTEXT}; background: transparent; border: none;")
        l.setFont(theme.qt_font(9))
        return l

    def _append_log(self, msg: str):
        self._log.appendPlainText(msg)
        doc = self._log.document()
        if doc.blockCount() > MAX_LOG_LINES:
            cursor = self._log.textCursor()
            cursor.movePosition(cursor.MoveOperation.Start)
            cursor.movePosition(cursor.MoveOperation.Down, cursor.MoveMode.KeepAnchor,
                                 doc.blockCount() - MAX_LOG_LINES)
            cursor.removeSelectedText()

    # ── Connection ────────────────────────────────────────────────────────

    def _set_connect_style(self, connected: bool):
        if connected:
            self._conn_btn.setText("Disconnect")
            self._conn_btn.setStyleSheet(
                f"QPushButton {{ background-color: {theme.RED}; color: {theme.BG}; border: none; padding: 4px 10px; }}"
            )
        else:
            self._conn_btn.setText("Connect")
            self._conn_btn.setStyleSheet(theme.accent_button_qss())

    def _toggle_connect(self):
        self.disconnect_engine() if self.engine.is_running else self.connect_engine()

    def connect_engine(self):
        host = self._listen_host.text().strip() or "127.0.0.1"
        try:
            port = int(self._listen_port.text().strip())
        except ValueError:
            self._append_log("⚠ Invalid listen port")
            return
        try:
            self.engine.start_listener(host, port)
        except OSError as exc:
            self._append_log(f"⚠ Could not start listener: {exc}")
            return
        self._status_dot.setStyleSheet(f"color: {theme.GREEN}; background: transparent; border: none;")
        self._set_connect_style(True)
        self._append_log(f"Listening on {host}:{port}")

    def disconnect_engine(self):
        self.engine.stop_listener()
        self._status_dot.setStyleSheet(f"color: {theme.SUBTEXT}; background: transparent; border: none;")
        self._set_connect_style(False)
        self._append_log("Disconnected")

    # ── Scripts ──────────────────────────────────────────────────────────

    def load_scripts(self):
        saved = self._cfg.get("scripts", [])
        if saved:
            for sd in saved:
                s = Script.from_dict(sd)
                self._uid_counter = max(self._uid_counter, s.uid)
                self._add_script_card(s)
        else:
            self._add_script()
        if self._cfg.get("auto_start"):
            self.connect_engine()

    def _add_script(self):
        self._uid_counter += 1
        s = default_script(self._uid_counter)
        self._add_script_card(s)
        self._save_cb()

    def _add_script_card(self, script: Script):
        card = ScriptCard(script, on_remove=self._remove_script)
        self._inner_layout.insertWidget(self._inner_layout.count() - 1, card)
        self.cards.append(card)

    def _remove_script(self, card: ScriptCard):
        self._inner_layout.removeWidget(card)
        card.setParent(None)
        card.deleteLater()
        self.cards.remove(card)
        self._save_cb()

    def _collect_scripts_for_engine(self) -> list[Script]:
        return [c.get_config() for c in self.cards]

    # ── Config I/O ────────────────────────────────────────────────────────

    def collect_scripts(self) -> list[dict]:
        return [c.get_config().to_dict() for c in self.cards]

    def destroy_all(self):
        self.engine.shutdown()
