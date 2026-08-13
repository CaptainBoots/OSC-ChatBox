"""
ui/app.py
─────────
Root window (Qt). Creates the two-tab notebook and wires together:
  - BuilderTab  (page/slot editor)
  - ChatboxTab  (live preview + controls)
  - OSC loop start/stop/restart
  - Config load/save
  - Settings dialog
  - Theme selection

The OSC loop itself (osc_loop.py) is unchanged — it's plain threading, no
UI-toolkit dependency. The only thing that changes is how its callbacks
reach the UI thread: a small QObject bridge re-emits them as Qt signals
(thread-safe by default across threads).
"""

from PySide6.QtCore import Qt, QObject, Signal
from PySide6.QtWidgets import (
    QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QLabel, QTabWidget,
    QMessageBox, QApplication,
)

from config import load_config, save_config, get_defaults
from osc_loop import start_loop, stop_loop
from state import AppState
from ui.builder import BuilderTab
from ui.chatbox_tab import ChatboxTab
from ui.help_dialog import open_help
from ui.settings_dialog import open_settings
from ui import theme

try:
    from main import VERSION
except ImportError:
    VERSION = "version error"


class _LoopBridge(QObject):
    """Re-emits osc_loop's callbacks as Qt signals so they land safely on
    the UI thread no matter which worker thread calls them."""
    status_signal  = Signal(str)
    preview_signal = Signal(str)


class App(QMainWindow):
    def __init__(self):
        super().__init__()
        self._cfg   = load_config()
        self._state = AppState()

        # Apply state flags from config
        self._state.slow_mode        = self._cfg.get("slow_mode", False)
        self._state.speed_mode       = self._cfg.get("speed_mode", False)
        self._state.media_title_trim = self._cfg.get("media_title_trim", True)
        self._state.cat_mode         = self._cfg.get("cat_mode", False)
        self._state.progress_filled  = self._cfg.get("progress_filled", self._state.progress_filled)
        self._state.progress_border  = self._cfg.get("progress_border", self._state.progress_border)
        self._state.progress_empty   = self._cfg.get("progress_empty",  self._state.progress_empty)

        self._bridge = _LoopBridge()

        self._build_root()
        self._build_tabs()

        self._bridge.status_signal.connect(self._chatbox_tab.set_status)
        self._bridge.preview_signal.connect(self._chatbox_tab.set_preview)

    # ── Root window ───────────────────────────────────────────────────────────

    def _build_root(self):
        self.setWindowTitle(f"{theme.TITLE_PREFIX} OSC-Chatbox")
        self.resize(1080, 720)
        self.setMinimumSize(150, 30)

        # ── Real background-only transparency ────────────────────────────────
        # This is the actual fix Qt gives us over Tkinter: WA_TranslucentBackground
        # gives the window an alpha channel, and only the bare background fill
        # (StripeBackground / central widget) uses a colour with alpha < 255.
        # Every other widget (header, buttons, entries, tab bar) is painted
        # with a fully opaque colour, so it stays 100% legible AND clickable —
        # Qt's hit-testing follows widget geometry, not pixel alpha, unlike
        # Tk's colour-key trick which made matching pixels click-through too.
        self.setAttribute(Qt.WA_TranslucentBackground, True)
        saved_opacity = self._cfg.get("transparency_opacity", 1.0)
        self._bg_alpha = max(0.0, min(1.0, float(saved_opacity)))

        central = theme.StripeBackground()
        central.set_bg_alpha(self._bg_alpha)
        self._bg_canvas = central
        root_layout = QVBoxLayout(central)
        root_layout.setContentsMargins(0, 0, 0, 0)
        root_layout.setSpacing(0)

        # Header bar
        header = QWidget()
        header.setStyleSheet(f"background-color: {theme.PANEL};")
        header_layout = QHBoxLayout(header)
        header_layout.setContentsMargins(14, 8, 14, 8)

        title_lbl = QLabel(f"{theme.TITLE_PREFIX} OSC-Chatbox")
        title_lbl.setStyleSheet(f"color: {theme.ACCENT2}; background: transparent;")
        title_lbl.setFont(theme.qt_font(13, bold=True))
        header_layout.addWidget(title_lbl)
        header_layout.addStretch(1)

        version_lbl = QLabel(f"v{VERSION}")
        version_lbl.setStyleSheet(f"color: {theme.SUBTEXT}; background: transparent;")
        version_lbl.setFont(theme.qt_font(9))
        header_layout.addWidget(version_lbl)

        root_layout.addWidget(header)

        divider = QWidget()
        divider.setFixedHeight(1)
        divider.setStyleSheet(f"background-color: {theme.BORDER};")
        root_layout.addWidget(divider)

        self._notebook = QTabWidget()
        self._notebook.setDocumentMode(True)
        root_layout.addWidget(self._notebook, 1)

        self.setCentralWidget(central)

    # ── Tabs ──────────────────────────────────────────────────────────────────

    def _build_tabs(self):
        self._chatbox_tab = ChatboxTab(
            self._cfg, self._state,
            save_cb     = self._save,
            start_cb    = self._start,
            stop_cb     = self._stop,
            restart_cb  = self._restart,
            settings_cb = self._open_settings,
            help_cb     = self._open_help,
        )
        self._builder_tab = BuilderTab(self._cfg, save_cb=self._save)

        self._notebook.addTab(self._chatbox_tab, "  Chatbox  ")
        self._notebook.addTab(self._builder_tab, "  Builder  ")

    # ── OSC loop control ──────────────────────────────────────────────────────

    def _start(self):
        if self._state.running:
            return
        self._sync_state_from_cfg()
        start_loop(
            cfg        = self._cfg,
            state      = self._state,
            status_cb  = lambda text: self._bridge.status_signal.emit(text),
            preview_cb = lambda text: self._bridge.preview_signal.emit(text),
        )

    def _stop(self):
        stop_loop(self._state)

    def _restart(self):
        self._stop()
        from PySide6.QtCore import QTimer
        QTimer.singleShot(1200, self._start)

    # ── Config sync ───────────────────────────────────────────────────────────

    def _sync_state_from_cfg(self):
        self._state.slow_mode        = self._cfg.get("slow_mode", False)
        self._state.speed_mode       = self._cfg.get("speed_mode", False)
        self._state.media_title_trim = self._cfg.get("media_title_trim", True)
        self._state.cat_mode         = self._cfg.get("cat_mode", False)
        self._state.progress_filled  = self._cfg.get("progress_filled", self._state.progress_filled)
        self._state.progress_border  = self._cfg.get("progress_border", self._state.progress_border)
        self._state.progress_empty   = self._cfg.get("progress_empty",  self._state.progress_empty)

    def _save(self):
        self._cfg["slow_mode"]        = self._state.slow_mode
        self._cfg["speed_mode"]       = self._state.speed_mode
        self._cfg["media_title_trim"] = self._state.media_title_trim
        self._cfg["cat_mode"]         = self._state.cat_mode
        self._cfg["progress_filled"]  = self._state.progress_filled
        self._cfg["progress_border"]  = self._state.progress_border
        self._cfg["progress_empty"]   = self._state.progress_empty
        save_config(self._cfg)

    # ── Settings dialog ───────────────────────────────────────────────────────

    def _open_settings(self):
        open_settings(
            parent        = self,
            state         = self._state,
            cfg           = self._cfg,
            save_cb       = self._save,
            reset_cb      = self._reset_to_defaults,
            theme_cb      = self._set_theme,
            opacity_cb    = self._set_opacity,
        )

    def _open_help(self):
        open_help(self)

    def _reset_to_defaults(self):
        defaults = get_defaults()
        keep = {k: self._cfg[k] for k in
                ("pages", "osc_ip", "osc_port", "interface", "location", "theme_mode")
                if k in self._cfg}
        self._cfg.clear()
        self._cfg.update(defaults)
        self._cfg.update(keep)
        self._sync_state_from_cfg()
        self._save()
        self._builder_tab.refresh()

    # ── Theme ─────────────────────────────────────────────────────────────────

    def _set_theme(self, mode: str):
        self._cfg["theme_mode"] = mode
        self._save()
        QMessageBox.information(
            self, "Theme Changed",
            "Theme will apply after restarting OSC-Chatbox."
        )

    # ── Opacity (background-only — see WA_TranslucentBackground above) ───────

    def _set_opacity(self, alpha_val: float):
        alpha_val = max(0.0, min(1.0, float(alpha_val)))
        self._cfg["transparency_opacity"] = alpha_val
        self._save()
        self._bg_alpha = alpha_val
        self._bg_canvas.set_bg_alpha(alpha_val)

    # ── Lifecycle ─────────────────────────────────────────────────────────────

    def closeEvent(self, event):
        self._stop()
        self._save()
        super().closeEvent(event)

    # ── Entry point ───────────────────────────────────────────────────────────
    # Kept so main.py's existing `app = App(); app.run()` still works
    # unchanged. If a QApplication doesn't exist yet (normal case, since
    # QMainWindow needs one to already exist before it's constructed), this
    # creates one.

    def run(self):
        self.show()
        instance = QApplication.instance()
        if instance is not None:
            instance.exec()