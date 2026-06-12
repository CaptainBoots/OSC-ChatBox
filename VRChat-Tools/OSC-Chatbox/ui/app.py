"""
ui/app.py
─────────
Root window. Creates the two-tab notebook and wires together:
  - BuilderTab  (page/slot editor)
  - ChatboxTab  (live preview + controls)
  - OSC loop start/stop/restart
  - Config load/save
  - Settings dialog
"""

import sys
import threading
import tkinter as tk
from tkinter import ttk, font as tkfont

from config import load_config, save_config, get_defaults
from state  import AppState
from osc_loop import start_loop, stop_loop

from ui.theme           import BG, PANEL, BORDER, ACCENT, ACCENT2, TEXT, SUBTEXT, FONT, TITLE_PREFIX
from ui.builder         import BuilderTab
from ui.chatbox_tab     import ChatboxTab
from ui.settings_dialog import open_settings
from ui.help_dialog     import open_help

VERSION = "8.0.1"


class App:
    def __init__(self):
        self._cfg   = load_config()
        self._state = AppState()

        # Apply state flags from config
        self._state.slow_mode         = self._cfg.get("slow_mode", False)
        self._state.speed_mode        = self._cfg.get("speed_mode", False)
        self._state.media_title_trim  = self._cfg.get("media_title_trim", True)
        self._state.cat_mode          = self._cfg.get("cat_mode", False)
        self._state.progress_filled   = self._cfg.get("progress_filled", self._state.progress_filled)
        self._state.progress_border   = self._cfg.get("progress_border", self._state.progress_border)
        self._state.progress_empty    = self._cfg.get("progress_empty",  self._state.progress_empty)

        self._build_root()
        self._build_tabs()

    # ── Root window ───────────────────────────────────────────────────────────

    def _build_root(self):
        self.root = tk.Tk()
        self.root.title(f"{TITLE_PREFIX} OSC-Chatbox")
        self.root.configure(bg=BG)
        self.root.geometry("900x680")
        self.root.minsize(700, 500)

        # Style notebook tabs
        style = ttk.Style()
        style.theme_use("default")
        style.configure(
            "Dark.TNotebook",
            background=BG, borderwidth=0, tabmargins=0,
        )
        style.configure(
            "Dark.TNotebook.Tab",
            background=PANEL, foreground=SUBTEXT,
            font=(FONT, 10), padding=(16, 6),
            borderwidth=0,
        )
        style.map(
            "Dark.TNotebook.Tab",
            background=[("selected", BG)],
            foreground=[("selected", ACCENT2)],
        )

        # Header bar
        header = tk.Frame(self.root, bg=PANEL, pady=8)
        header.pack(fill="x")
        tk.Label(
            header, text=f"{TITLE_PREFIX} OSC-Chatbox",
            bg=PANEL, fg=ACCENT2, font=(FONT, 13, "bold"),
        ).pack(side="left", padx=14)
        tk.Label(
            header, text=f"v{VERSION}",
            bg=PANEL, fg=SUBTEXT, font=(FONT, 9),
        ).pack(side="right", padx=14)
        tk.Frame(self.root, bg=BORDER, height=1).pack(fill="x")

        self.root.protocol("WM_DELETE_WINDOW", self._on_close)

    # ── Tabs ──────────────────────────────────────────────────────────────────

    def _build_tabs(self):
        self._notebook = ttk.Notebook(self.root, style="Dark.TNotebook")
        self._notebook.pack(fill="both", expand=True)

        self._builder_tab = BuilderTab(
            self._notebook, self._cfg, save_cb=self._save,
        )
        self._chatbox_tab = ChatboxTab(
            self._notebook, self._cfg, self._state,
            save_cb     = self._save,
            start_cb    = self._start,
            stop_cb     = self._stop,
            restart_cb  = self._restart,
            settings_cb = self._open_settings,
            help_cb     = self._open_help,
        )

        self._notebook.add(self._chatbox_tab, text="  Chatbox  ")
        self._notebook.add(self._builder_tab,  text="  Builder  ")

    # ── OSC loop control ──────────────────────────────────────────────────────

    def _start(self):
        if self._state.running:
            return
        self._sync_state_from_cfg()
        start_loop(
            cfg        = self._cfg,
            state      = self._state,
            status_cb  = self._on_status,
            preview_cb = self._on_preview,
        )

    def _stop(self):
        stop_loop(self._state)

    def _restart(self):
        self._stop()
        self.root.after(1200, self._start)

    def _on_status(self, text: str):
        self.root.after(0, lambda: self._chatbox_tab.set_status(text))

    def _on_preview(self, text: str):
        self.root.after(0, lambda: self._chatbox_tab.set_preview(text))

    # ── Config Sync ───────────────────────────────────────────────────────────

    def _sync_state_from_cfg(self):
        self._state.slow_mode        = self._cfg.get("slow_mode", False)
        self._state.speed_mode       = self._cfg.get("speed_mode", False)
        self._state.media_title_trim = self._cfg.get("media_title_trim", True)
        self._state.cat_mode         = self._cfg.get("cat_mode", False)
        self._state.progress_filled  = self._cfg.get("progress_filled", self._state.progress_filled)
        self._state.progress_border  = self._cfg.get("progress_border", self._state.progress_border)
        self._state.progress_empty   = self._cfg.get("progress_empty",  self._state.progress_empty)

    def _save(self):
        # Sync mutable state back into cfg before saving
        self._cfg["slow_mode"]        = self._state.slow_mode
        self._cfg["speed_mode"]       = self._state.speed_mode
        self._cfg["media_title_trim"] = self._state.media_title_trim
        self._cfg["cat_mode"]         = self._state.cat_mode
        self._cfg["progress_filled"]  = self._state.progress_filled
        self._cfg["progress_border"]  = self._state.progress_border
        self._cfg["progress_empty"]   = self._state.progress_empty
        save_config(self._cfg)

    # ── Settings Dialog ───────────────────────────────────────────────────────

    def _open_settings(self):
        open_settings(
            root    = self.root,
            state   = self._state,
            save_cb = self._save,
            reset_cb= self._reset_to_defaults,
        )

    def _open_help(self):
        open_help(self.root)

    def _reset_to_defaults(self):
        defaults = get_defaults()
        # Preserve pages and connection settings
        keep = {k: self._cfg[k] for k in ("pages", "osc_ip", "osc_port", "interface", "location") if k in self._cfg}
        self._cfg.clear()
        self._cfg.update(defaults)
        self._cfg.update(keep)
        self._sync_state_from_cfg()
        self._save()
        self._builder_tab.refresh()

    # ── Lifecycle ─────────────────────────────────────────────────────────────

    def _on_close(self):
        self._stop()
        self._save()
        self.root.destroy()

    def run(self):
        self.root.mainloop()