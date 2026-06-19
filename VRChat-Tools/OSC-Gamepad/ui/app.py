"""
ui/app.py
─────────
Root window for OSC-Gamepad.
Same structure and theme as OSC-Chatbox / OSC-Router:
header bar with title + version, dark notebook, single tab.
"""

import tkinter as tk
from tkinter import ttk, messagebox

from config import load_config, save_config, get_defaults
from ui.theme import BG, PANEL, BORDER, ACCENT, ACCENT2, TEXT, SUBTEXT, FONT, TITLE_PREFIX
from ui.gamepad_tab import GamepadTab
from ui.help_dialog import open_help
from ui.settings_dialog import open_settings

try:
    from main import VERSION
except ImportError:
    VERSION = "version error"


class App:
    def __init__(self):
        self._cfg      = load_config()

        self._build_root()
        self._build_tabs()
        self._gamepad_tab.load_pads()

        self.root.protocol("WM_DELETE_WINDOW", self._on_close)

    # ── Root window ───────────────────────────────────────────────────────────

    def _build_root(self):
        self.root = tk.Tk()
        self.root.title(f"{TITLE_PREFIX} OSC-Gamepad")
        self.root.configure(bg=BG)
        self.root.geometry("680x640")
        self.root.minsize(520, 420)

        style = ttk.Style()
        style.theme_use("clam")
        style.configure(
            "Dark.TNotebook",
            background=BG, borderwidth=0, tabmargins=(0, 0, 0, 0),
            bordercolor=BG, lightcolor=BG, darkcolor=BG,
        )
        # Clam draws a separate border around the notebook's client area
        # using its own layout element; remove that element entirely so
        # no light-coloured border shows around the tab content.
        style.layout("Dark.TNotebook", [
            ("Notebook.client", {"sticky": "nswe"})
        ])
        style.configure(
            "Dark.TNotebook.Tab",
            background=PANEL, foreground=SUBTEXT,
            font=(FONT, 10), padding=(16, 6), borderwidth=0,
            bordercolor=BG, lightcolor=PANEL, darkcolor=PANEL,
        )
        style.map(
            "Dark.TNotebook.Tab",
            background=[("selected", BG)],
            foreground=[("selected", ACCENT2)],
            lightcolor=[("selected", BG)],
            darkcolor=[("selected", BG)],
        )

        header = tk.Frame(self.root, bg=PANEL, pady=8)
        header.pack(fill="x")
        tk.Label(
            header, text=f"{TITLE_PREFIX} OSC-Gamepad",
            bg=PANEL, fg=ACCENT2, font=(FONT, 13, "bold"),
        ).pack(side="left", padx=14)
        tk.Label(
            header, text=f"v{VERSION}",
            bg=PANEL, fg=SUBTEXT, font=(FONT, 9),
        ).pack(side="right", padx=14)
        tk.Frame(self.root, bg=BORDER, height=1).pack(fill="x")

    # ── Tabs ──────────────────────────────────────────────────────────────────

    def _build_tabs(self):
        self._notebook = ttk.Notebook(self.root, style="Dark.TNotebook")
        self._notebook.pack(fill="both", expand=True)

        self._gamepad_tab = GamepadTab(
            self._notebook, self._cfg,
            save_cb     = self._save,
            help_cb     = self._open_help,
            settings_cb = self._open_settings,
        )
        self._notebook.add(self._gamepad_tab, text="  Pads  ")

    # ── Config ────────────────────────────────────────────────────────────────

    def _save(self):
        pads_data = self._gamepad_tab.collect_pads()
        self._cfg["pads"] = pads_data
        save_config(pads_data, self._cfg.get("theme_mode", "new"))

    def _reset_to_defaults(self):
        defaults = get_defaults()
        keep_pads  = self._cfg.get("pads", [])
        keep_theme = self._cfg.get("theme_mode", "new")
        self._cfg.clear()
        self._cfg.update(defaults)
        self._cfg["pads"]       = keep_pads
        self._cfg["theme_mode"] = keep_theme
        self._save()

    # ── Dialogs ───────────────────────────────────────────────────────────────

    def _open_settings(self):
        open_settings(
            root     = self.root,
            cfg      = self._cfg,
            save_cb  = self._save,
            reset_cb = self._reset_to_defaults,
            theme_cb = self._set_theme,
        )

    def _open_help(self):
        open_help(self.root)

    # ── Theme ─────────────────────────────────────────────────────────────────

    def _set_theme(self, mode: str):
        self._cfg["theme_mode"] = mode
        self._save()
        messagebox.showinfo(
            "Theme Changed",
            "Theme will apply after restarting OSC-Gamepad."
        )

    # ── Lifecycle ─────────────────────────────────────────────────────────────

    def _on_close(self):
        self._save()
        self._gamepad_tab.destroy_all()
        self.root.destroy()

    def run(self):
        self.root.mainloop()