"""
ui/app.py
─────────
Root window for OSC-Router.
"""

import tkinter as tk
from tkinter import ttk, messagebox

from config  import load_config, save_config, get_defaults
from core.router import OscRouter, OutputTarget
from core.source import OscSource

from ui.theme           import BG, PANEL, BORDER, ACCENT, ACCENT2, TEXT, SUBTEXT, FONT, TITLE_PREFIX
from ui.router_tab      import RouterTab
from ui.help_dialog     import open_help
from ui.settings_dialog import open_settings

try:
    from main import VERSION
except ImportError:
    VERSION = "version error"


class App:
    def __init__(self):
        self._cfg         = load_config()
        self._router      = OscRouter()
        self._colour_mode = self._cfg.get("colour_mode", "new")

        self._build_root()
        self._build_tabs()

    # ── Root window ───────────────────────────────────────────────────────────

    def _build_root(self):
        self.root = tk.Tk()
        self.root.title(f"{TITLE_PREFIX} OSC-Router")
        self.root.configure(bg=BG)
        self.root.geometry("640x720")
        self.root.minsize(500, 460)

        style = ttk.Style()
        style.theme_use("default")
        style.configure("Dark.TNotebook", background=BG, borderwidth=0, tabmargins=0)
        style.configure("Dark.TNotebook.Tab", background=PANEL, foreground=SUBTEXT,
                        font=(FONT, 10), padding=(16, 6), borderwidth=0)
        style.map("Dark.TNotebook.Tab",
                  background=[("selected", BG)],
                  foreground=[("selected", ACCENT2)])

        header = tk.Frame(self.root, bg=PANEL, pady=8)
        header.pack(fill="x")
        tk.Label(header, text=f"{TITLE_PREFIX} OSC-Router",
                 bg=PANEL, fg=ACCENT2, font=(FONT, 13, "bold")).pack(side="left", padx=14)
        tk.Label(header, text=f"v{VERSION}",
                 bg=PANEL, fg=SUBTEXT, font=(FONT, 9)).pack(side="right", padx=14)
        tk.Frame(self.root, bg=BORDER, height=1).pack(fill="x")

        self.root.protocol("WM_DELETE_WINDOW", self._on_close)

    # ── Tabs ──────────────────────────────────────────────────────────────────

    def _build_tabs(self):
        self._notebook = ttk.Notebook(self.root, style="Dark.TNotebook")
        self._notebook.pack(fill="both", expand=True)

        self._router_tab = RouterTab(
            self._notebook, self._cfg, self._router,
            save_cb     = self._save,
            start_cb    = self._start,
            stop_cb     = self._stop,
            restart_cb  = self._restart,
            settings_cb = self._open_settings,
            help_cb     = self._open_help,
        )
        self._notebook.add(self._router_tab, text="  Router  ")

        for src in self._cfg.get("sources", []):
            self._router_tab.add_source_row(src.get("name", "Source"), src.get("port", 9001))

        for out in self._cfg.get("outputs", []):
            self._router_tab.add_output_row(
                name       = out.get("name", "Output"),
                ip         = out.get("ip",   "127.0.0.1"),
                port       = out.get("port", 9000),
                subscribed = out.get("sources", []),
            )

    # ── Router control ────────────────────────────────────────────────────────

    def _start(self):
        if self._router.running:
            return

        cfg = self._router_tab.collect_config()
        self._cfg.update(cfg)
        self._save()

        self._router.sources = [OscSource(s["name"], s["port"]) for s in cfg["sources"]]
        self._router.outputs = [
            OutputTarget(
                name         = o["name"],
                ip           = o["ip"],
                port         = o["port"],
                source_names = o.get("sources", [s["name"] for s in cfg["sources"]]),
            )
            for o in cfg["outputs"]
        ]

        result = self._router.start()

        msgs = []
        if result["sources"]:
            msgs.append(f"Sources failed to bind: {', '.join(result['sources'])}")
        if result["outputs"]:
            msgs.append(f"Outputs failed to open: {', '.join(result['outputs'])}")
        if msgs:
            messagebox.showwarning("Start Issues", "\n".join(msgs))

        self._router_tab.set_status("Running" if self._router.running else "Failed")

    def _stop(self):
        self._router.stop()
        self._router_tab.set_status("Stopped")

    def _restart(self):
        self._stop()
        self.root.after(800, self._start)

    # ── Config ────────────────────────────────────────────────────────────────

    def _save(self):
        self._cfg["colour_mode"] = self._colour_mode
        save_config(self._cfg)

    def _reset_to_defaults(self):
        defaults = get_defaults()
        self._cfg.clear()
        self._cfg.update(defaults)
        self._save()

    # ── Dialogs ───────────────────────────────────────────────────────────────

    def _open_settings(self):
        open_settings(
            root           = self.root,
            apply_color_fn = self._apply_colour_mode,
            get_color_fn   = lambda: self._colour_mode,
            save_cb        = self._save,
            reset_cb       = self._reset_to_defaults
        )

    def _open_help(self):
        open_help(self.root)

    # ── Colour Mode ───────────────────────────────────────────────────────────

    def _apply_colour_mode(self, mode: str):
        self._colour_mode = mode

    # ── Lifecycle ─────────────────────────────────────────────────────────────

    def _on_close(self):
        self._save()
        self._stop()  # Cleanly stop the router instead of looking for gamepad tab
        self.root.destroy()

    def run(self):
        self.root.mainloop()