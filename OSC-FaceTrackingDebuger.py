"""
OSC Face Tracking Controller
Sends VRCFaceTracking-compatible OSC parameters via sliders.
Requires: python-osc
"""

# =============================================================================
# IMPORTS AND DEPENDENCY BOOTSTRAP
# =============================================================================

import importlib
import subprocess
import sys
import threading
import tkinter as tk
from tkinter import font, ttk
from typing import Optional


def install_if_missing(package: str, import_name: str | None = None) -> None:
    if import_name is None:
        import_name = package.split("==")[0].replace("-", "_")

    try:
        importlib.import_module(import_name)
    except ImportError:
        print(f"Installing {package}...")
        subprocess.check_call([sys.executable, "-m", "pip", "install", package])


install_if_missing("python-osc==1.9.3", "pythonosc")
from pythonosc import udp_client


# =============================================================================
# CONFIGURATION
# =============================================================================

VERSION = "0.1.2"

DEFAULT_OSC_IP = "127.0.0.1"
DEFAULT_OSC_PORT = "9000"
DEFAULT_OSC_PREFIX = "/avatar/parameters/"

BG = "#0f0f13"
PANEL = "#17171f"
BORDER = "#2a2a38"
ACCENT = "#7c5cfc"
ACCENT2 = "#a78bfa"
TEXT = "#e2e0f0"
SUBTEXT = "#7e7b9a"
TROUGH = "#252535"
GREEN = "#4ade80"
RED = "#f87171"

FONT_FAMILY = "Consolas"

FACE_PARAMS = {
    "\U0001F441  Eyes": [
        ("EyeLidRight", 0.0, 1.0, 1.0),
        ("EyeLidLeft", 0.0, 1.0, 1.0),
        ("EyeRightX", -1.0, 1.0, 0.0),
        ("EyeRightY", -1.0, 1.0, 0.0),
        ("EyeLeftX", -1.0, 1.0, 0.0),
        ("EyeLeftY", -1.0, 1.0, 0.0),
        ("EyeSquintRight", 0.0, 1.0, 0.0),
        ("EyeSquintLeft", 0.0, 1.0, 0.0),
        ("EyeWideRight", 0.0, 1.0, 0.0),
        ("EyeWideLeft", 0.0, 1.0, 0.0),
    ],
    "\U0001F928 Brows": [
        ("BrowInnerUp", 0.0, 1.0, 0.0),
        ("BrowOuterUpRight", 0.0, 1.0, 0.0),
        ("BrowOuterUpLeft", 0.0, 1.0, 0.0),
        ("BrowDownRight", 0.0, 1.0, 0.0),
        ("BrowDownLeft", 0.0, 1.0, 0.0),
        ("BrowExpressionRight", -1.0, 1.0, 0.0),
        ("BrowExpressionLeft", -1.0, 1.0, 0.0),
    ],
    "\U0001F444 Mouth": [
        ("JawOpen", 0.0, 1.0, 0.0),
        ("JawX", -1.0, 1.0, 0.0),
        ("JawForward", 0.0, 1.0, 0.0),
        ("MouthSmileRight", 0.0, 1.0, 0.0),
        ("MouthSmileLeft", 0.0, 1.0, 0.0),
        ("MouthSadRight", 0.0, 1.0, 0.0),
        ("MouthSadLeft", 0.0, 1.0, 0.0),
        ("MouthPout", 0.0, 1.0, 0.0),
        ("MouthRaiserUpper", 0.0, 1.0, 0.0),
        ("MouthRaiserLower", 0.0, 1.0, 0.0),
        ("LipSuckUpper", 0.0, 1.0, 0.0),
        ("LipSuckLower", 0.0, 1.0, 0.0),
        ("LipFunnelUpper", 0.0, 1.0, 0.0),
        ("LipFunnelLower", 0.0, 1.0, 0.0),
    ],
    "\U0001F624 Cheek / Nose": [
        ("CheekPuffRight", 0.0, 1.0, 0.0),
        ("CheekPuffLeft", 0.0, 1.0, 0.0),
        ("CheekSuckRight", 0.0, 1.0, 0.0),
        ("CheekSuckLeft", 0.0, 1.0, 0.0),
        ("NoseSneerRight", 0.0, 1.0, 0.0),
        ("NoseSneerLeft", 0.0, 1.0, 0.0),
    ],
    "\U0001F445 Tongue": [
        ("TongueOut", 0.0, 1.0, 0.0),
        ("TongueX", -1.0, 1.0, 0.0),
        ("TongueY", -1.0, 1.0, 0.0),
        ("TongueRoll", 0.0, 1.0, 0.0),
        ("TongueBendDown", 0.0, 1.0, 0.0),
        ("TongueCurlUp", 0.0, 1.0, 0.0),
        ("TongueSquish", 0.0, 1.0, 0.0),
        ("TongueFlat", 0.0, 1.0, 0.0),
    ],
}


# =============================================================================
# UI CONTROLLER
# =============================================================================

class OscFaceController(tk.Tk):
    def __init__(self):
        super().__init__()
        self._configure_window()
        self._init_state()
        self._build_fonts()
        self._build_ui()
        self._try_connect()

    # -------------------------------------------------------------------------
    # Setup
    # -------------------------------------------------------------------------
    def _configure_window(self) -> None:
        self.title("OSC Face Tracking Controller")
        self.configure(bg=BG)
        self.resizable(True, True)
        self.minsize(520, 400)

    def _init_state(self) -> None:
        self._client: Optional[udp_client.SimpleUDPClient] = None
        self._connected = False
        self._lock = threading.Lock()
        self._vars: dict[str, tk.DoubleVar] = {}

    def _build_fonts(self) -> None:
        self.f_title = font.Font(family=FONT_FAMILY, size=13, weight="bold")
        self.f_head = font.Font(family=FONT_FAMILY, size=10, weight="bold")
        self.f_label = font.Font(family=FONT_FAMILY, size=9)
        self.f_value = font.Font(family=FONT_FAMILY, size=9)
        self.f_small = font.Font(family=FONT_FAMILY, size=8)
        self.f_btn = font.Font(family=FONT_FAMILY, size=9, weight="bold")

    # -------------------------------------------------------------------------
    # UI Builders
    # -------------------------------------------------------------------------
    def _build_ui(self) -> None:
        self._build_title_bar()
        tk.Frame(self, bg=BORDER, height=1).pack(fill="x")
        self._build_connection_row()
        tk.Frame(self, bg=BORDER, height=1).pack(fill="x")
        self._build_notebook()

    def _build_title_bar(self) -> None:
        title_bar = tk.Frame(self, bg=PANEL, pady=10)
        title_bar.pack(fill="x")

        tk.Label(
            title_bar,
            text="\u25C8  OSC FACE TRACKING",
            font=self.f_title,
            bg=PANEL,
            fg=ACCENT2,
        ).pack(side="left", padx=16)

        self._status_dot = tk.Label(title_bar, text="●", font=self.f_head, bg=PANEL, fg=RED)
        self._status_dot.pack(side="right", padx=(0, 8))
        self._status_lbl = tk.Label(
            title_bar,
            text="disconnected",
            font=self.f_small,
            bg=PANEL,
            fg=SUBTEXT,
        )
        self._status_lbl.pack(side="right")

    def _build_connection_row(self) -> None:
        conn_row = tk.Frame(self, bg=BG, pady=8)
        conn_row.pack(fill="x", padx=12)

        tk.Label(conn_row, text="IP", font=self.f_label, bg=BG, fg=SUBTEXT).pack(side="left")
        self._ip_var = tk.StringVar(value=DEFAULT_OSC_IP)
        ip_entry = tk.Entry(
            conn_row,
            textvariable=self._ip_var,
            width=14,
            bg=PANEL,
            fg=TEXT,
            insertbackground=ACCENT,
            relief="flat",
            font=self.f_label,
            highlightthickness=1,
            highlightbackground=BORDER,
            highlightcolor=ACCENT,
        )
        ip_entry.pack(side="left", padx=(4, 12))

        tk.Label(conn_row, text="Port", font=self.f_label, bg=BG, fg=SUBTEXT).pack(side="left")
        self._port_var = tk.StringVar(value=DEFAULT_OSC_PORT)
        port_entry = tk.Entry(
            conn_row,
            textvariable=self._port_var,
            width=6,
            bg=PANEL,
            fg=TEXT,
            insertbackground=ACCENT,
            relief="flat",
            font=self.f_label,
            highlightthickness=1,
            highlightbackground=BORDER,
            highlightcolor=ACCENT,
        )
        port_entry.pack(side="left", padx=(4, 12))

        tk.Label(conn_row, text="Prefix", font=self.f_label, bg=BG, fg=SUBTEXT).pack(side="left")
        self._prefix_var = tk.StringVar(value=DEFAULT_OSC_PREFIX)
        prefix_entry = tk.Entry(
            conn_row,
            textvariable=self._prefix_var,
            width=22,
            bg=PANEL,
            fg=TEXT,
            insertbackground=ACCENT,
            relief="flat",
            font=self.f_label,
            highlightthickness=1,
            highlightbackground=BORDER,
            highlightcolor=ACCENT,
        )
        prefix_entry.pack(side="left", padx=(4, 12))

        tk.Button(
            conn_row,
            text="CONNECT",
            font=self.f_btn,
            bg=ACCENT,
            fg="#ffffff",
            relief="flat",
            activebackground=ACCENT2,
            activeforeground="#ffffff",
            padx=10,
            pady=3,
            cursor="hand2",
            command=self._try_connect,
        ).pack(side="left")

        tk.Button(
            conn_row,
            text="RESET ALL",
            font=self.f_btn,
            bg=PANEL,
            fg=SUBTEXT,
            relief="flat",
            activebackground=BORDER,
            activeforeground=TEXT,
            padx=10,
            pady=3,
            cursor="hand2",
            command=self._reset_all,
        ).pack(side="left", padx=(6, 0))

    def _build_notebook(self) -> None:
        self._notebook = ttk.Notebook(self)
        self._style_notebook()
        self._notebook.pack(fill="both", expand=True, padx=0, pady=0)

        for category, params in FACE_PARAMS.items():
            self._add_tab(category, params)

    def _style_notebook(self) -> None:
        style = ttk.Style(self)
        style.theme_use("default")
        style.configure("TNotebook", background=BG, borderwidth=0, tabmargins=[0, 0, 0, 0])
        style.configure(
            "TNotebook.Tab",
            background=PANEL,
            foreground=SUBTEXT,
            font=(FONT_FAMILY, 9, "bold"),
            padding=[12, 6],
            borderwidth=0,
        )
        style.map(
            "TNotebook.Tab",
            background=[("selected", BG)],
            foreground=[("selected", ACCENT2)],
        )

    def _add_tab(self, label: str, params: list[tuple[str, float, float, float]]) -> None:
        outer = tk.Frame(self._notebook, bg=BG)
        self._notebook.add(outer, text=label)

        canvas = tk.Canvas(outer, bg=BG, highlightthickness=0)
        scrollbar = tk.Scrollbar(outer, orient="vertical", command=canvas.yview)
        canvas.configure(yscrollcommand=scrollbar.set)

        scrollbar.pack(side="right", fill="y")
        canvas.pack(side="left", fill="both", expand=True)

        inner = tk.Frame(canvas, bg=BG)
        win_id = canvas.create_window((0, 0), window=inner, anchor="nw")

        def _on_frame_configure(_event) -> None:
            canvas.configure(scrollregion=canvas.bbox("all"))

        def _on_canvas_configure(event) -> None:
            canvas.itemconfig(win_id, width=event.width)

        def _on_mousewheel(event) -> None:
            canvas.yview_scroll(-1 * (event.delta // 120), "units")

        def _bind_wheel(_event) -> None:
            canvas.bind_all("<MouseWheel>", _on_mousewheel)

        def _unbind_wheel(_event) -> None:
            canvas.unbind_all("<MouseWheel>")

        inner.bind("<Configure>", _on_frame_configure)
        canvas.bind("<Configure>", _on_canvas_configure)
        canvas.bind("<Enter>", _bind_wheel)
        canvas.bind("<Leave>", _unbind_wheel)

        for name, lo, hi, default in params:
            self._add_slider_row(inner, name, lo, hi, default)

    def _add_slider_row(self, parent, name: str, lo: float, hi: float, default: float) -> None:
        var = tk.DoubleVar(value=default)
        self._vars[name] = var

        row = tk.Frame(parent, bg=BG, pady=4)
        row.pack(fill="x", padx=14)

        tk.Label(row, text=name, font=self.f_label, bg=BG, fg=TEXT, width=22, anchor="w").pack(side="left")
        tk.Label(row, text=f"{lo:.1f}", font=self.f_small, bg=BG, fg=SUBTEXT, width=4).pack(side="left")

        tk.Scale(
            row,
            variable=var,
            from_=lo,
            to=hi,
            resolution=0.001,
            orient="horizontal",
            showvalue=False,
            bg=BG,
            fg=TEXT,
            troughcolor=TROUGH,
            activebackground=ACCENT2,
            highlightthickness=0,
            sliderrelief="flat",
            length=260,
            cursor="hand2",
        ).pack(side="left", padx=4)

        tk.Label(row, text=f"{hi:.1f}", font=self.f_small, bg=BG, fg=SUBTEXT, width=4).pack(side="left")

        val_lbl = tk.Label(
            row,
            text=f"{default:.3f}",
            font=self.f_value,
            bg=PANEL,
            fg=ACCENT2,
            width=7,
            anchor="e",
            padx=4,
            pady=1,
        )
        val_lbl.pack(side="left", padx=(4, 0))

        def _reset(n=name, v=var, d=default, vl=val_lbl) -> None:
            v.set(d)
            vl.config(text=f"{d:.3f}")
            self._send_osc(n, d)

        tk.Button(
            row,
            text="↺",
            font=self.f_btn,
            bg=PANEL,
            fg=SUBTEXT,
            relief="flat",
            padx=4,
            pady=0,
            cursor="hand2",
            activebackground=BORDER,
            activeforeground=TEXT,
            command=_reset,
        ).pack(side="left", padx=(4, 0))

        def _on_change(*_args, n=name, v=var, vl=val_lbl) -> None:
            val = round(v.get(), 3)
            vl.config(text=f"{val:.3f}")
            self._send_osc(n, val)

        var.trace_add("write", _on_change)

        tk.Frame(parent, bg=BORDER, height=1).pack(fill="x", padx=14)

    # -------------------------------------------------------------------------
    # OSC
    # -------------------------------------------------------------------------
    def _try_connect(self) -> None:
        ip = self._ip_var.get().strip()
        port = int(self._port_var.get().strip())
        try:
            with self._lock:
                self._client = udp_client.SimpleUDPClient(ip, port)
                self._connected = True
            self._set_status(True, f"{ip}:{port}")
        except Exception as exc:
            self._connected = False
            self._set_status(False, str(exc))

    def _send_osc(self, param: str, value: float) -> None:
        if not self._connected or self._client is None:
            return
        address = self._prefix_var.get() + param
        try:
            with self._lock:
                self._client.send_message(address, value)
        except Exception:
            pass

    def _set_status(self, ok: bool, detail: str = "") -> None:
        if ok:
            self._status_dot.config(fg=GREEN)
            self._status_lbl.config(text=f"connected  {detail}", fg=SUBTEXT)
        else:
            self._status_dot.config(fg=RED)
            self._status_lbl.config(text=f"disconnected  {detail}", fg=RED)

    def _reset_all(self) -> None:
        for params in FACE_PARAMS.values():
            for name, _lo, _hi, default in params:
                if name in self._vars:
                    self._vars[name].set(default)


# =============================================================================
# ENTRY POINT
# =============================================================================

def main() -> None:
    app = OscFaceController()
    app.mainloop()


if __name__ == "__main__":
    main()
