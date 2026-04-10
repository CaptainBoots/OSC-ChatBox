# ═════════════════════════════════════════════════════════════════════════════════════════════════════════════════════#
#                                    OSC FaceTrackingController Script                                                 #
# ═════════════════════════════════════════════════════════════════════════════════════════════════════════════════════#
# Hi :3
# Wellcome to my code

# ═════════════════════════════════════════════════════════════════════════════════════════════════════════════════════#
# Imports
# ═════════════════════════════════════════════════════════════════════════════════════════════════════════════════════#

import importlib
import json
import os
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


# ═════════════════════════════════════════════════════════════════════════════════════════════════════════════════════#
# CONFIGURATION
# ═════════════════════════════════════════════════════════════════════════════════════════════════════════════════════#

VERSION = "0.3.0"

print("OSC FaceTrackingController")
print("Made By Boots")
print(f"Version {VERSION}")

DEFAULT_OSC_IP = "127.0.0.1"
DEFAULT_OSC_PORT = "9000"
DEFAULT_OSC_PREFIX = "/avatar/parameters/v2/"
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
CONFIG_DIR = SCRIPT_DIR
CONFIG_FILE = os.path.join(CONFIG_DIR, "face_tracking_config.json")

PREFIX_PRESETS = {
    "VRCFT v2 (default)": "/avatar/parameters/v2/",
    "Direct / v1": "/avatar/parameters/",
}


def get_default_config() -> dict:
    return {
        "osc_ip": DEFAULT_OSC_IP,
        "osc_port": DEFAULT_OSC_PORT,
        "osc_prefix": DEFAULT_OSC_PREFIX,
    }


def load_config() -> dict:
    defaults = get_default_config()
    try:
        with open(CONFIG_FILE, "r", encoding="utf-8") as f:
            return {**defaults, **json.load(f)}
    except (FileNotFoundError, json.JSONDecodeError, OSError):
        return defaults


def save_config(ip: str, port: str, prefix: str) -> None:
    os.makedirs(CONFIG_DIR, exist_ok=True)
    config = {
        "osc_ip": ip,
        "osc_port": str(port),
        "osc_prefix": prefix,
    }
    try:
        with open(CONFIG_FILE, "w", encoding="utf-8") as f:
            json.dump(config, f, indent=2)
    except OSError as e:
        print(f"[Config] Save failed: {e}")

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


# ═════════════════════════════════════════════════════════════════════════════════════════════════════════════════════#
# UI CONTROLLER
# ═════════════════════════════════════════════════════════════════════════════════════════════════════════════════════#
class OscFaceController(tk.Tk):
    def __init__(self):
        super().__init__()
        self._configure_window()
        self._init_state()
        self._build_fonts()
        self._build_ui()
        self._set_stopped()

# ═════════════════════════════════════════════════════════════════════════════════════════════════════════════════════#
# Setup
# ═════════════════════════════════════════════════════════════════════════════════════════════════════════════════════#
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
        self._config = load_config()

    def _build_fonts(self) -> None:
        self.f_title = font.Font(family=FONT_FAMILY, size=16, weight="bold")
        self.f_head = font.Font(family=FONT_FAMILY, size=10, weight="bold")
        self.f_label = font.Font(family=FONT_FAMILY, size=9)
        self.f_value = font.Font(family=FONT_FAMILY, size=9)
        self.f_small = font.Font(family=FONT_FAMILY, size=8)
        self.f_btn = font.Font(family=FONT_FAMILY, size=9, weight="bold")

# ═════════════════════════════════════════════════════════════════════════════════════════════════════════════════════#
# UI Builders
# ═════════════════════════════════════════════════════════════════════════════════════════════════════════════════════#
    @staticmethod
    def _square_button(parent: tk.Widget, text: str, command, base_size: int = 28) -> tk.Frame:
        container = tk.Frame(parent, bg=PANEL, highlightthickness=1, highlightbackground=BORDER)
        container.pack_propagate(False)

        button = tk.Button(
            container,
            text=text,
            command=command,
            bg=PANEL,
            fg=SUBTEXT,
            relief="flat",
            borderwidth=0,
            font=(FONT_FAMILY, 12),
            activebackground=BORDER,
            activeforeground=TEXT,
            cursor="hand2",
        )
        button.pack(fill="both", expand=True)
        container.config(width=base_size, height=base_size)
        return container

    def _build_ui(self) -> None:
        self._build_title_bar()
        tk.Frame(self, bg=BORDER, height=1).pack(fill="x")
        self._build_connection_row()
        self._build_action_row()
        tk.Frame(self, bg=BORDER, height=1).pack(fill="x")
        self._build_notebook()
        tk.Frame(self, bg=BORDER, height=1).pack(fill="x")
        self._build_footer_bar()

    def _build_title_bar(self) -> None:
        title_bar = tk.Frame(self, bg=PANEL, pady=14)
        title_bar.pack(fill="x")

        header_frame = tk.Frame(title_bar, bg=PANEL)
        header_frame.pack(fill="x", padx=16, expand=True)

        tk.Label(
            header_frame,
            text="◈  OSC FACE TRACKING",
            bg=PANEL,
            fg=ACCENT2,
            font=self.f_title,
        ).pack(side="left", anchor="w")

        version_label = tk.Label(
            header_frame,
            text=f"v{VERSION}",
            bg=PANEL,
            fg=SUBTEXT,
            font=self.f_small,
        )
        version_label.pack(side="right", anchor="e", padx=(32, 16))

        self._status_lbl = tk.Label(
            header_frame,
            text="Status: Stopped",
            bg=PANEL,
            fg=RED,
            font=(FONT_FAMILY, 10),
        )
        self._status_lbl.pack(side="right", anchor="e")

    def _build_connection_row(self) -> None:
        conn_row = tk.Frame(self, bg=BG, pady=8)
        conn_row.pack(fill="x", padx=12)

        tk.Label(conn_row, text="IP", font=self.f_label, bg=BG, fg=SUBTEXT).pack(side="left")
        self._ip_var = tk.StringVar(value=self._config.get("osc_ip", DEFAULT_OSC_IP))
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
        self._port_var = tk.StringVar(value=str(self._config.get("osc_port", DEFAULT_OSC_PORT)))
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
        self._prefix_var = tk.StringVar(value=self._config.get("osc_prefix", DEFAULT_OSC_PREFIX))
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
        prefix_entry.pack(side="left", padx=(4, 4))

        # Preset dropdown
        preset_var = tk.StringVar(value=list(PREFIX_PRESETS.keys())[0])

        def _on_preset_select(*_args) -> None:
            chosen = preset_var.get()
            if chosen in PREFIX_PRESETS:
                self._prefix_var.set(PREFIX_PRESETS[chosen])

        preset_var.trace_add("write", _on_preset_select)
        preset_menu = tk.OptionMenu(conn_row, preset_var, *PREFIX_PRESETS.keys())
        preset_menu.config(
            bg=PANEL, fg=SUBTEXT, activebackground=BORDER, activeforeground=TEXT,
            highlightthickness=0, relief="flat", font=self.f_small, cursor="hand2",
            indicatoron=True, bd=0,
        )
        preset_menu["menu"].config(bg=PANEL, fg=TEXT, activebackground=ACCENT, activeforeground="#fff")
        preset_menu.pack(side="left")

    def _build_action_row(self) -> None:
        button_row = tk.Frame(self, bg=BG)
        button_row.pack(fill="x", padx=12, pady=(0, 10))

        button_row.columnconfigure(0, weight=1)
        button_row.columnconfigure(1, weight=1)
        button_row.columnconfigure(2, weight=1)
        button_row.columnconfigure(3, weight=1)

        self._start_btn = tk.Button(
            button_row,
            text="Start",
            command=self._start_connection,
            bg=ACCENT,
            fg="#FFFFFF",
            disabledforeground="#FFFFFF",
            relief="flat",
            activebackground=ACCENT2,
            activeforeground="#FFFFFF",
            cursor="hand2",
            font=self.f_btn,
        )
        self._start_btn.grid(row=0, column=0, sticky="ew", padx=2)

        self._stop_btn = tk.Button(
            button_row,
            text="Stop",
            command=self._stop_connection,
            bg=PANEL,
            fg=SUBTEXT,
            relief="flat",
            activebackground=BORDER,
            activeforeground=TEXT,
            cursor="hand2",
            font=self.f_btn,
        )
        self._stop_btn.grid(row=0, column=1, sticky="ew", padx=2)

        self._restart_btn = tk.Button(
            button_row,
            text="Restart",
            command=self._restart_connection,
            bg=PANEL,
            fg=SUBTEXT,
            relief="flat",
            activebackground=BORDER,
            activeforeground=TEXT,
            cursor="hand2",
            font=self.f_btn,
        )
        self._restart_btn.grid(row=0, column=2, sticky="ew", padx=2)

        self._reset_btn = tk.Button(
            button_row,
            text="Reset All",
            command=self._reset_all,
            bg=PANEL,
            fg=SUBTEXT,
            relief="flat",
            activebackground=BORDER,
            activeforeground=TEXT,
            cursor="hand2",
            font=self.f_btn,
        )
        self._reset_btn.grid(row=0, column=3, sticky="ew", padx=2)
        self._update_connection_buttons()

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
        canvas.pack(fill="both", expand=True)

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

    def _build_footer_bar(self) -> None:
        footer_bar = tk.Frame(self, bg=PANEL, pady=8)
        footer_bar.pack(fill="x", side="bottom")
        footer_bar.columnconfigure(0, weight=1)

        help_btn = self._square_button(footer_bar, "?", self._open_help, base_size=28)
        help_btn.pack(side="left", padx=(8, 0))

        self._footer_label = tk.Label(
            footer_bar,
            text="Ready",
            bg=PANEL,
            fg=SUBTEXT,
            font=(FONT_FAMILY, 8),
        )
        self._footer_label.pack(side="left", padx=16)

    def _open_help(self) -> None:
        help_win = tk.Toplevel(self)
        help_win.title("OSC Face Tracking - Help")
        help_win.configure(bg=BG)
        help_win.resizable(True, True)

        self.update_idletasks()
        help_w = self.winfo_width()
        help_h = self.winfo_height()
        root_x = self.winfo_x()
        root_y = self.winfo_y()
        help_win.geometry(f"{help_w}x{help_h}+{root_x}+{root_y}")

        pages = [
            {
                "title": "Welcome",
                "content": (
                    "Welcome to OSC Face Tracking Controller!\n\n"
                    "This tool connects to VRChat and other applications to send face tracking data via OSC (Open Sound Control).\n\n"
                    "Use the tabs below to explore different facial parameters like eyes, mouth, brows, and more."
                ),
            },
            {
                "title": "Connection",
                "content": (
                    "Connection Setup:\n\n"
                    "1. IP Address: The target application's IP (usually 127.0.0.1 for local)\n"
                    "2. Port: OSC port number (default: 9000)\n"
                    "3. Prefix: OSC address prefix (VRCFT v2 or Direct/v1)\n\n"
                    "Click START to establish connection.\n"
                    "Green dot = running, Red dot = stopped or error"
                ),
            },
            {
                "title": "Parameters",
                "content": (
                    "Facial Parameters:\n\n"
                    "• Eyes: Eye movement, eyelid, blinking\n"
                    "• Brows: Eyebrow position and expression\n"
                    "• Mouth: Jaw, lips, smile, and other mouth shapes\n"
                    "• Cheek/Nose: Puffing and scrunching\n"
                    "• Tongue: Tongue position and movement\n\n"
                    "Use sliders to adjust values from 0 to 1 (or -1 to 1)."
                ),
            },
            {
                "title": "Tips",
                "content": (
                    "Tips & Tricks:\n\n"
                    "• RESET ALL button resets all parameters to defaults\n"
                    "• Changes are sent in real-time\n"
                    "• Use presets to quickly switch between prefix types\n"
                    "• The ↺ button resets individual parameters\n\n"
                    "For issues or questions, check your connection status and verify the target application is running."
                ),
            },
        ]

        current_page = [0]

        header = tk.Frame(help_win, bg=PANEL, pady=10)
        header.pack(fill="x")

        title_label = tk.Label(
            header, text="", bg=PANEL, fg=ACCENT2, font=(FONT_FAMILY, 12, "bold")
        )
        title_label.pack(side="left", padx=16)

        page_indicator = tk.Label(
            header, text="", bg=PANEL, fg=SUBTEXT, font=(FONT_FAMILY, 8)
        )
        page_indicator.pack(side="right", padx=16)

        tk.Frame(help_win, bg=BORDER, height=1).pack(fill="x")

        content_panel = tk.Frame(help_win, bg=PANEL, highlightthickness=1, highlightbackground=BORDER)
        content_panel.pack(padx=20, pady=(14, 0), fill="both", expand=True)

        content_text = tk.Label(
            content_panel, text="", bg=PANEL, fg=TEXT, anchor="nw", justify="left",
            font=(FONT_FAMILY, 9), wraplength=400
        )
        content_text.pack(padx=16, pady=16, fill="both", expand=True)

        def show_page(page_num):
            page = pages[page_num]
            title_label.config(text=page["title"])
            content_text.config(text=page["content"])
            page_indicator.config(text=f"Page {page_num + 1} / {len(pages)}")
            prev_btn.config(state="normal" if page_num > 0 else "disabled")
            is_last = page_num == len(pages) - 1
            next_btn.config(text="Finish" if is_last else "Next →")

        def prev_page() -> None:
            if current_page[0] > 0:
                current_page[0] -= 1
                show_page(current_page[0])

        def next_or_finish():
            if current_page[0] < len(pages) - 1:
                current_page[0] += 1
                show_page(current_page[0])
            else:
                help_win.destroy()

        nav_frame = tk.Frame(help_win, bg=BG)
        nav_frame.pack(fill="x", padx=20, pady=(0, 14))
        nav_frame.columnconfigure(1, weight=1)

        prev_btn = tk.Button(
            nav_frame, text="← Back", bg=PANEL, fg=SUBTEXT, relief="flat", width=10,
            command=prev_page
        )
        prev_btn.grid(row=0, column=0, sticky="w")
        prev_btn.configure(
            fg=SUBTEXT, activebackground=BORDER, activeforeground=TEXT,
            cursor="hand2", font=(FONT_FAMILY, 9, "bold"),
        )

        next_btn = tk.Button(
            nav_frame, text="Next →", bg=PANEL, fg=SUBTEXT, relief="flat", width=10,
            command=next_or_finish
        )
        next_btn.grid(row=0, column=2, sticky="e")
        next_btn.configure(
            bg=ACCENT, fg="#FFFFFF", activebackground=ACCENT2, activeforeground="#FFFFFF",
            cursor="hand2", font=(FONT_FAMILY, 9, "bold"),
        )

        show_page(0)

# ═════════════════════════════════════════════════════════════════════════════════════════════════════════════════════#
# OSC
# ═════════════════════════════════════════════════════════════════════════════════════════════════════════════════════#
    def _update_connection_buttons(self) -> None:
        if not hasattr(self, "_start_btn"):
            return
        self._start_btn.config(state="disabled" if self._connected else "normal")
        self._stop_btn.config(state="normal" if self._connected else "disabled")
        self._restart_btn.config(state="normal")

    def _start_connection(self) -> None:
        ip = self._ip_var.get().strip() or DEFAULT_OSC_IP
        port_text = self._port_var.get().strip() or DEFAULT_OSC_PORT
        prefix = self._normalize_prefix()
        self._ip_var.set(ip)
        self._port_var.set(port_text)
        self._prefix_var.set(prefix)

        try:
            port = int(port_text)
            if not (1 <= port <= 65535):
                raise ValueError("Port must be between 1 and 65535")

            with self._lock:
                self._client = udp_client.SimpleUDPClient(ip, port)
                self._connected = True

            self._port_var.set(str(port))
            save_config(ip, str(port), prefix)

            self._set_status(True, f"{ip}:{port}")
        except Exception as exc:
            self._client = None
            self._connected = False
            self._set_status(False, str(exc))

    def _stop_connection(self, set_status: bool = True) -> None:
        with self._lock:
            self._client = None
            self._connected = False
        if set_status:
            self._set_stopped()
        else:
            self._update_connection_buttons()

    def _restart_connection(self) -> None:
        self._stop_connection(set_status=False)
        self._start_connection()

    def _normalize_prefix(self) -> str:
        prefix = self._prefix_var.get().strip() or DEFAULT_OSC_PREFIX
        if not prefix.startswith("/"):
            prefix = "/" + prefix
        if not prefix.endswith("/"):
            prefix += "/"
        return prefix

    @staticmethod
    def _prefix_variants(normalized_prefix: str) -> list[str]:
        variants = [normalized_prefix]
        marker = "/avatar/parameters/"
        marker_pos = normalized_prefix.find(marker)
        if marker_pos == -1:
            return variants

        prefix_head = normalized_prefix[: marker_pos + len(marker)]
        prefix_tail = normalized_prefix[marker_pos + len(marker):]

        if prefix_tail.startswith("v2/"):
            alt = prefix_head + prefix_tail[len("v2/"):]
        else:
            alt = prefix_head + prefix_tail + "v2/"

        if alt not in variants:
            variants.append(alt)
        return variants

    @staticmethod
    def _param_payloads(param: str, value: float) -> list[tuple[str, float]]:
        payloads = [(param, value)]

        if param == "JawForward":
            payloads.append(("JawZ", value))
        elif param == "TongueBendDown":
            payloads.append(("TongueArchY", value))
        elif param == "TongueCurlUp":
            payloads.append(("TongueArchY", -value))
        elif param == "TongueSquish":
            payloads.append(("TongueShape", value))
        elif param == "TongueFlat":
            payloads.append(("TongueShape", -value))

        return payloads

    def _send_osc(self, param: str, value: float) -> None:
        if not self._connected or self._client is None:
            return
        prefixes = self._prefix_variants(self._normalize_prefix())
        payloads = self._param_payloads(param, value)
        sent_addresses: set[str] = set()
        try:
            with self._lock:
                for prefix in prefixes:
                    for param_name, param_value in payloads:
                        address = prefix + param_name
                        if address in sent_addresses:
                            continue
                        self._client.send_message(address, param_value)
                        sent_addresses.add(address)
        except (OSError, RuntimeError, ValueError):
            pass

    def _set_status(self, ok: bool, detail: str = "") -> None:
        if ok:
            self._status_lbl.config(text="Status: Running", fg=GREEN)
            self._footer_label.config(text="Running")
        else:
            self._status_lbl.config(text="Status: Error", fg=RED)
            self._footer_label.config(text=f"Error: {detail}" if detail else "Error")
        self._update_connection_buttons()

    def _set_stopped(self) -> None:
        self._status_lbl.config(text="Status: Stopped", fg=RED)
        self._footer_label.config(text="Stopped")
        self._update_connection_buttons()

    def _reset_all(self) -> None:
        for params in FACE_PARAMS.values():
            for name, _lo, _hi, default in params:
                if name in self._vars:
                    self._vars[name].set(default)


# ═════════════════════════════════════════════════════════════════════════════════════════════════════════════════════#
# ENTRY POINT
# ═════════════════════════════════════════════════════════════════════════════════════════════════════════════════════#

def main() -> None:
    app = OscFaceController()
    app.mainloop()


if __name__ == "__main__":
    main()
