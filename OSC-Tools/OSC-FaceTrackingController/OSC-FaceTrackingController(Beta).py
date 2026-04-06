# ═════════════════════════════════════════════════════════════════════════════════════════════════════════════════════#
#                                    OSC FaceTrackingController Script                                                 #
# ═════════════════════════════════════════════════════════════════════════════════════════════════════════════════════#
# Hi :3
# Wellcome to my code

# ═════════════════════════════════════════════════════════════════════════════════════════════════════════════════════#
# Imports
# ═════════════════════════════════════════════════════════════════════════════════════════════════════════════════════#

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


# ═════════════════════════════════════════════════════════════════════════════════════════════════════════════════════#
# CONFIGURATION
# ═════════════════════════════════════════════════════════════════════════════════════════════════════════════════════#

VERSION = "0.2.0"

print("OSC FaceTrackingController")
print("Made By Boots")
print(f"Version {VERSION}")

DEFAULT_OSC_IP = "127.0.0.1"
DEFAULT_OSC_PORT = "9000"
DEFAULT_OSC_PREFIX = "/avatar/parameters/v2/"

PREFIX_PRESETS = {
    "VRCFT v2 (default)": "/avatar/parameters/v2/",
    "Direct / v1": "/avatar/parameters/",
}

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
        self._try_connect()

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
        preset_menu.pack(side="left", padx=(0, 8))

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

    def _build_footer_bar(self) -> None:
        footer_bar = tk.Frame(self, bg=PANEL, pady=8)
        footer_bar.pack(fill="x", side="bottom")
        footer_bar.columnconfigure(0, weight=1)

        help_btn = self._square_button(footer_bar, "？", self._open_help, base_size=28)
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
        """Open help window with multipage tutorial."""
        help_window = tk.Toplevel(self)
        help_window.title("OSC Face Tracking - Help")
        help_window.geometry("600x500")
        help_window.configure(bg=BG)

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
                    "Click CONNECT to establish connection.\n"
                    "Green dot = connected, Red dot = disconnected"
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

        header_frame = tk.Frame(help_window, bg=PANEL, height=50)
        header_frame.pack(fill="x")
        header_frame.pack_propagate(False)

        page_title = tk.Label(
            header_frame,
            text=pages[0]["title"],
            font=self.f_title,
            bg=PANEL,
            fg=TEXT,
        )
        page_title.pack(pady=12)

        content_frame = tk.Frame(help_window, bg=BG)
        content_frame.pack(fill="both", expand=True, padx=20, pady=20)

        content_text = tk.Label(
            content_frame,
            text=pages[0]["content"],
            font=self.f_label,
            bg=BG,
            fg=TEXT,
            justify="left",
            wraplength=550,
        )
        content_text.pack(fill="both", expand=True)

        nav_frame = tk.Frame(help_window, bg=PANEL, height=50)
        nav_frame.pack(fill="x", side="bottom")
        nav_frame.pack_propagate(False)

        def update_page(direction: int) -> None:
            current_page[0] = max(0, min(current_page[0] + direction, len(pages) - 1))
            page = pages[current_page[0]]
            page_title.config(text=page["title"])
            content_text.config(text=page["content"])
            prev_btn.config(state="normal" if current_page[0] > 0 else "disabled")
            next_btn.config(state="normal" if current_page[0] < len(pages) - 1 else "disabled")

        prev_btn = tk.Button(
            nav_frame,
            text="← Previous",
            font=self.f_btn,
            bg=PANEL,
            fg=SUBTEXT,
            relief="flat",
            activebackground=BORDER,
            activeforeground=TEXT,
            cursor="hand2",
            command=lambda: update_page(-1),
        )
        prev_btn.pack(side="left", padx=10, pady=10)

        page_info = tk.Label(
            nav_frame,
            text=f"Page {current_page[0] + 1} of {len(pages)}",
            font=self.f_small,
            bg=PANEL,
            fg=SUBTEXT,
        )
        page_info.pack(side="left", expand=True)

        next_btn = tk.Button(
            nav_frame,
            text="Next →",
            font=self.f_btn,
            bg=PANEL,
            fg=SUBTEXT,
            relief="flat",
            activebackground=BORDER,
            activeforeground=TEXT,
            cursor="hand2",
            command=lambda: update_page(1),
        )
        next_btn.pack(side="right", padx=10, pady=10)

        def update_page_info() -> None:
            page_info.config(text=f"Page {current_page[0] + 1} of {len(pages)}")

        prev_btn.config(state="disabled")
        next_btn.config(state="normal" if len(pages) > 1 else "disabled")

# ═════════════════════════════════════════════════════════════════════════════════════════════════════════════════════#
# OSC
# ═════════════════════════════════════════════════════════════════════════════════════════════════════════════════════#
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
            self._status_lbl.config(text="Status: Connected", fg=GREEN)
            self._footer_label.config(text=f"Connected {detail}" if detail else "Connected")
        else:
            self._status_lbl.config(text="Status: Error", fg=RED)
            self._footer_label.config(text="Error")

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
