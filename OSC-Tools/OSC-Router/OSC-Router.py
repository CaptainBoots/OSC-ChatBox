# ═════════════════════════════════════════════════════════════════════════════════════════════════════════════════════#
#                                              OSC Router Script                                                       #
# ═════════════════════════════════════════════════════════════════════════════════════════════════════════════════════#
# Hi :3
# Wellcome to my code

# ═════════════════════════════════════════════════════════════════════════════════════════════════════════════════════#
# Imports
# ═════════════════════════════════════════════════════════════════════════════════════════════════════════════════════#

import subprocess
import sys
import importlib
import site
import json
import threading
import time
import os
import tkinter as tk
from tkinter import messagebox



def install_if_missing(package, import_name=None):
    if import_name is None:
        import_name = package.split("==")[0].replace("-", "_")
    try:
        importlib.import_module(import_name)
    except ImportError:
        print(f"Installing {package}...")
        attempts = [[sys.executable, "-m", "pip", "install", package]]
        if sys.platform != "win32":
            attempts += [
                [sys.executable, "-m", "pip", "install", package, "--break-system-packages"],
                [sys.executable, "-m", "pip", "install", package, "--user"],
            ]
        last_err = None
        for cmd in attempts:
            try:
                subprocess.check_call(cmd)
                last_err = None
                break
            except subprocess.CalledProcessError as e:
                last_err = e
        if last_err:
            raise last_err
        if sys.platform != "win32":
            user_site = site.getusersitepackages()
            if user_site and user_site not in sys.path:
                sys.path.insert(0, user_site)


install_if_missing("python-osc==1.9.3", "pythonosc")

from pythonosc.osc_server import ThreadingOSCUDPServer
from pythonosc.dispatcher import Dispatcher
from pythonosc.udp_client import SimpleUDPClient


# ═════════════════════════════════════════════════════════════════════════════════════════════════════════════════════#
# CONFIGURATION & GLOBAL VARIABLES
# ═════════════════════════════════════════════════════════════════════════════════════════════════════════════════════#

VERSION = "1.0.0"

print("OSC Router")
print("Made By Boots")
print(f"Version {VERSION}")

SCRIPT_DIR  = os.path.dirname(os.path.abspath(__file__))
CONFIG_DIR  = SCRIPT_DIR
CONFIG_FILE = os.path.join(CONFIG_DIR, "router_config.json")


def get_default_config() -> dict:
    return {
        "output_ip":   "127.0.0.1",
        "output_port": 9000,
        "sources": [
            {"name": "Chatbox",       "port": 9002},
            {"name": "Face Tracking", "port": 9002},
        ],
    }


def load_config() -> dict:
    defaults = get_default_config()
    try:
        with open(CONFIG_FILE, "r", encoding="utf-8") as f:
            return {**defaults, **json.load(f)}
    except (FileNotFoundError, json.JSONDecodeError, OSError):
        return defaults


def save_config(ip: str, port: int, sources: list) -> None:
    os.makedirs(CONFIG_DIR, exist_ok=True)
    try:
        with open(CONFIG_FILE, "w", encoding="utf-8") as f:
            json.dump({"output_ip": ip, "output_port": port, "sources": sources}, f, indent=2)
        print(f"[Config] Saved ({len(sources)} sources)")
    except OSError as e:
        print(f"[Config] Save failed: {e}")


# ═════════════════════════════════════════════════════════════════════════════════════════════════════════════════════#
# OSC CORE
# ═════════════════════════════════════════════════════════════════════════════════════════════════════════════════════#

class OscSource:
    """
    Listens on a UDP port and caches the latest value for every OSC address it receives.
    Thread-safe. Snapshot returns a point-in-time copy of the cache.
    """

    def __init__(self, name: str, port: int) -> None:
        self.name     = name
        self.port     = port
        self.running  = False
        self.rx_count = 0
        self._lock    = threading.Lock()
        self._values: dict[str, tuple] = {}
        self._server: ThreadingOSCUDPServer | None = None
        self._thread: threading.Thread       | None = None

    def _handle(self, address: str, *args) -> None:
        with self._lock:
            self._values[address] = args
            self.rx_count += 1

    def snapshot(self) -> dict[str, tuple]:
        """Return a shallow copy of the current address→args cache."""
        with self._lock:
            return dict(self._values)

    def start(self) -> bool:
        if self.running:
            return True
        try:
            d = Dispatcher()
            d.set_default_handler(self._handle)
            self._server = ThreadingOSCUDPServer(("127.0.0.1", self.port), d)
            self._thread = threading.Thread(target=self._server.serve_forever, daemon=True)
            self._thread.start()
            self.running = True
            print(f"[Router] '{self.name}' listening on :{self.port}")
            return True
        except OSError as e:
            print(f"[Router] Cannot bind '{self.name}' on :{self.port} — {e}")
            return False

    def stop(self) -> None:
        if self._server:
            try:
                self._server.shutdown()
            except Exception:
                pass
        self._server = None
        self._thread = None
        self.running  = False
        with self._lock:
            self._values.clear()
        self.rx_count = 0


class OscRouter:
    """
    Merges OSC values from multiple sources (priority order) and forwards changes
    to a single output at a fixed interval.

    Merge rules  (index 0 = highest priority):
        • Different addresses          → all forwarded  (compatible, no conflict)
        • Same address, same value     → forwarded once (compatible)
        • Same address, different value→ highest-priority source wins (conflict resolved)

    Only changed values are forwarded to avoid chattering the output.
    """

    INTERVAL = 0.05  # seconds between merge cycles (20 Hz)

    def __init__(self) -> None:
        self.sources: list[OscSource] = []
        self._client:  SimpleUDPClient | None = None
        self._running  = False
        self._thread:  threading.Thread | None = None
        self._last:    dict[str, tuple] = {}
        self._fw_lock  = threading.Lock()
        self.fwd_total = 0

    @property
    def live_conflicts(self) -> int:
        """Number of OSC addresses currently in conflict across sources."""
        seen: dict[str, set] = {}
        for src in self.sources:
            for addr, args in src.snapshot().items():
                try:
                    seen.setdefault(addr, set()).add(args)
                except TypeError:
                    pass  # unhashable edge-case – skip
        return sum(1 for vals in seen.values() if len(vals) > 1)

    def start(self, out_ip: str, out_port: int) -> list[str]:
        """
        Start listening and routing.
        Returns a list of source names that failed to bind their ports.
        """
        if self._running:
            return []
        try:
            self._client = SimpleUDPClient(out_ip, out_port)
        except Exception as e:
            print(f"[Router] Output client error: {e}")
            return [s.name for s in self.sources]

        failed = [s.name for s in self.sources if not s.start()]
        self._running = True
        self._thread  = threading.Thread(target=self._loop, daemon=True)
        self._thread.start()
        print(f"[Router] Routing → {out_ip}:{out_port}  ({len(self.sources)} sources, {len(self.sources) - len(failed)} active)")
        return failed

    def stop(self) -> None:
        self._running = False
        for s in self.sources:
            s.stop()
        self._client = None
        with self._fw_lock:
            self._last.clear()
        print("[Router] Stopped.")

    def _merge(self) -> dict[str, tuple]:
        """
        Build the merged address→args dict.
        Iterate sources lowest-priority-first so higher priority overwrites.
        """
        merged: dict[str, tuple] = {}
        for source in reversed(self.sources):
            merged.update(source.snapshot())
        return merged

    def _loop(self) -> None:
        while self._running:
            merged = self._merge()
            for addr, args in merged.items():
                with self._fw_lock:
                    last = self._last.get(addr)
                if last != args:
                    try:
                        if self._client:
                            # Single-value tuples unwrap to a scalar for clean OSC encoding
                            payload = args[0] if len(args) == 1 else list(args)
                            self._client.send_message(addr, payload)
                            self.fwd_total += 1
                    except Exception as e:
                        print(f"[Router] Forward error on {addr}: {e}")
                    with self._fw_lock:
                        self._last[addr] = args
            time.sleep(self.INTERVAL)


# ═════════════════════════════════════════════════════════════════════════════════════════════════════════════════════#
# GLOBALS
# ═════════════════════════════════════════════════════════════════════════════════════════════════════════════════════#

router = OscRouter()
cfg    = load_config()


# ═════════════════════════════════════════════════════════════════════════════════════════════════════════════════════#
# THEME
# ═════════════════════════════════════════════════════════════════════════════════════════════════════════════════════#

BG      = "#0f0f13"
PANEL   = "#17171f"
BORDER  = "#2a2a38"
ACCENT  = "#7c5cfc"
ACCENT2 = "#a78bfa"
TEXT    = "#e2e0f0"
SUBTEXT = "#7e7b9a"
GREEN   = "#4ade80"
RED     = "#f87171"
ROW_BG  = "#1a1a26"   # slightly lighter than BG for source cards
ENTRY_BG = "#1c1c2a"
UI_FONT  = "Consolas"


# ═════════════════════════════════════════════════════════════════════════════════════════════════════════════════════#
# ROOT WINDOW
# ═════════════════════════════════════════════════════════════════════════════════════════════════════════════════════#

root = tk.Tk()
root.title("OSC Router")
root.configure(bg=BG)
root.resizable(True, True)

# Header with title and version
_title_bar = tk.Frame(root, bg=PANEL, pady=14)
_title_bar.pack(fill="x")

header_frame = tk.Frame(_title_bar, bg=PANEL)
header_frame.pack(fill="x", padx=16, expand=True)

tk.Label(
    header_frame, text="◈  OSC ROUTER",
    bg=PANEL, fg=ACCENT2, font=(UI_FONT, 16, "bold"),
).pack(side="left", anchor="w")

version_label = tk.Label(
    header_frame,
    text=f"v{VERSION}",
    bg=PANEL,
    fg=SUBTEXT,
    font=(UI_FONT, 9)
)
version_label.pack(side="right", anchor="e", padx=(32, 16))

status_label = tk.Label(
    header_frame, text="Status: Stopped",
    bg=PANEL, fg=RED, font=(UI_FONT, 10),
)
status_label.pack(side="right", anchor="e")

tk.Frame(root, bg=BORDER, height=1).pack(fill="x")

# ── scrollable body with better padding ──────────────────────────────────────
_outer = tk.Frame(root, bg=BG)
_outer.pack(fill="both", expand=True, padx=16, pady=14)

_canvas = tk.Canvas(_outer, bg=BG, highlightthickness=0, bd=0)
_canvas.pack(side="left", fill="both", expand=True)

body    = tk.Frame(_canvas, bg=BG)
_bid    = _canvas.create_window((0, 0), window=body, anchor="nw")

body.bind("<Configure>",    lambda e: _canvas.configure(scrollregion=_canvas.bbox("all")))
_canvas.bind("<Configure>", lambda e: _canvas.itemconfig(_bid, width=e.width))

# Cross-platform scroll
root.bind_all("<MouseWheel>", lambda e: _canvas.yview_scroll(int(-1 * (e.delta / 120)), "units"))
root.bind_all("<Button-4>",   lambda e: _canvas.yview_scroll(-1, "units"))
root.bind_all("<Button-5>",   lambda e: _canvas.yview_scroll(1,  "units"))

# ── body helpers ─────────────────────────────────────────────────────────────
def _section(text: str) -> None:
    tk.Label(body, text=text, bg=BG, fg=ACCENT2,
             font=(UI_FONT, 11, "bold"), anchor="w",
             ).pack(fill="x", padx=12, pady=(10, 3))


def _field(label: str, default: str) -> tk.Entry:
    r = tk.Frame(body, bg=BG)
    r.pack(fill="x", padx=12, pady=2)
    tk.Label(r, text=label, bg=BG, fg=SUBTEXT,
             font=(UI_FONT, 9), width=18, anchor="w").pack(side="left")
    e = tk.Entry(
        r, bg=PANEL, fg=TEXT, insertbackground=ACCENT, relief="flat",
        font=(UI_FONT, 9), highlightthickness=1,
        highlightbackground=BORDER, highlightcolor=ACCENT,
    )
    e.insert(0, default)
    e.pack(side="left", fill="x", expand=True, ipady=3)
    return e


def _divider() -> None:
    tk.Frame(body, bg=BORDER, height=1).pack(fill="x", padx=12, pady=6)


def square_button(parent, text, command, base_size=28):
    container = tk.Frame(parent, bg=PANEL, highlightthickness=1, highlightbackground=BORDER)
    container.pack_propagate(False)

    button_widget = tk.Button(
        container,
        text=text,
        command=command,
        bg=PANEL,
        fg=SUBTEXT,
        relief="flat",
        borderwidth=0,
        font=(UI_FONT, 12),
        activebackground=BORDER,
        activeforeground=TEXT,
        cursor="hand2",
    )
    button_widget.pack(fill="both", expand=True)
    container.config(width=base_size, height=base_size)
    return container


# ═════════════════════════════════════════════════════════════════════════════════════════════════════════════════════#
# OUTPUT SECTION
# ═════════════════════════════════════════════════════════════════════════════════════════════════════════════════════#

_section("Output")
out_ip_entry   = _field("Output IP",   cfg.get("output_ip",   "127.0.0.1"))
out_port_entry = _field("Output Port", str(cfg.get("output_port", 9000)))
_divider()


# ═════════════════════════════════════════════════════════════════════════════════════════════════════════════════════#
# DRAGGABLE SOURCE LIST
#
# Each row dict:
#   frame          – tk.Frame  (the row container)
#   name_var       – tk.StringVar
#   port_var       – tk.StringVar
#   priority_label – tk.Label  (#1, #2, …)
#   stats_label    – tk.Label  (rx count / status)
# ═════════════════════════════════════════════════════════════════════════════════════════════════════════════════════#

_section("Sources  ·  drag ≡ to reorder  ·  top row = highest priority")

_src_frame = tk.Frame(body, bg=BG)
_src_frame.pack(fill="x", padx=12, pady=(0, 4))

_rows: list[dict] = []

# ── drag state ───────────────────────────────────────────────────────────────
_drag: dict = {
    "active":    False,
    "src_idx":   -1,
    "indicator": None,   # thin accent strip drawn between rows
}


def _indicator() -> tk.Frame:
    if _drag["indicator"] is None or not _drag["indicator"].winfo_exists():
        _drag["indicator"] = tk.Frame(_src_frame, bg=ACCENT, height=3)
    return _drag["indicator"]


def _idx_for_y(y_root: int) -> int:
    """Return the insert-before index that corresponds to cursor position y_root."""
    local_y = y_root - _src_frame.winfo_rooty()
    for i, row in enumerate(_rows):
        fw = row["frame"]
        if local_y < fw.winfo_y() + fw.winfo_height() // 2:
            return i
    return len(_rows)


def _repack() -> None:
    """Re-pack all rows in _rows order and refresh priority badges."""
    for row in _rows:
        row["frame"].pack_forget()
    for i, row in enumerate(_rows):
        row["frame"].pack(fill="x", pady=3)
        row["priority_label"].config(text=f"#{i + 1}")


def _drag_start(event: tk.Event, row_data: dict) -> None:
    idx = next((i for i, r in enumerate(_rows) if r is row_data), -1)
    if idx == -1:
        return
    _drag["active"]  = True
    _drag["src_idx"] = idx
    row_data["frame"].config(highlightbackground=ACCENT)


def _drag_motion(event: tk.Event) -> None:
    if not _drag["active"]:
        return
    tgt = _idx_for_y(event.y_root)
    ind = _indicator()

    if tgt < len(_rows):
        ref_y = _rows[tgt]["frame"].winfo_y()
    elif _rows:
        last = _rows[-1]["frame"]
        ref_y = last.winfo_y() + last.winfo_height() + 3
    else:
        ref_y = 4

    ind.place(x=0, y=ref_y, relwidth=1.0, height=3)
    ind.lift()


def _drag_release(event: tk.Event) -> None:
    if not _drag["active"]:
        return

    src_idx = _drag["src_idx"]
    tgt_idx = _idx_for_y(event.y_root)
    _drag["active"] = False

    # Restore border
    if 0 <= src_idx < len(_rows):
        _rows[src_idx]["frame"].config(highlightbackground=BORDER)

    # Hide indicator
    ind = _drag.get("indicator")
    if ind and ind.winfo_exists():
        ind.place_forget()

    # No meaningful change
    if tgt_idx in (src_idx, src_idx + 1):
        return

    row = _rows.pop(src_idx)
    if tgt_idx > src_idx:
        tgt_idx -= 1
    _rows.insert(tgt_idx, row)
    _repack()
    print(f"[Router] Priority updated: {[r['name_var'].get() for r in _rows]}")


def _remove_row(row_data: dict) -> None:
    if row_data in _rows:
        _rows.remove(row_data)
    row_data["frame"].destroy()
    _repack()


def add_row(name: str = "New Source", port: int = 9010) -> dict:
    """Create a source row, append it to _rows, return its data dict."""

    frm = tk.Frame(
        _src_frame, bg=ROW_BG,
        highlightthickness=1, highlightbackground=BORDER,
    )
    frm.pack(fill="x", pady=3)

    # Drag handle
    handle = tk.Label(frm, text="≡", bg=ROW_BG, fg=SUBTEXT,
                      font=(UI_FONT, 14), cursor="fleur", padx=6)
    handle.pack(side="left")

    # Priority badge
    pri_lbl = tk.Label(frm, text="#?", bg=ROW_BG, fg=ACCENT,
                       font=(UI_FONT, 8, "bold"), width=3)
    pri_lbl.pack(side="left")

    # Name entry
    name_var = tk.StringVar(value=name)
    tk.Entry(
        frm, textvariable=name_var, bg=ENTRY_BG, fg=TEXT,
        insertbackground=ACCENT, relief="flat", font=(UI_FONT, 9),
        highlightthickness=1, highlightbackground=BORDER, highlightcolor=ACCENT,
        width=14,
    ).pack(side="left", padx=(6, 8), ipady=4)

    # Port
    tk.Label(frm, text="Port", bg=ROW_BG, fg=SUBTEXT,
             font=(UI_FONT, 8)).pack(side="left")
    port_var = tk.StringVar(value=str(port))
    tk.Entry(
        frm, textvariable=port_var, bg=ENTRY_BG, fg=TEXT,
        insertbackground=ACCENT, relief="flat", font=(UI_FONT, 9),
        highlightthickness=1, highlightbackground=BORDER, highlightcolor=ACCENT,
        width=7,
    ).pack(side="left", padx=(4, 10), ipady=4)

    # Status / rx counter
    stats_lbl = tk.Label(frm, text="● offline", bg=ROW_BG, fg=SUBTEXT,
                          font=(UI_FONT, 7), width=13, anchor="w")
    stats_lbl.pack(side="left")

    # Remove button — forward-ref trick to capture row_data
    _self: list[dict] = []

    def _rm():
        if _self:
            _remove_row(_self[0])

    tk.Button(
        frm, text="×", command=_rm,
        bg=ROW_BG, fg=RED, relief="flat",
        activebackground=BORDER, activeforeground=RED,
        cursor="hand2", font=(UI_FONT, 11, "bold"), padx=8, pady=2,
    ).pack(side="right")

    row_data: dict = {
        "frame":          frm,
        "name_var":       name_var,
        "port_var":       port_var,
        "priority_label": pri_lbl,
        "stats_label":    stats_lbl,
    }
    _self.append(row_data)

    # Bind drag to the handle only
    handle.bind("<ButtonPress-1>",   lambda e: _drag_start(e, row_data))
    handle.bind("<B1-Motion>",       _drag_motion)
    handle.bind("<ButtonRelease-1>", _drag_release)

    _rows.append(row_data)
    _repack()
    return row_data


# ── Add Source button ─────────────────────────────────────────────────────────
tk.Button(
    body, text="＋  Add Source",
    command=lambda: add_row(),
    bg=PANEL, fg=ACCENT2, relief="flat",
    activebackground=BORDER, activeforeground=TEXT,
    cursor="hand2", font=(UI_FONT, 9, "bold"),
    padx=12, pady=6,
    highlightthickness=1, highlightbackground=BORDER,
).pack(fill="x", padx=12, pady=(4, 2))

_divider()


# ═════════════════════════════════════════════════════════════════════════════════════════════════════════════════════#
# STATS BAR
# ═════════════════════════════════════════════════════════════════════════════════════════════════════════════════════#

_stats = tk.Frame(body, bg=PANEL, highlightthickness=1, highlightbackground=BORDER)
_stats.pack(fill="x", padx=12, pady=(0, 6))

lbl_fwd      = tk.Label(_stats, text="Forwarded: —",    bg=PANEL, fg=SUBTEXT, font=(UI_FONT, 8))
lbl_conflict = tk.Label(_stats, text="Conflicts: —",     bg=PANEL, fg=SUBTEXT, font=(UI_FONT, 8))
lbl_sources  = tk.Label(_stats, text="Sources: 0 / 0",  bg=PANEL, fg=SUBTEXT, font=(UI_FONT, 8))

lbl_fwd.pack(side="left",  padx=12, pady=6)
lbl_conflict.pack(side="left",  padx=12, pady=6)
lbl_sources.pack(side="right", padx=12, pady=6)


# ═════════════════════════════════════════════════════════════════════════════════════════════════════════════════════#
# SCRIPT CONTROL
# ═════════════════════════════════════════════════════════════════════════════════════════════════════════════════════#

def _collect_sources() -> list[dict] | None:
    """Validate and return source dicts from UI, or None on validation error."""
    result = []
    for row in _rows:
        name = row["name_var"].get().strip() or "Source"
        try:
            port = int(row["port_var"].get().strip())
        except ValueError:
            messagebox.showerror("Error", f"Invalid port for source '{name}'.")
            return None
        result.append({"name": name, "port": port})
    return result


def start_routing() -> None:
    out_ip = out_ip_entry.get().strip()
    try:
        out_port = int(out_port_entry.get().strip())
    except ValueError:
        messagebox.showerror("Error", "Invalid output port.")
        footer_label.config(text="Error")
        return

    sources_data = _collect_sources()
    if sources_data is None:
        footer_label.config(text="Error")
        return

    save_config(out_ip, out_port, sources_data)

    # Re-build router sources in priority order (matches _rows order)
    router.sources = [OscSource(s["name"], s["port"]) for s in sources_data]

    failed = router.start(out_ip, out_port)
    if failed:
        messagebox.showwarning(
            "Some sources failed to start",
            "The following sources could not bind their ports:\n"
            + "\n".join(f"  • {n}" for n in failed)
            + "\n\nCheck that these ports are not already in use."
        )

    status_label.config(
        text="Status: Running" if router._running else "Status: Error",
        fg=GREEN if router._running else RED,
    )
    footer_label.config(text="Running" if router._running else "Error")


def stop_routing() -> None:
    router.stop()
    for row in _rows:
        row["stats_label"].config(text="● offline", fg=SUBTEXT)
    status_label.config(text="Status: Stopped", fg=RED)
    footer_label.config(text="Ready")


def restart_routing() -> None:
    stop_routing()
    root.after(600, start_routing)


def open_help():
    help_win = tk.Toplevel(root)
    help_win.title("OSC Router Tutorial")
    help_win.configure(bg=BG)
    help_win.resizable(True, True)

    root.update_idletasks()
    help_w = root.winfo_width()
    help_h = root.winfo_height()
    root_x = root.winfo_x()
    root_y = root.winfo_y()
    help_win.geometry(f"{help_w}x{help_h}+{root_x}+{root_y}")

    pages = [
        {
            "title": "Welcome",
            "content": (
                "OSC Router — Routes OSC messages between\n"
                "multiple sources and a single output.\n\n"
                "SOURCES — Input devices/applications\n"
                "that send OSC messages.\n\n"
                "OUTPUT — Where all messages are\n"
                "forwarded to (usually VRChat).\n\n"
                "Configure sources and output below,\n"
                "then click Start to begin routing."
            )
        },
        {
            "title": "Sources",
            "content": (
                "Each source needs:\n"
                "• Name — Friendly identifier\n"
                "• Port — UDP port it listens on\n\n"
                "Default sources:\n"
                "• Chatbox: port 9011\n"
                "• Face Tracking: port 9012\n\n"
                "You can add/remove sources as needed.\n"
                "Multiple sources can send to the same\n"
                "output without conflicts."
            )
        },
        {
            "title": "Output",
            "content": (
                "OUTPUT IP — Destination IP address\n"
                "Usually 127.0.0.1 (your local PC)\n\n"
                "OUTPUT PORT — Destination port\n"
                "Usually 9000 (VRChat's OSC port)\n\n"
                "Tip: If running VRChat on this machine,\n"
                "use 127.0.0.1:9000\n\n"
                "If routing to another PC on your\n"
                "network, use that PC's IP address."
            )
        },
        {
            "title": "Status",
            "content": (
                "FORWARDED — Total OSC messages routed\n\n"
                "CONFLICTS — Messages with same address\n"
                "from different sources (last one wins)\n\n"
                "SOURCES — How many are connected\n"
                "Active sources show: ● {count} rx\n\n"
                "Failed sources show: ✗ failed\n"
                "(Usually means the port is already in use)"
            )
        },
    ]

    current_page = [0]

    header = tk.Frame(help_win, bg=PANEL, pady=10)
    header.pack(fill="x")

    title_label = tk.Label(
        header, text="", bg=PANEL, fg=ACCENT2, font=(UI_FONT, 12, "bold")
    )
    title_label.pack(side="left", padx=16)

    page_indicator = tk.Label(
        header, text="", bg=PANEL, fg=SUBTEXT, font=(UI_FONT, 8)
    )
    page_indicator.pack(side="right", padx=16)

    tk.Frame(help_win, bg=BORDER, height=1).pack(fill="x")

    content_panel = tk.Frame(help_win, bg=PANEL, highlightthickness=1, highlightbackground=BORDER)
    content_panel.pack(padx=20, pady=(14, 0), fill="both", expand=True)

    content_text = tk.Label(
        content_panel, text="", bg=PANEL, fg=TEXT, anchor="nw", justify="left",
        font=(UI_FONT, 9), wraplength=400
    )
    content_text.pack(padx=16, pady=16, fill="both", expand=True)

    def show_page(page_num):
        page = pages[page_num]
        title_label.config(text=page["title"])
        content_text.config(text=page["content"])
        page_indicator.config(text=f"Page {page_num + 1} / {len(pages)}")
        is_last = page_num == len(pages) - 1
        next_btn.config(text="Finish" if is_last else "Next →")

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
        command=lambda: (current_page.__setitem__(0, current_page[0] - 1),
                         show_page(current_page[0]))
    )
    prev_btn.grid(row=0, column=0, sticky="w")
    prev_btn.configure(
        fg=SUBTEXT, activebackground=BORDER, activeforeground=TEXT,
        cursor="hand2", font=(UI_FONT, 9, "bold"),
    )

    next_btn = tk.Button(
        nav_frame, text="Next →", bg=PANEL, fg=SUBTEXT, relief="flat", width=10,
        command=next_or_finish
    )
    next_btn.grid(row=0, column=2, sticky="e")
    next_btn.configure(
        bg=ACCENT, fg="#FFFFFF", activebackground=ACCENT2, activeforeground="#FFFFFF",
        cursor="hand2", font=(UI_FONT, 9, "bold"),
    )

    show_page(0)


# ── Buttons ───────────────────────────────────────────────────────────────────
_btns = tk.Frame(body, bg=BG)
_btns.pack(fill="x", padx=12, pady=(4, 14))
_btns.columnconfigure(0, weight=1)
_btns.columnconfigure(1, weight=1)
_btns.columnconfigure(2, weight=1)

tk.Button(
    _btns, text="Start", command=start_routing,
    bg=ACCENT, fg="#FFFFFF", relief="flat",
    activebackground=ACCENT2, activeforeground="#FFFFFF",
    cursor="hand2", font=(UI_FONT, 9, "bold"),
).grid(row=0, column=0, sticky="ew", padx=2, ipady=5)

tk.Button(
    _btns, text="Stop", command=stop_routing,
    bg=PANEL, fg=SUBTEXT, relief="flat",
    activebackground=BORDER, activeforeground=TEXT,
    cursor="hand2", font=(UI_FONT, 9, "bold"),
).grid(row=0, column=1, sticky="ew", padx=2, ipady=5)

tk.Button(
    _btns, text="Restart", command=restart_routing,
    bg=PANEL, fg=SUBTEXT, relief="flat",
    activebackground=BORDER, activeforeground=TEXT,
    cursor="hand2", font=(UI_FONT, 9, "bold"),
).grid(row=0, column=2, sticky="ew", padx=2, ipady=5)


# ═════════════════════════════════════════════════════════════════════════════════════════════════════════════════════#
# LIVE STATS TICK
# ═════════════════════════════════════════════════════════════════════════════════════════════════════════════════════#

def _tick() -> None:
    if router._running:
        active = sum(1 for s in router.sources if s.running)
        total  = len(router.sources)

        lbl_fwd.config(text=f"Forwarded: {router.fwd_total:,}")
        lbl_conflict.config(text=f"Conflicts: {router.live_conflicts} live")
        lbl_sources.config(text=f"Sources: {active} / {total} active")

        # Per-row rx counts (matched by index — same order as router.sources)
        for i, row in enumerate(_rows):
            if i < len(router.sources):
                src = router.sources[i]
                if src.running:
                    row["stats_label"].config(text=f"● {src.rx_count:,} rx", fg=GREEN)
                else:
                    row["stats_label"].config(text="✗ failed", fg=RED)
    else:
        lbl_fwd.config(text="Forwarded: —")
        lbl_conflict.config(text="Conflicts: —")
        lbl_sources.config(text=f"Sources: 0 / {len(_rows)}")

    root.after(1000, _tick)


# ═════════════════════════════════════════════════════════════════════════════════════════════════════════════════════#
# INIT
# ═════════════════════════════════════════════════════════════════════════════════════════════════════════════════════#

def _on_close() -> None:
    router.stop()
    root.destroy()


root.protocol("WM_DELETE_WINDOW", _on_close)

for src in cfg.get("sources", []):
    add_row(src.get("name", "Source"), src.get("port", 9001))

# ── Footer with status ────────────────────────────────────────────────────────
footer_bar = tk.Frame(root, bg=PANEL, pady=8)
footer_bar.pack(fill="x", side="bottom")

footer_bar.columnconfigure(0, weight=1)

help_btn = square_button(footer_bar, "？", open_help, base_size=28)
help_btn.pack(side="left", padx=(8, 0))

footer_label = tk.Label(
    footer_bar,
    text="Ready",
    bg=PANEL,
    fg=SUBTEXT,
    font=(UI_FONT, 8)
)
footer_label.pack(side="left", padx=16)

# Set window size
root.geometry("580x700")
root.minsize(0, 0)

_tick()
root.mainloop()
