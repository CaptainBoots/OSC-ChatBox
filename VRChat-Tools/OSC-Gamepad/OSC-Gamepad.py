# ══════════════════════════════════════════════════════════════════════════════════════════#
#                                    OSC Gamepad                                           #
# ══════════════════════════════════════════════════════════════════════════════════════════#


import tkinter as tk
from tkinter import ttk, messagebox
import math
import threading
import time
import json
import os
import sys
import importlib
import subprocess
import site

def install_if_missing(package, import_name=None):
    if import_name is None:
        import_name = package.split("==")[0].replace("-", "_")
    try:
        importlib.import_module(import_name)
    except ImportError:
        print(f"Installing {package}...")
        attempts = [[sys.executable, "-m", "pip", "install", package]]
        if sys.platform != "win32":
            attempts.append([sys.executable, "-m", "pip", "install", package, "--break-system-packages"])
            attempts.append([sys.executable, "-m", "pip", "install", package, "--user"])
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

from pythonosc import udp_client

# ══════════════════════════════════════════════════════════════════════════════════════════#
# CONSTANTS
# ══════════════════════════════════════════════════════════════════════════════════════════#

VERSION = "0.0.3"

SCRIPT_DIR  = os.path.dirname(os.path.abspath(__file__))
CONFIG_FILE = os.path.join(SCRIPT_DIR, "gamepad_config.json")

# ── Colours (match OSC-Chatbox exactly) ───────────────────────────────────
BG       = "#0f0f13"
PANEL    = "#17171f"
BORDER   = "#2a2a38"
ACCENT   = "#7c5cfc"
ACCENT2  = "#a78bfa"
TEXT     = "#e2e0f0"
SUBTEXT  = "#7e7b9a"
GREEN    = "#4ade80"
RED      = "#f87171"
ENTRY_BG = PANEL
BTN_BG   = PANEL
BTN_FG   = TEXT
UI_FONT  = "Consolas"

# Pad-specific button colours
BTN_IDLE  = "#1e1e2e"
BTN_HOV   = "#2a2a3e"
BTN_ACT   = ACCENT

# ══════════════════════════════════════════════════════════════════════════════════════════#
# CONFIG
# ══════════════════════════════════════════════════════════════════════════════════════════#

DEFAULT_CONFIG = {
    "pads": []
}

def load_config():
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, "r") as f:
                data = json.load(f)
                # Merge with defaults so new keys always exist
                for k, v in DEFAULT_CONFIG.items():
                    data.setdefault(k, v)
                return data
        except Exception:
            pass
    return dict(DEFAULT_CONFIG)

def save_config(pads_data):
    cfg = {"pads": pads_data}
    try:
        with open(CONFIG_FILE, "w") as f:
            json.dump(cfg, f, indent=2)
    except Exception as e:
        print(f"Config save error: {e}")

# ══════════════════════════════════════════════════════════════════════════════════════════#
# PAD STATE (OSC logic)
# ══════════════════════════════════════════════════════════════════════════════════════════#

class PadState:
    def __init__(self, host, port):
        self.host = host
        self.port = port
        self.client = udp_client.SimpleUDPClient(host, port)
        self.axes_held  = set()
        self.btn_held   = set()
        self.btn_sent1  = set()
        self.seated     = False
        self.crouched   = False
        self.running    = True
        self.thread = threading.Thread(target=self._loop, daemon=True)
        self.thread.start()

    def press_axis(self, action):  self.axes_held.add(action)
    def release_axis(self, action): self.axes_held.discard(action)
    def press_btn(self, action):   self.btn_held.add(action)

    def release_btn(self, action):
        self.btn_held.discard(action)
        self.btn_sent1.discard(action)

    def toggle_avatar_param(self, param):
        if param == "seated":
            self.seated = not self.seated
            try: self.client.send_message("/avatar/parameters/Seated", self.seated)
            except Exception: pass
            return self.seated
        elif param == "crouched":
            self.crouched = not self.crouched
            try: self.client.send_message("/avatar/parameters/Crouching", self.crouched)
            except Exception: pass
            return self.crouched

    def _loop(self):
        while self.running:
            v = h = lh = lv = 0.0
            jump = grab = use = menu = voice = 0
            for a in list(self.axes_held):
                if a == "up":     v  =  1.0
                if a == "down":   v  = -1.0
                if a == "left":   h  = -1.0
                if a == "right":  h  =  1.0
                if a == "look_l": lh = -1.0
                if a == "look_r": lh =  1.0
                if a == "look_u": lv =  1.0
                if a == "look_d": lv = -1.0
            for btn in ["jump", "grab", "use", "menu", "voice"]:
                if btn in self.btn_held:
                    if btn not in self.btn_sent1:
                        self.btn_sent1.add(btn)
                        if btn == "jump":  jump  = 1
                        if btn == "grab":  grab  = 1
                        if btn == "use":   use   = 1
                        if btn == "menu":  menu  = 1
                        if btn == "voice": voice = 1
                    else:
                        if btn == "grab": grab = 1
                        if btn == "use":  use  = 1
                        if btn == "menu": menu = 1
            try:
                self.client.send_message("/input/Vertical",       v)
                self.client.send_message("/input/Horizontal",     h)
                self.client.send_message("/input/LookHorizontal", lh)
                self.client.send_message("/input/LookVertical",   lv)
                self.client.send_message("/input/Jump",           jump)
                self.client.send_message("/input/Grab",           grab)
                self.client.send_message("/input/Use",            use)
                self.client.send_message("/input/QuickMenuToggleLeft", menu)
                self.client.send_message("/input/Voice",          voice)
            except Exception: pass
            time.sleep(0.05)

    def stop(self):
        self.running = False

# ══════════════════════════════════════════════════════════════════════════════════════════#
# SHARED BUTTON FACTORIES
# ══════════════════════════════════════════════════════════════════════════════════════════#

def make_axis_btn(parent, label, action, state, font_size=13):
    b = tk.Label(parent, text=label, font=(UI_FONT, font_size, "bold"),
                 width=2, height=1, fg=TEXT, bg=BTN_IDLE,
                 relief="flat", cursor="hand2",
                 highlightthickness=1, highlightbackground=BORDER)
    b.bind("<ButtonPress-1>",   lambda e, a=action: state.press_axis(a))
    b.bind("<ButtonRelease-1>", lambda e, a=action: state.release_axis(a))
    b.bind("<Enter>",  lambda e, w=b: w.config(bg=BTN_HOV))
    b.bind("<Leave>",  lambda e, w=b, a=action: (state.release_axis(a), w.config(bg=BTN_IDLE)))
    return b

def make_action_btn(parent, label, action, colour, state, width=5, height=2):
    b = tk.Label(parent, text=label, font=(UI_FONT, 8, "bold"),
                 width=width, height=height, fg=colour, bg=BTN_IDLE,
                 relief="flat", cursor="hand2",
                 highlightthickness=1, highlightbackground=colour)
    b.bind("<ButtonPress-1>",   lambda e, a=action, w=b: (state.press_btn(a), w.config(bg=BTN_ACT)))
    b.bind("<ButtonRelease-1>", lambda e, a=action, w=b: (state.release_btn(a), w.config(bg=BTN_IDLE)))
    b.bind("<Enter>",  lambda e, w=b: w.config(bg=BTN_HOV))
    b.bind("<Leave>",  lambda e, w=b, a=action: (state.release_btn(a), w.config(bg=BTN_IDLE)))
    return b

def make_toggle_btn(parent, label, param, colour, state, width=5, height=2):
    active = [False]
    b = tk.Label(parent, text=label, font=(UI_FONT, 8, "bold"),
                 width=width, height=height, fg=colour, bg=BTN_IDLE,
                 relief="flat", cursor="hand2",
                 highlightthickness=1, highlightbackground=colour)
    def click(e):
        result = state.toggle_avatar_param(param)
        active[0] = result
        b.config(bg=BTN_ACT if result else BTN_IDLE)
    b.bind("<Button-1>", click)
    b.bind("<Enter>", lambda e, w=b: w.config(bg=BTN_HOV if not active[0] else BTN_ACT))
    b.bind("<Leave>", lambda e, w=b: w.config(bg=BTN_ACT if active[0] else BTN_IDLE))
    return b

def square_button(parent, text, command, base_size=28):
    container = tk.Frame(parent, bg=BTN_BG, highlightthickness=1,
                         highlightbackground=BORDER)
    container.pack_propagate(False)
    container.config(width=base_size, height=base_size)
    btn = tk.Button(container, text=text, command=command,
                    bg=BTN_BG, fg=SUBTEXT, relief="flat", borderwidth=0,
                    font=(UI_FONT, 12), activebackground=BORDER,
                    activeforeground=TEXT, cursor="hand2")
    btn.pack(fill="both", expand=True)
    return container

# ══════════════════════════════════════════════════════════════════════════════════════════#
# NES PAD WIDGET
# ══════════════════════════════════════════════════════════════════════════════════════════#

class NESPad(tk.Frame):
    def __init__(self, parent, state, **kwargs):
        super().__init__(parent, bg=PANEL, **kwargs)
        self.state = state
        self._build()

    def _build(self):
        dpad = tk.Frame(self, bg=PANEL)
        dpad.grid(row=0, column=0, padx=(8, 12), pady=8)

        make_axis_btn(dpad, "▲", "up",    self.state).grid(row=0, column=1, padx=2, pady=2)
        make_axis_btn(dpad, "◀", "left",  self.state).grid(row=1, column=0, padx=2, pady=2)
        make_axis_btn(dpad, "▶", "right", self.state).grid(row=1, column=2, padx=2, pady=2)
        make_axis_btn(dpad, "▼", "down",  self.state).grid(row=2, column=1, padx=2, pady=2)
        tk.Frame(dpad, width=44, height=44, bg=BG).grid(row=1, column=1, padx=2, pady=2)

        look = tk.Frame(self, bg=PANEL)
        look.grid(row=1, column=0, padx=(8, 12), pady=(0, 8))
        tk.Label(look, text="LOOK", font=(UI_FONT, 7), fg=SUBTEXT, bg=PANEL).pack()
        lh_row = tk.Frame(look, bg=PANEL); lh_row.pack()
        make_axis_btn(lh_row, "◀", "look_l", self.state, font_size=10).pack(side=tk.LEFT, padx=2)
        make_axis_btn(lh_row, "▶", "look_r", self.state, font_size=10).pack(side=tk.LEFT, padx=2)
        lv_row = tk.Frame(look, bg=PANEL); lv_row.pack(pady=(2,0))
        make_axis_btn(lv_row, "▲", "look_u", self.state, font_size=10).pack(side=tk.LEFT, padx=2)
        make_axis_btn(lv_row, "▼", "look_d", self.state, font_size=10).pack(side=tk.LEFT, padx=2)

        act = tk.Frame(self, bg=PANEL)
        act.grid(row=0, column=1, rowspan=2, padx=(0, 8), pady=8, sticky="n")
        btns = [
            ("JUMP",   "jump",     "#4ade80", False),
            ("GRAB",   "grab",     "#fb923c", False),
            ("USE",    "use",      "#38bdf8", False),
            ("MENU",   "menu",     "#e879f9", False),
            ("MUTE",   "voice",    "#facc15", False),
            ("SIT",    "seated",   "#f87171", True),
            ("CROUCH", "crouched", "#67e8f9", True),
        ]
        for i, (label, action, colour, is_toggle) in enumerate(btns):
            b = make_toggle_btn(act, label, action, colour, self.state) if is_toggle \
                else make_action_btn(act, label, action, colour, self.state)
            b.grid(row=i // 2, column=i % 2, padx=4, pady=4)

# ══════════════════════════════════════════════════════════════════════════════════════════#
# JOYSTICK PAD WIDGET
# ══════════════════════════════════════════════════════════════════════════════════════════#

class JoystickPad(tk.Frame):
    def __init__(self, parent, state, **kwargs):
        super().__init__(parent, bg=PANEL, **kwargs)
        self.state = state
        self._build()

    def _build(self):
        PS = 170; CEN = PS // 2; KNOB = 16

        c = tk.Canvas(self, width=PS, height=PS, bg=BG,
                      highlightthickness=1, highlightbackground=BORDER)
        c.grid(row=0, column=0, padx=8, pady=8)
        c.create_oval(4, 4, PS-4, PS-4, outline=BORDER, width=1)
        c.create_oval(CEN-5, CEN-5, CEN+5, CEN+5, outline=BORDER, fill=PANEL)
        knob = c.create_oval(CEN-KNOB, CEN-KNOB, CEN+KNOB, CEN+KNOB,
                              fill=ACCENT, outline=ACCENT2, width=2)

        def drag(ev):
            dx = ev.x - CEN; dy = ev.y - CEN
            dist = math.hypot(dx, dy)
            max_r = CEN - KNOB - 4
            if dist > max_r: dx = dx/dist*max_r; dy = dy/dist*max_r
            c.coords(knob, CEN+dx-KNOB, CEN+dy-KNOB, CEN+dx+KNOB, CEN+dy+KNOB)
            try:
                self.state.client.send_message("/input/Horizontal", round(max(-1.0,min(1.0,dx/max_r)),3))
                self.state.client.send_message("/input/Vertical",   round(max(-1.0,min(1.0,-dy/max_r)),3))
            except Exception: pass

        def rel(ev):
            c.coords(knob, CEN-KNOB, CEN-KNOB, CEN+KNOB, CEN+KNOB)
            try:
                self.state.client.send_message("/input/Horizontal", 0.0)
                self.state.client.send_message("/input/Vertical",   0.0)
            except Exception: pass

        c.bind("<B1-Motion>", drag)
        c.bind("<ButtonRelease-1>", rel)

        SW, SH, KW = PS, 36, 26

        def make_slider(row, osc_addr, lbl):
            sc = tk.Canvas(self, width=SW, height=SH, bg=BG,
                           highlightthickness=1, highlightbackground=BORDER)
            sc.grid(row=row, column=0, padx=8, pady=(0, 2))
            ty = SH // 2
            sc.create_line(KW//2, ty, SW-KW//2, ty, fill=BORDER, width=2)
            lk = sc.create_oval(SW//2-KW//2, ty-KW//2, SW//2+KW//2, ty+KW//2,
                                 fill=ACCENT, outline=ACCENT2, width=2)
            tk.Label(self, text=lbl, font=(UI_FONT, 7), fg=SUBTEXT,
                     bg=PANEL).grid(row=row+1, column=0)

            def sdrag(ev):
                usable = SW - KW
                x = max(KW//2, min(SW-KW//2, ev.x))
                sc.coords(lk, x-KW//2, ty-KW//2, x+KW//2, ty+KW//2)
                try: self.state.client.send_message(osc_addr, round((x-SW/2)/(usable/2),3))
                except Exception: pass

            def srel(ev):
                sc.coords(lk, SW//2-KW//2, ty-KW//2, SW//2+KW//2, ty+KW//2)
                try: self.state.client.send_message(osc_addr, 0.0)
                except Exception: pass

            sc.bind("<B1-Motion>", sdrag)
            sc.bind("<ButtonRelease-1>", srel)

        make_slider(1, "/input/LookHorizontal", "LOOK H")
        make_slider(3, "/input/LookVertical",   "LOOK V")

        act = tk.Frame(self, bg=PANEL)
        act.grid(row=0, column=1, rowspan=6, padx=(0,8), pady=8, sticky="n")
        btns = [
            ("JUMP",   "jump",     "#4ade80", False),
            ("GRAB",   "grab",     "#fb923c", False),
            ("USE",    "use",      "#38bdf8", False),
            ("MENU",   "menu",     "#e879f9", False),
            ("MUTE",   "voice",    "#facc15", False),
            ("SIT",    "seated",   "#f87171", True),
            ("CROUCH", "crouched", "#67e8f9", True),
        ]
        for i, (label, action, colour, is_toggle) in enumerate(btns):
            b = make_toggle_btn(act, label, action, colour, self.state) if is_toggle \
                else make_action_btn(act, label, action, colour, self.state)
            b.grid(row=i//2, column=i%2, padx=4, pady=4)

# ══════════════════════════════════════════════════════════════════════════════════════════#
# PAD CARD
# ══════════════════════════════════════════════════════════════════════════════════════════#

class PadCard(tk.Frame):
    def __init__(self, parent, index, on_remove, host="127.0.0.1", port="9000", style="nes", name="", **kwargs):
        super().__init__(parent, bg=PANEL,
                         highlightthickness=1, highlightbackground=BORDER, **kwargs)
        self.index     = index
        self.on_remove = on_remove
        self.state     = None
        self._default_host  = host
        self._default_port  = port
        self._default_style = style
        self._default_name  = name if name else f"Pad {index}"
        self._connected     = False
        self._build()

    def _build(self):
        # ── Header ────────────────────────────────────────────────────────
        hdr = tk.Frame(self, bg=PANEL)
        hdr.pack(fill=tk.X, padx=8, pady=(8, 4))

        tk.Label(hdr, text="◈", font=(UI_FONT, 10, "bold"),
                 fg=ACCENT2, bg=PANEL).pack(side=tk.LEFT)

        self._name_var = tk.StringVar(value=self._default_name)
        name_entry = tk.Entry(hdr, textvariable=self._name_var,
                              font=(UI_FONT, 10, "bold"), fg=ACCENT2, bg=PANEL,
                              insertbackground=ACCENT2, relief="flat",
                              highlightthickness=0, width=18)
        name_entry.pack(side=tk.LEFT, padx=(4, 0))

        rm_btn = tk.Label(hdr, text="✕", font=(UI_FONT, 10),
                          fg=RED, bg=PANEL, cursor="hand2")
        rm_btn.pack(side=tk.RIGHT, padx=4)
        rm_btn.bind("<Button-1>", lambda e: self.on_remove(self))

        tk.Frame(self, bg=BORDER, height=1).pack(fill=tk.X, padx=8)

        # ── Config row ────────────────────────────────────────────────────
        cfg_row = tk.Frame(self, bg=PANEL)
        cfg_row.pack(fill=tk.X, padx=8, pady=6)

        def lbl(text):
            tk.Label(cfg_row, text=text, font=(UI_FONT, 8), fg=SUBTEXT,
                     bg=PANEL).pack(side=tk.LEFT)

        def entry(default, width):
            e = tk.Entry(cfg_row, font=(UI_FONT, 9), width=width,
                         bg=ENTRY_BG, fg=TEXT, insertbackground=ACCENT,
                         relief="flat", highlightthickness=1,
                         highlightbackground=BORDER, highlightcolor=ACCENT)
            e.insert(0, default)
            e.pack(side=tk.LEFT, padx=(2, 8))
            return e

        lbl("Host:")
        self._host = entry(self._default_host, 13)
        lbl("Port:")
        self._port = entry(self._default_port, 6)

        self._style_var = tk.StringVar(value=self._default_style)
        for val, txt in [("nes", "NES"), ("joy", "Joystick")]:
            tk.Radiobutton(cfg_row, text=txt, variable=self._style_var,
                           value=val, font=(UI_FONT, 8), fg=TEXT, bg=PANEL,
                           selectcolor=PANEL, activebackground=PANEL,
                           command=self._rebuild_pad).pack(side=tk.LEFT, padx=2)

        # Connect/Disconnect button
        self._conn_btn = tk.Button(cfg_row, text="Connect", font=(UI_FONT, 8, "bold"),
                                   fg=TEXT, bg=BTN_BG, relief="flat",
                                   activebackground=ACCENT, activeforeground="#ffffff",
                                   cursor="hand2", padx=6,
                                   command=self._toggle_connect)
        self._conn_btn.pack(side=tk.LEFT, padx=(8, 0))

        self._status_dot = tk.Label(cfg_row, text="●", font=(UI_FONT, 10),
                                    fg=SUBTEXT, bg=PANEL)
        self._status_dot.pack(side=tk.LEFT, padx=4)

        # ── Pad area ──────────────────────────────────────────────────────
        self._pad_area = tk.Frame(self, bg=PANEL)
        self._pad_area.pack(padx=4, pady=(0, 8))
        tk.Label(self._pad_area, text="Press Connect to activate",
                 font=(UI_FONT, 8), fg=SUBTEXT, bg=PANEL).pack(pady=16)

    def _toggle_connect(self):
        if self._connected:
            self._disconnect()
        else:
            self._connect()

    def _connect(self):
        host = self._host.get().strip()
        try:
            port = int(self._port.get().strip())
        except ValueError:
            self._status_dot.config(fg=RED)
            return
        if self.state:
            self.state.stop()
        self.state = PadState(host, port)
        self._connected = True
        self._status_dot.config(fg=GREEN)
        self._conn_btn.config(text="Disconnect",
                              fg=RED, activebackground=RED,
                              activeforeground="#ffffff")
        self._rebuild_pad()

    def _disconnect(self):
        if self.state:
            self.state.stop()
            self.state = None
        self._connected = False
        self._status_dot.config(fg=SUBTEXT)
        self._conn_btn.config(text="Connect",
                              fg=TEXT, activebackground=ACCENT,
                              activeforeground="#ffffff")
        for w in self._pad_area.winfo_children():
            w.destroy()
        tk.Label(self._pad_area, text="Press Connect to activate",
                 font=(UI_FONT, 8), fg=SUBTEXT, bg=PANEL).pack(pady=16)

    def _rebuild_pad(self):
        if not self.state:
            return
        for w in self._pad_area.winfo_children():
            w.destroy()
        cls = NESPad if self._style_var.get() == "nes" else JoystickPad
        cls(self._pad_area, self.state).pack()

    def get_config(self):
        return {
            "host":  self._host.get().strip(),
            "port":  self._port.get().strip(),
            "style": self._style_var.get(),
            "name":  self._name_var.get().strip(),
        }

    def destroy_state(self):
        if self.state:
            self.state.stop()

# ══════════════════════════════════════════════════════════════════════════════════════════#
# SETTINGS WINDOW
# ══════════════════════════════════════════════════════════════════════════════════════════#

def open_settings(root):
    win = tk.Toplevel(root)
    win.title("OSC Gamepad — Settings")
    win.configure(bg=BG)
    win.resizable(False, False)
    win.geometry("360x220")

    hdr = tk.Frame(win, bg=PANEL, pady=10)
    hdr.pack(fill="x")
    tk.Label(hdr, text="◈  SETTINGS", font=(UI_FONT, 12, "bold"),
             fg=ACCENT2, bg=PANEL).pack(side="left", padx=16)
    tk.Frame(win, bg=BORDER, height=1).pack(fill="x")

    content = tk.Frame(win, bg=PANEL, highlightthickness=1,
                       highlightbackground=BORDER)
    content.pack(padx=20, pady=14, fill="both", expand=True)

    tk.Label(content,
             text="No settings available yet.\nMore options coming in a future version.",
             font=(UI_FONT, 9), fg=SUBTEXT, bg=PANEL,
             justify="center").pack(expand=True)

    tk.Frame(win, bg=BORDER, height=1).pack(fill="x")
    footer = tk.Frame(win, bg=PANEL, pady=8)
    footer.pack(fill="x")
    tk.Button(footer, text="Close", command=win.destroy,
              bg=BTN_BG, fg=SUBTEXT, relief="flat",
              activebackground=BORDER, activeforeground=TEXT,
              cursor="hand2", font=(UI_FONT, 9, "bold"),
              padx=12).pack(side="right", padx=8)

# ══════════════════════════════════════════════════════════════════════════════════════════#
# HELP WINDOW
# ══════════════════════════════════════════════════════════════════════════════════════════#

def open_help(root):
    win = tk.Toplevel(root)
    win.title("OSC Gamepad — Help")
    win.configure(bg=BG)
    win.resizable(False, False)
    win.geometry("420x380")

    hdr = tk.Frame(win, bg=PANEL, pady=10)
    hdr.pack(fill="x")
    tk.Label(hdr, text="◈  HELP", font=(UI_FONT, 12, "bold"),
             fg=ACCENT2, bg=PANEL).pack(side="left", padx=16)
    tk.Frame(win, bg=BORDER, height=1).pack(fill="x")

    content = tk.Frame(win, bg=PANEL, highlightthickness=1,
                       highlightbackground=BORDER)
    content.pack(padx=20, pady=14, fill="both", expand=True)

    lines = [
        ("SETUP", True),
        ("1. Enable OSC in VRChat: Action Menu → OSC → Enabled", False),
        ("2. Launch VRChat with --osc=9000:127.0.0.1:9001", False),
        ("   or use custom ports per pad below.", False),
        ("", False),
        ("PADS", True),
        ("+ Add Pad  — adds a new controller.", False),
        ("Host/Port  — OSC destination (default 127.0.0.1:9000).", False),
        ("NES        — D-pad style movement buttons.", False),
        ("Joystick   — analogue stick + look sliders.", False),
        ("Connect    — green dot = active.", False),
        ("", False),
        ("BUTTONS", True),
        ("SIT / CROUCH  — toggle avatar parameters (click once on/off).", False),
        ("MUTE          — pulses /input/Voice to toggle mute.", False),
        ("JUMP          — single pulse per press.", False),
        ("GRAB/USE/MENU — held while button is down.", False),
        ("", False),
        ("Config is saved automatically on exit.", False),
    ]

    text = tk.Text(content, bg=PANEL, fg=TEXT, font=(UI_FONT, 9),
                   relief="flat", cursor="arrow", wrap="word",
                   highlightthickness=0, padx=10, pady=10,
                   state="normal")
    text.pack(fill="both", expand=True)
    text.tag_config("heading", fg=ACCENT2, font=(UI_FONT, 9, "bold"))
    for line, is_heading in lines:
        if is_heading:
            text.insert("end", line + "\n", "heading")
        else:
            text.insert("end", line + "\n")
    text.config(state="disabled")

    tk.Frame(win, bg=BORDER, height=1).pack(fill="x")
    footer = tk.Frame(win, bg=PANEL, pady=8)
    footer.pack(fill="x")
    tk.Button(footer, text="Close", command=win.destroy,
              bg=BTN_BG, fg=SUBTEXT, relief="flat",
              activebackground=BORDER, activeforeground=TEXT,
              cursor="hand2", font=(UI_FONT, 9, "bold"),
              padx=12).pack(side="right", padx=8)

# ══════════════════════════════════════════════════════════════════════════════════════════#
# MAIN APP
# ══════════════════════════════════════════════════════════════════════════════════════════#

class App:
    def __init__(self):
        self.cfg = load_config()
        self.root = tk.Tk()
        self.root.title("OSC Gamepad")
        self.root.configure(bg=BG)
        self.root.resizable(True, True)
        self.root.geometry("660x560")
        self.cards = []
        self._pad_counter = 0
        self._build()
        self._load_pads()
        self.root.protocol("WM_DELETE_WINDOW", self._quit)

    def _build(self):
        # ── Title bar (matches OSC-Chatbox) ───────────────────────────────
        title_bar = tk.Frame(self.root, bg=PANEL, pady=14)
        title_bar.pack(fill="x")

        header_frame = tk.Frame(title_bar, bg=PANEL)
        header_frame.pack(fill="x", padx=16, expand=True)

        tk.Label(header_frame, text="◈  OSC GAMEPAD",
                 bg=PANEL, fg=ACCENT2,
                 font=(UI_FONT, 16, "bold")).pack(side="left", anchor="w")

        tk.Label(header_frame, text=f"v{VERSION}",
                 bg=PANEL, fg=SUBTEXT,
                 font=(UI_FONT, 9)).pack(side="right", anchor="e", padx=(32, 16))

        add_btn = tk.Label(header_frame, text="＋ Add Pad",
                           bg=PANEL, fg=ACCENT, cursor="hand2",
                           font=(UI_FONT, 9, "bold"))
        add_btn.pack(side="right", anchor="e")
        add_btn.bind("<Button-1>", lambda e: self._add_pad())
        add_btn.bind("<Enter>",    lambda e: add_btn.config(fg=ACCENT2))
        add_btn.bind("<Leave>",    lambda e: add_btn.config(fg=ACCENT))

        tk.Frame(self.root, bg=BORDER, height=1).pack(fill="x")

        # ── Scrollable content ────────────────────────────────────────────
        outer = tk.Frame(self.root, bg=BG)
        outer.pack(fill=tk.BOTH, expand=True, padx=0, pady=0)

        style = ttk.Style()
        style.theme_use("clam")
        style.configure("Dark.Vertical.TScrollbar",
                         background=PANEL, troughcolor=BG,
                         arrowcolor=SUBTEXT, bordercolor=BG,
                         lightcolor=PANEL, darkcolor=PANEL)

        self._canvas = tk.Canvas(outer, bg=BG, highlightthickness=0)
        scrollbar = ttk.Scrollbar(outer, orient="vertical",
                                   command=self._canvas.yview,
                                   style="Dark.Vertical.TScrollbar")
        self._canvas.configure(yscrollcommand=scrollbar.set)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self._canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        self._inner = tk.Frame(self._canvas, bg=BG)
        self._win_id = self._canvas.create_window((0, 0), window=self._inner, anchor="nw")

        self._inner.bind("<Configure>", lambda e: self._canvas.configure(
            scrollregion=self._canvas.bbox("all")))
        self._canvas.bind("<Configure>", lambda e: self._canvas.itemconfig(
            self._win_id, width=e.width))
        self._canvas.bind_all("<MouseWheel>",
                              lambda e: self._canvas.yview_scroll(-1*(e.delta//120), "units"))

        # ── Footer bar (matches OSC-Chatbox) ─────────────────────────────
        tk.Frame(self.root, bg=BORDER, height=1).pack(fill="x")
        footer_bar = tk.Frame(self.root, bg=PANEL, pady=8)
        footer_bar.pack(fill="x", side="bottom")

        help_btn = square_button(footer_bar, "？", lambda: open_help(self.root))
        help_btn.pack(side="left", padx=(8, 0))

        tk.Label(footer_bar, text=f"OSC Gamepad  v{VERSION}",
                 bg=PANEL, fg=SUBTEXT, font=(UI_FONT, 8)).pack(side="left", padx=16)

        settings_btn = square_button(footer_bar, "⚙", lambda: open_settings(self.root))
        settings_btn.pack(side="right", padx=(0, 8))

    def _load_pads(self):
        saved = self.cfg.get("pads", [])
        if saved:
            for pad_cfg in saved:
                self._add_pad(
                    host=pad_cfg.get("host", "127.0.0.1"),
                    port=pad_cfg.get("port", "9000"),
                    style=pad_cfg.get("style", "nes"),
                    name=pad_cfg.get("name", ""),
                )
        else:
            self._add_pad()

    def _add_pad(self, host="127.0.0.1", port="9000", style="nes", name=""):
        self._pad_counter += 1
        card = PadCard(self._inner, self._pad_counter, self._remove_pad,
                       host=host, port=str(port), style=style, name=name)
        card.pack(fill=tk.X, padx=12, pady=6)
        self.cards.append(card)

    def _remove_pad(self, card):
        card.destroy_state()
        card.destroy()
        self.cards.remove(card)

    def _save_and_quit(self):
        pads_data = [c.get_config() for c in self.cards]
        save_config(pads_data)

    def _quit(self):
        self._save_and_quit()
        for c in self.cards:
            c.destroy_state()
        self.root.destroy()

    def run(self):
        self.root.mainloop()


if __name__ == "__main__":
    print("OSC Gamepad")
    print("Made By Boots")
    print(f"Version {VERSION}")
    App().run()
