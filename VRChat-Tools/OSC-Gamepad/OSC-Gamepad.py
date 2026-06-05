import math
import threading
import time
import tkinter as tk
from tkinter import ttk

from pythonosc import udp_client


VERSION = "0.0.1"

BG = "#0e0e1c"
PANEL_BG = "#13132a"
BORDER = "#7b2fff"
ACCENT = "#a855f7"
BTN_IDLE = "#1e1e3a"
BTN_HOV = "#2d2d55"
BTN_ACT = "#7b2fff"
FG = "#e2d9f3"
FG_DIM = "#6b6b9a"
FONT = "Consolas"

# Axes  = float -1..1, reset to 0 when released
# Buttons = int 1 pressed / 0 released
# Avatar params (Sit/Crouch) are custom bools via /avatar/parameters/

BUTTON_ADDRESSES = {
    "jump": ("/input/Jump",),
    # VRChat only exposes hand-specific grab/use inputs, so send both hands.
    "grab": ("/input/GrabLeft", "/input/GrabRight"),
    "use": ("/input/UseLeft", "/input/UseRight"),
    "menu": ("/input/QuickMenuToggleLeft",),
    "voice": ("/input/Voice",),
}

PULSE_BUTTONS = {"jump", "menu"}

AVATAR_PARAM_TARGETS = {
    "sit": "Sit",
    "crouch": "Crouch",
}


class PadState:
    def __init__(self, host, port):
        self.host = host
        self.port = port
        self.client = udp_client.SimpleUDPClient(host, port)
        self.axes_held = set()
        self.avatar_states = {key: False for key in AVATAR_PARAM_TARGETS}
        self._pulse_tokens = {key: 0 for key in PULSE_BUTTONS}
        self._send_lock = threading.Lock()
        self.running = True
        self.thread = threading.Thread(target=self._loop, daemon=True)
        self.thread.start()

    def send_message(self, address, value):
        try:
            with self._send_lock:
                self.client.send_message(address, value)
        except Exception:
            pass

    def send_button(self, action, value):
        for address in BUTTON_ADDRESSES.get(action, ()):
            self.send_message(address, value)

    def press_axis(self, action):
        self.axes_held.add(action)

    def release_axis(self, action):
        self.axes_held.discard(action)

    def _pulse_release(self, action, token):
        if self._pulse_tokens.get(action) == token:
            self.send_button(action, 0)

    def press_btn(self, action):
        self.send_button(action, 1)
        if action in PULSE_BUTTONS:
            self._pulse_tokens[action] = self._pulse_tokens.get(action, 0) + 1
            token = self._pulse_tokens[action]
            timer = threading.Timer(0.08, self._pulse_release, args=(action, token))
            timer.daemon = True
            timer.start()

    def release_btn(self, action):
        if action not in PULSE_BUTTONS:
            self.send_button(action, 0)

    def toggle_avatar_param(self, param):
        target = AVATAR_PARAM_TARGETS.get(param, param)
        new_value = not self.avatar_states.get(param, False)
        self.avatar_states[param] = new_value
        # Built-in params like Seated are read-only in VRChat, so this targets
        # custom avatar parameters instead.
        self.send_message(f"/avatar/parameters/{target}", new_value)
        return new_value

    def _loop(self):
        while self.running:
            v = h = lh = lv = 0.0
            for a in list(self.axes_held):
                if a == "up":
                    v = 1.0
                if a == "down":
                    v = -1.0
                if a == "left":
                    h = -1.0
                if a == "right":
                    h = 1.0
                if a == "look_l":
                    lh = -1.0
                if a == "look_r":
                    lh = 1.0
                if a == "look_u":
                    lv = 1.0
                if a == "look_d":
                    lv = -1.0

            self.send_message("/input/Vertical", v)
            self.send_message("/input/Horizontal", h)
            self.send_message("/input/LookHorizontal", lh)
            self.send_message("/input/LookVertical", lv)
            time.sleep(0.05)

    def stop(self):
        self.running = False


def make_axis_btn(parent, label, action, state, size=44, font_size=14):
    b = tk.Label(
        parent,
        text=label,
        font=(FONT, font_size, "bold"),
        width=2,
        height=1,
        fg=FG,
        bg=BTN_IDLE,
        relief="flat",
        cursor="hand2",
    )

    def press(_event):
        state.press_axis(action)

    def release(_event=None):
        state.release_axis(action)
        b.config(bg=BTN_IDLE)

    b.bind("<ButtonPress-1>", press)
    b.bind("<ButtonRelease-1>", release)
    b.bind("<Enter>", lambda _event: b.config(bg=BTN_HOV))
    b.bind("<Leave>", lambda _event: release())
    return b


def make_action_btn(parent, label, action, colour, state, width=5, height=2):
    b = tk.Label(
        parent,
        text=label,
        font=(FONT, 8, "bold"),
        width=width,
        height=height,
        fg=colour,
        bg=BTN_IDLE,
        relief="flat",
        cursor="hand2",
        highlightthickness=1,
        highlightbackground=colour,
    )

    def press(_event):
        state.press_btn(action)
        b.config(bg=BTN_ACT)

    def release(_event=None):
        state.release_btn(action)
        b.config(bg=BTN_IDLE)

    b.bind("<ButtonPress-1>", press)
    b.bind("<ButtonRelease-1>", release)
    b.bind("<Enter>", lambda _event: b.config(bg=BTN_HOV))
    b.bind("<Leave>", lambda _event: release())
    return b


def make_toggle_btn(parent, label, param, colour, state, width=5, height=2):
    """Toggle button for custom avatar params."""
    active = [False]
    b = tk.Label(
        parent,
        text=label,
        font=(FONT, 8, "bold"),
        width=width,
        height=height,
        fg=colour,
        bg=BTN_IDLE,
        relief="flat",
        cursor="hand2",
        highlightthickness=1,
        highlightbackground=colour,
    )

    def click(_event):
        result = state.toggle_avatar_param(param)
        active[0] = result
        b.config(bg=BTN_ACT if result else BTN_IDLE)

    b.bind("<Button-1>", click)
    b.bind("<Enter>", lambda _event: b.config(bg=BTN_HOV if not active[0] else BTN_ACT))
    b.bind("<Leave>", lambda _event: b.config(bg=BTN_ACT if active[0] else BTN_IDLE))
    return b


class NESPad(tk.Frame):
    def __init__(self, parent, state, **kwargs):
        super().__init__(parent, bg=PANEL_BG, **kwargs)
        self.state = state
        self._build()

    def _build(self):
        dpad = tk.Frame(self, bg=PANEL_BG)
        dpad.grid(row=0, column=0, padx=(8, 12), pady=8)

        make_axis_btn(dpad, "▲", "up", self.state).grid(row=0, column=1, padx=2, pady=2)
        make_axis_btn(dpad, "◀", "left", self.state).grid(row=1, column=0, padx=2, pady=2)
        make_axis_btn(dpad, "▶", "right", self.state).grid(row=1, column=2, padx=2, pady=2)
        make_axis_btn(dpad, "▼", "down", self.state).grid(row=2, column=1, padx=2, pady=2)
        tk.Frame(dpad, width=44, height=44, bg="#0e0e1e").grid(row=1, column=1, padx=2, pady=2)

        look = tk.Frame(self, bg=PANEL_BG)
        look.grid(row=1, column=0, padx=(8, 12), pady=(0, 8))
        tk.Label(look, text="LOOK", font=(FONT, 7), fg=FG_DIM, bg=PANEL_BG).pack()
        lh = tk.Frame(look, bg=PANEL_BG)
        lh.pack()
        make_axis_btn(lh, "◀", "look_l", self.state, font_size=11).pack(side=tk.LEFT, padx=2)
        make_axis_btn(lh, "▶", "look_r", self.state, font_size=11).pack(side=tk.LEFT, padx=2)
        lv = tk.Frame(look, bg=PANEL_BG)
        lv.pack(pady=(2, 0))
        make_axis_btn(lv, "▲", "look_u", self.state, font_size=11).pack(side=tk.LEFT, padx=2)
        make_axis_btn(lv, "▼", "look_d", self.state, font_size=11).pack(side=tk.LEFT, padx=2)

        act = tk.Frame(self, bg=PANEL_BG)
        act.grid(row=0, column=1, rowspan=2, padx=(0, 8), pady=8, sticky="n")

        btns = [
            ("JUMP", "jump", "#2fff9a", False),
            ("GRAB", "grab", "#ff9a2f", False),
            ("USE", "use", "#2fc8ff", False),
            ("MENU", "menu", "#ff2fff", False),
            ("MUTE", "voice", "#ffff2f", False),
            ("SIT", "sit", "#ff7f7f", True),
            ("CROUCH", "crouch", "#7fffff", True),
        ]
        for i, (label, action, colour, is_toggle) in enumerate(btns):
            b = make_toggle_btn(act, label, action, colour, self.state) if is_toggle else make_action_btn(act, label, action, colour, self.state)
            b.grid(row=i // 2, column=i % 2, padx=4, pady=4)


class JoystickPad(tk.Frame):
    def __init__(self, parent, state, **kwargs):
        super().__init__(parent, bg=PANEL_BG, **kwargs)
        self.state = state
        self._build()

    def _build(self):
        ps = 170
        cen = ps // 2
        knob_size = 16

        canvas = tk.Canvas(
            self,
            width=ps,
            height=ps,
            bg="#0e0e1c",
            highlightthickness=2,
            highlightbackground=BORDER,
        )
        canvas.grid(row=0, column=0, padx=8, pady=8)
        canvas.create_oval(4, 4, ps - 4, ps - 4, outline="#2a2a4e", width=1)
        canvas.create_oval(cen - 5, cen - 5, cen + 5, cen + 5, outline="#2a2a4e", fill="#13132a")
        knob = canvas.create_oval(
            cen - knob_size,
            cen - knob_size,
            cen + knob_size,
            cen + knob_size,
            fill=BORDER,
            outline=ACCENT,
            width=2,
        )

        def drag(ev):
            dx = ev.x - cen
            dy = ev.y - cen
            dist = math.hypot(dx, dy)
            max_r = cen - knob_size - 4
            if dist > max_r:
                dx = dx / dist * max_r
                dy = dy / dist * max_r
            canvas.coords(knob, cen + dx - knob_size, cen + dy - knob_size, cen + dx + knob_size, cen + dy + knob_size)
            self.state.send_message("/input/Horizontal", round(max(-1.0, min(1.0, dx / max_r)), 3))
            self.state.send_message("/input/Vertical", round(max(-1.0, min(1.0, -dy / max_r)), 3))

        def release(_event):
            canvas.coords(knob, cen - knob_size, cen - knob_size, cen + knob_size, cen + knob_size)
            self.state.send_message("/input/Horizontal", 0.0)
            self.state.send_message("/input/Vertical", 0.0)

        canvas.bind("<B1-Motion>", drag)
        canvas.bind("<ButtonRelease-1>", release)

        sw, sh, kw = ps, 36, 26

        def make_slider(row, osc_addr, label):
            slider = tk.Canvas(
                self,
                width=sw,
                height=sh,
                bg="#0e0e1c",
                highlightthickness=2,
                highlightbackground=BORDER,
            )
            slider.grid(row=row, column=0, padx=8, pady=(0, 2))
            ty = sh // 2
            slider.create_line(kw // 2, ty, sw - kw // 2, ty, fill="#2a2a4e", width=2)
            handle = slider.create_oval(
                sw // 2 - kw // 2,
                ty - kw // 2,
                sw // 2 + kw // 2,
                ty + kw // 2,
                fill=BORDER,
                outline=ACCENT,
                width=2,
            )
            tk.Label(self, text=label, font=(FONT, 7), fg=FG_DIM, bg=PANEL_BG).grid(row=row + 1, column=0)

            def sdrag(ev):
                usable = sw - kw
                x = max(kw // 2, min(sw - kw // 2, ev.x))
                slider.coords(handle, x - kw // 2, ty - kw // 2, x + kw // 2, ty + kw // 2)
                self.state.send_message(osc_addr, round((x - sw / 2) / (usable / 2), 3))

            def srelease(_event):
                slider.coords(handle, sw // 2 - kw // 2, ty - kw // 2, sw // 2 + kw // 2, ty + kw // 2)
                self.state.send_message(osc_addr, 0.0)

            slider.bind("<B1-Motion>", sdrag)
            slider.bind("<ButtonRelease-1>", srelease)

        make_slider(1, "/input/LookHorizontal", "LOOK H")
        make_slider(3, "/input/LookVertical", "LOOK V")

        act = tk.Frame(self, bg=PANEL_BG)
        act.grid(row=0, column=1, rowspan=6, padx=(0, 8), pady=8, sticky="n")
        btns = [
            ("JUMP", "jump", "#2fff9a", False),
            ("GRAB", "grab", "#ff9a2f", False),
            ("USE", "use", "#2fc8ff", False),
            ("MENU", "menu", "#ff2fff", False),
            ("MUTE", "voice", "#ffff2f", False),
            ("SIT", "sit", "#ff7f7f", True),
            ("CROUCH", "crouch", "#7fffff", True),
        ]
        for i, (label, action, colour, is_toggle) in enumerate(btns):
            b = make_toggle_btn(act, label, action, colour, self.state) if is_toggle else make_action_btn(act, label, action, colour, self.state)
            b.grid(row=i // 2, column=i % 2, padx=4, pady=4)


class PadCard(tk.Frame):
    def __init__(self, parent, index, on_remove, **kwargs):
        super().__init__(parent, bg=PANEL_BG, highlightthickness=1, highlightbackground=BORDER, **kwargs)
        self.index = index
        self.on_remove = on_remove
        self.state = None
        self._build()

    def _build(self):
        hdr = tk.Frame(self, bg="#1a1a35")
        hdr.pack(fill=tk.X, padx=1, pady=(1, 0))

        tk.Label(hdr, text=f"PAD {self.index}", font=(FONT, 9, "bold"), fg=ACCENT, bg="#1a1a35").pack(side=tk.LEFT, padx=8, pady=4)

        rm_btn = tk.Label(hdr, text="✕", font=(FONT, 10, "bold"), fg="#ff4466", bg="#1a1a35", cursor="hand2")
        rm_btn.pack(side=tk.RIGHT, padx=8, pady=4)
        rm_btn.bind("<Button-1>", lambda _event: self.on_remove(self))

        cfg = tk.Frame(self, bg=PANEL_BG)
        cfg.pack(fill=tk.X, padx=8, pady=4)

        tk.Label(cfg, text="Host:", font=(FONT, 8), fg=FG_DIM, bg=PANEL_BG).pack(side=tk.LEFT)
        self._host = tk.Entry(cfg, font=(FONT, 9), width=12, bg="#0e0e1c", fg=FG, insertbackground=FG, relief="flat", highlightthickness=1, highlightbackground=BORDER)
        self._host.insert(0, "127.0.0.1")
        self._host.pack(side=tk.LEFT, padx=(2, 8))

        tk.Label(cfg, text="Port:", font=(FONT, 8), fg=FG_DIM, bg=PANEL_BG).pack(side=tk.LEFT)
        self._port = tk.Entry(cfg, font=(FONT, 9), width=6, bg="#0e0e1c", fg=FG, insertbackground=FG, relief="flat", highlightthickness=1, highlightbackground=BORDER)
        self._port.insert(0, "9000")
        self._port.pack(side=tk.LEFT, padx=(2, 8))

        self._style = tk.StringVar(value="nes")
        tk.Radiobutton(cfg, text="NES", variable=self._style, value="nes", font=(FONT, 8), fg=FG, bg=PANEL_BG, selectcolor=PANEL_BG, activebackground=PANEL_BG, command=self._rebuild_pad).pack(side=tk.LEFT, padx=2)
        tk.Radiobutton(cfg, text="Joystick", variable=self._style, value="joy", font=(FONT, 8), fg=FG, bg=PANEL_BG, selectcolor=PANEL_BG, activebackground=PANEL_BG, command=self._rebuild_pad).pack(side=tk.LEFT, padx=2)

        conn_btn = tk.Label(cfg, text="CONNECT", font=(FONT, 8, "bold"), fg="#2fff9a", bg=BTN_IDLE, cursor="hand2", padx=6, pady=2, highlightthickness=1, highlightbackground="#2fff9a")
        conn_btn.pack(side=tk.LEFT, padx=(8, 0))
        conn_btn.bind("<Button-1>", lambda _event: self._connect())

        self._status = tk.Label(cfg, text="●", font=(FONT, 10), fg=FG_DIM, bg=PANEL_BG)
        self._status.pack(side=tk.LEFT, padx=4)

        self._pad_area = tk.Frame(self, bg=PANEL_BG)
        self._pad_area.pack(padx=4, pady=(0, 6))
        tk.Label(self._pad_area, text="Press CONNECT to activate", font=(FONT, 8), fg=FG_DIM, bg=PANEL_BG).pack(pady=16)

    def _connect(self):
        host = self._host.get().strip()
        try:
            port = int(self._port.get().strip())
        except ValueError:
            self._status.config(fg="#ff4466")
            return

        if self.state:
            self.state.stop()

        self.state = PadState(host, port)
        self._status.config(fg="#2fff9a")
        self._rebuild_pad()

    def _rebuild_pad(self):
        if not self.state:
            return
        for widget in self._pad_area.winfo_children():
            widget.destroy()
        if self._style.get() == "nes":
            NESPad(self._pad_area, self.state).pack()
        else:
            JoystickPad(self._pad_area, self.state).pack()

    def destroy_state(self):
        if self.state:
            self.state.stop()


class App:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("VRC OSC Gamepad Manager")
        self.root.configure(bg=BG)
        self.root.resizable(True, True)
        self.cards = []
        self._pad_counter = 0
        self._build()
        self.root.protocol("WM_DELETE_WINDOW", self._quit)

    def _build(self):
        title = tk.Frame(self.root, bg=BG)
        title.pack(fill=tk.X, padx=12, pady=(10, 4))
        tk.Label(title, text="VRC OSC GAMEPAD", font=(FONT, 13, "bold"), fg=ACCENT, bg=BG).pack(side=tk.LEFT)
        add_btn = tk.Label(title, text="＋ ADD PAD", font=(FONT, 9, "bold"), fg=BORDER, bg=BG, cursor="hand2")
        add_btn.pack(side=tk.RIGHT)
        add_btn.bind("<Button-1>", lambda _event: self._add_pad())
        add_btn.bind("<Enter>", lambda _event: add_btn.config(fg=ACCENT))
        add_btn.bind("<Leave>", lambda _event: add_btn.config(fg=BORDER))

        tk.Frame(self.root, bg=BORDER, height=1).pack(fill=tk.X, padx=12, pady=(0, 8))

        outer = tk.Frame(self.root, bg=BG)
        outer.pack(fill=tk.BOTH, expand=True, padx=8, pady=(0, 8))

        self._canvas = tk.Canvas(outer, bg=BG, highlightthickness=0)
        scrollbar = ttk.Scrollbar(outer, orient="vertical", command=self._canvas.yview)
        self._canvas.configure(yscrollcommand=scrollbar.set)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self._canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        self._inner = tk.Frame(self._canvas, bg=BG)
        self._win_id = self._canvas.create_window((0, 0), window=self._inner, anchor="nw")
        self._inner.bind("<Configure>", lambda _event: self._canvas.configure(scrollregion=self._canvas.bbox("all")))
        self._canvas.bind("<Configure>", lambda event: self._canvas.itemconfig(self._win_id, width=event.width))
        self._canvas.bind_all("<MouseWheel>", lambda event: self._canvas.yview_scroll(-1 * (event.delta // 120), "units"))
        self._add_pad()

    def _add_pad(self):
        self._pad_counter += 1
        card = PadCard(self._inner, self._pad_counter, self._remove_pad)
        card.pack(fill=tk.X, padx=6, pady=6)
        self.cards.append(card)

    def _remove_pad(self, card):
        card.destroy_state()
        card.destroy()
        self.cards.remove(card)

    def _quit(self):
        for card in self.cards:
            card.destroy_state()
        self.root.destroy()

    def run(self):
        self.root.mainloop()


if __name__ == "__main__":
    App().run()
