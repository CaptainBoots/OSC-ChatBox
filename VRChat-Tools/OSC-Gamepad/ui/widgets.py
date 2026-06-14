"""
ui/widgets.py
─────────────
Reusable button factories (axis/action/toggle/square) and the
NES / Joystick pad layout widgets.
"""

import math
import tkinter as tk

from ui.theme import BG, PANEL, BORDER, ACCENT, ACCENT2, TEXT, SUBTEXT, GREEN, RED, YELLOW, ORANGE, CYAN, FONT


# ── Button factories ──────────────────────────────────────────────────────────

def make_axis_btn(parent, label: str, action: str, state, font_size: int = 13) -> tk.Label:
    b = tk.Label(
        parent, text=label, font=(FONT, font_size, "bold"),
        width=2, height=1, fg=TEXT, bg=PANEL,
        relief="flat", cursor="hand2",
        highlightthickness=1, highlightbackground=BORDER,
    )
    b.bind("<ButtonPress-1>",   lambda e, a=action: state.press_axis(a))
    b.bind("<ButtonRelease-1>", lambda e, a=action: state.release_axis(a))
    b.bind("<Enter>", lambda e, w=b: w.config(bg=BORDER))
    b.bind("<Leave>", lambda e, w=b, a=action: (state.release_axis(a), w.config(bg=PANEL)))
    return b


def make_action_btn(parent, label: str, action: str, colour: str, state,
                     width: int = 5, height: int = 2) -> tk.Label:
    b = tk.Label(
        parent, text=label, font=(FONT, 8, "bold"),
        width=width, height=height, fg=colour, bg=PANEL,
        relief="flat", cursor="hand2",
        highlightthickness=1, highlightbackground=colour,
    )
    b.bind("<ButtonPress-1>",   lambda e, a=action, w=b: (state.press_btn(a), w.config(bg=ACCENT)))
    b.bind("<ButtonRelease-1>", lambda e, a=action, w=b: (state.release_btn(a), w.config(bg=PANEL)))
    b.bind("<Enter>", lambda e, w=b: w.config(bg=BORDER))
    b.bind("<Leave>", lambda e, w=b, a=action: (state.release_btn(a), w.config(bg=PANEL)))
    return b


def make_toggle_btn(parent, label: str, param: str, colour: str, state,
                     width: int = 5, height: int = 2) -> tk.Label:
    active = [False]
    b = tk.Label(
        parent, text=label, font=(FONT, 8, "bold"),
        width=width, height=height, fg=colour, bg=PANEL,
        relief="flat", cursor="hand2",
        highlightthickness=1, highlightbackground=colour,
    )

    def click(_event):
        result = state.toggle_avatar_param(param)
        active[0] = result
        b.config(bg=ACCENT if result else PANEL)

    b.bind("<Button-1>", click)
    b.bind("<Enter>", lambda e, w=b: w.config(bg=BORDER if not active[0] else ACCENT))
    b.bind("<Leave>", lambda e, w=b: w.config(bg=ACCENT if active[0] else PANEL))
    return b


def square_button(parent, text: str, command, base_size: int = 28) -> tk.Frame:
    container = tk.Frame(parent, bg=PANEL, highlightthickness=1, highlightbackground=BORDER)
    container.pack_propagate(False)
    container.config(width=base_size, height=base_size)
    btn = tk.Button(
        container, text=text, command=command,
        bg=PANEL, fg=SUBTEXT, relief="flat", borderwidth=0,
        font=(FONT, 12), activebackground=BORDER,
        activeforeground=TEXT, cursor="hand2",
    )
    btn.pack(fill="both", expand=True)
    return container


# ── Action button definitions shared by both pad styles ──────────────────────

ACTION_BUTTONS = [
    ("JUMP",   "jump",     GREEN, False),
    ("GRAB",   "grab",     ORANGE, False),
    ("USE",    "use",      ACCENT2, False),
    ("MENU",   "menu",     ACCENT, False),
    ("MUTE",   "voice",    YELLOW, False),
    ("SIT",    "seated",   RED, True),
    ("CROUCH", "crouched", CYAN, True),
]


def build_action_grid(parent, state) -> tk.Frame:
    act = tk.Frame(parent, bg=PANEL)
    for i, (label, action, colour, is_toggle) in enumerate(ACTION_BUTTONS):
        b = (make_toggle_btn(act, label, action, colour, state) if is_toggle
             else make_action_btn(act, label, action, colour, state))
        b.grid(row=i // 2, column=i % 2, padx=4, pady=4)
    return act


# ── NES-style D-pad layout ────────────────────────────────────────────────────

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
        tk.Label(look, text="LOOK", font=(FONT, 7), fg=SUBTEXT, bg=PANEL).pack()

        lh_row = tk.Frame(look, bg=PANEL)
        lh_row.pack()
        make_axis_btn(lh_row, "◀", "look_l", self.state, font_size=10).pack(side=tk.LEFT, padx=2)
        make_axis_btn(lh_row, "▶", "look_r", self.state, font_size=10).pack(side=tk.LEFT, padx=2)

        lv_row = tk.Frame(look, bg=PANEL)
        lv_row.pack(pady=(2, 0))
        make_axis_btn(lv_row, "▲", "look_u", self.state, font_size=10).pack(side=tk.LEFT, padx=2)
        make_axis_btn(lv_row, "▼", "look_d", self.state, font_size=10).pack(side=tk.LEFT, padx=2)

        build_action_grid(self, self.state).grid(row=0, column=1, rowspan=2, padx=(0, 8), pady=8, sticky="n")


# ── Joystick-style analogue layout ────────────────────────────────────────────

class JoystickPad(tk.Frame):
    def __init__(self, parent, state, **kwargs):
        super().__init__(parent, bg=PANEL, **kwargs)
        self.state = state
        self._build()

    def _build(self):
        PS, KNOB = 170, 16
        CEN = PS // 2

        c = tk.Canvas(self, width=PS, height=PS, bg=BG,
                      highlightthickness=1, highlightbackground=BORDER)
        c.grid(row=0, column=0, padx=8, pady=8)
        c.create_oval(4, 4, PS - 4, PS - 4, outline=BORDER, width=1)
        c.create_oval(CEN - 5, CEN - 5, CEN + 5, CEN + 5, outline=BORDER, fill=PANEL)
        knob = c.create_oval(CEN - KNOB, CEN - KNOB, CEN + KNOB, CEN + KNOB,
                             fill=ACCENT, outline=ACCENT2, width=2)

        def drag(ev):
            dx, dy = ev.x - CEN, ev.y - CEN
            dist   = math.hypot(dx, dy)
            max_r  = CEN - KNOB - 4
            if dist > max_r:
                dx, dy = dx / dist * max_r, dy / dist * max_r
            c.coords(knob, CEN + dx - KNOB, CEN + dy - KNOB, CEN + dx + KNOB, CEN + dy + KNOB)
            self.state._safe_send("/input/Horizontal", round(max(-1.0, min(1.0,  dx / max_r)), 3))
            self.state._safe_send("/input/Vertical",   round(max(-1.0, min(1.0, -dy / max_r)), 3))

        def rel(_ev):
            c.coords(knob, CEN - KNOB, CEN - KNOB, CEN + KNOB, CEN + KNOB)
            self.state._safe_send("/input/Horizontal", 0.0)
            self.state._safe_send("/input/Vertical",   0.0)

        c.bind("<B1-Motion>", drag)
        c.bind("<ButtonRelease-1>", rel)

        SW, SH, KW = PS, 36, 26

        def make_slider(row, osc_addr, label):
            sc = tk.Canvas(self, width=SW, height=SH, bg=BG,
                          highlightthickness=1, highlightbackground=BORDER)
            sc.grid(row=row, column=0, padx=8, pady=(0, 2))
            ty = SH // 2
            sc.create_line(KW // 2, ty, SW - KW // 2, ty, fill=BORDER, width=2)
            lk = sc.create_oval(SW // 2 - KW // 2, ty - KW // 2, SW // 2 + KW // 2, ty + KW // 2,
                                fill=ACCENT, outline=ACCENT2, width=2)
            tk.Label(self, text=label, font=(FONT, 7), fg=SUBTEXT, bg=PANEL).grid(row=row + 1, column=0)

            def sdrag(ev):
                usable = SW - KW
                x = max(KW // 2, min(SW - KW // 2, ev.x))
                sc.coords(lk, x - KW // 2, ty - KW // 2, x + KW // 2, ty + KW // 2)
                self.state._safe_send(osc_addr, round((x - SW / 2) / (usable / 2), 3))

            def srel(_ev):
                sc.coords(lk, SW // 2 - KW // 2, ty - KW // 2, SW // 2 + KW // 2, ty + KW // 2)
                self.state._safe_send(osc_addr, 0.0)

            sc.bind("<B1-Motion>", sdrag)
            sc.bind("<ButtonRelease-1>", srel)

        make_slider(1, "/input/LookHorizontal", "LOOK H")
        make_slider(3, "/input/LookVertical",   "LOOK V")

        build_action_grid(self, self.state).grid(row=0, column=1, rowspan=6, padx=(0, 8), pady=8, sticky="n")
