"""
ui/help_dialog.py
─────────────────
Paginated help window for OSC-Gamepad.
"""

import tkinter as tk

from ui.theme import BG, PANEL, BORDER, ACCENT, ACCENT2, TEXT, SUBTEXT, FONT

HELP_PAGES = [
    {
        "title": "Getting Started",
        "content": (
            "OSC Gamepad lets you control your VRChat avatar using\n"
            "on-screen buttons and joysticks, sent over OSC.\n\n"
            "1. Make sure VRChat is running with OSC enabled.\n"
            "   (Action Menu → Options → OSC → Enabled)\n\n"
            "2. Click + Add Pad to create a new pad.\n\n"
            "3. Set the Host and Port to match your VRChat OSC\n"
            "   settings. Default is 127.0.0.1 : 9000 for a\n"
            "   local game instance.\n\n"
            "4. Click Connect. The status dot turns green\n"
            "   when active.\n\n"
            "5. Use the pad controls to move and interact.\n\n"
            "Your pad layout and connection settings are saved\n"
            "automatically when you close the app."
        ),
    },
    {
        "title": "NES Pad Mode",
        "content": (
            "NES mode gives you a classic D-pad layout.\n\n"
            "D-PAD (top-left)\n"
            "  ▲ ▼ ◀ ▶ — Move forward, back, left, right.\n\n"
            "LOOK (bottom-left)\n"
            "  ◀ ▶ — Rotate camera left / right.\n"
            "  ▲ ▼ — Look up / down.\n"
            "  Hold a button to keep looking that direction.\n\n"
            "ACTION BUTTONS (right side)\n"
            "  JUMP — Jump. Fires once per press.\n"
            "  GRAB — Hold to grab objects or players.\n"
            "  USE  — Interact with world objects.\n"
            "  MENU — Toggle the Quick Menu.\n"
            "  MUTE — Toggle microphone mute.\n\n"
            "TOGGLE BUTTONS\n"
            "  SIT    — Toggles Seated avatar parameter.\n"
            "  CROUCH — Toggles Crouching avatar parameter.\n"
            "  Active toggles stay highlighted in purple."
        ),
    },
    {
        "title": "Joystick Mode",
        "content": (
            "Joystick mode replaces the D-pad with an analogue\n"
            "stick and sliders.\n\n"
            "ANALOGUE STICK (circle canvas)\n"
            "  Click and drag inside the circle to move.\n"
            "  Snaps back to centre on release.\n"
            "  Movement is proportional — drag further for\n"
            "  faster movement.\n\n"
            "LOOK H / LOOK V SLIDERS\n"
            "  Drag to rotate camera / look up-down.\n"
            "  Returns to centre on release.\n\n"
            "ACTION BUTTONS (right side)\n"
            "  Same as NES mode — JUMP, GRAB, USE, MENU,\n"
            "  MUTE, SIT, CROUCH.\n\n"
            "Useful for smoother, variable-speed movement\n"
            "instead of binary on/off inputs."
        ),
    },
    {
        "title": "Multiple Pads",
        "content": (
            "You can run as many pads as you like at once.\n\n"
            "Each pad is independent and can have its own:\n"
            "  • Custom name (click the name field to edit)\n"
            "  • Host and Port\n"
            "  • NES or Joystick style\n\n"
            "USE CASES\n"
            "  • One pad for movement, another for actions.\n"
            "  • Control two VRChat instances on the same PC\n"
            "    (e.g. one on port 9000, one on port 9001).\n"
            "  • Send OSC to another app on a different port\n"
            "    alongside VRChat.\n\n"
            "REMOVING A PAD\n"
            "  Click the ✕ in the pad's header.\n"
            "  This also disconnects the OSC client cleanly.\n\n"
            "All pad configs are saved to gamepad_config.json\n"
            "and restored on next launch."
        ),
    },
    {
        "title": "OSC Reference",
        "content": (
            "OSC addresses used by this app:\n\n"
            "  /input/Vertical          Float -1.0 to 1.0\n"
            "  /input/Horizontal        Float -1.0 to 1.0\n"
            "  /input/LookHorizontal    Float -1.0 to 1.0\n"
            "  /input/LookVertical      Float -1.0 to 1.0\n"
            "  /input/Jump              Int   0 or 1\n"
            "  /input/Grab              Int   0 or 1\n"
            "  /input/Use               Int   0 or 1\n"
            "  /input/QuickMenuToggleLeft  Int  0 or 1\n"
            "  /input/Voice             Int   0 or 1\n\n"
            "  /avatar/parameters/Seated     Bool\n"
            "  /avatar/parameters/Crouching  Bool\n\n"
            "Axis and button messages are sent on a 50ms loop\n"
            "(20Hz) while the pad is connected. Toggle\n"
            "parameters are sent once on click.\n\n"
            "Default VRChat OSC port: 9000 (incoming to VRChat)"
        ),
    },
]


def open_help(root):
    win = tk.Toplevel(root)
    win.title("OSC-Gamepad Help")
    win.configure(bg=BG)
    win.resizable(True, True)
    root.update_idletasks()
    win.geometry(f"{root.winfo_width()}x{root.winfo_height()}+{root.winfo_x()}+{root.winfo_y()}")

    current = [0]

    hdr = tk.Frame(win, bg=PANEL, pady=10)
    hdr.pack(fill="x")
    title_lbl = tk.Label(hdr, text="", bg=PANEL, fg=ACCENT2, font=(FONT, 12, "bold"))
    title_lbl.pack(side="left", padx=16)
    page_lbl = tk.Label(hdr, text="", bg=PANEL, fg=SUBTEXT, font=(FONT, 8))
    page_lbl.pack(side="right", padx=16)
    tk.Frame(win, bg=BORDER, height=1).pack(fill="x")

    content_panel = tk.Frame(win, bg=PANEL, highlightthickness=1, highlightbackground=BORDER)
    content_panel.pack(padx=20, pady=(14, 0), fill="both", expand=True)
    content_lbl = tk.Label(
        content_panel, text="", bg=PANEL, fg=TEXT,
        justify="left", wraplength=460, anchor="nw", font=(FONT, 10),
    )
    content_lbl.pack(padx=16, pady=14, fill="both", expand=True)

    def show(idx):
        p = HELP_PAGES[idx]
        title_lbl.config(text=p["title"])
        content_lbl.config(text=p["content"])
        page_lbl.config(text=f"{idx + 1} / {len(HELP_PAGES)}")
        prev_btn.config(state="normal" if idx > 0 else "disabled")
        next_btn.config(text="Close" if idx == len(HELP_PAGES) - 1 else "Next →")

    nav = tk.Frame(win, bg=BG)
    nav.pack(fill="x", padx=20, pady=(8, 14))
    nav.columnconfigure(1, weight=1)

    prev_btn = tk.Button(
        nav, text="← Back", bg=PANEL, fg=SUBTEXT, relief="flat", width=10,
        activebackground=BORDER, activeforeground=TEXT, cursor="hand2", font=(FONT, 9, "bold"),
        command=lambda: (current.__setitem__(0, current[0] - 1), show(current[0])),
    )
    prev_btn.grid(row=0, column=0, sticky="w")

    def next_or_close():
        if current[0] < len(HELP_PAGES) - 1:
            current[0] += 1
            show(current[0])
        else:
            win.destroy()

    next_btn = tk.Button(
        nav, text="Next →", bg=ACCENT, fg=BG, relief="flat", width=10,
        activebackground=ACCENT2, activeforeground=BG, cursor="hand2", font=(FONT, 9, "bold"),
        command=next_or_close,
    )
    next_btn.grid(row=0, column=2, sticky="e")

    show(0)
