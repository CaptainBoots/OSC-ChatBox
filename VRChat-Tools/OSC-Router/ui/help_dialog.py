"""
ui/help_dialog.py
─────────────────
Paginated help window for OSC-Router.
"""

import tkinter as tk
from ui.theme import BG, PANEL, BORDER, ACCENT, ACCENT2, TEXT, SUBTEXT, FONT


HELP_PAGES = [
    {
        "title": "Getting Started",
        "content": (
            "Welcome to OSC-Router!\n\n"
            "OSC-Router takes OSC messages from multiple\n"
            "sources and merges them into a single output\n"
            "stream — usually VRChat on port 9000.\n\n"
            "Quick start:\n"
            "  1. Add your source apps and their ports\n"
            "  2. Set the Output IP and Port\n"
            "     (defaults: 127.0.0.1 / 9000)\n"
            "  3. Press ▶ Start\n\n"
            "Any OSC app (Chatbox, Face Tracking, etc.)\n"
            "can send to the router instead of VRChat\n"
            "directly, and the router handles the merge."
        ),
    },
    {
        "title": "Sources",
        "content": (
            "Each source is an app or device that sends\n"
            "OSC messages to the router.\n\n"
            "Each source needs:\n"
            "  Name — a friendly label (e.g. 'Chatbox')\n"
            "  Port — the UDP port it listens on\n\n"
            "Defaults:\n"
            "  Chatbox:       port 9011\n"
            "  Face Tracking: port 9012\n\n"
            "Point your OSC apps at those ports instead\n"
            "of VRChat's 9000 port. The router will\n"
            "merge them and forward to the output.\n\n"
            "Click + Add Source to add more.\n"
            "Click ✕ to remove a source."
        ),
    },
    {
        "title": "Output",
        "content": (
            "Output IP — destination IP address.\n"
            "  Same PC:       127.0.0.1 (default)\n"
            "  Different PC:  use that machine's LAN IP\n\n"
            "Output Port — destination port.\n"
            "  VRChat default: 9000\n\n"
            "All merged OSC messages are forwarded\n"
            "to this single IP:Port destination.\n\n"
            "Tip: If VRChat is on this machine,\n"
            "leave both fields at their defaults."
        ),
    },
    {
        "title": "Priority & Conflicts",
        "content": (
            "Sources are listed with a priority number\n"
            "(#1 is highest priority).\n\n"
            "Merge rules:\n"
            "  • Different OSC addresses → all forwarded\n"
            "  • Same address, same value → sent once\n"
            "  • Same address, different value →\n"
            "    highest priority source wins\n\n"
            "Live Conflicts in the stats bar shows how\n"
            "many addresses are currently being contested\n"
            "between sources right now.\n\n"
            "The router runs at 20 Hz (every 50ms) and\n"
            "only forwards values that have changed,\n"
            "so it won't chatter the output."
        ),
    },
    {
        "title": "Live Stats",
        "content": (
            "The stats bar at the top shows:\n\n"
            "Forwarded — total OSC messages routed\n"
            "since the router was last started.\n\n"
            "Conflicts — number of OSC addresses\n"
            "currently being sent by more than one\n"
            "source with different values.\n\n"
            "Sources — how many are active vs total.\n\n"
            "Each source row shows:\n"
            "  ● 1,234 rx  — running, message count\n"
            "  ✗ failed    — could not bind the port\n"
            "                (port already in use)\n\n"
            "Stats update every second."
        ),
    },
]


def open_help(root):
    win = tk.Toplevel(root)
    win.title("OSC-Router Help")
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
