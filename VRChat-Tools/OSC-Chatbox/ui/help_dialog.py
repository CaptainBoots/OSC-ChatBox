"""
ui/help_dialog.py
─────────────────
"""

import tkinter as tk
from ui.theme import BG, PANEL, BORDER, ACCENT, ACCENT2, TEXT, SUBTEXT, FONT


HELP_PAGES = [
    {
        "title": "Getting Started",
        "content": (
            "Welcome to OSC-Chatbox v2!\n\n"
            "OSC-Chatbox sends live system stats, weather, and media\n"
            "info to your VRChat chatbox via OSC.\n\n"
            "Quick start:\n"
            "  1. Set your OSC IP and Port in the Chatbox tab\n"
            "     (defaults: 127.0.0.1 / 9000)\n"
            "  2. Make sure OSC is enabled in VRChat\n"
            "     (Action Menu → Options → OSC → Enable)\n"
            "  3. Press ▶ Start\n\n"
            "The Builder tab lets you fully customise what each\n"
            "page shows and how long it stays on screen."
        ),
    },
    {
        "title": "The Builder Tab",
        "content": (
            "The Builder tab is where you design your pages.\n\n"
            "  All available data modules grouped by category.\n\n"
            "Right panel — Pages:\n"
            "  Each page is a card. Slots inside it are the lines\n"
            "  sent to VRChat in order.\n\n"
            "  ▲ ▼  — reorder slots\n"
            "  ✕    — remove a slot\n"
            "  ⠿    — drag to reorder (grab and move up/down)\n\n"
            "  Custom Text slots have an inline text box so you\n"
            "  can type a fixed line directly in the slot.\n\n"
            "+ Add Page creates a new blank page."
        ),
    },
    {
        "title": "Pages & Duration",
        "content": (
            "Each page has its own Duration (in seconds).\n\n"
            "The chatbox rotates through your enabled pages,\n"
            "spending exactly that many seconds on each one\n"
            "before moving to the next.\n\n"
            "You can set different durations per page — for\n"
            "example, show your hardware stats for 30 seconds\n"
            "but your weather page for only 10 seconds.\n\n"
            "The checkbox on each page header enables or\n"
            "disables that page without deleting it.\n\n"
            "If ALL pages are disabled the chatbox will show\n"
            "'No pages enabled' until you re-enable one."
        ),
    },
    {
        "title": "OSC & Network Config",
        "content": (
            "OSC IP — IP address to send messages to.\n"
            "  • Same PC: 127.0.0.1 (default)\n"
            "  • Different PC on LAN: use that PC's local IP\n\n"
            "OSC Port — VRChat listens on 9000 by default.\n"
            "  Don't change this unless you know you need to.\n\n"
            "Network Interface — The adapter to monitor for\n"
            "upload/download speed modules.\n"
            "  Open Task Manager → Performance tab to find\n"
            "  your adapter name (e.g. Ethernet, Wi-Fi).\n\n"
            "Interval (s) — Default duration for new pages.\n"
            "  Existing pages use their own per-page duration."
        ),
    },
    {
        "title": "LibreHardwareMonitor",
        "content": (
            "CPU/GPU temperature, wattage, and load modules\n"
            "require LibreHardwareMonitor (LHM) to be running.\n\n"
            "Setup:\n"
            "  1. Download LHM from GitHub:\n"
            "     github.com/LibreHardwareMonitor/LibreHardwareMonitor\n"
            "  2. Run LibreHardwareMonitor.exe as Administrator\n"
            "  3. Options → Web Server → Run\n"
            "     (default port 8085)\n\n"
            "LHM URL in the config should be:\n"
            "  http://localhost:8085/data.json\n\n"
            "Without LHM, CPU/GPU stat modules will show N/A\n"
            "but everything else continues working normally."
        ),
    },
    {
        "title": "Weather",
        "content": (
            "Weather modules use the free Open-Meteo API.\n"
            "No API key needed.\n\n"
            "Location — Enter your coordinates as:\n"
            "   latitude,longitude\n\n"
            "Examples:\n"
            "   53.4,-2.2   (Manchester, UK)\n"
            "   51.5,-0.1   (London, UK)\n"
            "   40.7,-74.0  (New York, USA)\n"
            "   35.6,139.7  (Tokyo, Japan)\n\n"
            "To find your coordinates:\n"
            "  Google Maps → right-click your location\n"
            "  The first two numbers are lat,lon.\n\n"
            "Weather refreshes every 5 minutes."
        ),
    },
    {
        "title": "Media Modules",
        "content": (
            "Media modules show your currently playing music.\n\n"
            "On Windows these use the Windows Media Transport\n"
            "Controls (the same system as the taskbar overlay).\n"
            "Any app that reports to Windows media controls\n"
            "will work: Spotify, Chrome, Firefox, VLC, etc.\n\n"
            "Available media modules:\n"
            "  Media Title     — song name (with ▶/⏸ icon)\n"
            "  Media Artist    — artist name\n"
            "  Media Album     — album name\n"
            "  Media Source    — app name (e.g. Spotify)\n"
            "  Media Progress  — visual progress bar\n"
            "  Media Time      — 2:14 / 3:45\n"
            "  Media Detail    — track, time & source combined\n\n"
            "Trim Media Titles (in Settings) removes clutter\n"
            "like '(Official Video)' and '[Lyrics]'."
        ),
    },
    {
        "title": "Forced Text",
        "content": (
            "The Forced Text field on the Chatbox tab lets you\n"
            "override all pages instantly.\n\n"
            "While anything is typed in that field, the chatbox\n"
            "will send only that text — ignoring all pages.\n\n"
            "Leave it blank (or clear it) to return to the\n"
            "normal rotating page display.\n\n"
            "Useful for:\n"
            "  • Sending a quick custom message to the world\n"
            "  • Temporarily hiding your stats\n"
            "  • Testing what a line looks like in VRChat\n\n"
            "VRChat's chatbox limit is 144 characters.\n"
            "Any text longer than that is automatically trimmed."
        ),
    },
    {
        "title": "Settings",
        "content": (
            "Open Settings with the ⚙ Settings button.\n\n"
            "UI Scale — Resize the entire app window.\n"
            "  Drag the slider; changes apply live.\n\n"
            "Progress Bar Characters:\n"
            "  Filled / Border / Empty — the three characters\n"
            "  used to draw the Media Progress bar module.\n"
            "  Defaults: ▓ ▒ ░\n"
            "  Type any character; preview updates live.\n\n"
            "Trim Media Titles — strips words like 'official',\n"
            "  'lyrics', 'video' from song titles.\n\n"
            "Slow Mode — updates every 5 seconds.\n"
            "Speed Mode — updates every 0.1 seconds.\n"
            "  (Both off = 1 second update interval)\n\n"
            "Reset to Defaults restores all settings but\n"
            "keeps your pages and connection config."
        ),
    },
]


def open_help(root):
    win = tk.Toplevel(root)
    win.title("OSC-Chatbox Help")
    win.configure(bg=BG)
    win.resizable(True, True)
    root.update_idletasks()
    win.geometry(f"{root.winfo_width()}x{root.winfo_height()}+{root.winfo_x()}+{root.winfo_y()}")

    current = [0]

    # Header
    hdr = tk.Frame(win, bg=PANEL, pady=10)
    hdr.pack(fill="x")
    title_lbl = tk.Label(hdr, text="", bg=PANEL, fg=ACCENT2, font=(FONT, 12, "bold"))
    title_lbl.pack(side="left", padx=16)
    page_lbl = tk.Label(hdr, text="", bg=PANEL, fg=SUBTEXT, font=(FONT, 8))
    page_lbl.pack(side="right", padx=16)
    tk.Frame(win, bg=BORDER, height=1).pack(fill="x")

    # Content
    content_panel = tk.Frame(win, bg=PANEL, highlightthickness=1, highlightbackground=BORDER)
    content_panel.pack(padx=20, pady=(14, 0), fill="both", expand=True)
    content_lbl = tk.Label(
        content_panel, text="", bg=PANEL, fg=TEXT,
        justify="left", wraplength=500, anchor="nw", font=(FONT, 10),
    )
    content_lbl.pack(padx=16, pady=14, fill="both", expand=True)

    def show(idx):
        p = HELP_PAGES[idx]
        title_lbl.config(text=p["title"])
        content_lbl.config(text=p["content"])
        page_lbl.config(text=f"{idx + 1} / {len(HELP_PAGES)}")
        prev_btn.config(state="normal" if idx > 0 else "disabled")
        next_btn.config(text="Close" if idx == len(HELP_PAGES) - 1 else "Next →")

    # Nav
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
