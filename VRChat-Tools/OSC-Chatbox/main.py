import os
import subprocess
import sys

VERSION = "8.3.6"

# ── Dependency bootstrap ──────────────────────────────────────────────────────

REQUIRED = [
    "pythonosc",
    "psutil",
    "requests",
    "Pillow",
    "openvr",
]

# Windows-only
if sys.platform == "win32":
    REQUIRED.append("winrt-runtime")
    REQUIRED.append("winrt-Windows.Media.Control")


def _install(pkg: str):
    print(f"[setup] Installing {pkg}...")
    subprocess.check_call(
        [sys.executable, "-m", "pip", "install", "--quiet", pkg],
        stdout=subprocess.DEVNULL,
    )


def _ensure_deps():
    import importlib
    mappings = {
        "pythonosc":                          "pythonosc",
        "psutil":                             "psutil",
        "requests":                           "requests",
        "Pillow":                             "PIL",
        "openvr":                             "openvr",
        "winrt-runtime":                      "winrt",
        "winrt-Windows.Media.Control":        "winrt.windows.media.control",
    }
    for pkg, import_name in mappings.items():
        if pkg not in REQUIRED:
            continue
        try:
            importlib.import_module(import_name)
        except ImportError:
            _install(pkg)


# ── Ensure we can find our own modules ───────────────────────────────────────

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))


# ── LHM startup helpers ───────────────────────────────────────────────────────

def _lhm_exe_path() -> str:
    """Resolve the LHM exe path relative to the VRChat-Tools root."""
    script_dir = os.path.dirname(os.path.abspath(__file__))
    tools_root = os.path.dirname(script_dir)          # VRChat-Tools/
    toolbox_root = os.path.dirname(tools_root)         # same level as VRChat-Tools/
    # Try both: tools root sibling and tools root child
    candidates = [
        os.path.join(tools_root, "LibreHardwareMonitor", "LibreHardwareMonitor.exe"),
        os.path.join(toolbox_root, "VRChat-Tools", "LibreHardwareMonitor", "LibreHardwareMonitor.exe"),
    ]
    for c in candidates:
        if os.path.isfile(c):
            return c
    return candidates[0]  # return primary path even if missing (will error on launch)


def _launch_lhm():
    exe = _lhm_exe_path()
    if not os.path.isfile(exe):
        print(f"[LHM] exe not found at {exe} — skipping launch")
        return
    try:
        if sys.platform == "win32":
            import ctypes
            ret = ctypes.windll.shell32.ShellExecuteW(
                None, "runas", exe, None, os.path.dirname(exe), 1
            )
            if ret <= 32:
                print(f"[LHM] ShellExecuteW returned {ret} (elevation denied or failed)")
            else:
                print(f"[LHM] Launched with admin elevation")
        else:
            subprocess.Popen([exe], cwd=os.path.dirname(exe))
            print(f"[LHM] Launched")
    except Exception as e:
        print(f"[LHM] Launch failed: {e}")


def _show_lhm_prompt(cfg: dict, save_cfg_cb) -> bool:
    """
    Show a popup asking whether to start LHM.
    Returns True if LHM should be launched.
    Saves preference back to config if user picks always/never.
    """
    import tkinter as tk

    result = {"launch": False}

    popup = tk.Tk()
    popup.withdraw()  # hide briefly while we style it

    # Import theme values (theme already set before this call)
    try:
        from ui.theme import BG, PANEL, BORDER, ACCENT, ACCENT2, TEXT, SUBTEXT, FONT
    except Exception:
        BG = "#0f0f13"; PANEL = "#1f102a"; BORDER = "#2a2a38"
        ACCENT = "#9D00FF"; ACCENT2 = "#b44bff"
        TEXT = "#e2e0f0"; SUBTEXT = "#7e7b9a"; FONT = "Consolas"

    popup.title("Libre Hardware Monitor")
    popup.configure(bg=BG)
    popup.resizable(False, False)

    # Header
    hdr = tk.Frame(popup, bg=PANEL, pady=10)
    hdr.pack(fill="x")
    tk.Label(hdr, text="Libre Hardware Monitor", bg=PANEL, fg=ACCENT2,
             font=(FONT, 12, "bold")).pack(side="left", padx=16)
    tk.Frame(popup, bg=BORDER, height=1).pack(fill="x")

    # Body
    body = tk.Frame(popup, bg=BG, padx=24, pady=16)
    body.pack(fill="both", expand=True)
    tk.Label(
        body,
        text="Would you like to start Libre Hardware Monitor?\n"
             "LHM provides GPU & CPU temperature data for the ChatBox.",
        bg=BG, fg=TEXT, font=(FONT, 9), justify="left", wraplength=320,
    ).pack(anchor="w", pady=(0, 16))

    # Buttons
    btn_frame = tk.Frame(body, bg=BG)
    btn_frame.pack(fill="x")

    def _do(choice: str):
        """choice: 'start' | 'always' | 'dismiss' | 'never'"""
        if choice in ("start", "always"):
            result["launch"] = True
        if choice == "always":
            cfg["lhm_prompt"] = "always"
            save_cfg_cb(cfg)
        elif choice == "never":
            cfg["lhm_prompt"] = "never"
            save_cfg_cb(cfg)
        popup.destroy()

    btn_cfg = dict(relief="flat", font=(FONT, 9, "bold"), cursor="hand2", padx=14, pady=6)

    tk.Button(btn_frame, text="▶  Start LHM",
              bg=ACCENT, fg=TEXT,
              activebackground=ACCENT2, activeforeground=TEXT,
              command=lambda: _do("start"), **btn_cfg,
              ).grid(row=0, column=0, padx=(0, 6), pady=4, sticky="ew")

    tk.Button(btn_frame, text="▶  Always Start",
              bg=PANEL, fg=TEXT,
              activebackground=BORDER, activeforeground=TEXT,
              command=lambda: _do("always"), **btn_cfg,
              ).grid(row=0, column=1, padx=6, pady=4, sticky="ew")

    tk.Button(btn_frame, text="✕  Dismiss",
              bg=PANEL, fg=SUBTEXT,
              activebackground=BORDER, activeforeground=TEXT,
              command=lambda: _do("dismiss"), **btn_cfg,
              ).grid(row=1, column=0, padx=(0, 6), pady=4, sticky="ew")

    tk.Button(btn_frame, text="✕  Never Ask Again",
              bg=PANEL, fg=SUBTEXT,
              activebackground=BORDER, activeforeground=TEXT,
              command=lambda: _do("never"), **btn_cfg,
              ).grid(row=1, column=1, padx=6, pady=4, sticky="ew")

    btn_frame.columnconfigure(0, weight=1)
    btn_frame.columnconfigure(1, weight=1)

    # Centre on screen
    popup.update_idletasks()
    w = popup.winfo_reqwidth()
    h = popup.winfo_reqheight()
    sw = popup.winfo_screenwidth()
    sh = popup.winfo_screenheight()
    popup.geometry(f"{w}x{h}+{(sw - w) // 2}+{(sh - h) // 2}")
    popup.deiconify()
    popup.lift()
    popup.attributes("-topmost", True)
    popup.after(100, lambda: popup.attributes("-topmost", False))

    popup.mainloop()
    return result["launch"]


def _handle_lhm_startup(cfg: dict, save_cfg_cb):
    """Check lhm_prompt preference and act accordingly."""
    pref = cfg.get("lhm_prompt", "ask")
    if pref == "always":
        print("[LHM] Auto-starting (always)")
        _launch_lhm()
    elif pref == "never":
        print("[LHM] Skipping (never)")
    else:  # "ask"
        if _show_lhm_prompt(cfg, save_cfg_cb):
            _launch_lhm()


# ── Main ──────────────────────────────────────────────────────────────────────

if __name__ == "__main__":
    _ensure_deps()

    from config import load_config, save_config
    from ui import theme
    cfg = load_config()
    theme.set_theme(cfg.get("theme_mode", "rich_purple"))

    _handle_lhm_startup(cfg, save_config)

    from monitors import steamvr, vrchat
    steamvr.start()
    vrchat.start()

    from ui.app import App
    app = App()
    app.run()