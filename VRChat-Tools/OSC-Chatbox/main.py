"""
main.py
───────
Entry point for OSC-Chatbox.

Auto-installs any missing pip packages on first run, then launches the UI.
"""

import subprocess
import sys
import os

VERSION = "8.1.0"

# ── Dependency bootstrap ──────────────────────────────────────────────────────

REQUIRED = [
    "pythonosc",
    "psutil",
    "requests",
    "Pillow",
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


# ── Main ──────────────────────────────────────────────────────────────────────

if __name__ == "__main__":
    _ensure_deps()

    # Apply the saved theme BEFORE any ui module is imported, since
    # ui.theme colours are read at import time by every ui module.
    from config import load_config
    from ui import theme
    theme.set_theme(load_config().get("theme_mode", "new"))

    from ui.app import App

    app = App()
    app.run()