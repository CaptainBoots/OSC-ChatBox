"""
main.py
───────
Entry point for OSC-Gamepad.
Auto-installs missing packages then launches the UI.
"""

import os
import subprocess
import sys

VERSION = "1.1.2"

REQUIRED = [
    ("python-osc==1.9.3", "pythonosc"),
]


def _install(pkg: str):
    print(f"[setup] Installing {pkg}...")
    attempts = [[sys.executable, "-m", "pip", "install", "--quiet", pkg]]
    if sys.platform != "win32":
        attempts += [
            [sys.executable, "-m", "pip", "install", "--quiet", pkg, "--break-system-packages"],
            [sys.executable, "-m", "pip", "install", "--quiet", pkg, "--user"],
        ]
    for cmd in attempts:
        try:
            subprocess.check_call(cmd, stdout=subprocess.DEVNULL)
            return
        except subprocess.CalledProcessError:
            continue


def _ensure_deps():
    import importlib
    for pkg, import_name in REQUIRED:
        try:
            importlib.import_module(import_name)
        except ImportError:
            _install(pkg)


sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

if __name__ == "__main__":
    _ensure_deps()

    from config import load_config
    from ui import theme
    theme.set_theme(load_config().get("theme_mode", "new"))

    from ui.app import App
    App().run()