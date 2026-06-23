import os
import subprocess
import sys

VERSION = "8.3.1"

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


# ── Main ──────────────────────────────────────────────────────────────────────

if __name__ == "__main__":
    _ensure_deps()

    from config import load_config
    from ui import theme
    theme.set_theme(load_config().get("theme_mode", "rich_purple"))

    from monitors import steamvr, vrchat
    steamvr.start()
    vrchat.start()

    from ui.app import App
    app = App()
    app.run()