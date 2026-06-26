# ═════════════════════════════════════════════════════════════════════════════════════════════════════════════════════#
#                                              Boot's ToolBox Script                                                      #
# ═════════════════════════════════════════════════════════════════════════════════════════════════════════════════════#
# Hi :3
# Welcome to my code

# ═════════════════════════════════════════════════════════════════════════════════════════════════════════════════════#
# Imports
# ═════════════════════════════════════════════════════════════════════════════════════════════════════════════════════#

import importlib
import io
import json
import os
import re
import site
import subprocess
import sys
import time
import tkinter as tk
import tkinter.font as font
import zipfile
import webbrowser
import threading
from tkinter import messagebox


def install_if_missing(package, import_name=None):
    if import_name is None:
        import_name = package.split("==")[0].replace("-", "_")

    try:
        importlib.import_module(import_name)
    except ImportError:
        print(f"Installing {package}...")

        install_attempts = [
            [sys.executable, "-m", "pip", "install", package],
        ]
        if sys.platform != "win32":
            install_attempts.append([sys.executable, "-m", "pip", "install", package, "--break-system-packages"])
            install_attempts.append([sys.executable, "-m", "pip", "install", package, "--user"])

        last_error = None
        for cmd in install_attempts:
            try:
                subprocess.check_call(cmd)
                last_error = None
                break
            except subprocess.CalledProcessError as e:
                last_error = e

        if last_error is not None:
            raise last_error

        if sys.platform != "win32":
            user_site = site.getusersitepackages()
            if user_site and user_site not in sys.path:
                sys.path.insert(0, user_site)


install_if_missing("requests==2.32.5", "requests")

import requests

# ═════════════════════════════════════════════════════════════════════════════════════════════════════════════════════#
# CONFIGURATION & GLOBAL VARIABLES
# ═════════════════════════════════════════════════════════════════════════════════════════════════════════════════════#

_processes = []
VERSION = "9.4.8"

# Default selected branch tracking variable
UPDATE_BRANCH = "main"


def get_github_raw_url():
    return f"https://raw.githubusercontent.com/CaptainBoots/VRChat-ToolBox/{UPDATE_BRANCH}/VRChat-ToolBox.py"


def get_github_base_url():
    return f"https://raw.githubusercontent.com/CaptainBoots/VRChat-ToolBox/{UPDATE_BRANCH}/VRChat-Tools/"


GITHUB_EXE_RELEASE_BASE_URL = "https://github.com/CaptainBoots/VRChat-ToolBox/releases/latest/download/"

if getattr(sys, 'frozen', False):
    SCRIPT_DIR = os.path.dirname(sys.executable)
else:
    SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))

TOOLS_ROOT_DIR = os.path.join(SCRIPT_DIR, "VRChat-Tools")
TOOLBOX_CONFIG_DIR = os.path.join(TOOLS_ROOT_DIR, "VRChat-Toolbox")
TOOLBOX_CONFIG_FILE = os.path.join(TOOLBOX_CONFIG_DIR, "toolbox_config.json")
LEGACY_TOOLBOX_CONFIG_FILES = [
    os.path.join(TOOLS_ROOT_DIR, "osc_config.json"),
    os.path.join(TOOLS_ROOT_DIR, "chatbox_config.json"),
    os.path.join(TOOLS_ROOT_DIR, "toolbox_config.json"),
]

os.makedirs(TOOLBOX_CONFIG_DIR, exist_ok=True)

print(f"[Config] Script directory: {SCRIPT_DIR}")
print(f"[Config] Config directory: {TOOLBOX_CONFIG_DIR}")
print(f"[Config] Config file: {TOOLBOX_CONFIG_FILE}")

if VERSION == "9.4.0":  # upate when adding tools or dependencies
    if os.path.exists(TOOLBOX_CONFIG_FILE):
        try:
            os.remove(TOOLBOX_CONFIG_FILE)
            print(f"[Config] Version config change detected. Forced clean reset of: {TOOLBOX_CONFIG_FILE}")
        except OSError as e:
            print(f"[Config] Failed to force-delete config: {e}")

DEFAULT_MANAGED_SCRIPTS = [
    {"filename": "VRChat-Launcher/main.py", "label": "VRChat Launcher(Beta)"},
    {"filename": "LibreHardwareMonitor/LibreHardwareMonitor.exe", "label": "Libre Hardware Monitor"},
    {"filename": "OSC-Router/main.py", "label": "Router"},
    {"filename": "OSC-Chatbox/main.py", "label": "ChatBox"},
    {"filename": "OSC-Gamepad/main.py", "label": "Gamepad"},
    {"filename": "OSC-FaceTrackingController/main.py", "label": "Face Tracking Controller(Beta)"},
    {"filename": "OSC-ParameterBrowser/main.py", "label": "Parameter Browser(Beta)"},
    {"filename": "OSC-ScriptMaker/main.py", "label": "Script Maker(Placeholder)"},
    {"filename": "VRChat-LocalFavorites/main.py", "label": "VRChat Local Favorites(Placeholder)"},
    {"filename": "VRChat-SocialLogger/main.py", "label": "VRChat SocialLogger(Placeholder)"},
]


def load_managed_scripts():
    global UPDATE_BRANCH
    if os.path.exists(TOOLBOX_CONFIG_FILE):
        try:
            with open(TOOLBOX_CONFIG_FILE, "r", encoding="utf-8") as f:
                config = json.load(f)

            UPDATE_BRANCH = config.get("update_branch", "main")

            # Verify the configuration version matches the current app version
            config_version = config.get("version")
            if config_version == VERSION:
                return config.get("managed_scripts", DEFAULT_MANAGED_SCRIPTS)
            else:
                print(
                    f"[Config] Version mismatch (Config: {config_version}, App: {VERSION}). Wiping and regenerating config...")
        except Exception as e:
            print(f"[Config] Error loading config: {e}")

    save_managed_scripts(DEFAULT_MANAGED_SCRIPTS)
    return DEFAULT_MANAGED_SCRIPTS


def save_managed_scripts(scripts):
    try:
        os.makedirs(TOOLBOX_CONFIG_DIR, exist_ok=True)
        config = {
            "version": VERSION,
            "update_branch": UPDATE_BRANCH,
            "managed_scripts": scripts
        }
        with open(TOOLBOX_CONFIG_FILE, "w", encoding="utf-8") as f:
            json.dump(config, f, indent=2)
        print(f"[Config] Saved {len(scripts)} managed scripts (v{VERSION}) to {TOOLBOX_CONFIG_FILE}")
    except Exception as e:
        print(f"[Config] Error saving config: {e}")
        print(f"[Config] Attempted path: {TOOLBOX_CONFIG_FILE}")


MANAGED_SCRIPTS = load_managed_scripts()

print("Boot's ToolBox")
print("Made By Boots")
print(f"Version {VERSION}")

TOOLS_ROOT_DIR = os.path.join(SCRIPT_DIR, "VRChat-Tools")
TOOLBOX_CONFIG_DIR = os.path.join(TOOLS_ROOT_DIR, "VRChat-Toolbox")
TOOLBOX_CONFIG_FILE = os.path.join(TOOLBOX_CONFIG_DIR, "toolbox_config.json")
BACKUP_DIR = os.path.join(TOOLBOX_CONFIG_DIR, "ToolBox Backup")

SUBFOLDER_SCRIPT_MAP = {
    "VRChat-Launcher/main.py": {
        "remote_path": "VRChat-Launcher/main.py",
        "local_path": os.path.join("VRChat-Launcher", "main.py"),
    },
    "OSC-Router/main.py": {
        "remote_path": "OSC-Router/main.py",
        "local_path": os.path.join("OSC-Router", "main.py"),
    },
    "OSC-Chatbox/main.py": {
        "remote_path": "OSC-Chatbox/main.py",
        "local_path": os.path.join("OSC-Chatbox", "main.py"),
    },
    "OSC-Gamepad/main.py": {
        "remote_path": "OSC-Gamepad/main.py",
        "local_path": os.path.join("OSC-Gamepad", "main.py"),
    },
    "OSC-FaceTrackingController/main.py": {
        "remote_path": "OSC-FaceTrackingController/main.py",
        "local_path": os.path.join("OSC-FaceTrackingController", "main.py"),
    },
    "OSC-ParameterBrowser/main.py": {
        "remote_path": "OSC-ParameterBrowser/main.py",
        "local_path": os.path.join("OSC-ParameterBrowser", "main.py"),
    },
    "OSC-ScriptMaker/main.py": {
        "remote_path": "OSC-ScriptMaker/main.py",
        "local_path": os.path.join("OSC-ScriptMaker", "main.py"),
    },
    "VRChat-LocalFavorites/main.py": {
        "remote_path": "VRChat-LocalFavorites/main.py",
        "local_path": os.path.join("VRChat-LocalFavorites", "main.py"),
    },
    "VRChat-SocialLogger/main.py": {
        "remote_path": "VRChat-SocialLogger/main.py",
        "local_path": os.path.join("VRChat-SocialLogger", "main.py"),
    },
}

# Per-tool config files to wipe on update (paths relative to TOOLS_ROOT_DIR).
# This ensures users always get a clean config after a breaking update.
TOOL_CONFIG_WIPE_MAP: dict[str, list[str]] = {
    "OSC-Chatbox/main.py": [
        os.path.join("OSC-Chatbox", "chatbox_config.json"),
    ],
}

TOOL_DEPENDENCIES_MAP = {
    "VRChat-Launcher/main.py": [
    ],
    "OSC-Router/main.py": [
        # Sibling files in the main folder
        "OSC-Router/__init__.py",
        "OSC-Router/config.py",

        # Core
        "OSC-Router/core/__init__.py",
        "OSC-Router/core/router.py",
        "OSC-Router/core/source.py",

        # UI module
        "OSC-Router/ui/__init__.py",
        "OSC-Router/ui/app.py",
        "OSC-Router/ui/circle_toggle.py",
        "OSC-Router/ui/router_tab.py",
        "OSC-Router/ui/help_dialog.py",
        "OSC-Router/ui/settings_dialog.py",
        "OSC-Router/ui/theme.py",
    ],
    "OSC-Chatbox/main.py": [
        # Sibling files in the main folder
        "OSC-Chatbox/__init__.py",
        "OSC-Chatbox/config.py",
        "OSC-Chatbox/gpu_ids.py",
        "OSC-Chatbox/osc_loop.py",
        "OSC-Chatbox/state.py",

        # Assets
        "OSC-Chatbox/assets/__init__.py",
        "OSC-Chatbox/assets/discord.png",
        "OSC-Chatbox/assets/github.png",

        # Hardware module
        "OSC-Chatbox/hardware/__init__.py",
        "OSC-Chatbox/hardware/cpu.py",
        "OSC-Chatbox/hardware/gpu.py",
        "OSC-Chatbox/hardware/lhm.py",
        "OSC-Chatbox/hardware/memory.py",

        # Modules module
        "OSC-Chatbox/modules/__init__.py",
        "OSC-Chatbox/modules/registry.py",

        # Monitors module
        "OSC-Chatbox/monitors/__init__.py",
        "OSC-Chatbox/monitors/media.py",
        "OSC-Chatbox/monitors/network.py",
        "OSC-Chatbox/monitors/weather.py",
        "OSC-Chatbox/monitors/steamvr.py",
        "OSC-Chatbox/monitors/vrchat.py",

        # UI module
        "OSC-Chatbox/ui/__init__.py",
        "OSC-Chatbox/ui/app.py",
        "OSC-Chatbox/ui/circle_toggle.py",
        "OSC-Chatbox/ui/builder.py",
        "OSC-Chatbox/ui/chatbox_tab.py",
        "OSC-Chatbox/ui/help_dialog.py",
        "OSC-Chatbox/ui/dev_menu.py",
        "OSC-Chatbox/ui/settings_dialog.py",
        "OSC-Chatbox/ui/theme.py",
    ],
    "OSC-Gamepad/main.py": [
        # Sibling files in the main folder
        "OSC-Gamepad/__init__.py",
        "OSC-Gamepad/config.py",

        # Core
        "OSC-Gamepad/core/__init__.py",
        "OSC-Gamepad/core/pad_state.py",

        # UI module
        "OSC-Gamepad/ui/__init__.py",
        "OSC-Gamepad/ui/app.py",
        "OSC-Gamepad/ui/circle_toggle.py",
        "OSC-Gamepad/ui/gamepad_tab.py",
        "OSC-Gamepad/ui/pad_card.py",
        "OSC-Gamepad/ui/help_dialog.py",
        "OSC-Gamepad/ui/settings_dialog.py",
        "OSC-Gamepad/ui/theme.py",
        "OSC-Gamepad/ui/widgets.py",
    ],
    "OSC-FaceTrackingController/main.py": [
    ],
    "OSC-ParameterBrowser/main.py": [
    ],
    "OSC-ScriptMaker/main.py": [
    ],
    "VRChat-LocalFavorites/main.py": [
    ],
    "VRChat-SocialLogger/main.py": [
    ],

}

# ─── Libre Hardware Monitor (EXE tool, downloaded from GitHub Releases) ───────
LHM_FOLDER = "LibreHardwareMonitor"
LHM_EXE_NAME = "LibreHardwareMonitor.exe"
LHM_FILENAME = f"{LHM_FOLDER}/{LHM_EXE_NAME}"
LHM_RELEASE_URL = "https://github.com/LibreHardwareMonitor/LibreHardwareMonitor/releases/latest/download/LibreHardwareMonitor.zip"


def _lhm_exe_path() -> str:
    return os.path.join(TOOLS_ROOT_DIR, LHM_FOLDER, LHM_EXE_NAME)


def ensure_lhm(show_errors: bool = False) -> bool:
    """Download and extract the full LibreHardwareMonitor package if not already present."""
    dest = _lhm_exe_path()
    lhm_dir = os.path.dirname(dest)
    if os.path.isfile(dest):
        return True

    os.makedirs(lhm_dir, exist_ok=True)
    print(f"[LHM] Downloading from {LHM_RELEASE_URL} ...")
    try:
        resp = requests.get(LHM_RELEASE_URL, timeout=60)
        resp.raise_for_status()
        zdata = io.BytesIO(resp.content)
        with zipfile.ZipFile(zdata) as zf:
            members = zf.namelist()

            # Detect whether the ZIP has a single top-level subfolder (common GitHub pattern)
            top_dirs = {m.split("/")[0] for m in members if "/" in m}
            single_root = (
                    len(top_dirs) == 1 and
                    all(m.startswith(next(iter(top_dirs)) + "/") or m == next(iter(top_dirs)) + "/"
                        for m in members)
            )
            strip_prefix = (next(iter(top_dirs)) + "/") if single_root else ""

            exe_members = [m for m in members if m.endswith(LHM_EXE_NAME)]
            if not exe_members:
                raise FileNotFoundError(f"{LHM_EXE_NAME} not found in release ZIP")

            # Extract everything (exe + all DLLs and supporting files) into lhm_dir
            for member in members:
                if member.endswith("/"):
                    continue
                rel_path = member[len(strip_prefix):] if strip_prefix and member.startswith(strip_prefix) else member
                out_path = os.path.join(lhm_dir, rel_path.replace("/", os.sep))
                os.makedirs(os.path.dirname(out_path), exist_ok=True)
                with zipfile.ZipFile(zdata).open(member) as src, open(out_path, "wb") as dst:
                    dst.write(src.read())
                print(f"[LHM] Extracted: {rel_path}")

        print(f"[LHM] All files extracted to {lhm_dir}")
        return True
    except Exception as e:
        print(f"[LHM] Download failed: {e}")
        if show_errors:
            messagebox.showerror(
                "Libre Hardware Monitor",
                f"Could not download LibreHardwareMonitor.\n\nCheck your internet connection and try again.\n\nDetails:\n{e}"
            )
        return False


def _patch_lhm_config() -> None:
    """
    Ensure the LHM .config file has the required keys set before launch.
    Sets:
      runWebServerMenuItem = true   (enables the web API on port 8085)
      startMinMenuItem     = true   (starts minimised to tray)
    Creates the config from scratch if it doesn't exist yet.
    """
    import xml.etree.ElementTree as ET

    lhm_dir = os.path.dirname(_lhm_exe_path())
    cfg_path = os.path.join(lhm_dir, "LibreHardwareMonitor.config")

    REQUIRED = {
        "runWebServerMenuItem": "true",
        "startMinMenuItem": "true",
    }

    # ── Build / load the XML tree ─────────────────────────────────────────────
    if os.path.isfile(cfg_path):
        try:
            tree = ET.parse(cfg_path)
            root = tree.getroot()
        except ET.ParseError as e:
            print(f"[LHM] Config parse error ({e}), will recreate.")
            root = ET.Element("configuration")
            tree = ET.ElementTree(root)
    else:
        print("[LHM] No config found, creating one.")
        root = ET.Element("configuration")
        tree = ET.ElementTree(root)

    # ── Find or create <appSettings> ─────────────────────────────────────────
    app_settings = root.find("appSettings")
    if app_settings is None:
        app_settings = ET.SubElement(root, "appSettings")

    # ── Update / insert each required key ────────────────────────────────────
    for key, value in REQUIRED.items():
        node = app_settings.find(f"./add[@key='{key}']")
        if node is not None:
            if node.get("value") != value:
                print(f"[LHM] Config: setting {key} = {value} (was {node.get('value')})")
                node.set("value", value)
        else:
            print(f"[LHM] Config: inserting {key} = {value}")
            ET.SubElement(app_settings, "add", key=key, value=value)

    # ── Write back ────────────────────────────────────────────────────────────
    try:
        tree.write(cfg_path, encoding="utf-8", xml_declaration=True)
        print(f"[LHM] Config written to {cfg_path}")
    except Exception as e:
        print(f"[LHM] Could not write config: {e}")


def _show_lhm_started_popup() -> None:
    """Small non-blocking confirmation that LHM launched successfully."""

    def _popup():
        win = tk.Toplevel(root)
        win.title("Libre Hardware Monitor")
        win.configure(bg=BG)
        win.resizable(False, False)
        win.attributes("-topmost", True)

        # Import theme colours (already loaded at this point)
        try:
            from ui.theme import PANEL, BORDER, ACCENT2, TEXT, SUBTEXT, FONT as _FONT
        except Exception:
            PANEL = "#1f102a"
            BORDER = "#2a2a38"
            ACCENT2 = "#b44bff"
            TEXT = "#e2e0f0"
            SUBTEXT = "#7e7b9a"
            _FONT = "Consolas"

        hdr = tk.Frame(win, bg=PANEL, pady=8)
        hdr.pack(fill="x")
        tk.Label(hdr, text="Libre Hardware Monitor", bg=PANEL, fg=ACCENT2,
                 font=(_FONT, 11, "bold")).pack(side="left", padx=14)
        tk.Frame(win, bg=BORDER, height=1).pack(fill="x")

        body = tk.Frame(win, bg=BG, padx=20, pady=14)
        body.pack()
        tk.Label(body,
                 text="✓  LHM started successfully.\n\nIt will appear in your system tray shortly.\nThe UAC prompt may have appeared behind this window.",
                 bg=BG, fg=TEXT, font=(_FONT, 9), justify="left").pack()

        tk.Button(
            body, text="OK", bg=ACCENT2, fg=BG, relief="flat",
            font=(_FONT, 9, "bold"), padx=16, pady=4, cursor="hand2",
            activebackground=ACCENT2, activeforeground=BG,
            command=win.destroy,
        ).pack(pady=(10, 0))

        # Centre on screen
        win.update_idletasks()
        sw = win.winfo_screenwidth()
        sh = win.winfo_screenheight()
        w = win.winfo_reqwidth()
        h = win.winfo_reqheight()
        win.geometry(f"{w}x{h}+{(sw - w) // 2}+{(sh - h) // 2}")
        win.after(100, lambda: win.attributes("-topmost", False))

    # Schedule on the main tkinter thread
    root.after(0, _popup)


def launch_lhm() -> None:
    """Patch the LHM config, launch the exe with admin elevation, confirm success."""
    footer_label.config(text="Starting up Libre Hardware Monitor...")
    root.update_idletasks()

    if not ensure_lhm(show_errors=True):
        footer_label.config(text="Error preparing Libre Hardware Monitor")
        return

    # Patch config before every launch so the settings are always correct
    _patch_lhm_config()

    dest = _lhm_exe_path()
    try:
        if sys.platform == "win32":
            import ctypes
            ret = ctypes.windll.shell32.ShellExecuteW(
                None, "runas", dest, None, os.path.dirname(dest), 1
            )
            if ret <= 32:
                raise OSError(f"ShellExecuteW returned {ret} (elevation may have been denied)")
            print(f"[LHM] Launched with admin elevation via ShellExecuteW")
        else:
            p = subprocess.Popen([dest], cwd=os.path.dirname(dest))
            _processes.append(p)
            print(f"[LHM] Launched (PID: {p.pid})")

        _show_lhm_started_popup()
        footer_label.config(text="Ready")
    except Exception as e:
        print(f"[LHM] Launch failed: {e}")
        footer_label.config(text="Error launching Libre Hardware Monitor")
        messagebox.showerror("Launch Error", f"Failed to start LibreHardwareMonitor.\n\nDetails:\n{e}")


# ──────────────────────────────────────────────────────────────────────────────

SCRIPT_FOLDER_MAP = {
    "OSC-Router.py": "OSC-Router",
    "OSC-FaceTrackingController.py": "OSC-FaceTrackingController",
    "OSC-Gamepad.py": "OSC-Gamepad",
    "OSC-ScriptMaker.py": "OSC-ScriptMaker",
    "OSC-ParameterBrowser.py": "OSC-ParameterBrowser",
    "VRChat-Launcher.py": "VRChat-Launcher",
    "VRChat-LocalFavorites.py": "VRChat-LocalFavorites",
    "VRChat-SocialLogger.py": "VRChat-SocialLogger",
}


# ═════════════════════════════════════════════════════════════════════════════════════════════════════════════════════#
# MANAGED SCRIPT HELPERS
# ═════════════════════════════════════════════════════════════════════════════════════════════════════════════════════#

def _ensure_layout_dirs() -> None:
    os.makedirs(TOOLS_ROOT_DIR, exist_ok=True)
    os.makedirs(TOOLBOX_CONFIG_DIR, exist_ok=True)
    os.makedirs(BACKUP_DIR, exist_ok=True)
    for folder in SCRIPT_FOLDER_MAP.values():
        os.makedirs(os.path.join(TOOLS_ROOT_DIR, folder), exist_ok=True)


def _migrate_legacy_layout() -> None:
    for script_name, folder in SCRIPT_FOLDER_MAP.items():
        legacy_script = os.path.join(TOOLS_ROOT_DIR, script_name)
        target_script = os.path.join(TOOLS_ROOT_DIR, folder, script_name)
        if not os.path.isfile(legacy_script) or os.path.isfile(target_script):
            continue
        try:
            os.makedirs(os.path.dirname(target_script), exist_ok=True)
            os.replace(legacy_script, target_script)
            print(f"[Layout] Moved {script_name} -> {folder}\\")
        except OSError as e:
            print(f"[Layout] Could not move {script_name}: {e}")

    legacy_backup_dir = os.path.join(TOOLS_ROOT_DIR, "ToolBox Backup")
    if os.path.isdir(legacy_backup_dir) and os.path.abspath(legacy_backup_dir) != os.path.abspath(BACKUP_DIR):
        try:
            os.makedirs(BACKUP_DIR, exist_ok=True)
            for backup_name in os.listdir(legacy_backup_dir):
                src = os.path.join(legacy_backup_dir, backup_name)
                dst = os.path.join(BACKUP_DIR, backup_name)
                if os.path.isfile(src) and not os.path.exists(dst):
                    os.replace(src, dst)
            if not os.listdir(legacy_backup_dir):
                os.rmdir(legacy_backup_dir)
        except OSError as e:
            print(f"[Layout] Could not migrate legacy backups: {e}")


def _is_path_like(filename: str) -> bool:
    """True only for absolute paths or paths with explicit OS separators that aren't in our known maps."""
    if filename in SUBFOLDER_SCRIPT_MAP:
        return False
    return os.path.isabs(filename) or ("/" in filename) or ("\\" in filename)


def _script_remote_urls(filename: str) -> list[str]:
    if filename in SUBFOLDER_SCRIPT_MAP:
        remote_path = SUBFOLDER_SCRIPT_MAP[filename]["remote_path"]
        return [f"{get_github_base_url()}{remote_path}"]
    if _is_path_like(filename):
        return []
    script_name = os.path.basename(filename)
    urls: list[str] = []
    folder = SCRIPT_FOLDER_MAP.get(script_name)
    if folder:
        urls.append(f"{get_github_base_url()}{folder}/{script_name}")
    urls.append(f"{get_github_base_url()}{script_name}")
    return list(dict.fromkeys(urls))


def _script_bundle_candidates(filename: str) -> list[str]:
    if filename in SUBFOLDER_SCRIPT_MAP:
        local_path = SUBFOLDER_SCRIPT_MAP[filename]["local_path"]
        return [os.path.join(TOOLS_ROOT_DIR, local_path)]
    if _is_path_like(filename):
        resolved = filename if os.path.isabs(filename) else os.path.join(SCRIPT_DIR, filename)
        return [os.path.normpath(resolved)]
    script_name = os.path.basename(filename)
    candidates: list[str] = []
    folder = SCRIPT_FOLDER_MAP.get(script_name)
    if folder:
        candidates.append(os.path.join(SCRIPT_DIR, "VRChat-Tools", folder, script_name))
    candidates.append(os.path.join(SCRIPT_DIR, "VRChat-Tools", script_name))
    candidates.append(os.path.join(SCRIPT_DIR, script_name))
    return list(dict.fromkeys(candidates))


def launch_script(filename: str) -> None:
    """Ensures the target script exists/is updated, then launches it in a separate process."""
    # Route LHM to its dedicated launcher
    if filename == LHM_FILENAME:
        launch_lhm()
        return

    footer_label.config(text=f"Starting up {filename}...")
    root.update_idletasks()

    # 1. Make sure the script and its dependencies exist locally
    if not ensure_script(filename, show_errors=True):
        footer_label.config(text="Error preparing script")
        return

    # 2. Resolve local execution paths
    _, dest_path = _script_paths(filename)
    script_dir = os.path.dirname(dest_path)

    try:
        # 3. Launch script via current Python interpreter in a detached environment
        # Uses sys.executable to ensure it runs on the exact same Python env (like virtual envs)
        p = subprocess.Popen(
            [sys.executable, os.path.basename(dest_path)],
            cwd=script_dir,
            creationflags=subprocess.CREATE_NEW_CONSOLE if sys.platform == "win32" else 0
        )
        _processes.append(p)  # Keep track of the process handle

        print(f"[Launcher] Successfully started {filename} (PID: {p.pid})")
        footer_label.config(text="Ready")

    except Exception as e:
        print(f"[Launcher] Failed to execute {filename}: {e}")
        footer_label.config(text="Error launching script")
        messagebox.showerror(
            "Launch Error",
            f"Failed to start {filename}.\n\nTechnical details:\n{e}"
        )


def _script_paths(filename: str) -> tuple[str, str]:
    if filename in SUBFOLDER_SCRIPT_MAP:
        local_path = SUBFOLDER_SCRIPT_MAP[filename]["local_path"]
        dest_path = os.path.join(TOOLS_ROOT_DIR, local_path)
        return dest_path, dest_path
    if _is_path_like(filename):
        dest_path = filename if os.path.isabs(filename) else os.path.join(SCRIPT_DIR, filename)
        dest_path = os.path.normpath(dest_path)
        return dest_path, dest_path
    script_name = os.path.basename(filename)
    folder = SCRIPT_FOLDER_MAP.get(script_name)
    if folder:
        dest_path = os.path.join(TOOLS_ROOT_DIR, folder, script_name)
    else:
        dest_path = os.path.join(TOOLS_ROOT_DIR, script_name)
    bundle_candidates = _script_bundle_candidates(script_name)
    bundled_path = next((p for p in bundle_candidates if os.path.isfile(p)), bundle_candidates[0])
    return bundled_path, dest_path


_ensure_layout_dirs()
_migrate_legacy_layout()


def ensure_script(filename: str, show_errors: bool = False) -> bool:
    # LHM is an exe tool — handled by ensure_lhm, not this function
    if filename == LHM_FILENAME:
        return ensure_lhm(show_errors=show_errors)

    bundled_path, dest_path = _script_paths(filename)
    dependencies = TOOL_DEPENDENCIES_MAP.get(filename, [])

    # Check if the main file and all of its dependency files exist locally
    main_file_exists = os.path.isfile(dest_path)
    deps_exist = all(os.path.isfile(os.path.join(TOOLS_ROOT_DIR, dep.replace("/", os.sep))) for dep in dependencies)

    if main_file_exists and deps_exist:
        return True

    try:
        os.makedirs(os.path.dirname(dest_path), exist_ok=True)
    except OSError as e:
        print(f"[{filename}] Could not create destination directory: {e}")
        if show_errors:
            messagebox.showerror(f"{filename} Error", f"Could not create target layout directory:\n{e}")
        return False

    # Try downloading from remote repository first
    download_success = False
    for url in _script_remote_urls(filename):
        print(f"[{filename}] Target asset broken or missing. Pulling from Github: {url}")
        try:
            resp = requests.get(
                url, timeout=15, params={"_": int(time.time())},
                headers={"Cache-Control": "no-cache", "Pragma": "no-cache"}
            )
            resp.raise_for_status()

            with open(dest_path, "w", encoding="utf-8") as lf:
                lf.write(resp.text)

            # Download sub-dependencies concurrently if any are registered
            for dep in dependencies:
                dep_url = f"{get_github_base_url()}{dep}"
                dep_dest = os.path.join(TOOLS_ROOT_DIR, dep.replace("/", os.sep))
                os.makedirs(os.path.dirname(dep_dest), exist_ok=True)

                dep_resp = requests.get(
                    dep_url, timeout=15, params={"_": int(time.time())},
                    headers={"Cache-Control": "no-cache", "Pragma": "no-cache"}
                )
                dep_resp.raise_for_status()
                with open(dep_dest, "w", encoding="utf-8") as df:
                    df.write(dep_resp.text)
                print(f"[{filename}] Dependency added: {dep}")

            download_success = True
            print(f"[{filename}] Application code stack prepared successfully.")
            break
        except Exception as e:
            print(f"[{filename}] Remote connection vector faulted: {e}")

    if download_success:
        return True

    # Check for local bundled fallback assets
    if os.path.isfile(bundled_path) and bundled_path != dest_path:
        print(f"[{filename}] Defaulting to local bundle mirror fallback asset.")
        try:
            with open(bundled_path, "r", encoding="utf-8") as sf, open(dest_path, "w", encoding="utf-8") as df:
                df.write(sf.read())
            return True
        except OSError as e:
            print(f"[{filename}] Could not copy bundled fallback: {e}")
            if show_errors:
                messagebox.showerror(
                    f"{filename} Error",
                    f"Could not prepare {filename}.\nCheck your internet connection and try again.",
                )
            return False
    else:
        if show_errors:
            messagebox.showerror(
                f"{filename} Error",
                f"Could not prepare {filename}.\nCheck your internet connection and try again.",
            )
        return False


def check_for_script_updates(filename: str, silent: bool = False) -> bool:
    # LHM is an exe tool — no update check, just ensure it's present
    if filename == LHM_FILENAME:
        ensure_lhm(show_errors=not silent)
        return False

    if not ensure_script(filename, show_errors=not silent):
        return False

    _, dest_path = _script_paths(filename)
    remote_text = None
    remote_version = None

    for url in _script_remote_urls(filename):
        remote_text, remote_version, _ = _fetch_remote_script(url, timeout=10)
        if remote_text is not None:
            break

    if remote_text is None:
        if not silent:
            messagebox.showinfo(
                f"{filename} Update",
                f"Could not reach GitHub to check updates for {filename}."
            )
        return False

    remote_version = remote_version or "0.0.0"
    try:
        with open(dest_path, "r", encoding="utf-8") as lf:
            local_text = lf.read()
    except OSError:
        local_text = ""

    local_version = _extract_version_from_source(local_text) or "0.0.0"
    print(f"[{filename}] Checking... (local: {local_version} remote: {remote_version})")

    if _parse_version(remote_version) <= _parse_version(local_version):
        print(f"[{filename}] Up to date ({local_version})")
        return False

    try:
        with open(dest_path, "w", encoding="utf-8") as lf:
            lf.write(remote_text)
        print(f"[{filename}] Updated: {local_version} -> {remote_version}")

        # Update dependency submodules as well
        dependencies = TOOL_DEPENDENCIES_MAP.get(filename, [])
        for dep in dependencies:
            dep_url = f"{get_github_base_url()}{dep}"
            dep_dest = os.path.join(TOOLS_ROOT_DIR, dep.replace("/", os.sep))
            os.makedirs(os.path.dirname(dep_dest), exist_ok=True)
            try:
                dep_resp = requests.get(
                    dep_url, timeout=10, params={"_": int(time.time())},
                    headers={"Cache-Control": "no-cache", "Pragma": "no-cache"},
                )
                dep_resp.raise_for_status()
                with open(dep_dest, "w", encoding="utf-8") as df:
                    df.write(dep_resp.text)
            except Exception as e:
                print(f"[{filename}] Dependency update error ({dep}): {e}")

        return True
    except OSError as e:
        print(f"[{filename}] Failed writing updated local source tree context: {e}")
        return False


def _parse_version(v_str: str) -> tuple[int, ...]:
    try:
        return tuple(map(int, (v_str.split("."))))
    except ValueError:
        return (0, 0, 0)


def _extract_version_from_source(source_text: str) -> str | None:
    for line in source_text.splitlines():
        if line.strip().startswith("VERSION"):
            match = re.search(r'["\']([^"\']+)["\']', line)
            if match:
                return match.group(1)
    return None


def _fetch_remote_script(url: str, timeout: int = 10) -> tuple[str | None, str | None, str | None]:
    try:
        resp = requests.get(
            url, timeout=timeout, params={"_": int(time.time())},
            headers={"Cache-Control": "no-cache", "Pragma": "no-cache"},
        )
        resp.raise_for_status()
        return resp.text, _extract_version_from_source(resp.text), url
    except requests.RequestException:
        return None, None, None


def get_remote_script_info() -> dict[str, str] | None:
    urls = [get_github_raw_url()]
    errors = []
    best: dict[str, str] | None = None

    for url in urls:
        text, remote_version, used_url = _fetch_remote_script(url, timeout=10)
        if text is None:
            errors.append(url)
            continue

        info: dict[str, str] = {
            "text": text,
            "version": remote_version or "0.0.0",
            "url": used_url or url,
        }

        if best is None:
            best = info
            continue

        best_version: str = best["version"] if best else "0.0.0"
        if _parse_version(info["version"]) > _parse_version(best_version):
            best = info

    if best is not None:
        return best

    print(f"[Updater] Could not reach GitHub URLs: {errors}")
    return None


def perform_update(remote_text=None, source_url=None):
    try:
        if getattr(sys, 'frozen', False):
            messagebox.showinfo(
                "Update Disabled",
                "Automatic EXE self-update has been disabled. Please update the executable manually from GitHub."
            )
            return

        if remote_text is None:
            info = get_remote_script_info()
            if not info:
                raise RuntimeError("No remote script source available")
            remote_text = info["text"]
            source_url = info["url"]

        script_path = os.path.abspath(__file__)
        os.makedirs(BACKUP_DIR, exist_ok=True)
        script_name = os.path.splitext(os.path.basename(script_path))[0]
        backup_path = os.path.join(BACKUP_DIR, f"{script_name} {VERSION}.bak")

        with open(script_path, "r", encoding="utf-8") as f_src:
            current_code = f_src.read()
        with open(backup_path, "w", encoding="utf-8") as f_dst:
            f_dst.write(current_code)
        print(f"[Updater] Created rollback backup pointing at: {backup_path}")

        with open(script_path, "w", encoding="utf-8") as f_upper:
            f_upper.write(remote_text)
        print(f"[Updater] Main system assembly updated successfully from {source_url}.")

        # Wipe targeted configurations on version shift
        for tool_key, configs_to_wipe in TOOL_CONFIG_WIPE_MAP.items():
            for relative_cfg in configs_to_wipe:
                full_cfg_path = os.path.join(TOOLS_ROOT_DIR, relative_cfg)
                if os.path.exists(full_cfg_path):
                    try:
                        os.remove(full_cfg_path)
                        print(f"[Updater] Wiped breaking config layout targets: {relative_cfg}")
                    except Exception as ex:
                        print(f"[Updater] Error cleaning target configuration profile: {ex}")

        messagebox.showinfo(
            "Update Complete",
            f"ToolBox updated to the latest available software build on branch '{UPDATE_BRANCH}'.\n\nThe system will now restart automatically."
        )

        root.destroy()
        subprocess.Popen([sys.executable, script_path], cwd=os.path.dirname(script_path))
        sys.exit(0)

    except Exception as e:
        print(f"[Updater] Self-update failed catastrophically: {e}")
        messagebox.showerror(
            "Update Failed",
            f"An error occurred during updating processing:\n\n{e}"
        )


def check_for_main_updates(silent: bool = True):
    if not silent:
        footer_label.config(text="Connecting to repository update server nodes...")
        root.update_idletasks()

    info = get_remote_script_info()
    if not info:
        if not silent:
            messagebox.showerror("Update Connection Fault", "Unable to pull validation info records from GitHub.")
        footer_label.config(text="Ready")
        return

    remote_text = info["text"]
    remote_version = info["version"]
    remote_url = info["url"]

    try:
        with open(__file__, "r", encoding="utf-8", errors="ignore") as f:
            local_text = f.read()
    except Exception:
        local_text = ""

    local_norm = local_text.replace("\r\n", "\n")
    remote_norm = remote_text.replace("\r\n", "\n")

    remote_newer = _parse_version(remote_version) > _parse_version(VERSION)
    content_differs = remote_norm != local_norm
    main_update_available = remote_newer or content_differs

    print(f"[VRChat-Tools] Checking... (local: {VERSION} remote: {remote_version} Branch: {UPDATE_BRANCH})")

    if main_update_available:
        if remote_newer:
            print(f"[VRChat-Tools] Update available: {VERSION} -> {remote_version}")
            if os.path.exists(TOOLBOX_CONFIG_FILE):
                try:
                    os.remove(TOOLBOX_CONFIG_FILE)
                except Exception:
                    pass
            prompt = (
                f"New version {remote_version} is available on branch '{UPDATE_BRANCH}' (you have {VERSION}).\n\n"
                "Update and restart now?"
            )
        else:
            print(f"[VRChat-Tools] Remote content differs (version string unchanged at {VERSION})")
            prompt = (
                f"A remote script update is available on branch '{UPDATE_BRANCH}' (content changed,\n"
                "but version string may not have been bumped).\n\n"
                "Update and restart now?"
            )

        if messagebox.askyesno("Update Available", prompt):
            perform_update(remote_text=remote_text, source_url=remote_url)
        else:
            print(f"[VRChat-Tools] Update skipped by user")
    else:
        print(f"[VRChat-Tools] Up to date ({VERSION})")

    any_tool_updated = False
    for script_entry in MANAGED_SCRIPTS:
        if check_for_script_updates(script_entry["filename"], silent=silent):
            any_tool_updated = True

    if not silent and not main_update_available and not any_tool_updated:
        messagebox.showinfo(
            "Up to Date", f"You're on the latest version ({VERSION}) for branch '{UPDATE_BRANCH}'."
        )


def force_update_all_scripts():
    """Wipes the cached/downloaded files and force-downloads them from the current branch live."""

    def _update_task():
        footer_label.config(text=f"Switching branch to '{UPDATE_BRANCH}' & updating...")
        root.update_idletasks()

        success = True
        # Wiping local script cache and force-downloading afresh from the newly assigned branch context
        for script in MANAGED_SCRIPTS:
            filename = script["filename"]
            if filename == LHM_FILENAME:
                continue

            bundled_path, dest_path = _script_paths(filename)
            if os.path.isfile(dest_path):
                try:
                    os.remove(dest_path)
                except Exception:
                    pass

            dependencies = TOOL_DEPENDENCIES_MAP.get(filename, [])
            for dep in dependencies:
                dep_path = os.path.join(TOOLS_ROOT_DIR, dep.replace("/", os.sep))
                if os.path.isfile(dep_path):
                    try:
                        os.remove(dep_path)
                    except Exception:
                        pass

            if not ensure_script(filename, show_errors=False):
                success = False

        if success:
            footer_label.config(text=f"Successfully switched to branch '{UPDATE_BRANCH}'!")
            messagebox.showinfo(
                "Branch Updated",
                f"All scripts have been successfully updated to match the '{UPDATE_BRANCH}' branch structure.",
            )
        else:
            footer_label.config(text="Error updating branch assets")
            messagebox.showerror(
                "Branch Update Error",
                "Failed to fully re-download some script assets from the chosen branch.",
            )

    threading.Thread(target=_update_task, daemon=True).start()


# ═════════════════════════════════════════════════════════════════════════════════════════════════════════════════════#
# GUI
# ═════════════════════════════════════════════════════════════════════════════════════════════════════════════════════#

FONT = "Consolas"
TITLE_PREFIX = "◈"  # new (default)
BG = "#0f0f13"
PANEL = "#1f102a"
BORDER = "#2a2a38"
ACCENT = "#9D00FF"
ACCENT2 = "#b44bff"
TEXT = "#e2e0f0"
TEXT2 = "#ffffff"
SUBTEXT = "#7e7b9a"
GREEN = "#00ffcc"
RED = "#ff4b72"

root = tk.Tk()
root.title("VRChat-ToolBox")
root.configure(bg=BG)
root.geometry("580x600")
root.minsize(0, 0)

# Window Setup Core Sizing Constraints
root.update_idletasks()

# Context DPI Engine Scaling Configurations
scale = 1.0
scalable_widgets = []
square_widgets = []


def register_scalable(widget, base_size, extras=()):
    scalable_widgets.append((widget, base_size, extras))


def apply_ui_scaling(new_scale):
    global scale
    scale = new_scale
    for widget, base_size, extras in scalable_widgets:
        try:
            widget.configure(font=(FONT, max(6, int(base_size * scale))) + extras)
        except tk.TclError:
            pass
    for container, base_size, btn_widget in square_widgets:
        size = int(base_size * scale)
        container.config(width=size, height=size)
        btn_widget.config(font=(FONT, max(8, int(12 * scale))))


# Factory Function: Standard UI Subtext Label
def dark_label(text, r, **kwargs):
    lbl = tk.Label(main_frame, text=text, bg=BG, fg=SUBTEXT, anchor="w", font=(FONT, 9))
    lbl.grid(row=r, column=0, sticky="w", pady=6, **kwargs)
    return lbl


# Factory Function: Standard Form Input Field
def dark_entry(r, default=""):
    e = tk.Entry(
        main_frame, bg=PANEL, fg=TEXT, insertbackground=ACCENT, relief="flat", font=(FONT, 9),
        highlightthickness=1, highlightbackground=BORDER, highlightcolor=ACCENT,
    )
    e.insert(0, default)
    e.grid(row=r, column=1, pady=6, sticky="ew", padx=(8, 0))
    return e


# Factory Function: Fixed Aspect-Ratio Square Button
def square_button(parent, text, command, base_size=32):
    container = tk.Frame(parent, bg=PANEL, highlightthickness=1, highlightbackground=BORDER)
    container.pack_propagate(False)
    button_widget = tk.Button(
        container, text=text, command=command, bg=PANEL, fg=SUBTEXT, relief="flat", borderwidth=0, font=(FONT, 12),
        activebackground=BORDER, activeforeground=TEXT, cursor="hand2",
    )
    button_widget.pack(fill="both", expand=True)
    square_widgets.append((container, base_size, button_widget))
    container.config(width=base_size, height=base_size)
    return container


def _show_beta_popup():
    """Elegant non-blocking modal modal overlay thanking beta testers and promoting the discord server."""
    win = tk.Toplevel(root)
    win.title("Beta Branch Active")
    win.geometry("420x260")
    win.configure(bg=BG)
    win.resizable(False, False)
    win.transient(root)
    win.grab_set()

    hdr = tk.Frame(win, bg=PANEL, pady=10)
    hdr.pack(fill="x")
    tk.Label(hdr, text="✨ Beta Testing Active", bg=PANEL, fg=ACCENT2, font=(FONT, 12, "bold")).pack(side="left", padx=16)
    tk.Frame(win, bg=BORDER, height=1).pack(fill="x")

    body = tk.Frame(win, bg=BG, padx=20, pady=16)
    body.pack(fill="both", expand=True)

    msg = (
        "Thank you for participating in our beta test program!\n\n"
        "Your bug reports and active testing cycles help optimize these tools for everyone.\n\n"
        "Join our active developer hub to report issues, suggest modifications, or track roadmap updates live!"
    )
    tk.Label(body, text=msg, bg=BG, fg=TEXT, font=(FONT, 9), justify="left", wraplength=380).pack(anchor="w")

    btn_row = tk.Frame(body, bg=BG)
    btn_row.pack(fill="x", pady=(18, 0))

    tk.Button(
        btn_row, text="Join Discord Server", bg=ACCENT, fg=TEXT2, relief="flat", font=(FONT, 9, "bold"),
        padx=14, pady=6, cursor="hand2", activebackground=ACCENT2, activeforeground=TEXT2,
        command=lambda: (webbrowser.open("https://discord.gg/VWeTPh3m8Q"), win.destroy()),
    ).pack(side="left")

    tk.Button(
        btn_row, text="Dismiss", bg=PANEL, fg=SUBTEXT, relief="flat", font=(FONT, 9, "bold"),
        padx=14, pady=6, cursor="hand2", activebackground=BORDER, activeforeground=TEXT,
        command=win.destroy,
    ).pack(side="right")

    win.update_idletasks()
    sw = win.winfo_screenwidth()
    sh = win.winfo_screenheight()
    win.geometry(f"+{(sw - 420) // 2}+{(sh - 260) // 2}")


def open_help():
    help_win = tk.Toplevel(root)
    help_win.title("Documentation & Guide")
    help_win.geometry("520x460")
    help_win.configure(bg=BG)
    help_win.resizable(False, False)

    current_page = [0]
    pages = [
        {
            "title": "Welcome to ToolBox",
            "content": (
                "This control center manages and runs various modular optimization tools "
                "tailored for VRChat OSC network tracking.\n\n"
                "Features include:\n"
                "• Automated system update patches on initialization cycles.\n"
                "• Sandbox virtual execution container environments.\n"
                "• Fast preference configuration overlays."
            ),
        },
        {
            "title": "Status Indicator",
            "content": (
                "The status shelf located across the footer displays active telemetry feedback:\n"
                "• 'Ready' — waiting for action\n"
                "• 'Starting up (ScriptName)' — launching\n"
                "• 'Up to date' — version check complete\n"
                "• 'Error' — something went wrong"
            ),
        },
        {
            "title": "Available Scripts",
            "content": (
                "▶ Router(Beta) — Manages OSC routing\n"
                " Forwards OSC messages between sources\n"
                " and destinations.\n\n"
                "▶ ChatBox — Sends data to VRChat over OSC\n"
                " Displays system info, weather, music,\n"
                " and custom messages.\n\n"
                "▶ Face Tracking Controller(Beta) — Control\n"
                " face tracking features in VRChat ."
            ),
        },
        {
            "title": "Status Bar",
            "content": (
                "The top bar of each script shows:\n\n"
                "Left: Script name and icon\n"
                "Center: Version number\n"
                "Right: Current status\n\n"
                "Status Examples:\n"
                "• Status: Running — Script is active\n"
                "• Status: Stopped — Script is inactive\n"
                "• Status: Error — Something failed"
            ),
        },
        {
            "title": "Adding a Script",
            "content": (
                "1. Click the ⚙ (gear) button in the footer\n"
                "2. Click '+ Add Script' button\n"
                "3. Enter a label (button text)\n"
                "4. Enter filename or full path\n"
                "5. Click 'Add' to save\n\n"
                "Your new script button appears in\n"
                "'MANAGED SCRIPTS' section immediately!"
            ),
        },
        {
            "title": "Removing a Script",
            "content": (
                "1. Click the ⚙ (gear) button\n"
                "2. Find the script in the list\n"
                "3. Click the '✕ Remove' button\n"
                "4. Script removed from buttons\n\n"
                "Changes save automatically. Close and\n"
                "reopen ToolBox to fully refresh if needed."
            ),
        },
        {
            "title": "Tips",
            "content": (
                "• Always start Router first, then ChatBox\n\n"
                "• Each script remembers its settings\n"
                " between sessions\n\n"
                "• Check your internet connection if\n"
                " scripts fail to start\n\n"
                "• Run scripts from the ToolBox for\n"
                " proper management"
            ),
        },
    ]

    # Header Panel Layout
    header_panel = tk.Frame(help_win, bg=PANEL, pady=12)
    header_panel.pack(fill="x")
    title_label = tk.Label(header_panel, text="", bg=PANEL, fg=ACCENT2, font=(FONT, 12, "bold"))
    title_label.pack(side="left", padx=20)

    tk.Frame(help_win, bg=BORDER, height=1).pack(fill="x")

    # Content Display Panel Block Frame
    content_panel = tk.Frame(help_win, bg=PANEL)
    content_panel.pack(fill="both", expand=True, padx=20, pady=16)

    # Body Label
    content_label = tk.Label(
        content_panel, text="", bg=PANEL, fg=TEXT, justify="left", wraplength=460, anchor="nw", font=(FONT, 10)
    )
    content_label.pack(padx=14, pady=14, fill="both", expand=True)

    # Pagination View Engine Configuration Block
    def show_page(idx):
        p = pages[idx]
        title_label.config(text=p["title"])
        content_label.config(text=p["content"])
        page_indicator.config(text=f"Page {idx + 1} of {len(pages)}")
        prev_btn.config(state="normal" if idx > 0 else "disabled")
        is_last = idx == len(pages) - 1
        next_btn.config(text="Finish" if is_last else "Next →")

    # Help Window Lower Navigation Dock Frame
    nav_frame = tk.Frame(help_win, bg=BG)
    nav_frame.pack(fill="x", padx=20, pady=(0, 14))
    nav_frame.columnconfigure(1, weight=1)

    # Help Window Previous Page Pagination Control Button
    prev_btn = tk.Button(
        nav_frame, text="← Back", bg=PANEL, fg=TEXT, relief="flat", width=10,
        command=lambda: (current_page.__setitem__(0, current_page[0] - 1), show_page(current_page[0])),
    )
    prev_btn.grid(row=0, column=0, sticky="w")
    prev_btn.configure(
        fg=SUBTEXT, activebackground=BORDER, activeforeground=TEXT, cursor="hand2", font=(FONT, 9, "bold")
    )

    # Execution Link Logic Block for Next/Finish Routines
    def next_or_finish():
        if current_page[0] < len(pages) - 1:
            current_page[0] += 1
            show_page(current_page[0])
        else:
            help_win.destroy()

    # Help Window Next Page/Finish Progression Action Button
    next_btn = tk.Button(
        nav_frame, text="Next →", bg=PANEL, fg=TEXT, relief="flat", width=10, command=next_or_finish
    )
    next_btn.grid(row=0, column=2, sticky="e")
    next_btn.configure(
        bg=ACCENT, fg=TEXT2, activebackground=ACCENT2, activeforeground=TEXT2, cursor="hand2", font=(FONT, 9, "bold")
    )

    # Center Numeric Page Tracker Information Metric Text
    page_indicator = tk.Label(nav_frame, text="", bg=BG, fg=SUBTEXT, font=(FONT, 9))
    page_indicator.grid(row=0, column=1, sticky="center")

    show_page(0)


# Window View: Core App Preference Management Overlay Window
def open_settings():
    global MANAGED_SCRIPTS, UPDATE_BRANCH
    settings_win = tk.Toplevel(root)
    settings_win.title("Settings")
    settings_win.configure(bg=BG)
    settings_win.resizable(True, True)
    settings_win.geometry("520x460")

    # Settings Panel Top Header Frame
    header = tk.Frame(settings_win, bg=PANEL, pady=10)
    header.pack(fill="x")

    # Settings Window Structural Header Text Label
    title_label = tk.Label(
        header, text="Manage Scripts & Settings", bg=PANEL, fg=ACCENT2, font=(FONT, 12, "bold")
    )

    tk.Frame(settings_win, bg=BORDER, height=1).pack(fill="x")

    # Lower Form Shell Sub-Block Packaging Layer
    body_frame = tk.Frame(settings_win, bg=BG, padx=20, pady=14)
    body_frame.pack(fill="both", expand=True)

    # --- Live Branch Selection Trace Dropdown Area ---
    branch_frame = tk.Frame(body_frame, bg=BG)
    branch_frame.pack(fill="x", pady=(0, 10))

    tk.Label(branch_frame, text="Update Branch Context:", bg=BG, fg=TEXT, font=(FONT, 9, "bold")).pack(side="left")

    branch_var = tk.StringVar(value=UPDATE_BRANCH)

    def on_branch_change(*args):
        global UPDATE_BRANCH
        new_branch = branch_var.get()
        if new_branch != UPDATE_BRANCH:
            UPDATE_BRANCH = new_branch
            save_managed_scripts(MANAGED_SCRIPTS)  # Commit update_branch string context to storage configurations
            force_update_all_scripts()  # Instantly fire asynchronous live updates swapping code logic branches

    branch_var.trace_add("write", on_branch_change)

    branch_dropdown = tk.OptionMenu(branch_frame, branch_var, "main", "dev", "beta")
    branch_dropdown.configure(
        bg=PANEL, fg=TEXT, relief="flat", highlightthickness=1, highlightbackground=BORDER,
        font=(FONT, 9), activebackground=BORDER, activeforeground=TEXT, cursor="hand2",
    )
    branch_dropdown["menu"].configure(bg=PANEL, fg=TEXT, font=(FONT, 9), selectcolor=ACCENT)
    branch_dropdown.pack(side="left", padx=(10, 0))

    # Scrollable Canvas List Container Layout Control Set
    list_panel = tk.Frame(body_frame, bg=PANEL, highlightthickness=1, highlightbackground=BORDER)
    list_panel.pack(fill="both", expand=True, pady=(4, 14))

    canvas_width = 460
    inner_canvas = tk.Canvas(list_panel, bg=PANEL, bd=0, highlightthickness=0, width=canvas_width)
    scrollbar = tk.Scrollbar(list_panel, orient="vertical", command=inner_canvas.yview)
    inner_canvas.configure(yscrollcommand=scrollbar.set)

    scrollbar.pack(side="right", fill="y")
    inner_canvas.pack(side="left", fill="both", expand=True)

    # Settings Inner Grid Core Packaging Frame Layout Box
    inner_frame = tk.Frame(inner_canvas, bg=PANEL)
    inner_canvas.create_window((0, 0), window=inner_frame, anchor="nw")

    # Adjust window width dynamically on layout configurations
    inner_frame.bind("<Configure>", lambda e: inner_canvas.configure(scrollregion=inner_canvas.bbox("all")))

    # Content Generator Rendering Core Function
    def refresh_script_list():
        for widget in inner_frame.winfo_children():
            widget.destroy()

        for idx, script in enumerate(MANAGED_SCRIPTS):
            # Row Boundary Segment Wrapper Frame
            script_row = tk.Frame(inner_frame, bg=BG)
            script_row.pack(fill="x", padx=10, pady=6)

            # Row Entry Display Label Description Header
            tk.Label(script_row, text=f"{script['label']}", bg=BG, fg=TEXT, font=(FONT, 9, "bold")).pack(
                side="left", fill="x", expand=True
            )

            # Row Entry Meta-Info Technical String Subtext Label
            tk.Label(script_row, text=f"({script['filename']})", bg=BG, fg=SUBTEXT, font=(FONT, 8)).pack(
                side="left", padx=(10, 0)
            )

            # Entry Item Deletion/Removal Management Interceptor Button
            remove_btn = tk.Button(
                script_row, text="✕ Remove", bg=PANEL, fg=RED, relief="flat", font=(FONT, 8, "bold"), cursor="hand2",
                command=lambda i=idx: remove_script(i),
            )
            remove_btn.pack(side="right", padx=(10, 0))
            remove_btn.configure(activebackground=BORDER, activeforeground=RED)

    # Entry Erasure Logic Array Mutator
    def remove_script(idx):
        MANAGED_SCRIPTS.pop(idx)
        save_managed_scripts(MANAGED_SCRIPTS)
        refresh_script_list()
        refresh_main_buttons()

    # View Component Context: Modal Form Window Container View
    def add_script():
        add_win = tk.Toplevel(settings_win)
        add_win.title("Add Script")
        add_win.configure(bg=BG)
        add_win.geometry("400x200")
        add_win.resizable(False, False)

        tk.Label(add_win, text="Script Display Label:", bg=BG, fg=TEXT, font=(FONT, 9)).grid(
            row=0, column=0, padx=14, pady=14, sticky="w"
        )
        label_entry = tk.Entry(
            add_win, bg=PANEL, fg=TEXT, insertbackground=ACCENT, relief="flat", font=(FONT, 9),
            highlightthickness=1, highlightbackground=BORDER, highlightcolor=ACCENT,
        )
        label_entry.grid(row=0, column=1, padx=14, pady=14, sticky="ew")

        tk.Label(add_win, text="Filename / Resource Path:", bg=BG, fg=TEXT, font=(FONT, 9)).grid(
            row=1, column=0, padx=14, pady=6, sticky="w"
        )
        file_entry = tk.Entry(
            add_win, bg=PANEL, fg=TEXT, insertbackground=ACCENT, relief="flat", font=(FONT, 9),
            highlightthickness=1, highlightbackground=BORDER, highlightcolor=ACCENT,
        )
        file_entry.grid(row=1, column=1, padx=14, pady=6, sticky="ew")

        def save_new_script():
            lbl = label_entry.get().strip()
            flm = file_entry.get().strip()
            if not lbl or not flm:
                messagebox.showwarning("Validation Error", "All entry parameters must be populated.")
                return

            MANAGED_SCRIPTS.append({"filename": flm, "label": lbl})
            save_managed_scripts(MANAGED_SCRIPTS)
            refresh_script_list()
            refresh_main_buttons()
            add_win.destroy()

        submit_btn = tk.Button(
            add_win, text="Save Script", bg=ACCENT, fg=TEXT2, relief="flat", font=(FONT, 9, "bold"),
            command=save_new_script, cursor="hand2", activebackground=ACCENT2, activeforeground=TEXT2,
        )
        submit_btn.grid(row=2, column=1, padx=14, pady=14, sticky="e")

    # Settings Control System Bottom Action Row Dock
    nav_frame = tk.Frame(body_frame, bg=BG)
    nav_frame.pack(fill="x")

    # Application Modal Creation Trigger Action Button
    add_btn = tk.Button(
        nav_frame, text="+ Add Script", bg=ACCENT, fg=TEXT2, relief="flat", width=15, command=add_script
    )
    add_btn.pack(side="left")
    add_btn.configure(activebackground=ACCENT2, activeforeground=TEXT2, cursor="hand2", font=(FONT, 9, "bold"))

    # Settings Panel Termination UI Dismiss Command Execution Button
    close_btn = tk.Button(
        nav_frame, text="Close", bg=PANEL, fg=SUBTEXT, relief="flat", width=10, command=settings_win.destroy
    )
    close_btn.pack(side="right")
    close_btn.configure(activebackground=BORDER, activeforeground=TEXT, cursor="hand2", font=(FONT, 9, "bold"))

    refresh_script_list()


# Main Layout System Workspace Structural Containers
header_frame = tk.Frame(root, bg=PANEL, pady=12)
header_frame.pack(fill="x")

title_font = font.Font(family=FONT, size=16, weight="bold")
title_lbl = tk.Label(header_frame, text=f"{TITLE_PREFIX} VRChat-ToolBox", bg=PANEL, fg=ACCENT2, font=title_font)
title_lbl.pack(side="left", padx=20)

tk.Frame(root, bg=BORDER, height=1).pack(fill="x")

main_frame = tk.Frame(root, bg=BG, padx=24, pady=16)
main_frame.pack(fill="both", expand=True)
main_frame.columnconfigure(0, weight=1)

main_frame.columnconfigure(1, weight=1)

# ── Tool buttons section with label ────────────────────────────────────────
# Main View Dashboard Content Partition Text Heading Label
tools_label = tk.Label(main_frame, text="MANAGED SCRIPTS", bg=BG, fg=ACCENT, font=(FONT, 9, "bold"))
tools_label.grid(row=0, column=0, sticky="w", pady=(0, 10))

# Subview Scrollable Button Container Frame Engine Component Block
buttons_container = tk.Frame(main_frame, bg=BG)
buttons_container.grid(row=1, column=0, sticky="nsew", columnspan=2)
buttons_container.columnconfigure(0, weight=1)

script_buttons = {}


def refresh_main_buttons():
    for widget in buttons_container.winfo_children():
        widget.destroy()

    script_buttons.clear()

    for i, script in enumerate(MANAGED_SCRIPTS):
        btn = tk.Button(
            buttons_container, text=script["label"], command=lambda f=script["filename"]: launch_script(f),
            bg=PANEL, fg=TEXT, relief="flat", borderwidth=0, highlightthickness=1, highlightbackground=BORDER,
            activebackground=ACCENT, activeforeground=TEXT2, cursor="hand2", font=(FONT, 10, "bold"), padx=20, pady=8,
        )
        btn.grid(row=i, column=0, padx=0, pady=4, sticky="ew")
        script_buttons[i] = btn

    btn_count = len(MANAGED_SCRIPTS)
    root.geometry(f"580x{440 + btn_count * 52}")
    root.minsize(0, 0)


refresh_main_buttons()

# ── Footer with update info ────────────────────────────────────────────────
# App Window Core Status Shelf Base Layout Frame
footer_bar = tk.Frame(root, bg=PANEL, pady=8)
footer_bar.pack(fill="x", side="bottom")

footer_bar.columnconfigure(0, weight=1)

# Status Footer Navigation System Information Callout Entry Utility Button
help_btn = square_button(footer_bar, "？", open_help, base_size=28)
help_btn.pack(side="left", padx=(8, 0))

# Status Footer App Preferences Component Settings Navigation Button Entry Widget
settings_btn = square_button(footer_bar, "⚙", open_settings, base_size=28)
settings_btn.pack(side="right", padx=(0, 8))

# Status Footer System Operational Message Text Display Feedback Label
footer_label = tk.Label(
    footer_bar, text="Checking for updates on startup...", bg=PANEL, fg=SUBTEXT, font=(FONT, 8)
)
footer_label.pack(side="bottom", fill="x", pady=(2, 0))

# Automatically kick off startup network validation threads asynchronously
threading.Thread(target=lambda: check_for_main_updates(silent=True), daemon=True).start()

# Conditional Beta Modal Promotion Injection
if UPDATE_BRANCH == "beta":
    root.after(800, _show_beta_popup)

root.mainloop()