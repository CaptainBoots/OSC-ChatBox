# ═════════════════════════════════════════════════════════════════════════════════════════════════════════════════════#
#                                              Boot's ToolBox Script                                                      #
# ═════════════════════════════════════════════════════════════════════════════════════════════════════════════════════#
# Hi :3
# Welcome to my code

# ═════════════════════════════════════════════════════════════════════════════════════════════════════════════════════#
# Imports
# ═════════════════════════════════════════════════════════════════════════════════════════════════════════════════════#

import subprocess
import sys
import importlib
import re
import time
import tkinter as tk
from tkinter import messagebox, ttk
import tkinter.font as font
import os
import site
import json


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
VERSION = "9.3.2"

GITHUB_RAW_URL = "https://raw.githubusercontent.com/CaptainBoots/VRChat-ToolBox/main/VRChat-ToolBox.py"
GITHUB_BASE_URL = "https://raw.githubusercontent.com/CaptainBoots/VRChat-ToolBox/main/VRChat-Tools/"
GITHUB_EXE_RELEASE_BASE_URL = "https://github.com/CaptainBoots/VRChat-ToolBox/releases/latest/download/"
GITHUB_EXE_RAW_BASE_URL = "https://raw.githubusercontent.com/CaptainBoots/VRChat-ToolBox/main/"

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

if VERSION == "9.3.2":
    if os.path.exists(TOOLBOX_CONFIG_FILE):
        try:
            os.remove(TOOLBOX_CONFIG_FILE)
            print(f"[Config] Version config change detected. Forced clean reset of: {TOOLBOX_CONFIG_FILE}")
        except OSError as e:
            print(f"[Config] Failed to force-delete config: {e}")

DEFAULT_MANAGED_SCRIPTS = [
    {"filename": "VRChat-Launcher/main.py", "label": "VRChat Launcher(Beta)"},
    {"filename": "OSC-Router/main.py", "label": "Router"},
    {"filename": "OSC-Chatbox/main.py", "label": "ChatBox"},
    {"filename": "OSC-Gamepad/main.py", "label": "Gamepad(Beta)"},
    {"filename": "OSC-FaceTrackingController/main.py", "label": "Face Tracking Controller(Beta)"},
    {"filename": "OSC-ParameterBrowser/main.py", "label": "Parameter Browser(Beta)"},
    {"filename": "OSC-ScriptMaker/main.py", "label": "Script Maker(Placeholder)"},
    {"filename": "VRChat-LocalFavorites/main.py", "label": "VRChat Local Favorites(Placeholder)"},
    {"filename": "VRChat-SocialLogger/main.py", "label": "VRChat SocialLogger(Placeholder)"},
]


def load_managed_scripts():
    if os.path.exists(TOOLBOX_CONFIG_FILE):
        try:
            with open(TOOLBOX_CONFIG_FILE, "r", encoding="utf-8") as f:
                config = json.load(f)

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
        "local_path":  os.path.join("VRChat-Launcher", "main.py"),
    },
    "OSC-Router/main.py": {
        "remote_path": "OSC-Router/main.py",
        "local_path":  os.path.join("OSC-Router", "main.py"),
    },
    "OSC-Chatbox/main.py": {
        "remote_path": "OSC-Chatbox/main.py",
        "local_path":  os.path.join("OSC-Chatbox", "main.py"),
    },
    "OSC-Gamepad/main.py": {
        "remote_path": "OSC-Gamepad/main.py",
        "local_path":  os.path.join("OSC-Gamepad", "main.py"),
    },
    "OSC-FaceTrackingController/main.py": {
        "remote_path": "OSC-FaceTrackingController/main.py",
        "local_path":  os.path.join("OSC-FaceTrackingController", "main.py"),
    },
    "OSC-ParameterBrowser/main.py": {
        "remote_path": "OSC-ParameterBrowser/main.py",
        "local_path":  os.path.join("OSC-ParameterBrowser", "main.py"),
    },
    "OSC-ScriptMaker/main.py": {
        "remote_path": "OSC-ScriptMaker/main.py",
        "local_path":  os.path.join("OSC-ScriptMaker", "main.py"),
    },
    "VRChat-LocalFavorites/main.py": {
        "remote_path": "VRChat-LocalFavorites/main.py",
        "local_path":  os.path.join("VRChat-LocalFavorites", "main.py"),
    },
    "VRChat-SocialLogger/main.py": {
        "remote_path": "VRChat-SocialLogger/main.py",
        "local_path":  os.path.join("VRChat-SocialLogger", "main.py"),
    },
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

        # UI module
        "OSC-Chatbox/ui/__init__.py",
        "OSC-Chatbox/ui/app.py",
        "OSC-Chatbox/ui/builder.py",
        "OSC-Chatbox/ui/chatbox_tab.py",
        "OSC-Chatbox/ui/help_dialog.py",
        "OSC-Chatbox/ui/settings_dialog.py",
        "OSC-Chatbox/ui/theme.py",
    ],
    "OSC-Gamepad/main.py": [
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

# Per-tool config files to wipe on update (paths relative to TOOLS_ROOT_DIR).
# This ensures users always get a clean config after a breaking update.
TOOL_CONFIG_WIPE_MAP: dict[str, list[str]] = {
    "OSC-Chatbox/main.py": [
        os.path.join("OSC-Chatbox", "chatbox_config.json"),
    ],
}

SCRIPT_FOLDER_MAP = {
    "OSC-Router.py": "OSC-Router",
    "OSC-FaceTrackingController.py": "OSC-FaceTrackingController",
    "OSC-Gamepad.py": "OSC-Gamepad",
    "OSC-ScriptMaker.py": "OSC-ScriptMaker",
    "OSC-ParameterBrowser.py": "OSC-ParameterBrowser",
    "VRChat-Launcher.py": "VRChat-Launcher",
    "VRChat-LocalFavorites.py": "LocalFavorites",
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
        return [f"{GITHUB_BASE_URL}{remote_path}"]
    if _is_path_like(filename):
        return []
    script_name = os.path.basename(filename)
    urls: list[str] = []
    folder = SCRIPT_FOLDER_MAP.get(script_name)
    if folder:
        urls.append(f"{GITHUB_BASE_URL}{folder}/{script_name}")
    urls.append(f"{GITHUB_BASE_URL}{script_name}")
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
            messagebox.showerror(f"{filename} Error", f"Could not create destination directory:\n{e}")
        return False

    remote_text = None
    for url in _script_remote_urls(filename):
        try:
            response = requests.get(
                url,
                timeout=10,
                params={"_": int(time.time())},
                headers={"Cache-Control": "no-cache", "Pragma": "no-cache"},
            )
            response.raise_for_status()
            remote_text = response.text
            print(f"[{filename}] Download source: {url}")
            break
        except requests.RequestException:
            continue

    if remote_text:
        try:
            with open(dest_path, "w", encoding="utf-8") as f:
                f.write(remote_text)
            print(f"[{filename}] Downloaded to {dest_path}")

            # Download all associated project dependency files
            for dep in dependencies:
                dep_url = f"{GITHUB_BASE_URL}{dep}"
                dep_dest = os.path.join(TOOLS_ROOT_DIR, dep.replace("/", os.sep))
                os.makedirs(os.path.dirname(dep_dest), exist_ok=True)
                try:
                    dep_resp = requests.get(
                        dep_url,
                        timeout=10,
                        params={"_": int(time.time())},
                        headers={"Cache-Control": "no-cache", "Pragma": "no-cache"},
                    )
                    dep_resp.raise_for_status()
                    with open(dep_dest, "w", encoding="utf-8") as df:
                        df.write(dep_resp.text)
                    print(f"[{filename}] Downloaded dependency: {dep_dest}")
                except requests.RequestException as dep_err:
                    print(f"[{filename}] Failed to download dependency {dep_url}: {dep_err}")

            return True
        except OSError as e:
            print(f"[{filename}] Could not save downloaded script: {e}")

    if os.path.isfile(bundled_path):
        try:
            with open(bundled_path, "r", encoding="utf-8") as src:
                with open(dest_path, "w", encoding="utf-8") as dst:
                    dst.write(src.read())
            print(f"[{filename}] Using bundled fallback copy from {bundled_path}")
            return True
        except OSError as e:
            print(f"[{filename}] Could not copy bundled fallback: {e}")

    if show_errors:
        messagebox.showerror(
            f"{filename} Error",
            f"Could not prepare {filename}.\nCheck your internet connection and try again.",
        )
    return False


def check_for_script_updates(filename: str, silent: bool = False) -> bool:
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

    print(f"[{filename}] Checking... (local: {local_version}  remote: {remote_version})")

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
            dep_url = f"{GITHUB_BASE_URL}{dep}"
            dep_dest = os.path.join(TOOLS_ROOT_DIR, dep.replace("/", os.sep))
            os.makedirs(os.path.dirname(dep_dest), exist_ok=True)
            try:
                dep_resp = requests.get(
                    dep_url,
                    timeout=10,
                    params={"_": int(time.time())},
                    headers={"Cache-Control": "no-cache", "Pragma": "no-cache"},
                )
                dep_resp.raise_for_status()
                with open(dep_dest, "w", encoding="utf-8") as df:
                    df.write(dep_resp.text)
                print(f"[{filename}] Updated dependency: {dep_dest}")
            except requests.RequestException as dep_err:
                print(f"[{filename}] Failed to update dependency {dep_url}: {dep_err}")

        # Wipe any tool-specific config files so stale config doesn't break the new version
        for relative_cfg in TOOL_CONFIG_WIPE_MAP.get(filename, []):
            cfg_path = os.path.join(TOOLS_ROOT_DIR, relative_cfg)
            if os.path.exists(cfg_path):
                try:
                    os.remove(cfg_path)
                    print(f"[{filename}] Wiped stale config: {cfg_path}")
                except OSError as wipe_err:
                    print(f"[{filename}] Could not wipe config {cfg_path}: {wipe_err}")

        if not silent:
            messagebox.showinfo(
                f"{filename} Updated",
                f"{filename} updated from {local_version} to {remote_version}."
            )
        return True
    except OSError as e:
        print(f"[{filename}] Could not save updated script: {e}")
        return False

# ═════════════════════════════════════════════════════════════════════════════════════════════════════════════════════#
# UPDATER
# ═════════════════════════════════════════════════════════════════════════════════════════════════════════════════════#

def _set_footer_ready(_unused=None) -> None:
    footer_label.config(text="Ready")


def _parse_version(v: str):
    nums = [int(x) for x in re.findall(r"\d+", (v or "").strip())]
    if not nums:
        return 0, 0, 0
    while len(nums) < 3:
        nums.append(0)
    return tuple(nums[:3])


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
            url,
            timeout=timeout,
            params={"_": int(time.time())},
            headers={"Cache-Control": "no-cache", "Pragma": "no-cache"},
        )
        resp.raise_for_status()
        return resp.text, _extract_version_from_source(resp.text), url
    except requests.RequestException:
        return None, None, None


def get_remote_script_info() -> dict[str, str] | None:
    urls = [GITHUB_RAW_URL]
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

        with open(script_path, "r", encoding="utf-8") as f_cur:
            with open(backup_path, "w", encoding="utf-8") as f_bak:
                f_bak.write(f_cur.read())

        backup_files = [
            os.path.join(BACKUP_DIR, fn)
            for fn in os.listdir(BACKUP_DIR)
            if fn.lower().endswith(".bak")
        ]
        if len(backup_files) > 5:
            oldest_backup = min(backup_files, key=os.path.getmtime)
            if os.path.abspath(oldest_backup) != os.path.abspath(backup_path):
                os.remove(oldest_backup)
                print(f"[Updater] Removed old backup: {os.path.basename(oldest_backup)}")

        with open(script_path, "w", encoding="utf-8") as f:
            f.write(remote_text)

        print(f"[Updater] Update downloaded from {source_url or GITHUB_RAW_URL}. Restarting...")
        subprocess.Popen([sys.executable, script_path])
        root.destroy()
        sys.exit(0)

    except Exception as e:
        messagebox.showerror(
            "Update Failed",
            f"Could not apply update:\n{e}\n\nYour original file is intact."
        )
        print(f"[Updater] Update failed: {e}")


def check_for_updates(silent=False):
    info = get_remote_script_info()
    if not info:
        if not silent:
            messagebox.showinfo(
                "Update Check",
                "Could not reach GitHub.\nCheck your internet connection."
            )
        for script_entry in MANAGED_SCRIPTS:
            check_for_script_updates(script_entry["filename"], silent=silent)
        return

    remote_version = info["version"]
    remote_text = info["text"]
    remote_url = info["url"]

    try:
        # FIX: Changed to os.path.realpath to follow links, and added errors="ignore"
        # to prevent crashing on computers with non-English system locales/encodings.
        with open(os.path.realpath(__file__), "r", encoding="utf-8", errors="ignore") as f:
            local_text = f.read()
    except Exception:
        local_text = ""

    local_norm = local_text.replace("\r\n", "\n")
    remote_norm = remote_text.replace("\r\n", "\n")
    remote_newer = _parse_version(remote_version) > _parse_version(VERSION)
    content_differs = remote_norm != local_norm
    main_update_available = remote_newer or content_differs

    print(f"[VRChat-Tools] Checking... (local: {VERSION}  remote: {remote_version})")

    if main_update_available:
        if remote_newer:
            print(f"[VRChat-Tools] Update available: {VERSION} -> {remote_version}")
            if os.path.exists(TOOLBOX_CONFIG_FILE):
                try:
                    os.remove(TOOLBOX_CONFIG_FILE)
                except Exception:
                    pass
            prompt = (
                f"New version {remote_version} is available (you have {VERSION}).\n\n"
                "Update and restart now?"
            )
        else:
            print(f"[VRChat-Tools] Remote content differs (version string unchanged at {VERSION})")
            prompt = (
                "A remote script update is available (content changed,\n"
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
        messagebox.showinfo("Up to Date", f"You're on the latest version ({VERSION}).")

# ═════════════════════════════════════════════════════════════════════════════════════════════════════════════════════#
# GUI
# ═════════════════════════════════════════════════════════════════════════════════════════════════════════════════════#

# UI Color Palette Configurations
BG = "#0f0f13"
PANEL = "#1f102a"
BORDER = "#2a2a38"
ACCENT = "#9D00FF"
ACCENT2 = "#b44bff"
TEXT = "#e2e0f0"
TEXT2 = "#E0E0E0"
SUBTEXT = "#7e7b9a"
GREEN = "#4ade80"
RED = "#f87171"
FG = TEXT
ENTRY_BG = PANEL
BTN_BG = PANEL
BTN_FG = TEXT
UI_FONT = "Consolas"

# Dynamic Scaling Variables
ui_scale = 1.0
scalable_widgets = []
square_widgets = []

# Main Application Window Initialization
root = tk.Tk()
root.title("Boot's ToolBox")
root.configure(bg=BG)
root.resizable(True, True)

# Top Bar Container Frame
title_bar = tk.Frame(root, bg=PANEL, pady=14)
title_bar.pack(fill="x")

# Header Padding and Alignment Frame
header_frame = tk.Frame(title_bar, bg=PANEL)
header_frame.pack(fill="x", padx=16, expand=True)

# Main Application Brand Title Label
header_title_label = tk.Label(
    header_frame,
    text="◈  Boot's TOOLBOX",
    bg=PANEL,
    fg=ACCENT2,
    font=(UI_FONT, 16, "bold")
)
header_title_label.pack(side="left", anchor="w")

# Application Version Display Label
version_label = tk.Label(
    header_frame,
    text=f"v{VERSION}",
    bg=PANEL,
    fg=SUBTEXT,
    font=(UI_FONT, 9)
)
version_label.pack(side="right", anchor="e", padx=(32, 0))

# Visual Separator Line (Border)
tk.Frame(root, bg=BORDER, height=1).pack(fill="x")

# Central Content Wrapper Frame
main_frame = tk.Frame(root, bg=BG)
main_frame.pack(fill="both", expand=True, padx=16, pady=14)


# ── Scaling ────────────────────────────────────────────────────────────────
def apply_scale(scale):
    global ui_scale
    ui_scale = scale

    base_default = 9
    new_default = max(7, int(base_default * scale))

    font.nametofont("TkDefaultFont").configure(size=new_default)
    font.nametofont("TkTextFont").configure(size=new_default)
    font.nametofont("TkFixedFont").configure(size=new_default)

    for widget, base_size, extras in scalable_widgets:
        try:
            widget.configure(font=(UI_FONT, max(6, int(base_size * scale))) + extras)
        except tk.TclError:
            pass

    for container, base_size, btn_widget in square_widgets:
        size = int(base_size * scale)
        container.config(width=size, height=size)
        btn_widget.config(font=(UI_FONT, max(8, int(12 * scale))))


# Factory Function: Standard UI Subtext Label
def dark_label(text, r, **kwargs):
    lbl = tk.Label(main_frame, text=text, bg=BG, fg=SUBTEXT, anchor="w", font=(UI_FONT, 9))
    lbl.grid(row=r, column=0, sticky="w", pady=6, **kwargs)
    return lbl


# Factory Function: Standard Form Input Field
def dark_entry(r, default=""):
    e = tk.Entry(
        main_frame,
        bg=ENTRY_BG,
        fg=TEXT,
        insertbackground=ACCENT,
        relief="flat",
        font=(UI_FONT, 9),
        highlightthickness=1,
        highlightbackground=BORDER,
        highlightcolor=ACCENT,
    )
    e.insert(0, default)
    e.grid(row=r, column=1, pady=6, sticky="ew", padx=(8, 0))
    return e


# Factory Function: Fixed Aspect-Ratio Square Button
def square_button(parent, text, command, base_size=32):
    container = tk.Frame(parent, bg=BTN_BG, highlightthickness=1, highlightbackground=BORDER)
    container.pack_propagate(False)

    button_widget = tk.Button(
        container,
        text=text,
        command=command,
        bg=BTN_BG,
        fg=SUBTEXT,
        relief="flat",
        borderwidth=0,
        font=(UI_FONT, 12),
        activebackground=BORDER,
        activeforeground=TEXT,
        cursor="hand2",
    )
    button_widget.pack(fill="both", expand=True)

    square_widgets.append((container, base_size, button_widget))
    container.config(width=base_size, height=base_size)

    return container


# Window View: Help and Tutorial Modal Overlay
def open_help():
    help_win = tk.Toplevel(root)
    help_win.title("Tutorial")
    help_win.configure(bg=BG)
    help_win.resizable(True, True)

    root.update_idletasks()
    help_w = root.winfo_width()
    help_h = root.winfo_height()
    root_x = root.winfo_x()
    root_y = root.winfo_y()
    help_win.geometry(f"{help_w}x{help_h}+{root_x}+{root_y}")

    pages = [
        {
            "title": "Welcome",
            "content": (
                "Boot's ToolBox — The main control center for\n"
                "managing all of Boot's scripts.\n\n"
                "MANAGED SCRIPTS\n"
                "Click any script button to launch it.\n\n"
                "The footer at the bottom shows:\n"
                "• 'Ready' — waiting for action\n"
                "• 'Starting up (ScriptName)' — launching\n"
                "• 'Up to date' — version check complete\n"
                "• 'Error' — something went wrong"
            )
        },
        {
            "title": "Available Scripts",
            "content": (
                "▶ Router(Beta) — Manages OSC routing\n"
                "  Forwards OSC messages between sources\n"
                "  and destinations.\n\n"
                "▶ ChatBox — Sends data to VRChat over OSC\n"
                "  Displays system info, weather, music,\n"
                "  and custom messages.\n\n"
                "▶ Face Tracking Controller(Beta) — Control\n"
                "  face tracking features in VRChat ."
            )
        },
        {
            "title": "Status Bar",
            "content": (
                "The top bar of each script shows:\n\n"
                "Left: Script name and icon\n"
                "Center: Version number\n"
                "Right: Current status\n\n"
                "Status Examples:\n"
                "● Status: Running — Script is active\n"
                "● Status: Stopped — Script is inactive\n"
                "● Status: Error — Something failed"
            )
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
            )
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
            )
        },
        {
            "title": "Tips",
            "content": (
                "• Always start Router first, then ChatBox\n\n"
                "• Each script remembers its settings\n"
                "  between sessions\n\n"
                "• Check your internet connection if\n"
                "  scripts fail to start\n\n"
                "• Run scripts from the ToolBox for\n"
                "  proper management"
            )
        },
    ]

    current_page = [0]

    # Help Window Top Header Frame
    header = tk.Frame(help_win, bg=PANEL, pady=10)
    header.pack(fill="x")

    # Help Window Section Title Label
    title_label = tk.Label(
        header, text="", bg=PANEL, fg=ACCENT2, font=(UI_FONT, 12, "bold")
    )
    title_label.pack(side="left", padx=16)

    # Help Window Pagination Tracker Label
    page_indicator = tk.Label(
        header, text="", bg=PANEL, fg=SUBTEXT, font=(UI_FONT, 8)
    )
    page_indicator.pack(side="right", padx=16)

    # Help Window Top Visual Separator Line
    tk.Frame(help_win, bg=BORDER, height=1).pack(fill="x")

    # Help Window Content Card Outer Boundary Panel
    content_panel = tk.Frame(help_win, bg=PANEL, highlightthickness=1, highlightbackground=BORDER)
    content_panel.pack(padx=20, pady=(14, 0), fill="both", expand=True)

    # Help Window Main Text Reader Body Label
    content_label = tk.Label(
        content_panel,
        text="",
        bg=PANEL, fg=TEXT,
        justify="left",
        wraplength=460,
        anchor="nw",
        font=(UI_FONT, 10)
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
        nav_frame, text="← Back", bg=BTN_BG, fg=BTN_FG, relief="flat", width=10,
        command=lambda: (current_page.__setitem__(0, current_page[0] - 1),
                         show_page(current_page[0]))
    )
    prev_btn.grid(row=0, column=0, sticky="w")
    prev_btn.configure(
        fg=SUBTEXT, activebackground=BORDER, activeforeground=TEXT,
        cursor="hand2", font=(UI_FONT, 9, "bold"),
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
        nav_frame, text="Next →", bg=BTN_BG, fg=BTN_FG, relief="flat", width=10,
        command=next_or_finish
    )
    next_btn.grid(row=0, column=2, sticky="e")
    next_btn.configure(
        bg=ACCENT, fg=TEXT2, activebackground=ACCENT2, activeforeground=TEXT2,
        cursor="hand2", font=(UI_FONT, 9, "bold"),
    )

    show_page(0)


# Window View: Core App Preference Management Overlay Window
def open_settings():
    global MANAGED_SCRIPTS

    settings_win = tk.Toplevel(root)
    settings_win.title("Settings")
    settings_win.configure(bg=BG)
    settings_win.resizable(True, True)

    root.update_idletasks()
    help_w = root.winfo_width()
    help_h = root.winfo_height()
    root_x = root.winfo_x()
    root_y = root.winfo_y()
    settings_win.geometry(f"{help_w}x{help_h}+{root_x}+{root_y}")

    # Settings Panel Top Header Frame
    header = tk.Frame(settings_win, bg=PANEL, pady=10)
    header.pack(fill="x")

    # Settings Window Structural Header Text Label
    title_label = tk.Label(
        header, text="Manage Scripts", bg=PANEL, fg=ACCENT2, font=(UI_FONT, 12, "bold")
    )
    title_label.pack(side="left", padx=16)

    # Settings Panel Top Visual Border Partition
    tk.Frame(settings_win, bg=BORDER, height=1).pack(fill="x")

    # Settings Scrollable View Core Container Panel
    content_panel = tk.Frame(settings_win, bg=PANEL, highlightthickness=1, highlightbackground=BORDER)
    content_panel.pack(padx=20, pady=(14, 0), fill="both", expand=True)

    # Settings Dynamic Rendering Scroll Engine Canvas Viewport
    inner_canvas = tk.Canvas(content_panel, bg=PANEL, highlightthickness=0)

    # Settings Scroller Bar Component Interceptor Link
    scrollbar = tk.Scrollbar(content_panel, orient="vertical", command=inner_canvas.yview)
    inner_canvas.configure(yscrollcommand=scrollbar.set)

    scrollbar.pack(side="right", fill="y")
    inner_canvas.pack(side="left", fill="both", expand=True)

    # Settings Inner Grid Core Packaging Frame Layout Box
    inner_frame = tk.Frame(inner_canvas, bg=PANEL)
    inner_canvas.create_window((0, 0), window=inner_frame, anchor="nw")

    # Tracking Event Loop Boundary Adjuster Call
    def on_frame_configure(event=None):
        inner_canvas.configure(scrollregion=inner_canvas.bbox("all"))

    inner_frame.bind("<Configure>", on_frame_configure)

    # Content Generator Rendering Core Function
    def refresh_script_list():
        for widget in inner_frame.winfo_children():
            widget.destroy()

        for idx, script in enumerate(MANAGED_SCRIPTS):
            # Row Boundary Segment Wrapper Frame
            script_row = tk.Frame(inner_frame, bg=BG)
            script_row.pack(fill="x", padx=10, pady=6)

            # Row Entry Display Label Description Header
            tk.Label(script_row, text=f"{script['label']}", bg=BG, fg=TEXT, font=(UI_FONT, 9, "bold")).pack(side="left",
                                                                                                            fill="x",
                                                                                                            expand=True)
            # Row Entry Meta-Info Technical String Subtext Label
            tk.Label(script_row, text=f"({script['filename']})", bg=BG, fg=SUBTEXT, font=(UI_FONT, 8)).pack(side="left",
                                                                                                            padx=(10,
                                                                                                                  0))

            # Entry Item Deletion/Removal Management Interceptor Button
            remove_btn = tk.Button(
                script_row,
                text="✕ Remove",
                bg=PANEL,
                fg=RED,
                relief="flat",
                font=(UI_FONT, 8, "bold"),
                cursor="hand2",
                command=lambda i=idx: remove_script(i)
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

        # Form Display Script Label Section Title Header
        tk.Label(add_win, text="Script Label:", bg=BG, fg=TEXT, font=(UI_FONT, 9)).pack(pady=(10, 0), padx=10,
                                                                                        anchor="w")
        # Form Form-Field Value Entry Interface Text Box
        label_entry = tk.Entry(add_win, bg=PANEL, fg=TEXT, font=(UI_FONT, 9), relief="flat", insertbackground=ACCENT,
                               highlightthickness=1, highlightbackground=BORDER, highlightcolor=ACCENT)
        label_entry.pack(pady=(0, 10), padx=10, fill="x")

        # Form System Storage Path / Resource Filename Title Label
        tk.Label(add_win, text="Filename/Path:", bg=BG, fg=TEXT, font=(UI_FONT, 9)).pack(pady=(10, 0), padx=10,
                                                                                         anchor="w")
        # Form System Target Parameter String Input Data Box
        file_entry = tk.Entry(add_win, bg=PANEL, fg=TEXT, font=(UI_FONT, 9), relief="flat", insertbackground=ACCENT,
                              highlightthickness=1, highlightbackground=BORDER, highlightcolor=ACCENT)
        file_entry.pack(pady=(0, 20), padx=10, fill="x")

        # Action Execution Validation Sequence Mutator Method
        def save_new_script():
            label = label_entry.get().strip()
            filename = file_entry.get().strip()
            if label and filename:
                MANAGED_SCRIPTS.append({"label": label, "filename": filename})
                save_managed_scripts(MANAGED_SCRIPTS)
                refresh_script_list()
                refresh_main_buttons()
                add_win.destroy()

        # Modal Form Action Button Alignment Layout Box Frame
        btn_frame = tk.Frame(add_win, bg=BG)
        btn_frame.pack(pady=10)

        # Form Operations Commitment Target Dispatch Execution Button
        tk.Button(
            btn_frame, text="Add", bg=ACCENT, fg=TEXT2, relief="flat",
            font=(UI_FONT, 9, "bold"), cursor="hand2",
            activebackground=ACCENT2, activeforeground=TEXT2,
            command=save_new_script
        ).pack(side="left", padx=5)

        # Form Operations Termination Abort View Modal Close Button
        tk.Button(
            btn_frame, text="Cancel", bg=PANEL, fg=SUBTEXT, relief="flat",
            font=(UI_FONT, 9, "bold"), cursor="hand2",
            activebackground=BORDER, activeforeground=TEXT,
            command=add_win.destroy
        ).pack(side="left", padx=5)

    refresh_script_list()

    # Settings Window Footer Navigation Command Panel Frame
    nav_frame = tk.Frame(settings_win, bg=BG)
    nav_frame.pack(fill="x", padx=20, pady=(0, 14))
    nav_frame.columnconfigure(0, weight=1)

    # Settings Panel Trigger Interface New Entry Creation Action Button
    add_btn = tk.Button(
        nav_frame, text="+ Add Script", bg=ACCENT, fg=TEXT2, relief="flat", width=15,
        command=add_script
    )
    add_btn.pack(side="left")
    add_btn.configure(
        activebackground=ACCENT2, activeforeground=TEXT2,
        cursor="hand2", font=(UI_FONT, 9, "bold"),
    )

    # Settings Panel Termination UI Dismiss Command Execution Button
    close_btn = tk.Button(
        nav_frame, text="Close", bg=PANEL, fg=SUBTEXT, relief="flat", width=10,
        command=settings_win.destroy
    )
    close_btn.pack(side="right")
    close_btn.configure(
        activebackground=BORDER, activeforeground=TEXT,
        cursor="hand2", font=(UI_FONT, 9, "bold"),
    )


main_frame.columnconfigure(1, weight=1)

# ── Tool buttons section with label ────────────────────────────────────────
# Main View Dashboard Content Partition Text Heading Label
tools_label = tk.Label(
    main_frame,
    text="MANAGED SCRIPTS",
    bg=BG,
    fg=ACCENT,
    font=(UI_FONT, 10, "bold")
)
tools_label.grid(row=14, column=0, columnspan=2, sticky="w", pady=(16, 8))

# Main Dashboard Dynamic Scripts Grid Allocation Box Layout Frame
bottom_bar = tk.Frame(main_frame, bg=BG)
bottom_bar.grid(row=15, column=0, columnspan=2, pady=(0, 6), sticky="ew")
bottom_bar.columnconfigure(0, weight=1)

script_buttons = {}


# Component Refresher Task Invocation Engine
def refresh_main_buttons():
    global script_buttons
    for btn in script_buttons.values():
        btn.destroy()
    script_buttons.clear()

    for i, entry in enumerate(MANAGED_SCRIPTS):
        script_filename = entry["filename"]
        label = entry["label"]

        # Dashboard Interface Target Application Launch/Initialization Command Button
        btn = tk.Button(
            bottom_bar,
            text=f"▶  {label}",
            command=lambda fn=script_filename: launch_script(fn),
            bg=ACCENT,
            fg=TEXT2,
            relief="flat",
            activebackground=ACCENT2,
            activeforeground=TEXT2,
            cursor="hand2",
            font=(UI_FONT, 10, "bold"),
            padx=20,
            pady=8,
        )
        btn.grid(row=i, column=0, padx=0, pady=4, sticky="ew")
        script_buttons[i] = btn

    btn_count = len(MANAGED_SCRIPTS)
    root.geometry(f"580x{400 + btn_count * 52}")
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
    footer_bar,
    text="Checking for updates on startup...",
    bg=PANEL,
    fg=SUBTEXT,
    font=(UI_FONT, 8)
)
footer_label.pack(side="left", padx=16)


# ── Startup update check ───────────────────────────────────────────────────
def run_startup_update_check(_unused=None):
    check_for_updates(silent=True)
    footer_label.config(text="Up to date")


root.after(2000, run_startup_update_check, None)

root.mainloop()