# ═════════════════════════════════════════════════════════════════════════════════════════════════════════════════════#
#                                              OSC Python Script                                                       #
# ═════════════════════════════════════════════════════════════════════════════════════════════════════════════════════#
# Hi :3
# Wellcome to my code

# ═════════════════════════════════════════════════════════════════════════════════════════════════════════════════════#
# Imports
# ═════════════════════════════════════════════════════════════════════════════════════════════════════════════════════#

import subprocess
import sys
import importlib
import re
import time
import tkinter as tk
from tkinter import messagebox
import tkinter.font as font
import os
import site

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

VERSION = "8.1.3"
GITHUB_RAW_URL = "https://raw.githubusercontent.com/CaptainBoots/OSC-ChatBox/main/OSC-ToolBox.py"
GITHUB_BASE_URL = "https://raw.githubusercontent.com/CaptainBoots/OSC-ChatBox/main/OSC-Tools/"


# ─────────────────────────────────────────────────────────────────────────────
MANAGED_SCRIPTS = [
    {"filename": "OSC-Chatbox.py",                     "label": "ChatBox"},
    {"filename": "OSC-FaceTrackingController(Beta).py", "label": "Face Tracking Controller"},
    {"filename": "OSC-Router.py", "label": "Router"},
]
# ─────────────────────────────────────────────────────────────────────────────


print("OSC ToolBox")
print("Made By Boots")
print(f"Version {VERSION}")


def rename_self_to_toolbox():
    try:
        current_path = os.path.abspath(__file__)
        directory = os.path.dirname(current_path)
        new_name = "OSC-Tools.py"
        new_path = os.path.join(directory, new_name)
        if os.path.basename(current_path) == new_name:
            return
        if os.path.exists(new_path):
            print("[Rename] Target file already exists, skipping rename.")
            return
        os.rename(current_path, new_path)
        print(f"[Rename] Renamed script to {new_name}")
    except Exception as e:
        print(f"[Rename] Failed: {e}")

rename_self_to_toolbox()

SCRIPT_DIR  = os.path.dirname(os.path.abspath(__file__))
CONFIG_DIR  = os.path.join(SCRIPT_DIR, "OSC-Tools")
CONFIG_FILE = os.path.join(CONFIG_DIR, "osc_config.json")
BACKUP_DIR  = os.path.join(CONFIG_DIR, "ToolBox Backup")   # ← all .bak files live here

# Running process handles, keyed by filename
_processes: dict[str, subprocess.Popen | None] = {}


# ═════════════════════════════════════════════════════════════════════════════════════════════════════════════════════#
# MANAGED SCRIPT HELPERS
# ═════════════════════════════════════════════════════════════════════════════════════════════════════════════════════#

def _script_paths(filename: str) -> tuple[str, str]:
    """Return (bundled_path, config_dir_path) for a managed script filename."""
    return (
        os.path.join(SCRIPT_DIR, filename),
        os.path.join(CONFIG_DIR, filename),
    )


def ensure_script(filename: str, show_errors: bool = False) -> bool:
    """
    Make sure *filename* exists in CONFIG_DIR.
    Tries: already present → download from GitHub → bundled copy.
    Returns True if the script is ready to run.
    """
    bundled_path, dest_path = _script_paths(filename)

    if os.path.isfile(dest_path):
        return True

    try:
        os.makedirs(CONFIG_DIR, exist_ok=True)
    except OSError as e:
        print(f"[{filename}] Could not create config directory: {e}")
        if show_errors:
            messagebox.showerror(f"{filename} Error", f"Could not create config directory:\n{e}")
        return False

    url = GITHUB_BASE_URL + filename
    remote_text = None
    try:
        response = requests.get(
            url,
            timeout=10,
            params={"_": int(time.time())},
            headers={"Cache-Control": "no-cache", "Pragma": "no-cache"},
        )
        response.raise_for_status()
        remote_text = response.text
    except requests.RequestException as e:
        print(f"[{filename}] Download failed: {e}")

    if remote_text:
        try:
            with open(dest_path, "w", encoding="utf-8") as f:
                f.write(remote_text)
            print(f"[{filename}] Downloaded to {dest_path}")
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
            f"Could not download {filename}.\nCheck your internet connection and try again.",
        )
    return False


def check_for_script_updates(filename: str, silent: bool = False) -> bool:
    """
    Check GitHub for a newer version of *filename* and update if found.
    Returns True if an update was applied.
    """
    if not ensure_script(filename, show_errors=not silent):
        return False

    _, dest_path = _script_paths(filename)
    url = GITHUB_BASE_URL + filename

    remote_text, remote_version, _ = _fetch_remote_script(url, timeout=10)
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
        if not silent:
            messagebox.showinfo(
                f"{filename} Updated",
                f"{filename} updated from {local_version} to {remote_version}."
            )
        return True
    except OSError as e:
        print(f"[{filename}] Failed to write update: {e}")
        if not silent:
            messagebox.showerror(
                f"{filename} Update Failed",
                f"Could not update {filename}:\n{e}"
            )
        return False


def launch_script(filename: str) -> None:
    """Launch *filename* as a subprocess, or warn if already running."""
    global _processes

    if not ensure_script(filename, show_errors=True):
        return

    proc = _processes.get(filename)
    if proc is not None and proc.poll() is None:
        messagebox.showinfo(filename, f"{filename} is already running.")
        return

    _, dest_path = _script_paths(filename)
    try:
        _processes[filename] = subprocess.Popen(
            [sys.executable, dest_path],
            cwd=CONFIG_DIR,
        )
    except Exception as e:
        _processes[filename] = None
        messagebox.showerror(f"{filename} Error", f"Failed to start {filename}:\n{e}")


# ═════════════════════════════════════════════════════════════════════════════════════════════════════════════════════#
# UPDATER  (self-update for OSC-Tools.py itself)
# ═════════════════════════════════════════════════════════════════════════════════════════════════════════════════════#

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
        if _parse_version(info["version"]) > _parse_version(best["version"]):
            best = info

    if best is not None:
        return best

    print(f"[Updater] Could not reach GitHub URLs: {errors}")
    return None


def perform_update(remote_text=None, source_url=None):
    try:
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
        for entry in MANAGED_SCRIPTS:
            check_for_script_updates(entry["filename"], silent=silent)
        return

    remote_version = info["version"]
    remote_text = info["text"]
    remote_url = info["url"]

    try:
        with open(os.path.abspath(__file__), "r", encoding="utf-8") as f:
            local_text = f.read()
    except OSError:
        local_text = ""

    local_norm = local_text.replace("\r\n", "\n")
    remote_norm = remote_text.replace("\r\n", "\n")
    remote_newer = _parse_version(remote_version) > _parse_version(VERSION)
    content_differs = remote_norm != local_norm
    main_update_available = remote_newer or content_differs

    print(f"[OSC-Tools] Checking... (local: {VERSION}  remote: {remote_version})")

    if main_update_available:
        if remote_newer:
            print(f"[OSC-Tools] Update available: {VERSION} -> {remote_version}")
            prompt = (
                f"New version {remote_version} is available (you have {VERSION}).\n\n"
                "Update and restart now?"
            )
        else:
            print(f"[OSC-Tools] Remote content differs (version string unchanged at {VERSION})")
            prompt = (
                "A remote script update is available (content changed,\n"
                "but version string may not have been bumped).\n\n"
                "Update and restart now?"
            )
        if messagebox.askyesno("Update Available", prompt):
            perform_update(remote_text=remote_text, source_url=remote_url)
        else:
            print(f"[OSC-Tools] Update skipped by user")
    else:
        print(f"[OSC-Tools] Up to date ({VERSION})")

    any_tool_updated = False
    for entry in MANAGED_SCRIPTS:
        if check_for_script_updates(entry["filename"], silent=silent):
            any_tool_updated = True

    if not silent and not main_update_available and not any_tool_updated:
        messagebox.showinfo("Up to Date", f"You're on the latest version ({VERSION}).")


# ═════════════════════════════════════════════════════════════════════════════════════════════════════════════════════#
# GUI
# ═════════════════════════════════════════════════════════════════════════════════════════════════════════════════════#

BG = "#0f0f13"
PANEL = "#17171f"
BORDER = "#2a2a38"
ACCENT = "#7c5cfc"
ACCENT2 = "#a78bfa"
TEXT = "#e2e0f0"
SUBTEXT = "#7e7b9a"
GREEN = "#4ade80"
RED = "#f87171"
FG = TEXT
ENTRY_BG = PANEL
BTN_BG = PANEL
BTN_FG = TEXT
UI_FONT = "Consolas"

ui_scale = 1.0
scalable_widgets = []
square_widgets = []

root = tk.Tk()
root.title("OSC ToolBox")
root.configure(bg=BG)
root.resizable(True, True)

title_bar = tk.Frame(root, bg=PANEL, pady=10)
title_bar.pack(fill="x")

header_title_label = tk.Label(
    title_bar,
    text="◈  OSC TOOLBOX",
    bg=PANEL,
    fg=ACCENT2,
    font=(UI_FONT, 13, "bold")
)
header_title_label.pack(side="left", padx=16)

tk.Frame(root, bg=BORDER, height=1).pack(fill="x")

frame = tk.Frame(root, bg=BG)
frame.pack(fill="both", expand=True, padx=12, pady=10)


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

    for container, base_size, btn in square_widgets:
        size = int(base_size * scale)
        container.config(width=size, height=size)
        btn.config(font=(UI_FONT, max(8, int(12 * scale))))


def dark_label(text, r):
    lbl = tk.Label(frame, text=text, bg=BG, fg=SUBTEXT, anchor="w", font=(UI_FONT, 9))
    lbl.grid(row=r, column=0, sticky="w", pady=4)
    return lbl


def dark_entry(r, default=""):
    e = tk.Entry(
        frame,
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
    e.grid(row=r, column=1, pady=4, sticky="ew")
    return e


def square_button(parent, text, command, base_size=32):
    container = tk.Frame(parent, bg=BTN_BG, highlightthickness=1, highlightbackground=BORDER)
    container.pack_propagate(False)

    btn = tk.Button(
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
    btn.pack(fill="both", expand=True)

    square_widgets.append((container, base_size, btn))
    container.config(width=base_size, height=base_size)

    return container


frame.columnconfigure(1, weight=1)

# ── Tool buttons (auto-generated from MANAGED_SCRIPTS) ────────────────────
bottom_bar = tk.Frame(frame, bg=BG)
bottom_bar.grid(row=15, column=0, columnspan=2, pady=6, sticky="ew")
bottom_bar.columnconfigure(0, weight=1)

for i, entry in enumerate(MANAGED_SCRIPTS):
    filename = entry["filename"]
    label = entry["label"]

    btn = tk.Button(
        bottom_bar,
        text=label,
        command=lambda fn=filename: launch_script(fn),
        bg=ACCENT,
        fg="#FFFFFF",
        relief="flat",
        activebackground=ACCENT2,
        activeforeground="#FFFFFF",
        cursor="hand2",
        font=(UI_FONT, 10, "bold"),
        padx=18,
        pady=6,
    )
    btn.grid(row=i, column=0, padx=6, pady=4, sticky="ew")


# ── Startup update check ───────────────────────────────────────────────────
def run_startup_update_check(_unused=None):
    check_for_updates(silent=True)


btn_count = len(MANAGED_SCRIPTS)
root.geometry(f"560x{460 + btn_count * 48}")
root.minsize(520, 400 + btn_count * 48)

root.after(2000, run_startup_update_check, None)

root.mainloop()