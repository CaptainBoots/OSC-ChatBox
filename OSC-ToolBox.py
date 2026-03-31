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

VERSION = "8.0.2"
GITHUB_RAW_URL = "https://raw.githubusercontent.com/CaptainBoots/OSC-ChatBox/main/OSC-ToolBox.py"


print("OSC ToolBox")
print("Made By Boots")

def rename_self_to_toolbox():
    try:
        current_path = os.path.abspath(__file__)
        directory = os.path.dirname(current_path)

        # Target filename
        new_name = "OSC-ToolBox.py"
        new_path = os.path.join(directory, new_name)

        # If already named correctly, do nothing
        if os.path.basename(current_path) == new_name:
            return

        # If target already exists, avoid overwriting
        if os.path.exists(new_path):
            print("[Rename] Target file already exists, skipping rename.")
            return

        os.rename(current_path, new_path)
        print(f"[Rename] Renamed script to {new_name}")

    except Exception as e:
        print(f"[Rename] Failed: {e}")

rename_self_to_toolbox()

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
CONFIG_DIR = os.path.join(SCRIPT_DIR, "OSC-ToolBox")
CONFIG_FILE = os.path.join(CONFIG_DIR, "osc_config.json")
FACE_TRACKING_DEBUGGER_NAME = "OSC-FaceTrackingController.py"
FACE_TRACKING_DEBUGGER_SOURCE_URL = (
    "https://raw.githubusercontent.com/CaptainBoots/OSC-ChatBox/main/OSC-FaceTrackingController.py"
)
FACE_TRACKING_DEBUGGER_BUNDLED_SCRIPT = os.path.join(SCRIPT_DIR, FACE_TRACKING_DEBUGGER_NAME)
FACE_TRACKING_DEBUGGER_SCRIPT = os.path.join(CONFIG_DIR, FACE_TRACKING_DEBUGGER_NAME)

CHATBOX_NAME = "OSC-Chatbox.py"
CHATBOX_SOURCE_URL = (
    "https://raw.githubusercontent.com/CaptainBoots/OSC-ChatBox/main/OSC-Chatbox.py"
)
CHATBOX_BUNDLED_SCRIPT = os.path.join(SCRIPT_DIR, CHATBOX_NAME)
CHATBOX_SCRIPT = os.path.join(CONFIG_DIR, CHATBOX_NAME)

face_tracking_debugger_process = None
chatbox_process = None


def rename_self_to_toolbox():
    try:
        current_path = os.path.abspath(__file__)
        directory = os.path.dirname(current_path)

        # Target filename
        new_name = "OSC-ToolBox.py"
        new_path = os.path.join(directory, new_name)

        # If already named correctly, do nothing
        if os.path.basename(current_path) == new_name:
            return

        # If target already exists, avoid overwriting
        if os.path.exists(new_path):
            print("[Rename] Target file already exists, skipping rename.")
            return

        os.rename(current_path, new_path)
        print(f"[Rename] Renamed script to {new_name}")

    except Exception as e:
        print(f"[Rename] Failed: {e}")


def ensure_face_tracking_debugger_script(show_errors=False):
    if os.path.isfile(FACE_TRACKING_DEBUGGER_SCRIPT):
        return True

    try:
        os.makedirs(CONFIG_DIR, exist_ok=True)
    except OSError as e:
        print(f"[Face Controller] Could not create config directory: {e}")
        if show_errors:
            messagebox.showerror("Face Controller Error", f"Could not create config directory:\n{e}")
        return False

    remote_text = None
    try:
        response = requests.get(
            FACE_TRACKING_DEBUGGER_SOURCE_URL,
            timeout=10,
            params={"_": int(time.time())},
            headers={"Cache-Control": "no-cache", "Pragma": "no-cache"},
        )
        response.raise_for_status()
        remote_text = response.text
    except requests.RequestException as e:
        print(f"[Face Controller] Download failed: {e}")

    if remote_text:
        try:
            with open(FACE_TRACKING_DEBUGGER_SCRIPT, "w", encoding="utf-8") as f:
                f.write(remote_text)
            print(f"[Face Controller] Downloaded to {FACE_TRACKING_DEBUGGER_SCRIPT}")
            return True
        except OSError as e:
            print(f"[Face Controller] Could not save downloaded script: {e}")

    if os.path.isfile(FACE_TRACKING_DEBUGGER_BUNDLED_SCRIPT):
        try:
            with open(FACE_TRACKING_DEBUGGER_BUNDLED_SCRIPT, "r", encoding="utf-8") as src:
                with open(FACE_TRACKING_DEBUGGER_SCRIPT, "w", encoding="utf-8") as dst:
                    dst.write(src.read())
            print(
                "[Face Controller] Using bundled fallback copy "
                f"from {FACE_TRACKING_DEBUGGER_BUNDLED_SCRIPT}"
            )
            return True
        except OSError as e:
            print(f"[Face Controller] Could not copy bundled fallback: {e}")

    if show_errors:
        messagebox.showerror(
            "Face Controller Error",
            "Could not download Face Tracking Controller.\n"
            "Check your internet connection and try again.",
        )
    return False

def ensure_chatbox_script(show_errors=False):
    if os.path.isfile(CHATBOX_SCRIPT):
        return True

    try:
        os.makedirs(CONFIG_DIR, exist_ok=True)
    except OSError as e:
        print(f"[Chatbox] Could not create config directory: {e}")
        if show_errors:
            messagebox.showerror("Chatbox Error", f"Could not create config directory:\n{e}")
        return False

    remote_text = None
    try:
        response = requests.get(
            CHATBOX_SOURCE_URL,
            timeout=10,
            params={"_": int(time.time())},
            headers={"Cache-Control": "no-cache", "Pragma": "no-cache"},
        )
        response.raise_for_status()
        remote_text = response.text
    except requests.RequestException as e:
        print(f"[Chatbox] Download failed: {e}")

    if remote_text:
        try:
            with open(CHATBOX_SCRIPT, "w", encoding="utf-8") as f:
                f.write(remote_text)
            print(f"[Chatbox] Downloaded to {CHATBOX_SCRIPT}")
            return True
        except OSError as e:
            print(f"[Chatbox] Could not save downloaded script: {e}")

    if os.path.isfile(CHATBOX_SCRIPT):
        try:
            with open(CHATBOX_SCRIPT, "r", encoding="utf-8") as src:
                with open(CHATBOX_SCRIPT, "w", encoding="utf-8") as dst:
                    dst.write(src.read())
            print(
                "[Chatbox] Using bundled fallback copy "
                f"from {CHATBOX_SCRIPT}"
            )
            return True
        except OSError as e:
            print(f"[Chatbox] Could not copy bundled fallback: {e}")

    if show_errors:
        messagebox.showerror(
            "Chatbox Error",
            "Could not download Chatbox.\n"
            "Check your internet connection and try again.",
        )
    return False


# ═════════════════════════════════════════════════════════════════════════════════════════════════════════════════════#
# UPDATER
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
        current_best_version = best["version"] if best is not None else "0.0.0"
        if _parse_version(info["version"]) > _parse_version(current_best_version):
            best = info

    if best is not None:
        return best

    print(f"[Updater] Could not reach GitHub URLs: {errors}")
    return None


def check_for_face_tracking_debugger_updates(silent=False):
    if not ensure_face_tracking_debugger_script(show_errors=not silent):
        return False

    remote_text, remote_version, remote_url = _fetch_remote_script(
        FACE_TRACKING_DEBUGGER_SOURCE_URL,
        timeout=10,
    )
    if remote_text is None:
        if not silent:
            messagebox.showinfo(
                "Face Controller Update",
                "Could not reach GitHub to check Face Tracking Controller updates."
            )
        return False

    remote_version = remote_version or "0.0.0"

    try:
        with open(FACE_TRACKING_DEBUGGER_SCRIPT, "r", encoding="utf-8") as local_file:
            local_text = local_file.read()
    except OSError:
        local_text = ""

    local_version = _extract_version_from_source(local_text) or "0.0.0"

    if _parse_version(remote_version) <= _parse_version(local_version):
        if silent:
            print(
                f"[Face Controller] Up to date ({local_version}) "
                f"vs remote ({remote_version}) from {remote_url}"
            )
        return False

    try:
        with open(FACE_TRACKING_DEBUGGER_SCRIPT, "w", encoding="utf-8") as local_file:
            local_file.write(remote_text)
        print(
            f"[Face Controller] Updated from {local_version} to {remote_version} "
            f"from {remote_url}"
        )
        if not silent:
            messagebox.showinfo(
                "Face Controller Updated",
                f"Face Tracking Controller updated from {local_version} to {remote_version}."
            )
        return True
    except OSError as e:
        print(f"[Face Controller] Failed to write update: {e}")
        if not silent:
            messagebox.showerror(
                "Face Controller Update Failed",
                f"Could not update Face Tracking Controller:\n{e}"
            )
        return False


def check_for_chatbox_updates(silent=False):
    if not ensure_chatbox_script(show_errors=not silent):
        return False

    remote_text, remote_version, remote_url = _fetch_remote_script(
        CHATBOX_SOURCE_URL,
        timeout=10,
    )
    if remote_text is None:
        if not silent:
            messagebox.showinfo(
                "Chatbox Update",
                "Could not reach GitHub to check Chatbox updates."
            )
        return False

    remote_version = remote_version or "0.0.0"

    try:
        with open(CHATBOX_SCRIPT, "r", encoding="utf-8") as local_file:
            local_text = local_file.read()
    except OSError:
        local_text = ""

    local_version = _extract_version_from_source(local_text) or "0.0.0"

    if _parse_version(remote_version) <= _parse_version(local_version):
        if silent:
            print(
                f"[Chatbox] Up to date ({local_version}) "
                f"vs remote ({remote_version}) from {remote_url}"
            )
        return False

    try:
        with open(CHATBOX_SCRIPT, "w", encoding="utf-8") as local_file:
            local_file.write(remote_text)
        print(
            f"[Chatbox] Updated from {local_version} to {remote_version} "
            f"from {remote_url}"
        )
        if not silent:
            messagebox.showinfo(
                "Chatbox Updated",
                f"Chatbox updated from {local_version} to {remote_version}."
            )
        return True
    except OSError as e:
        print(f"[Chatbox] Failed to write update: {e}")
        if not silent:
            messagebox.showerror(
                "Chatbox Update Failed",
                f"Could not update ChatboxController:\n{e}"
            )
        return False

def perform_update(remote_text=None, source_url=None):
    try:
        if remote_text is None:
            info = get_remote_script_info()
            if not info:
                raise RuntimeError("No remote script source available")
            remote_text = info["text"]
            source_url = info["url"]

        script_path = os.path.abspath(__file__)
        config_dir = os.path.dirname(os.path.abspath(CONFIG_FILE))
        os.makedirs(config_dir, exist_ok=True)
        script_name = os.path.splitext(os.path.basename(script_path))[0]
        backup_path = os.path.join(config_dir, f"{script_name} {VERSION}.bak")

        with open(script_path, "r", encoding="utf-8") as f_cur:
            with open(backup_path, "w", encoding="utf-8") as f_bak:
                f_bak.write(f_cur.read())

        backup_files = [
            os.path.join(config_dir, backup_file_name)
            for backup_file_name in os.listdir(config_dir)
            if backup_file_name.lower().endswith(".bak")
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
        check_for_face_tracking_debugger_updates(silent=silent)
        check_for_chatbox_updates(silent=silent)
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
    if remote_newer or content_differs:
        if remote_newer:
            prompt = (
                f"New version {remote_version} is available (you have {VERSION}).\n\n"
                "Update and restart now?"
            )
        else:
            prompt = (
                "A remote script update is available (content changed,\n"
                "but version string may not have been bumped).\n\n"
                "Update and restart now?"
            )
        answer = messagebox.askyesno(
            "Update Available",
            prompt
        )
        if answer:
            perform_update(remote_text=remote_text, source_url=remote_url)

    if not main_update_available and silent:
        print(f"[Updater] Up to date ({VERSION}) vs remote ({remote_version}) from {remote_url}")

    face_updated = check_for_face_tracking_debugger_updates(silent=silent)

    if not silent and not main_update_available and not face_updated:
        messagebox.showinfo("Up to Date", f"You're on the latest version ({VERSION}).")



# ═════════════════════════════════════════════════════════════════════════════════════════════════════════════════════#
# SCRIPT CONTROL
# ═════════════════════════════════════════════════════════════════════════════════════════════════════════════════════#

def start_face_tracking_debugger():
    global face_tracking_debugger_process

    if not ensure_face_tracking_debugger_script(show_errors=True):
        return

    if face_tracking_debugger_process is not None and face_tracking_debugger_process.poll() is None:
        messagebox.showinfo("Face Controller", "Face Tracking Controller is already running.")
        return

    try:
        face_tracking_debugger_process = subprocess.Popen(
            [sys.executable, FACE_TRACKING_DEBUGGER_SCRIPT],
            cwd=CONFIG_DIR,
        )
    except Exception as e:
        face_tracking_debugger_process = None
        messagebox.showerror("Face Controller Error", f"Failed to start Face Tracking Controller:\n{e}")

def start_chatbox():
    global chatbox_process

    if not ensure_chatbox_script(show_errors=True):
        return

    if chatbox_process is not None and chatbox_process.poll() is None:
        messagebox.showinfo("Face Controller", "Face Tracking Controller is already running.")
        return

    try:
        chatbox_process = subprocess.Popen(
            [sys.executable, CHATBOX_SCRIPT],
            cwd=CONFIG_DIR,
        )
    except Exception as e:
        chatbox_process = None
        messagebox.showerror("Face Controller Error", f"Failed to start Face Tracking Controller:\n{e}")

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
square_widgets = []  # store square buttons separately

root = tk.Tk()
root.title("OSC ToolBox")
root.geometry("560x620")
root.minsize(520, 560)
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
    new_default  = max(7, int(base_default * scale))

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

bottom_bar = tk.Frame(frame, bg=BG)
bottom_bar.grid(row=15, column=0, columnspan=2, pady=6, sticky="ew")

bottom_bar.columnconfigure(1, weight=2)  # center (TOOLS bigger)


Face_btn = tk.Button(
    bottom_bar,
    text="Face Tracking Controller",
    command=start_face_tracking_debugger,
    bg=ACCENT,
    fg="#FFFFFF",
    relief="flat",
    activebackground=ACCENT2,
    activeforeground="#FFFFFF",
    cursor="hand2",
    font=(UI_FONT, 10, "bold"),
    padx=18,
    pady=6
)

Chat_btn = tk.Button(
    bottom_bar,
    text="ChatBox",
    command=start_chatbox,
    bg=ACCENT,
    fg="#FFFFFF",
    relief="flat",
    activebackground=ACCENT2,
    activeforeground="#FFFFFF",
    cursor="hand2",
    font=(UI_FONT, 10, "bold"),
    padx=18,
    pady=6
)

Face_btn.grid(row=0, column=1, padx=6)
Chat_btn.grid(row=2, column=1, padx=6)


# ── Startup update check ───────────────────────────────────────────────────
def run_startup_update_check(_unused=None):
    check_for_updates(silent=True)


root.after(2000, run_startup_update_check, None)

root.mainloop()