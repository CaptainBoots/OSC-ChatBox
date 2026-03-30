"""
Legacy compatibility shim.

This file exists so older builds that still fetch the legacy entry script
are redirected to the `OSC-PC.py` script.
"""

import os
import subprocess
import sys
import tempfile
import urllib.error
import urllib.request


VERSION = "7.3.2"
TARGET_NAME = "OSC-PC.py"
TARGET_URL = "https://raw.githubusercontent.com/CaptainBoots/OSC-ChatBox/main/OSC-PC.py"


def _download_text(url, timeout=20):
    request = urllib.request.Request(
        url,
        headers={"Cache-Control": "no-cache", "Pragma": "no-cache"},
    )
    with urllib.request.urlopen(request, timeout=timeout) as response:
        return response.read().decode("utf-8")


def _write_atomic(path, content):
    directory = os.path.dirname(path) or "."
    fd, temp_path = tempfile.mkstemp(prefix="osc_pc_", suffix=".tmp", dir=directory)
    try:
        with os.fdopen(fd, "w", encoding="utf-8", newline="\n") as temp_file:
            temp_file.write(content)
        os.replace(temp_path, path)
    finally:
        if os.path.exists(temp_path):
            os.remove(temp_path)


def _launch(script_path):
    subprocess.Popen([sys.executable, script_path])


def main():
    script_dir = os.path.dirname(os.path.abspath(__file__))
    target_path = os.path.join(script_dir, TARGET_NAME)

    try:
        remote_source = _download_text(TARGET_URL)
        if "def run_osc_loop" not in remote_source:
            raise RuntimeError("Downloaded file does not look like OSC-PC.py")

        _write_atomic(target_path, remote_source)
        print("[Legacy Redirect] OSC-PC.py updated. Launching...")
        _launch(target_path)
        return 0

    except (urllib.error.URLError, RuntimeError, OSError) as error:
        if os.path.exists(target_path):
            print(f"[Legacy Redirect] Download failed ({error}). Launching cached OSC-PC.py.")
            _launch(target_path)
            return 0

        print(f"[Legacy Redirect] Failed to fetch OSC-PC.py: {error}")
        print(f"[Legacy Redirect] Please download manually: {TARGET_URL}")
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
