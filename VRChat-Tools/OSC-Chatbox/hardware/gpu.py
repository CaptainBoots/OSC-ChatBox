"""
hardware/gpu.py
───────────────
GPU name detection (via gpu_ids.py) and individual sensor readers.
"""

import re
import subprocess
import sys
from typing import Optional

from gpu_ids import GPU_ID_MAP, AMBIGUOUS_IDS
from hardware.lhm import hw_nodes, is_gpu, numeric


# ── Detection ─────────────────────────────────────────────────────────────────

_AMD_IGPU_KEYWORDS = ("radeon graphics", "vega", "raphael", "rembrandt", "phoenix", "hawk point")
_AMD_DGPU_KEYWORDS = ("radeon rx", "rx ")


# Priority tiers (lower = better):
#   0 = discrete NVIDIA or AMD discrete
#   1 = unknown AMD (no name match, assume discrete)
#   2 = AMD iGPU (APU)
#   3 = Intel iGPU or unknown
def _vendor_priority(vid: str, name: str = "") -> int:
    n = name.lower()
    if vid == "10de":
        return 0
    if vid == "1002":
        if any(k in n for k in _AMD_DGPU_KEYWORDS):
            return 0
        if any(k in n for k in _AMD_IGPU_KEYWORDS):
            return 2
        return 1
    return 3


def _pci_id() -> Optional[str]:
    if sys.platform == "win32":
        try:
            out = subprocess.check_output(
                ["powershell", "-NoProfile", "-Command",
                 "Get-CimInstance Win32_VideoController | Select-Object -First 100 PNPDeviceID,Name"
                 " | Format-List"],
                encoding="utf-8", stderr=subprocess.DEVNULL, timeout=5,
            )
            # Parse blocks of "PNPDeviceID : ...\nName         : ..."
            best_pid, best_pri = None, 99
            blocks = re.split(r"\n\s*\n", out.strip())
            for block in blocks:
                id_m = re.search(r"VEN_([0-9A-Fa-f]{4}).*DEV_([0-9A-Fa-f]{4})", block)
                name_m = re.search(r"Name\s*:\s*(.+)", block)
                if not id_m:
                    continue
                vid, did = id_m.group(1).lower(), id_m.group(2).lower()
                name = name_m.group(1).strip() if name_m else ""
                pri = _vendor_priority(vid, name)
                if pri < best_pri:
                    best_pri, best_pid = pri, f"{vid}:{did}"
            return best_pid
        except Exception:
            return None
    try:
        out = subprocess.check_output(
            ["lspci", "-nn"], encoding="utf-8", stderr=subprocess.DEVNULL
        )
        best_pid, best_pri = None, 99
        for line in out.splitlines():
            if "VGA" not in line and "3D controller" not in line:
                continue
            m = re.search(r"\[(\w{4}):(\w{4})\]", line)
            if not m:
                continue
            vid, did = m.group(1).lower(), m.group(2).lower()
            pri = _vendor_priority(vid, line)
            if pri < best_pri:
                best_pri, best_pid = pri, f"{vid}:{did}"
        return best_pid
    except Exception:
        return None


def _gpu_name_from_os(pid: str) -> Optional[str]:
    """Ask the OS for the display name of the GPU with the given PCI ID."""
    vid, did = pid.split(":")
    if sys.platform == "win32":
        try:
            out = subprocess.check_output(
                ["powershell", "-NoProfile", "-Command",
                 f"Get-CimInstance Win32_VideoController | Where-Object {{"
                 f" $_.PNPDeviceID -match 'VEN_{vid.upper()}.*DEV_{did.upper()}' }}"
                 f" | Select-Object -ExpandProperty Name"],
                encoding="utf-8", stderr=subprocess.DEVNULL, timeout=5,
            ).strip()
            return out or None
        except Exception:
            return None
    # Linux: name is already in the lspci line, extract the human-readable part
    try:
        out = subprocess.check_output(
            ["lspci", "-nn"], encoding="utf-8", stderr=subprocess.DEVNULL
        )
        for line in out.splitlines():
            if f"[{vid}:{did}]" in line.lower():
                # Strip the address and class prefix, keep everything before the [id] tag
                m = re.match(r"[^:]+:\s+[^:]+:\s+(.+?)\s+\[[\w:]+\]", line)
                if m:
                    return m.group(1).strip()
    except Exception:
        pass
    return None


def detect_gpu() -> str:
    pid = _pci_id()
    if not pid:
        name = _gpu_name_from_os(pid)
        if name:
            return name
    if pid in AMBIGUOUS_IDS:
        name = _gpu_name_from_os(pid)
        if name:
            return name
    if pid in GPU_ID_MAP:
        return GPU_ID_MAP[pid]
    return f"Unknown GPU ({pid})"


def detect_vram_type(gpu_name: str) -> str:
    n = gpu_name.lower()
    if any(x in n for x in ["5090", "5080", "5070", "5060"]):
        return "GDDR7"
    if any(x in n for x in ["4090", "4080", "3090", "3080"]):
        return "GDDR6X"
    if any(x in n for x in ["rx 9", "rx9", "rx 7", "rx7", "rx 6", "rx6", "rx 5", "rx5"]):
        return "GDDR6"
    if any(x in n for x in ["1080"]):
        return "GDDR5X"
    return "GDDR6"


# ── LHM readers ───────────────────────────────────────────────────────────────

def get_gpu_temp(data) -> int:
    if sys.platform != "win32":
        return _linux_gpu_stat("temp")
    best = None
    try:
        for hw in hw_nodes(data):
            if not is_gpu(hw.get("Text", "")):
                continue
            for cat in hw.get("Children", []):
                if "temperature" not in cat.get("Text", "").lower():
                    continue
                for sensor in cat.get("Children", []):
                    st = sensor.get("Text", "").lower()
                    if "distance" in st or "memory" in st:
                        continue
                    if "gpu core" in st or "gpu temperature" in st:
                        try:
                            best = numeric(sensor.get("Value", 0))
                        except ValueError:
                            pass
    except Exception:
        pass
    return int(best) if best is not None else 0


def get_gpu_power(data) -> int:
    if sys.platform != "win32":
        return _linux_gpu_stat("power")
    try:
        for hw in hw_nodes(data):
            if not is_gpu(hw.get("Text", "")):
                continue
            for cat in hw.get("Children", []):
                if "power" not in cat.get("Text", "").lower():
                    continue
                for sensor in cat.get("Children", []):
                    st = sensor.get("Text", "").lower()
                    if any(x in st for x in ("gpu package", "gpu total", "board power")):
                        try:
                            return int(numeric(sensor.get("Value", 0)))
                        except ValueError:
                            pass
    except Exception:
        pass
    return 0


def get_gpu_load(data) -> int:
    if sys.platform != "win32":
        return _linux_gpu_stat("load")
    try:
        for hw in hw_nodes(data):
            if not is_gpu(hw.get("Text", "")):
                continue
            for cat in hw.get("Children", []):
                if "load" not in cat.get("Text", "").lower():
                    continue
                for sensor in cat.get("Children", []):
                    if "gpu core" in sensor.get("Text", "").lower():
                        try:
                            return int(numeric(sensor.get("Value", 0)))
                        except ValueError:
                            pass
    except Exception:
        pass
    return 0


# ── Linux fallback ────────────────────────────────────────────────────────────

def _linux_gpu_stat(kind: str) -> int:
    import glob
    for card in glob.glob("/sys/class/drm/card*/device"):
        if kind == "temp":
            for hwmon in glob.glob(f"{card}/hwmon/hwmon*"):
                try:
                    v = int(open(f"{hwmon}/temp1_input").read().strip())
                    return v // 1000
                except (OSError, ValueError):
                    pass
        elif kind == "power":
            for hwmon in glob.glob(f"{card}/hwmon/hwmon*"):
                for pf in ("power1_average", "power1_input"):
                    try:
                        v = int(open(f"{hwmon}/{pf}").read().strip())
                        return v // 1_000_000
                    except (OSError, ValueError):
                        pass
        elif kind == "load":
            try:
                return int(open(f"{card}/gpu_busy_percent").read().strip())
            except (OSError, ValueError):
                pass
    return 0