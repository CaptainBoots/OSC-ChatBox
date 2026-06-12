"""
hardware/gpu.py
───────────────
GPU name detection (via gpu_ids.py) and individual sensor readers.
"""

import re
import subprocess
import sys
from typing import Optional

from gpu_ids import GPU_ID_MAP
from hardware.lhm import hw_nodes, is_gpu, numeric


# ── Detection ─────────────────────────────────────────────────────────────────

def _pci_id() -> Optional[str]:
    if sys.platform == "win32":
        try:
            out = subprocess.check_output(
                ["powershell", "-NoProfile", "-Command",
                 "Get-CimInstance Win32_VideoController | Select-Object -ExpandProperty PNPDeviceID"],
                encoding="utf-8", stderr=subprocess.DEVNULL, timeout=5,
            )
            intel = None
            for line in out.splitlines():
                m = re.search(r"VEN_([0-9A-Fa-f]{4}).*DEV_([0-9A-Fa-f]{4})", line)
                if not m:
                    continue
                vid, did = m.group(1).lower(), m.group(2).lower()
                pid = f"{vid}:{did}"
                if vid == "8086":
                    intel = pid
                else:
                    return pid
            return intel
        except Exception:
            return None
    try:
        out = subprocess.check_output(
            ["lspci", "-nn"], encoding="utf-8", stderr=subprocess.DEVNULL
        )
        intel = None
        for line in out.splitlines():
            if "VGA" in line or "3D controller" in line:
                m = re.search(r"\[(\w{4}:\w{4})]", line)
                if m:
                    pid = m.group(1).lower()
                    if pid.startswith("8086"):
                        intel = pid
                    else:
                        return pid
        return intel
    except Exception:
        return None


def detect_gpu() -> str:
    pid = _pci_id()
    if pid and pid in GPU_ID_MAP:
        return GPU_ID_MAP[pid]
    return f"Unknown GPU ({pid})" if pid else "GPU Unknown"


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
