"""
hardware/memory.py
──────────────────
DRAM and VRAM individual stat readers.
"""

import subprocess
import sys
from typing import Optional

from hardware.lhm import hw_nodes, is_cpu, is_gpu, numeric


def detect_dram_type() -> str:
    if sys.platform != "win32":
        return "DDR"
    try:
        out = subprocess.check_output(
            ["powershell", "-Command",
             "(Get-CimInstance Win32_PhysicalMemory | Select-Object -First 1).SMBIOSMemoryType"],
            encoding="utf-8", stderr=subprocess.DEVNULL, timeout=5,
        ).strip()
        return {"24": "DDR3", "26": "DDR4", "34": "DDR5", "35": "DDR5"}.get(out, "DDR")
    except Exception:
        return "DDR"


def _psutil_ram():
    import psutil
    vm = psutil.virtual_memory()
    return round(vm.used / (1024**3), 1), _fmt_gb(vm.total / (1024**3))


def _fmt_gb(gb: float) -> str:
    return str(int(gb)) if float(gb).is_integer() else f"{gb:.1f}"


def get_dram_used(data) -> float:
    if sys.platform != "win32" or not data:
        return _psutil_ram()[0]
    try:
        for hw in hw_nodes(data):
            if not is_cpu(hw.get("Text", "")):
                continue
            for cat in hw.get("Children", []):
                for sensor in cat.get("Children", []):
                    if "memory used" in sensor.get("Text", "").lower() and \
                       "virtual" not in sensor.get("Text", "").lower():
                        return round(numeric(sensor.get("Value", 0)), 1)
    except Exception:
        pass
    return _psutil_ram()[0]


def get_dram_total(data) -> str:
    import psutil
    return _fmt_gb(psutil.virtual_memory().total / (1024**3))


def get_vram_used(data) -> float:
    if sys.platform != "win32" or not data:
        return _linux_vram()[0]
    try:
        for hw in hw_nodes(data):
            if not is_gpu(hw.get("Text", "")):
                continue
            for cat in hw.get("Children", []):
                for sensor in cat.get("Children", []):
                    if "gpu memory used" in sensor.get("Text", "").lower():
                        raw = numeric(sensor.get("Value", 0))
                        gb = raw / 1024 if raw > 500 else raw
                        return round(gb, 1)
    except Exception:
        pass
    return 0.0


def get_vram_total(data) -> str:
    if sys.platform != "win32" or not data:
        total = _linux_vram()[1]
        return _fmt_gb(total) if total else "?"
    try:
        for hw in hw_nodes(data):
            if not is_gpu(hw.get("Text", "")):
                continue
            total_mb = used_mb = free_mb = None
            for cat in hw.get("Children", []):
                for sensor in cat.get("Children", []):
                    st = sensor.get("Text", "").lower()
                    try:
                        val = numeric(sensor.get("Value", 0))
                    except ValueError:
                        continue
                    if "gpu memory total" in st:
                        total_mb = val
                    elif "gpu memory used" in st:
                        used_mb = val
                    elif "gpu memory free" in st:
                        free_mb = val
            if total_mb and total_mb > 0:
                gb = total_mb / 1024 if total_mb > 500 else total_mb
                return _fmt_gb(gb)
            if used_mb is not None and free_mb is not None:
                total = used_mb + free_mb
                gb = total / 1024 if total > 500 else total
                return _fmt_gb(gb)
    except Exception:
        pass
    return "?"


def _linux_vram():
    import glob
    for card in glob.glob("/sys/class/drm/card*/device"):
        try:
            used  = int(open(f"{card}/mem_info_vram_used").read().strip())
            total = int(open(f"{card}/mem_info_vram_total").read().strip())
            return round(used / (1024**3), 1), total / (1024**3)
        except (OSError, ValueError):
            pass
    return 0.0, None
