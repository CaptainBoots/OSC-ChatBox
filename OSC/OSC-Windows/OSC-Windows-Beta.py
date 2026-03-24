# ═════════════════════════════════════════════════════════════════════════════════════════════════════════════════════#
#                                              OSC Python Script                                                      #
# ═════════════════════════════════════════════════════════════════════════════════════════════════════════════════════#
# Hi :3
# Wellcome to my code

# ═════════════════════════════════════════════════════════════════════════════════════════════════════════════════════#
# Imports
# ═════════════════════════════════════════════════════════════════════════════════════════════════════════════════════#

import subprocess
import sys
import importlib
import json
import asyncio
import re
import threading
import time
import tkinter as tk
from tkinter import messagebox
from enum import Enum

def install_if_missing(package, import_name=None):
    if import_name is None:
        import_name = package.split("==")[0].replace("-", "_")

    try:
        importlib.import_module(import_name)
    except ImportError:
        print(f"Installing {package}...")
        subprocess.check_call([sys.executable, "-m", "pip", "install", package])

install_if_missing("python-osc==1.9.3", "pythonosc")
install_if_missing("psutil==7.2.2", "psutil")
install_if_missing("winrt-Windows.Media.Control==3.2.1", "winrt")
install_if_missing("winrt-windows.foundation==3.2.1", "winrt.windows.foundation")
install_if_missing("requests==2.32.5", "requests")

import psutil
from pythonosc.udp_client import SimpleUDPClient
import winrt.windows.media.control as wmc
import requests

# ═════════════════════════════════════════════════════════════════════════════════════════════════════════════════════#
# CONFIGURATION & GLOBAL VARIABLES
# ═════════════════════════════════════════════════════════════════════════════════════════════════════════════════════#

class CPUManufacturer(Enum):
    INTEL = "Intel"
    AMD = "AMD"
    UNKNOWN = "Unknown"


print("OSC Chatbox")
print("Made By Boots")

OSC_IP = "127.0.0.1"
OSC_PORT = 9000

INTERFACE = "Ethernet"

SWITCH_INTERVAL = 20

LHM_REST_API = "http://localhost:8085/data.json"

client = None
running = False

page1_line1_text = "-enter text-"
page2_line1_text = "-enter text-"
page3_line1_text = "-enter text-"
page4_line1_text = "-enter text-"
error_text = "Error: Page error value exceeding limit"

cpu_wattage = "error"
cpu_temp = "error"
gpu_wattage = "error"
gpu_temp = "error"
dram_load = "error"
vram_load = "error"

cpu_manufacturer = CPUManufacturer.UNKNOWN

# Weather globals
weather_temp     = "?"
weather_humidity = "?"
weather_desc     = "Unknown"

# ═════════════════════════════════════════════════════════════════════════════════════════════════════════════════════#
# WMO WEATHER CODE MAP
# ═════════════════════════════════════════════════════════════════════════════════════════════════════════════════════#

WMO_CODES = {
    0:  "Clear sky",
    1:  "Mainly clear",
    2:  "Partly cloudy",
    3:  "Overcast",
    45: "Foggy",
    48: "Icy fog",
    51: "Light drizzle",
    53: "Drizzle",
    55: "Heavy drizzle",
    61: "Light rain",
    63: "Rain",
    65: "Heavy rain",
    66: "Freezing rain",
    67: "Heavy freezing rain",
    71: "Light snow",
    73: "Snow",
    75: "Heavy snow",
    77: "Snow grains",
    80: "Light showers",
    81: "Showers",
    82: "Heavy showers",
    85: "Snow showers",
    86: "Heavy snow showers",
    95: "Thunderstorm",
    96: "Thunderstorm w/ hail",
    99: "Thunderstorm w/ heavy hail",
}

# ═════════════════════════════════════════════════════════════════════════════════════════════════════════════════════#
# CONFIG FILE
# ═════════════════════════════════════════════════════════════════════════════════════════════════════════════════════#

CONFIG_FILE = "osc_config.json"

def load_config():
    defaults = {
        "osc_ip":          OSC_IP,
        "osc_port":        OSC_PORT,
        "interface":       INTERFACE,
        "switch_interval": SWITCH_INTERVAL,
        "lhm_api":         LHM_REST_API,
        "location":        "0,0",
        "page1_text":      "Thx for using Boots's osc code",
        "page2_text":      "Join the discord server at https://discord.gg/XdfKAWu6Ph",
        "page3_text":      "hi put your text here :3",
        "page4_text":      "Local Weather",
    }
    try:
        with open(CONFIG_FILE, "r") as f:
            return {**defaults, **json.load(f)}
    except (FileNotFoundError, json.JSONDecodeError):
        return defaults

def save_config():
    config = {
        "osc_ip":          ip_entry.get(),
        "osc_port":        port_entry.get(),
        "interface":       iface_entry.get(),
        "switch_interval": interval_entry.get(),
        "lhm_api":         lhm_entry.get(),
        "location":        location_entry.get(),
        "page1_text":      page1_entry.get(),
        "page2_text":      page2_entry.get(),
        "page3_text":      page3_entry.get(),
        "page4_text":      page4_entry.get(),
    }
    with open(CONFIG_FILE, "w") as f:
        json.dump(config, f, indent=2)

# ═════════════════════════════════════════════════════════════════════════════════════════════════════════════════════#
# HARDWARE MONITORING
# ═════════════════════════════════════════════════════════════════════════════════════════════════════════════════════#

def _clean_name(name: str):
    name = re.sub(r"\(.*?\)|\[.*?]|\{.*?}", "", name)
    name = name.split("@")[0]
    name = re.sub(r"\s+", " ", name).strip()
    return name

def detect_cpu():
    global cpu_manufacturer

    try:
        cpu_name = subprocess.check_output(
            ["powershell", "-Command", "(Get-CimInstance Win32_Processor).Name"],
            encoding="utf-8",
            stderr=subprocess.DEVNULL,
            timeout=5
        ).strip()

        clean_cpu_name = _clean_name(cpu_name)

        if "intel" in cpu_name.lower():
            cpu_manufacturer = CPUManufacturer.INTEL
        elif "amd" in cpu_name.lower():
            cpu_manufacturer = CPUManufacturer.AMD
        else:
            cpu_manufacturer = CPUManufacturer.UNKNOWN

        return clean_cpu_name
    except (subprocess.CalledProcessError, UnicodeDecodeError, subprocess.TimeoutExpired):
        return "CPU Unknown"


def detect_gpu():
    try:
        gpu_name = subprocess.check_output(
            ["powershell", "-Command", "(Get-CimInstance Win32_VideoController).Name"],
            encoding="utf-8",
            stderr=subprocess.DEVNULL,
            timeout=5
        ).strip()

        gpu_lines = [
            line for line in gpu_name.splitlines()
            if "virtual desktop" not in line.lower() and "virtual monitor" not in line.lower()
        ]
        gpu_name = "\n".join(gpu_lines).strip()

        return _clean_name(gpu_name)
    except (subprocess.CalledProcessError, UnicodeDecodeError, subprocess.TimeoutExpired):
        return "GPU Unknown"


# ═════════════════════════════════════════════════════════════════════════════════════════════════════════════════════#
#  LHM HELPERS
# ═════════════════════════════════════════════════════════════════════════════════════════════════════════════════════#

def _numeric(sensor_value) -> float:
    s = re.sub(r'[^\d.-]', '', str(sensor_value))
    if not s or s in ('.', '-'):
        raise ValueError(f"No numeric content in: {sensor_value!r}")
    return float(s)


def _walk_sensors(node):
    children = node.get("Children", [])
    if not children:
        yield node
    else:
        for child in children:
            yield from _walk_sensors(child)


def _is_cpu(text: str) -> bool:
    t = text.lower()
    return ("intel" in t or "amd" in t) and "radeon" not in t


def _is_gpu(text: str) -> bool:
    t = text.lower()
    return "radeon" in t or "nvidia" in t or "geforce" in t or "rtx" in t or "gtx" in t or "rx " in t


def _get_hardware_nodes(data):
    for top in data.get("Children", []):          # computer
        for hw in top.get("Children", []):         # hardware
            yield hw

def diagnose_lhm():
    try:
        response = requests.get(LHM_REST_API, timeout=5)
        if response.status_code == 200:
            print("[DIAGNOSTIC] ✓ API Connection: SUCCESS")
            data = response.json()
            print(f"[DIAGNOSTIC] ✓ Sensors Found")
            return True
        else:
            print(f"[DIAGNOSTIC] ✗ API returned status: {response.status_code}")
    except requests.ConnectionError:
        print("[DIAGNOSTIC] ✗ Cannot connect to LibreHardwareMonitor REST API")
        print("[DIAGNOSTIC] FIX 1: Make sure LibreHardwareMonitor.exe is RUNNING")
        print("[DIAGNOSTIC] FIX 2: Enable web server in LHM (Options → Web server)")
        print("[DIAGNOSTIC] FIX 3: Check port is 8085 (default)")
    except requests.Timeout:
        print("[DIAGNOSTIC] ✗ REST API query timed out")
    except Exception as e:
        print(f"[DIAGNOSTIC] ✗ Error: {e}")
    return False


def get_lhm_data():
    try:
        response = requests.get(LHM_REST_API, timeout=5)
        if response.status_code == 200:
            return response.json()
    except (requests.ConnectionError, requests.Timeout, json.JSONDecodeError):
        pass
    return None

# ═════════════════════════════════════════════════════════════════════════════════════════════════════════════════════#
#  CPU SENSORS
# ═════════════════════════════════════════════════════════════════════════════════════════════════════════════════════#

def get_cpu_temp_from_lhm(data) -> int:
    if not data:
        return 0

    primary   = ("cpu package", "tdie")
    secondary = ("core average", "cpu core", "core max")
    best: dict[str, float | None] = {"primary": None, "secondary": None}

    try:
        for hw in _get_hardware_nodes(data):
            if not _is_cpu(hw.get("Text", "")):
                continue
            for cat in hw.get("Children", []):
                if "temperature" not in cat.get("Text", "").lower():
                    continue
                for sensor in cat.get("Children", []):
                    st = sensor.get("Text", "").lower()
                    if "distance" in st:
                        continue
                    try:
                        val = _numeric(sensor.get("Value", 0))
                    except ValueError:
                        continue
                    if any(p in st for p in primary):
                        best["primary"] = val
                    elif any(s in st for s in secondary):
                        if best["secondary"] is None:
                            best["secondary"] = val
    except (KeyError, TypeError, AttributeError):
        pass

    result = best["primary"] if best["primary"] is not None else best["secondary"]
    if result is None:
        return 0
    return int(result)


def get_cpu_power_from_lhm(data) -> int:
    if not data:
        return 0

    primary   = ("cpu package",)
    secondary = ("cpu cores", "cpu total", "package")
    best: dict[str, float | None] = {"primary": None, "secondary": None}

    try:
        for hw in _get_hardware_nodes(data):
            if not _is_cpu(hw.get("Text", "")):
                continue
            for cat in hw.get("Children", []):
                cat_text = cat.get("Text", "").lower()
                if "power" not in cat_text and "watt" not in cat_text:
                    continue
                for sensor in cat.get("Children", []):
                    st = sensor.get("Text", "").lower()
                    try:
                        val = _numeric(sensor.get("Value", 0))
                    except ValueError:
                        continue
                    if any(p in st for p in primary):
                        best["primary"] = val
                    elif any(s in st for s in secondary):
                        if best["secondary"] is None:
                            best["secondary"] = val
    except (KeyError, TypeError, AttributeError):
        pass

    result = best["primary"] if best["primary"] is not None else best["secondary"]
    if result is None:
        return 0
    return int(result)

def get_cpu_load_from_lhm(data) -> int:
    if not data:
        return 0

    try:
        for hw in _get_hardware_nodes(data):
            if not _is_cpu(hw.get("Text", "")):
                continue
            for cat in hw.get("Children", []):
                if "load" not in cat.get("Text", "").lower():
                    continue
                for sensor in cat.get("Children", []):
                    if "cpu total" in sensor.get("Text", "").lower():
                        try:
                            return int(_numeric(sensor.get("Value", 0)))
                        except ValueError:
                            pass
    except (KeyError, TypeError, AttributeError, ValueError):
        pass

    return 0


# ═════════════════════════════════════════════════════════════════════════════════════════════════════════════════════#
#  GPU SENSORS
# ═════════════════════════════════════════════════════════════════════════════════════════════════════════════════════#

def get_gpu_temp_from_lhm(data) -> int:
    if not data:
        return 0

    primary   = ("gpu core", "gpu temperature", "gpu temp")
    secondary = ("gpu hot spot", "gpu")
    best: dict[str, float | None] = {"primary": None, "secondary": None}

    try:
        for hw in _get_hardware_nodes(data):
            if not _is_gpu(hw.get("Text", "")):
                continue
            for cat in hw.get("Children", []):
                if "temperature" not in cat.get("Text", "").lower():
                    continue
                for sensor in cat.get("Children", []):
                    st = sensor.get("Text", "").lower()
                    if "distance" in st or "memory" in st:
                        continue
                    try:
                        val = _numeric(sensor.get("Value", 0))
                    except ValueError:
                        continue
                    if any(p in st for p in primary):
                        best["primary"] = val
                    elif any(s in st for s in secondary):
                        if best["secondary"] is None:
                            best["secondary"] = val
    except (KeyError, TypeError, AttributeError):
        pass

    result = best["primary"] if best["primary"] is not None else best["secondary"]
    if result is None:
        return 0
    return int(result)


def get_gpu_power_from_lhm(data) -> int:
    if not data:
        return 0

    labels = ("gpu package", "gpu total", "board power", "total board power", "gpu power", "gpu")

    try:
        for hw in _get_hardware_nodes(data):
            if not _is_gpu(hw.get("Text", "")):
                continue
            for cat in hw.get("Children", []):
                cat_text = cat.get("Text", "").lower()
                if "power" not in cat_text and "watt" not in cat_text:
                    continue
                for sensor in cat.get("Children", []):
                    st = sensor.get("Text", "").lower()
                    if any(lbl in st for lbl in labels):
                        try:
                            return int(_numeric(sensor.get("Value", 0)))
                        except ValueError:
                            pass
    except (KeyError, TypeError, AttributeError, ValueError):
        pass

    return 0


def get_gpu_load_from_lhm(data) -> int:
    if not data:
        return 0

    try:
        for hw in _get_hardware_nodes(data):
            if not _is_gpu(hw.get("Text", "")):
                continue
            for cat in hw.get("Children", []):
                if "load" not in cat.get("Text", "").lower():
                    continue
                for sensor in cat.get("Children", []):
                    if "gpu core" in sensor.get("Text", "").lower():
                        try:
                            return int(_numeric(sensor.get("Value", 0)))
                        except ValueError:
                            pass
    except (KeyError, TypeError, AttributeError, ValueError):
        pass

    return 0


# ═════════════════════════════════════════════════════════════════════════════════════════════════════════════════════#
#  MEMORY SENSORS
# ═════════════════════════════════════════════════════════════════════════════════════════════════════════════════════#

def _parse_gb(sensor_value) -> float:
    numeric_str = re.sub(r'[^\d.-]', '', str(sensor_value))
    return float(numeric_str)


def _fmt_gb(gb: float) -> str:
    rounded = round(gb)
    for nice in [2, 4, 6, 8, 10, 12, 16, 20, 24, 32, 48, 64, 96, 128]:
        if abs(rounded - nice) <= max(1, int(nice * 0.10)):
            return f"{nice}"
    return f"{rounded}"


def get_dram_total_from_lhm(data) -> str:
    try:
        total_bytes = psutil.virtual_memory().total
        total_gb = total_bytes / (1024 ** 3)
        return f"{_fmt_gb(total_gb)}"
    except (OSError, AttributeError):
        return "error"

def get_dram_used_from_lhm(data) -> float:
    if not data:
        return 0.0

    try:
        for hw in _get_hardware_nodes(data):
            if "total memory" not in hw.get("Text", "").lower():
                continue
            for cat in hw.get("Children", []):
                if "data" not in cat.get("Text", "").lower():
                    continue
                for sensor in cat.get("Children", []):
                    if "memory used" in sensor.get("Text", "").lower():
                        try:
                            return round(_numeric(sensor.get("Value", 0)), 1)
                        except ValueError:
                            pass
    except (KeyError, TypeError, AttributeError, ValueError):
        pass

    return 0.0


def get_vram_used_from_lhm(data) -> float:
    if not data:
        return 0.0

    try:
        for hw in _get_hardware_nodes(data):
            if not _is_gpu(hw.get("Text", "")):
                continue
            for cat in hw.get("Children", []):
                for sensor in cat.get("Children", []):
                    st = sensor.get("Text", "").lower()
                    if "gpu memory used" in st:
                        try:
                            raw = _numeric(sensor.get("Value", 0))
                            gb = raw / 1024 if raw > 500 else raw
                            return round(gb, 1)
                        except ValueError:
                            pass
    except (KeyError, TypeError, AttributeError, ValueError):
        pass

    return 0.0


def get_vram_total_from_lhm(data) -> str:
    if not data:
        return "error"

    try:
        for hw in _get_hardware_nodes(data):
            if not _is_gpu(hw.get("Text", "")):
                continue

            total_mb = used_mb = free_mb = None

            for cat in hw.get("Children", []):
                for sensor in cat.get("Children", []):
                    st = sensor.get("Text", "").lower()
                    try:
                        val = _numeric(sensor.get("Value", 0))
                    except ValueError:
                        continue
                    if "gpu memory total" in st:
                        total_mb = val
                    elif "gpu memory used" in st and total_mb is None:
                        used_mb = val
                    elif "gpu memory free" in st and total_mb is None:
                        free_mb = val

            if total_mb and total_mb > 0:
                gb = total_mb / 1024 if total_mb > 500 else total_mb
                return _fmt_gb(gb)
            if used_mb is not None and free_mb is not None:
                total = used_mb + free_mb
                gb = total / 1024 if total > 500 else total
                return _fmt_gb(gb)

    except (KeyError, TypeError, AttributeError, ValueError):
        pass

    return "error"

def parse_lhm_data(data):
    return (
        get_cpu_temp_from_lhm(data),
        get_cpu_power_from_lhm(data),
        get_gpu_temp_from_lhm(data),
        get_gpu_power_from_lhm(data),
    )

def detect_dram_type() -> str:
    try:
        result = subprocess.check_output(
            ["powershell", "-Command",
             "(Get-CimInstance Win32_PhysicalMemory | Select-Object -First 1).SMBIOSMemoryType"],
            encoding="utf-8", stderr=subprocess.DEVNULL, timeout=5
        ).strip()
        type_map = {
            "17": "SDRAM", "18": "SDRAM", "19": "SDRAM",
            "20": "DDR",   "21": "DDR2",  "22": "DDR2",
            "23": "DDR2",  "24": "DDR3",  "26": "DDR4",
            "34": "DDR5",  "35": "DDR5"
        }
        return type_map.get(result, "DDR")
    except (subprocess.CalledProcessError, UnicodeDecodeError, subprocess.TimeoutExpired):
        return "DDR"


def detect_vram_type(gpu_name: str) -> str:
    name = gpu_name.lower()

    if any(x in name for x in ["5090", "5080", "5070 ti", "5070ti", "5070", "5060 ti", "5060ti", "5060"]):
        return "GDDR7"
    if any(x in name for x in ["4090", "4080", "4070 ti", "4070ti", "4070 super", "4070super"]):
        return "GDDR6X"
    if any(x in name for x in ["4070", "4060 ti", "4060ti", "4060"]):
        return "GDDR6"
    if any(x in name for x in ["3090", "3080"]):
        return "GDDR6X"
    if any(x in name for x in ["3070", "3060", "rtx 30", "rtx30"]):
        return "GDDR6"
    if any(x in name for x in ["rtx 20", "rtx20", "2080", "2070", "2060"]):
        return "GDDR6"
    if any(x in name for x in ["1080 ti", "1080ti", "1080"]):
        return "GDDR5X"
    if any(x in name for x in ["gtx 10", "gtx10", "1070", "1060", "1050"]):
        return "GDDR5"
    if any(x in name for x in ["980 ti", "980ti", "980", "970"]):
        return "GDDR5"
    if any(x in name for x in ["rx 9", "rx9"]):
        return "GDDR6"
    if any(x in name for x in ["rx 7", "rx7"]):
        return "GDDR6"
    if any(x in name for x in ["rx 6", "rx6"]):
        return "GDDR6"
    if any(x in name for x in ["rx 5", "rx5"]):
        return "GDDR6"
    if any(x in name for x in ["rx 5", "rx 4", "rx5", "rx4", "rx 580", "rx 570", "rx 480", "rx 470"]):
        return "GDDR5"
    if any(x in name for x in ["radeon", "nvidia", "geforce"]):
        return "GDDR6"

    return "GDDR"

# ═════════════════════════════════════════════════════════════════════════════════════════════════════════════════════#
# NETWORK MONITORING
# ═════════════════════════════════════════════════════════════════════════════════════════════════════════════════════#

def fmt(bps):
    if bps >= 1_000_000:
        return f"{bps / 1_000_000:.2f} Mb/s"
    return f"{bps / 1_000:.1f} Kb/s"


def get_network_usage(prev, prev_time):
    now = time.time()
    try:
        cur = psutil.net_io_counters(pernic=True)[INTERFACE]
        elapsed = now - prev_time
        if elapsed > 0:
            up = (cur.bytes_sent - prev.bytes_sent) / elapsed
            down = (cur.bytes_recv - prev.bytes_recv) / elapsed
        else:
            up, down = 0, 0
        return cur, up, down, now
    except (KeyError, ZeroDivisionError):
        return prev, 0, 0, now


# ═════════════════════════════════════════════════════════════════════════════════════════════════════════════════════#
# WEATHER MONITORING  (Open-Meteo — no API key required)
# ═════════════════════════════════════════════════════════════════════════════════════════════════════════════════════#

def fetch_weather(lat_lon_str: str):
    global weather_temp, weather_humidity, weather_desc

    try:
        parts = [p.strip() for p in lat_lon_str.split(",")]
        if len(parts) != 2:
            raise ValueError("Expected format: latitude,longitude  (e.g. 51.5,-0.1)")
        lat, lon = float(parts[0]), float(parts[1])
    except (ValueError, AttributeError) as e:
        print(f"[Weather] Bad location format: {e}")
        weather_temp     = "?"
        weather_humidity = "?"
        weather_desc     = "Bad location"
        return False

    url = (
        "https://api.open-meteo.com/v1/forecast"
        f"?latitude={lat}&longitude={lon}"
        "&current=temperature_2m,relative_humidity_2m,weather_code"
        "&temperature_unit=celsius"
        "&wind_speed_unit=mph"
        "&timezone=auto"
    )

    try:
        resp = requests.get(url, timeout=10)
        resp.raise_for_status()
        data = resp.json()

        current          = data["current"]
        weather_temp     = round(current["temperature_2m"])
        weather_humidity = current["relative_humidity_2m"]
        weather_desc     = WMO_CODES.get(current["weather_code"], "Unknown")

        print(f"[Weather] {weather_temp}°C  {weather_humidity}%  {weather_desc}")
        return True

    except requests.ConnectionError:
        print("[Weather] ✗ No internet connection")
    except requests.Timeout:
        print("[Weather] ✗ Request timed out")
    except (KeyError, ValueError, requests.HTTPError) as e:
        print(f"[Weather] ✗ Parse error: {e}")
    except Exception as e:
        print(f"[Weather] ✗ Unexpected error: {e}")

    weather_temp     = "?"
    weather_humidity = "?"
    weather_desc     = "Unavailable"
    return False


# ═════════════════════════════════════════════════════════════════════════════════════════════════════════════════════#
# MEDIA MONITORING
# ═════════════════════════════════════════════════════════════════════════════════════════════════════════════════════#

async def get_media_info():
    try:
        manager = await wmc.GlobalSystemMediaTransportControlsSessionManager.request_async()
        session = manager.get_current_session()
        if session:
            props = await session.try_get_media_properties_async()
            timeline = session.get_timeline_properties()
            playback = session.get_playback_info()
            pos = timeline.position.total_seconds() * 1000
            dur = timeline.end_time.total_seconds() * 1000
            is_paused = (
                playback.playback_status ==
                wmc.GlobalSystemMediaTransportControlsSessionPlaybackStatus.PAUSED
            )
            return props.title, props.artist, pos, dur, is_paused
    except (OSError, AttributeError, RuntimeError):
        pass
    return None, None, 0, 0, False


def clean_title(raw_title):
    if not raw_title:
        return ""

    title = re.sub(r"\(.*?\)|\[.*?]|\{.*?}", "", raw_title)
    junk_words = [
        "official", "video", "lyrics", "audio", "hd", "4k", "remastered",
        "live", "visualizer", "explicit", "clean", "version", "mix"
    ]
    pattern = r"\b(" + "|".join(junk_words) + r")\b"
    title = re.sub(pattern, "", title, flags=re.IGNORECASE)
    title = re.sub(r"\b(ft\.|feat\.|featuring).*", "", title, flags=re.IGNORECASE)
    parts = [p.strip() for p in re.split(r"[-–|•]", title) if len(p.strip()) > 2]
    title = parts[0] if parts else title

    return re.sub(r"\s+", " ", title).strip()


def create_progress_bar(position_ms, duration_ms, length=13):
    if duration_ms <= 0:
        return "No music playing︎"
    percent = min(max(position_ms / duration_ms, 0), 1)
    filled_len = int(length * percent)
    return "■" * filled_len + "□" * (length - filled_len)


# ═════════════════════════════════════════════════════════════════════════════════════════════════════════════════════#
# MAIN OSC LOOP
# ═════════════════════════════════════════════════════════════════════════════════════════════════════════════════════#

def run_osc_loop():
    global running, cpu_wattage, cpu_temp, gpu_wattage, gpu_temp, dram_load, vram_load, client

    all_stats = psutil.net_io_counters(pernic=True)
    if INTERFACE not in all_stats:
        print(f"Error: {INTERFACE} not found. Available: {list(all_stats.keys())}")
        running = False
        return

    prev = all_stats[INTERFACE]
    prev_time = time.time()

    cpu_detect = detect_cpu()
    gpu_detect = detect_gpu()

    lhm_data = get_lhm_data()
    startup_lhm = get_lhm_data()
    dram_detect = get_dram_total_from_lhm(startup_lhm)
    vram_detect = get_vram_total_from_lhm(startup_lhm)
    dram = detect_dram_type()
    vram = detect_vram_type(gpu_detect)

    fetch_weather(location_entry.get().strip())

    print(f"\n{'=' * 60}")
    print(f"CPU:  {cpu_detect} ({cpu_manufacturer.value})")
    print(f"GPU:  {gpu_detect}")
    print(f"Dram Total: {dram_detect}")
    print(f"VRAM Total: {vram_detect}")
    print(f"{'=' * 60}")

    query_cooldown   = 0
    weather_cooldown = 0
    WEATHER_INTERVAL = 60

    cpu_load = 0
    gpu_load = 0
    dram_load = 0
    vram_load = 0

    while running:
        try:
            song, artist, pos, dur, is_paused = asyncio.run(get_media_info())
            clean_song = clean_title(song)

            query_cooldown += 1
            if query_cooldown >= 3:
                lhm_data = get_lhm_data()
                if lhm_data:
                    cpu_temp, cpu_wattage, gpu_temp, gpu_wattage = parse_lhm_data(lhm_data)
                    cpu_load  = get_cpu_load_from_lhm(lhm_data)
                    gpu_load  = get_gpu_load_from_lhm(lhm_data)
                    dram_load = get_dram_used_from_lhm(lhm_data)
                    vram_load = get_vram_used_from_lhm(lhm_data)
                else:
                    cpu_load = gpu_load = 0
                    dram_load = vram_load = 0.0
                query_cooldown = 0

            weather_cooldown += 1
            if weather_cooldown >= WEATHER_INTERVAL:
                fetch_weather(location_entry.get().strip())
                weather_cooldown = 0

            if vram_detect == "error":
                vram_detect = get_vram_total_from_lhm(lhm_data)
            if dram_detect == "error":
                dram_detect = get_dram_total_from_lhm(lhm_data)

            prev, up_raw, down_raw, prev_time = get_network_usage(prev, prev_time)

            cur_time_str = time.strftime("%I:%M %p")
            progress_bar = create_progress_bar(pos, dur)

            if clean_song:
                if is_paused:
                   display_artist = f"- {artist}" if artist else ""
                   display_song = f"⏸ {clean_song}" if clean_song else ""
                else:
                    display_artist = f"- {artist}" if artist else ""
                    display_song = f"🎵 {clean_song}" if clean_song else ""
            else:
                display_artist = "⏸"
                display_song = ""

            page_index = int((time.time() // SWITCH_INTERVAL) % 4)

            if forced_text.get().strip() == "":

                if page_index == 0:
                    text = (
                        f"{page1_line1_text}\n"
                        f"{cur_time_str}\n"
                        f"Download {fmt(down_raw)}\n"
                        f"Upload {fmt(up_raw)}\n"
                        f"{progress_bar}\n"
                        f"{display_song} {display_artist}"
                    )
                elif page_index == 1:
                    text = (
                        f"{page2_line1_text}\n"
                        f"{cur_time_str}\n"
                        f"{cpu_detect} {cpu_load}%\n"
                        f"{cpu_wattage}w {cpu_temp}℃\n"
                        f"{gpu_detect} {gpu_load}%\n"
                        f"{gpu_wattage}w {gpu_temp}℃\n"
                    )
                elif page_index == 2:
                    text = (
                        f"{page3_line1_text}\n"
                        f"{cur_time_str}\n"
                        f"{dram} {dram_load}GB/{dram_detect}GB\n"
                        f"{vram} {vram_load}GB/{vram_detect}GB\n"
                        f"{progress_bar}\n"
                        f"{display_song} {display_artist}"
                    )
                elif page_index == 3:
                    text = (
                        f"{page4_line1_text}\n"
                        f"{cur_time_str}\n"
                        f"{weather_temp}℃  {weather_humidity}% humidity\n"
                        f"{weather_desc}\n"
                        f"{progress_bar}\n"
                        f"{display_song} {display_artist}"
                    )
                else:
                    text = f"{error_text}"
            else:
                text = forced_text.get()

            print(text)

            if client is not None:
                client.send_message("/chatbox/input", [text, True])  # type: ignore
            else:
                print("Warning: OSC client not initialized")

            time.sleep(5.0)

        except Exception as e:
            print(f"Error: OSC loop Error {e}")
            time.sleep(1)

# ═════════════════════════════════════════════════════════════════════════════════════════════════════════════════════#
# SCRIPT CONTROL
# ═════════════════════════════════════════════════════════════════════════════════════════════════════════════════════#

def start_script():
    global running, client, OSC_IP, OSC_PORT, INTERFACE, SWITCH_INTERVAL, LHM_REST_API
    global page1_line1_text, page2_line1_text, page3_line1_text, page4_line1_text, error_text

    if running:
        return

    try:
        OSC_IP          = ip_entry.get()
        OSC_PORT        = int(port_entry.get())
        INTERFACE       = iface_entry.get()
        SWITCH_INTERVAL = int(interval_entry.get())
        LHM_REST_API    = lhm_entry.get()

        page1_line1_text = page1_entry.get()
        page2_line1_text = page2_entry.get()
        page3_line1_text = page3_entry.get()
        page4_line1_text = page4_entry.get()

        save_config()
        error_text = "Error: Page error value exceeding limit"
        client = SimpleUDPClient(OSC_IP, OSC_PORT)

        running = True
        status_label.config(text="Status: Running", fg="#4CFF4C")

        diagnose_lhm()

        thread = threading.Thread(target=run_osc_loop, daemon=True)
        thread.start()

    except ValueError as e:
        messagebox.showerror("Error", f"Invalid input: {e}")
        status_label.config(text="Status: Error", fg="#FF4C4C")
    except Exception as e:
        messagebox.showerror("Error", f"Failed to start: {e}")
        status_label.config(text="Status: Error", fg="#FF4C4C")


def stop_script():
    global running
    running = False
    status_label.config(text="Status: Stopped", fg="#FF4C4C")


def restart_script():
    stop_script()
    time.sleep(1)
    start_script()


# ═════════════════════════════════════════════════════════════════════════════════════════════════════════════════════#
# GUI
# ═════════════════════════════════════════════════════════════════════════════════════════════════════════════════════#

cfg = load_config()

BG       = "#121212"
FG       = "#E0E0E0"
ENTRY_BG = "#1E1E1E"
BTN_BG   = "#2A2A2A"
BTN_FG   = "#FFFFFF"

root = tk.Tk()
root.title("OSC Chatbox")
root.geometry("450x480")
root.configure(bg=BG)
root.resizable(True, True)

frame = tk.Frame(root, bg=BG)
frame.pack(fill="both", expand=True, padx=10, pady=10)


def dark_label(text, r):
    lbl = tk.Label(frame, text=text, bg=BG, fg=FG, anchor="w")
    lbl.grid(row=r, column=0, sticky="w", pady=4)
    return lbl


def dark_entry(r, default=""):
    e = tk.Entry(frame, bg=ENTRY_BG, fg=FG, insertbackground=FG, relief="flat")
    e.insert(0, default)
    e.grid(row=r, column=1, pady=4, sticky="ew")
    return e

def open_help():
    help_win = tk.Toplevel(root)
    help_win.title("OSC Chatbox Tutorial")
    help_win.configure(bg=BG)
    help_win.resizable(True, True)

    help_w, help_h = 460, 480
    root_x = root.winfo_x()
    root_y = root.winfo_y()
    root_w = root.winfo_width()
    root_h = root.winfo_height()
    x = root_x + (root_w // 2) - (help_w // 2)
    y = root_y + (root_h // 2) - (help_h // 2)
    help_win.geometry(f"{help_w}x{help_h}+{x}+{y}")

    pages = [
        {
            "title": "Data Config",
            "content": (
                "OSC IP — The IP address to send OSC messages to.\n"
                "Usually 127.0.0.1 (your own PC).\n\n"
                "OSC Port — The port VRChat listens on.\n"
                "Default is 9000. Don't change unless needed.\n\n"
                "Network Interface — The name of your network\n"
                "adapter to monitor (e.g. Ethernet, Wi-Fi).\n"
                "Open Task Manager → Performance to find yours.\n\n"
                "Switch Interval — How many seconds before the\n"
                "chatbox rotates to the next page (e.g. 20)."
            )
        },
        {
            "title": "LHM Interface",
            "content": (
                "LHM Interface — The URL for the LibreHardwareMonitor\n"
                "REST API. Default: http://localhost:8085/data.json\n\n"
                "This is used to read CPU/GPU temperatures,\n"
                "wattage, and load percentages.\n\n"
                "To enable it:\n"
                "1. Open LibreHardwareMonitor.exe\n"
                "2. Go to Options → Web Server\n"
                "3. Click 'Run' and make sure port is 8085\n"
            )
        },
        {
            "title": "Page Text",
            "content": (
                "Page 1 / 2 / 3 / 4 Text — The first line shown\n"
                "on each rotating chatbox page.\n\n"
                "Page 1 also shows: time, network speed, and\n"
                "currently playing song.\n\n"
                "Page 2 also shows: CPU & GPU usage, temps,\n"
                "and wattage.\n\n"
                "Page 3 also shows: RAM & VRAM usage and\n"
                "currently playing song.\n\n"
                "Page 4 shows your local weather — see the\n"
                "next page for setup details."
            )
        },
        {
            "title": "Weather",
            "content": (
                "Page 4 shows live weather from Open-Meteo\n\n"
                "Location — Enter your coordinates as:\n"
                "   latitude,longitude\n"
                "Example: 51.5,-0.1  (London)\n"
                "         40.7,-74.0  (New York)\n\n"
                "To find your coordinates:\n"
                "  Google Maps → right-click your location\n"
                "  → the first line shown is lat,lon.\n\n"
                "Weather refreshes every 5 minutes.\n"
                "Page 4 displays: temperature (°C),\n"
                "humidity (%), and a short condition\n"
                "description (e.g. 'Partly cloudy')."
            )
        },
        {
            "title": "Text Message",
            "content": (
                "Text Message — If you type anything here,\n"
                "it overrides ALL pages and sends only this\n"
                "text to the chatbox.\n\n"
                "Leave it blank (or just spaces) to go back\n"
                "to the normal rotating pages.\n\n"
                "Useful for quickly sending a custom message\n"
                "without stopping the script."
            )
        },
    ]

    current_page = [0]

    title_label = tk.Label(help_win, text="", bg=BG, fg="#FFFFFF",
                           font=("Segoe UI", 16, "bold"))
    title_label.pack(pady=(14, 4))

    content_label = tk.Label(help_win, text="", bg=BG, fg=FG,
                             justify="left", wraplength=400,
                             font=("Segoe UI", 10))
    content_label.pack(padx=20, fill="both", expand=True)

    page_indicator = tk.Label(help_win, text="", bg=BG, fg="#888888",
                              font=("Segoe UI", 8))
    page_indicator.pack(pady=(0, 4))

    def show_page(idx):
        p = pages[idx]
        title_label.config(text=p["title"])
        content_label.config(text=p["content"])
        page_indicator.config(text=f"Page {idx + 1} of {len(pages)}")
        prev_btn.config(state="normal" if idx > 0 else "disabled")
        is_last = idx == len(pages) - 1
        next_btn.config(text="Finish" if is_last else "Next →")

    nav_frame = tk.Frame(help_win, bg=BG)
    nav_frame.pack(fill="x", padx=20, pady=(0, 14))
    nav_frame.columnconfigure(1, weight=1)

    prev_btn = tk.Button(nav_frame, text="← Back", bg=BTN_BG, fg=BTN_FG,
                         relief="flat", width=10,
                         command=lambda: (current_page.__setitem__(0, current_page[0] - 1),
                                          show_page(current_page[0])))
    prev_btn.grid(row=0, column=0, sticky="w")

    def next_or_finish():
        if current_page[0] < len(pages) - 1:
            current_page[0] += 1
            show_page(current_page[0])
        else:
            help_win.destroy()

    next_btn = tk.Button(nav_frame, text="Next →", bg=BTN_BG, fg=BTN_FG,
                         relief="flat", width=10, command=next_or_finish)
    next_btn.grid(row=0, column=2, sticky="e")

    show_page(0)

frame.columnconfigure(1, weight=1)

# ── Data Config section ────────────────────────────────────────────────────────
data_config_label = tk.Label(frame, text="Data config", bg=BG, fg="#FFFFFF")
data_config_label.grid(row=0, column=0, columnspan=2)

dark_label("OSC IP", 1)
ip_entry = dark_entry(1, cfg["osc_ip"])

dark_label("OSC Port", 2)
port_entry = dark_entry(2, cfg["osc_port"])

dark_label("Network Interface", 3)
iface_entry = dark_entry(3, cfg["interface"])

dark_label("Switch Interval", 4)
interval_entry = dark_entry(4, cfg["switch_interval"])

dark_label("LHM Interface", 5)
lhm_entry = dark_entry(5, cfg["lhm_api"])

dark_label("Location (lat,lon)", 6)
location_entry = dark_entry(6, cfg["location"])

# ── Page Text section ──────────────────────────────────────────────────────────
page_text_label = tk.Label(frame, text="Page Text", bg=BG, fg="#FFFFFF")
page_text_label.grid(row=7, column=0, columnspan=2)

dark_label("Page 1 Text", 8)
page1_entry = dark_entry(8, cfg["page1_text"])

dark_label("Page 2 Text", 9)
page2_entry = dark_entry(9, cfg["page2_text"])

dark_label("Page 3 Text", 10)
page3_entry = dark_entry(10, cfg["page3_text"])

dark_label("Page 4 Text", 11)
page4_entry = dark_entry(11, cfg["page4_text"])

dark_label("Text Message", 12)
forced_text = dark_entry(12, " ")

# ── Buttons ────────────────────────────────────────────────────────────────────
button_frame = tk.Frame(frame, bg=BG)
button_frame.grid(row=13, column=0, columnspan=2, pady=15, sticky="ew")
button_frame.columnconfigure(0, weight=1)
button_frame.columnconfigure(1, weight=1)
button_frame.columnconfigure(2, weight=1)

start_btn = tk.Button(button_frame, text="Start", command=start_script,
                      bg=BTN_BG, fg=BTN_FG, relief="flat")
start_btn.grid(row=0, column=0, sticky="ew", padx=2)

stop_btn = tk.Button(button_frame, text="Stop", command=stop_script,
                     bg=BTN_BG, fg=BTN_FG, relief="flat")
stop_btn.grid(row=0, column=1, sticky="ew", padx=2)

restart_btn = tk.Button(button_frame, text="Restart", command=restart_script,
                        bg=BTN_BG, fg=BTN_FG, relief="flat")
restart_btn.grid(row=0, column=2, sticky="ew", padx=2)

status_label = tk.Label(frame, text="Status: Stopped", bg=BG, fg="#FF4C4C")
status_label.grid(row=14, column=0, columnspan=2)

help_btn = tk.Button(frame, text=" ？ ", command=open_help,
                     bg=BTN_BG, fg="#FFFFFF", relief="flat",
                     font=("Segoe UI", 9), cursor="question_arrow")
help_btn.grid(row=14, column=0, sticky="w", padx=2)

root.mainloop()