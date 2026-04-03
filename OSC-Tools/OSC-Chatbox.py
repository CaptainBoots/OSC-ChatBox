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
import json
import asyncio
import re
import threading
import time
import tkinter as tk
from tkinter import messagebox
from enum import Enum
from typing import Optional
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


install_if_missing("python-osc==1.9.3", "pythonosc")
install_if_missing("psutil==7.2.2", "psutil")
install_if_missing("requests==2.32.5", "requests")

if sys.platform == "win32":
    install_if_missing("winrt-Windows.Media.Control==3.2.1", "winrt")
    install_if_missing("winrt-windows.foundation==3.2.1", "winrt.windows.foundation")

import psutil
from pythonosc.udp_client import SimpleUDPClient
import requests

if sys.platform == "win32":
    import winrt.windows.media.control as wmc
else:
    wmc = None


# ═════════════════════════════════════════════════════════════════════════════════════════════════════════════════════#
# CONFIGURATION & GLOBAL VARIABLES
# ═════════════════════════════════════════════════════════════════════════════════════════════════════════════════════#

VERSION = "7.6.2"


class CPUManufacturer(Enum):
    INTEL = "Intel"
    AMD = "AMD"
    UNKNOWN = "Unknown"


cpu_manufacturer = CPUManufacturer.UNKNOWN

print("OSC Chatbox")
print("Made By Boots")

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
CONFIG_DIR = os.path.join(SCRIPT_DIR, "OSC-Chatbox")
CONFIG_FILE = os.path.join(CONFIG_DIR, "osc_config.json")
OSC_IP = "error"
OSC_PORT = "error"
INTERFACE = "error"
SWITCH_INTERVAL = "error"
LHM_REST_API = "error"

client: Optional[SimpleUDPClient] = None
running = False

page1_line1_text = "error"
page2_line1_text = "error"
page3_line1_text = "error"
page4_line1_text = "error"
error_text = "Error: Page error value exceeding limit"

cpu_wattage = "error"
cpu_temp = "error"
gpu_wattage = "error"
gpu_temp = "error"
dram_load = "error"
vram_load = "error"

weather_temp = "error"
weather_humidity = "error"
weather_desc = "error"

page_toggles = []
DEFAULT_PROGRESS_FILLED_CHAR = "\u2592"
DEFAULT_PROGRESS_BORDER_CHAR = "\u2593"
DEFAULT_PROGRESS_EMPTY_CHAR = "\u2591"
progress_filled_char: str
progress_border_char: str
progress_empty_char: str


def normalize_progress_char(value, fallback):
    text = "" if value is None else str(value).strip()
    return text[0] if text else fallback


def detect_default_interface():
    fallback = "Ethernet" if sys.platform == "win32" else "eth0"
    try:
        interfaces = list(psutil.net_io_counters(pernic=True).keys())
        if not interfaces:
            return fallback

        if sys.platform == "win32":
            for preferred in ("Ethernet", "Wi-Fi", "WiFi"):
                if preferred in interfaces:
                    return preferred

        for iface in interfaces:
            lower = iface.lower()
            if lower.startswith("lo") or "loopback" in lower:
                continue
            return iface

        return interfaces[0]
    except (OSError, AttributeError):
        return fallback


# ═════════════════════════════════════════════════════════════════════════════════════════════════════════════════════#
# CONFIG FILE
# ═════════════════════════════════════════════════════════════════════════════════════════════════════════════════════#

def get_default_config():
    return {
        "osc_ip": "127.0.0.1",
        "osc_port": 9000,
        "interface": detect_default_interface(),
        "switch_interval": 20,
        "lhm_api": "http://localhost:8085/data.json",
        "location": "0,0",
        "page1_text": "Thx for using Boots's osc code",
        "page2_text": "Join the discord server at https://discord.gg/XdfKAWu6Ph",
        "page3_text": "hi put your text here :3",
        "page4_text": "Local Weather",
        "page1_enabled": True,
        "page2_enabled": True,
        "page3_enabled": True,
        "page4_enabled": True,
        "progress_filled_char": DEFAULT_PROGRESS_FILLED_CHAR,
        "progress_border_char": DEFAULT_PROGRESS_BORDER_CHAR,
        "progress_empty_char": DEFAULT_PROGRESS_EMPTY_CHAR,
    }


def _legacy_config_dirs():
    legacy_dirs = []
    current_config_dir_name = os.path.basename(CONFIG_DIR)
    try:
        for entry in os.scandir(SCRIPT_DIR):
            if not entry.is_dir():
                continue
            if entry.name == current_config_dir_name or not entry.name.startswith("OSC-"):
                continue
            if os.path.isfile(os.path.join(entry.path, "osc_config.json")):
                legacy_dirs.append(entry.path)
    except OSError:
        return []
    legacy_dirs.sort()
    return legacy_dirs


def migrate_legacy_config_directory():
    if os.path.isdir(CONFIG_DIR):
        return

    legacy_dirs = _legacy_config_dirs()
    if not legacy_dirs:
        return

    legacy_dir = legacy_dirs[0]
    try:
        os.rename(legacy_dir, CONFIG_DIR)
        print(
            f"[Config] Migrated config directory: "
            f"{os.path.basename(legacy_dir)} -> {os.path.basename(CONFIG_DIR)}"
        )
        return
    except OSError:
        pass

    legacy_config_file = os.path.join(legacy_dir, "osc_config.json")
    try:
        os.makedirs(CONFIG_DIR, exist_ok=True)
        with open(legacy_config_file, "r", encoding="utf-8") as src:
            with open(CONFIG_FILE, "w", encoding="utf-8") as dst:
                dst.write(src.read())
        print(
            f"[Config] Copied legacy config into {os.path.basename(CONFIG_DIR)} "
            f"from {os.path.basename(legacy_dir)}"
        )
    except OSError:
        pass


def load_config():
    defaults = get_default_config()
    migrate_legacy_config_directory()

    candidate_paths = [CONFIG_FILE]
    for legacy_dir in _legacy_config_dirs():
        candidate = os.path.join(legacy_dir, "osc_config.json")
        if candidate not in candidate_paths:
            candidate_paths.append(candidate)

    for path in candidate_paths:
        try:
            with open(path, "r", encoding="utf-8") as f:
                loaded = json.load(f)

            if path != CONFIG_FILE and not os.path.exists(CONFIG_FILE):
                os.makedirs(os.path.dirname(CONFIG_FILE), exist_ok=True)
                with open(CONFIG_FILE, "w", encoding="utf-8") as migrated:
                    json.dump(loaded, migrated, indent=2)

            return {**defaults, **loaded}
        except (FileNotFoundError, json.JSONDecodeError, OSError):
            continue
    return defaults


def save_config():
    os.makedirs(os.path.dirname(CONFIG_FILE), exist_ok=True)
    config = {
        "osc_ip": ip_entry.get(),
        "osc_port": port_entry.get(),
        "interface": iface_entry.get(),
        "switch_interval": interval_entry.get(),
        "lhm_api": lhm_entry.get(),
        "location": location_entry.get(),
        "page1_text": page1_entry.get(),
        "page2_text": page2_entry.get(),
        "page3_text": page3_entry.get(),
        "page4_text": page4_entry.get(),
        "page1_enabled": page_toggles[0].get() if page_toggles else True,
        "page2_enabled": page_toggles[1].get() if page_toggles else True,
        "page3_enabled": page_toggles[2].get() if page_toggles else True,
        "page4_enabled": page_toggles[3].get() if page_toggles else True,
        "progress_filled_char": progress_filled_char,
        "progress_border_char": progress_border_char,
        "progress_empty_char": progress_empty_char,
    }
    with open(CONFIG_FILE, "w", encoding="utf-8") as f:
        json.dump(config, f, indent=2)


def reset_to_defaults():
    global progress_filled_char, progress_border_char, progress_empty_char
    defaults = get_default_config()

    ip_entry.delete(0, tk.END)
    ip_entry.insert(0, defaults["osc_ip"])

    port_entry.delete(0, tk.END)
    port_entry.insert(0, str(defaults["osc_port"]))

    iface_entry.delete(0, tk.END)
    iface_entry.insert(0, defaults["interface"])

    interval_entry.delete(0, tk.END)
    interval_entry.insert(0, str(defaults["switch_interval"]))

    lhm_entry.delete(0, tk.END)
    lhm_entry.insert(0, defaults["lhm_api"])

    location_entry.delete(0, tk.END)
    location_entry.insert(0, defaults["location"])

    page1_entry.delete(0, tk.END)
    page1_entry.insert(0, defaults["page1_text"])

    page2_entry.delete(0, tk.END)
    page2_entry.insert(0, defaults["page2_text"])

    page3_entry.delete(0, tk.END)
    page3_entry.insert(0, defaults["page3_text"])

    page4_entry.delete(0, tk.END)
    page4_entry.insert(0, defaults["page4_text"])

    keys = ["page1_enabled", "page2_enabled", "page3_enabled", "page4_enabled"]
    for i, cfg_key in enumerate(keys):
        page_toggles[i].set(defaults[cfg_key])

    progress_filled_char = defaults["progress_filled_char"]
    progress_border_char = defaults["progress_border_char"]
    progress_empty_char = defaults["progress_empty_char"]

    forced_text.delete(0, tk.END)

    save_config()


# ═════════════════════════════════════════════════════════════════════════════════════════════════════════════════════#
#  HARDWARE MONITORING HELPERS
# ═════════════════════════════════════════════════════════════════════════════════════════════════════════════════════#

def _clean_name(text: str):
    text = re.sub(r"\(.*?\)|\[.*?]|\{.*?}", "", text)
    text = text.split("@")[0]
    text = re.sub(r"\s+", " ", text).strip()
    return text


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
    for top in data.get("Children", []):
        for hw in top.get("Children", []):
            yield hw


linux_gpu_id_map = {

    # ══════════════════════════════════════════════════════════════
    # NVIDIA — RTX 50 series (Blackwell)
    # ══════════════════════════════════════════════════════════════
    "10de:2c04": "GeForce RTX 5090",
    "10de:2c03": "GeForce RTX 5080",
    "10de:2c02": "GeForce RTX 5070 Ti",
    "10de:2c01": "GeForce RTX 5070",
    "10de:2c00": "GeForce RTX 5060 Ti",

    # ══════════════════════════════════════════════════════════════
    # NVIDIA — RTX 40 series (Ada Lovelace)
    # ══════════════════════════════════════════════════════════════
    "10de:2684": "GeForce RTX 4090",
    "10de:2704": "GeForce RTX 4080",
    "10de:2702": "GeForce RTX 4080 Super",
    "10de:2782": "GeForce RTX 4070 Ti",
    "10de:2705": "GeForce RTX 4070 Ti Super",
    "10de:2786": "GeForce RTX 4070",
    "10de:2860": "GeForce RTX 4070 Super",
    "10de:2803": "GeForce RTX 4060 Ti",
    "10de:2882": "GeForce RTX 4060",
    "10de:27ba": "GeForce RTX 4090 Laptop",
    "10de:27b8": "GeForce RTX 4080 Laptop",
    "10de:27b0": "GeForce RTX 4070 Ti Laptop",
    "10de:27b9": "GeForce RTX 4070 Laptop",
    "10de:28e0": "GeForce RTX 4060 Laptop",
    "10de:28a1": "GeForce RTX 4050 Laptop",

    # ══════════════════════════════════════════════════════════════
    # NVIDIA — RTX 30 series (Ampere)
    # ══════════════════════════════════════════════════════════════
    "10de:2203": "GeForce RTX 3090 Ti",
    "10de:2204": "GeForce RTX 3090",
    "10de:2208": "GeForce RTX 3080 Ti",
    "10de:2206": "GeForce RTX 3080",
    "10de:2216": "GeForce RTX 3080 12GB",
    "10de:2482": "GeForce RTX 3070 Ti",
    "10de:2484": "GeForce RTX 3070",
    "10de:2489": "GeForce RTX 3060 Ti",
    "10de:2503": "GeForce RTX 3060",
    "10de:2507": "GeForce RTX 3050",
    "10de:2420": "GeForce RTX 3080 Ti Laptop",
    "10de:2460": "GeForce RTX 3080 Laptop",
    "10de:24b9": "GeForce RTX 3070 Ti Laptop",
    "10de:24dd": "GeForce RTX 3070 Laptop",
    "10de:249d": "GeForce RTX 3070 Laptop GPU",
    "10de:2520": "GeForce RTX 3060 Laptop",
    "10de:25a0": "GeForce RTX 3050 Ti Laptop",
    "10de:25a2": "GeForce RTX 3050 Laptop",

    # ══════════════════════════════════════════════════════════════
    # NVIDIA — RTX 20 series (Turing)
    # ══════════════════════════════════════════════════════════════
    "10de:1e04": "GeForce RTX 2080 Ti",
    "10de:1e87": "GeForce RTX 2080",
    "10de:1e84": "GeForce RTX 2080 Super",
    "10de:1f02": "GeForce RTX 2070",
    "10de:1e91": "GeForce RTX 2070 Super",
    "10de:1f08": "GeForce RTX 2060",
    "10de:1f06": "GeForce RTX 2060 Super",
    "10de:1e90": "GeForce RTX 2080 Laptop",
    "10de:1f91": "GeForce RTX 2070 Laptop",
    "10de:1f15": "GeForce RTX 2060 Laptop",

    # ══════════════════════════════════════════════════════════════
    # NVIDIA — GTX 16 series (Turing)
    # ══════════════════════════════════════════════════════════════
    "10de:2182": "GeForce GTX 1660 Ti",
    "10de:2184": "GeForce GTX 1660",
    "10de:21c4": "GeForce GTX 1660 Super",
    "10de:1f82": "GeForce GTX 1650",
    "10de:2187": "GeForce GTX 1650 Super",
    "10de:1f9d": "GeForce GTX 1650 Ti Laptop",
    "10de:1f99": "GeForce GTX 1650 Laptop",

    # ══════════════════════════════════════════════════════════════
    # NVIDIA — GTX 10 series (Pascal)
    # ══════════════════════════════════════════════════════════════
    "10de:1b06": "GeForce GTX 1080 Ti",
    "10de:1b80": "GeForce GTX 1080",
    "10de:1b81": "GeForce GTX 1070",
    "10de:1b82": "GeForce GTX 1070 Ti",
    "10de:1c03": "GeForce GTX 1060 6GB",
    "10de:1c02": "GeForce GTX 1060 3GB",
    "10de:1c82": "GeForce GTX 1050 Ti",
    "10de:1c81": "GeForce GTX 1050",
    "10de:1c8d": "GeForce GTX 1050 Laptop",
    "10de:1be1": "GeForce GTX 1080 Laptop",
    "10de:1be0": "GeForce GTX 1070 Laptop",
    "10de:1c20": "GeForce GTX 1060 Laptop",

    # ══════════════════════════════════════════════════════════════
    # NVIDIA — GTX 900 series (Maxwell)
    # ══════════════════════════════════════════════════════════════
    "10de:17c8": "GeForce GTX 980 Ti",
    "10de:13c0": "GeForce GTX 980",
    "10de:13c2": "GeForce GTX 970",
    "10de:1401": "GeForce GTX 960",
    "10de:1402": "GeForce GTX 950",
    "10de:1617": "GeForce GTX 980M",
    "10de:1618": "GeForce GTX 970M",
    "10de:1619": "GeForce GTX 960M",

    # ══════════════════════════════════════════════════════════════
    # NVIDIA — GTX 700 series (Kepler)
    # ══════════════════════════════════════════════════════════════
    "10de:1004": "GeForce GTX 780 Ti",
    "10de:1005": "GeForce GTX 780",
    "10de:1187": "GeForce GTX 770",
    "10de:1189": "GeForce GTX 760",

    # ══════════════════════════════════════════════════════════════
    # AMD — RX 9000 series (RDNA 4)
    # ══════════════════════════════════════════════════════════════
    "1002:7518": "Radeon RX 9070 XT",
    "1002:7580": "Radeon RX 9070",
    "1002:7590": "Radeon RX 9060 XT",

    # ══════════════════════════════════════════════════════════════
    # AMD — RX 7000 series (RDNA 3)
    # ══════════════════════════════════════════════════════════════
    "1002:744c": "Radeon RX 7900 XTX",
    "1002:7448": "Radeon RX 7900 XT",
    "1002:744e": "Radeon RX 7900 GRE",
    "1002:747e": "Radeon RX 7800 XT",
    "1002:7483": "Radeon RX 7700 XT",
    "1002:7489": "Radeon RX 7700",
    "1002:7422": "Radeon RX 7600",
    "1002:7424": "Radeon RX 7600 XT",
    "1002:7466": "Radeon RX 7600M Laptop",
    "1002:7474": "Radeon RX 7600S Laptop",

    # ══════════════════════════════════════════════════════════════
    # AMD — RX 6000 series (RDNA 2)
    # ══════════════════════════════════════════════════════════════
    "1002:73bf": "Radeon RX 6950 XT",
    "1002:73b7": "Radeon RX 6800 XT",
    "1002:73b8": "Radeon RX 6800",
    "1002:73ef": "Radeon RX 6750 XT",
    "1002:73df": "Radeon RX 6700 XT",
    "1002:73e3": "Radeon RX 6700",
    "1002:73e4": "Radeon RX 6750 GRE",
    "1002:73ff": "Radeon RX 6600 XT",
    "1002:73a0": "Radeon RX 6600",
    "1002:73a4": "Radeon RX 6500 XT",
    "1002:743f": "Radeon RX 6400",
    "1002:73e1": "Radeon RX 6700M Laptop",
    "1002:7360": "Radeon RX 6600M Laptop",

    # ══════════════════════════════════════════════════════════════
    # AMD — RX 5000 series (RDNA 1)
    # ══════════════════════════════════════════════════════════════
    "1002:731f": "Radeon RX 5700 XT",
    "1002:7310": "Radeon RX 5700",
    "1002:7362": "Radeon RX 5600 XT",
    "1002:7340": "Radeon RX 5500 XT",
    "1002:7341": "Radeon RX 5500",
    "1002:7347": "Radeon RX 5500M Laptop",

    # ══════════════════════════════════════════════════════════════
    # AMD — RX 500 / Vega series (Polaris / Vega)
    # ══════════════════════════════════════════════════════════════
    "1002:687f": "Radeon RX Vega 64",
    "1002:6863": "Radeon RX Vega 56",
    "1002:67df": "Radeon RX 580",
    "1002:67ef": "Radeon RX 570",
    "1002:67ff": "Radeon RX 560",
    "1002:699f": "Radeon RX 550",

    # ══════════════════════════════════════════════════════════════
    # Intel — Arc (Alchemist)
    # ══════════════════════════════════════════════════════════════
    "8086:56a0": "Intel Arc A770",
    "8086:56a1": "Intel Arc A750",
    "8086:56a5": "Intel Arc A580",
    "8086:56a6": "Intel Arc A380",
    "8086:56b0": "Intel Arc A770M Laptop",
    "8086:56b1": "Intel Arc A730M Laptop",
    "8086:56b2": "Intel Arc A550M Laptop",
    "8086:56c0": "Intel Arc A370M Laptop",
    "8086:56c1": "Intel Arc A350M Laptop",
    "8086:5690": "Intel Arc A370M Laptop",
    "8086:5691": "Intel Arc A350M Laptop",

    # ══════════════════════════════════════════════════════════════
    # Intel — Iris Xe (Tiger / Alder / Raptor Lake)
    # ══════════════════════════════════════════════════════════════
    "8086:9a40": "Intel Iris Xe Graphics",
    "8086:9a49": "Intel Iris Xe Graphics",
    "8086:9a60": "Intel Iris Xe Graphics",
    "8086:9a68": "Intel Iris Xe Graphics",
    "8086:9a70": "Intel Iris Xe Graphics",
    "8086:9a78": "Intel Iris Xe Graphics",
    "8086:46a6": "Intel Iris Xe Graphics",
    "8086:46a8": "Intel Iris Xe Graphics",
    "8086:4626": "Intel Iris Xe Graphics",
    "8086:4628": "Intel Iris Xe Graphics",
    "8086:46d0": "Intel Iris Xe Graphics",
    "8086:46d1": "Intel Iris Xe Graphics",
    "8086:46d2": "Intel Iris Xe Graphics",
    "8086:a780": "Intel Iris Xe Graphics",
    "8086:a781": "Intel Iris Xe Graphics",
    "8086:a788": "Intel Iris Xe Graphics",
    "8086:a789": "Intel Iris Xe Graphics",

    # ══════════════════════════════════════════════════════════════
    # Intel — UHD (integrated)
    # ══════════════════════════════════════════════════════════════
    "8086:3e92": "Intel UHD Graphics 630",
    "8086:3e9b": "Intel UHD Graphics 630",
    "8086:3e98": "Intel UHD Graphics 630",
    "8086:9bc5": "Intel UHD Graphics 630",
    "8086:9bc8": "Intel UHD Graphics 630",
    "8086:4680": "Intel UHD Graphics 770",
    "8086:4682": "Intel UHD Graphics 770",
    "8086:4692": "Intel UHD Graphics 730",
    "8086:4693": "Intel UHD Graphics 730",
    "8086:4698": "Intel UHD Graphics 710",
    "8086:4699": "Intel UHD Graphics 710",
}


def _linux_detect_gpu_pci_id():
    try:
        output = subprocess.check_output(
            ["lspci", "-nn"], encoding="utf-8", stderr=subprocess.DEVNULL
        )
        intel_id = None
        for line in output.splitlines():
            if "VGA" in line or "3D controller" in line:
                match = re.search(r"\[(\w{4}:\w{4})]", line)
                if match:
                    pci_id = match.group(1).lower()
                    if pci_id.startswith("8086"):
                        intel_id = pci_id
                    else:
                        return pci_id
        return intel_id
    except (OSError, subprocess.CalledProcessError, UnicodeDecodeError):
        pass
    return None


def _linux_detect_gpu_name():
    import glob

    for proc_gpu_info in glob.glob("/proc/driver/nvidia/gpus/*/information"):
        try:
            with open(proc_gpu_info, "r", encoding="utf-8") as f:
                for line in f:
                    if line.lower().startswith("model:"):
                        return line.split(":", 1)[1].strip()
        except OSError:
            continue

    try:
        output = subprocess.check_output(
            ["lspci", "-nn"],
            encoding="utf-8",
            stderr=subprocess.DEVNULL,
            timeout=2,
        )
        intel_name = None
        for line in output.splitlines():
            if "VGA" not in line and "3D controller" not in line:
                continue

            match = re.search(r":\s(.+?)\s\[[0-9a-fA-F]{4}:[0-9a-fA-F]{4}](?:\s\(rev .+\))?$", line)
            if not match:
                continue

            gpu_name = match.group(1).strip()
            gpu_name = re.sub(
                r"^(NVIDIA Corporation|Intel Corporation|Advanced Micro Devices, Inc\. \[AMD/ATI])\s+",
                "",
                gpu_name,
            )
            if "[8086:" in line:
                intel_name = gpu_name
            else:
                return gpu_name
        return intel_name
    except (OSError, subprocess.CalledProcessError, subprocess.TimeoutExpired, UnicodeDecodeError):
        return None


def _read_sysfs(path):
    try:
        with open(path, "r") as f:
            return int(f.read().strip())
    except (OSError, ValueError):
        return None


def _read_text(path):
    try:
        with open(path, "r") as f:
            return f.read().strip()
    except OSError:
        return None


def _parse_first_number(value):
    if value is None:
        return None
    if isinstance(value, (int, float)):
        return float(value)
    match = re.search(r"[-+]?\d*\.?\d+", str(value))
    if not match:
        return None
    try:
        return float(match.group(0))
    except ValueError:
        return None


def _safe_linux_probe(probe, default):
    try:
        return probe()
    except Exception as e:
        print(f"[LINUX] Probe failed for {getattr(probe, '__name__', 'probe')}: {e}")
        return default


def _linux_powercap_energy_files():
    base_dir = "/sys/class/powercap"
    energy_files = []

    try:
        entries = list(os.scandir(base_dir))
    except OSError:
        return energy_files

    for entry in entries:
        if not entry.is_dir(follow_symlinks=True):
            continue

        direct_energy = os.path.join(entry.path, "energy_uj")
        if os.path.isfile(direct_energy):
            energy_files.append(direct_energy)

        try:
            children = list(os.scandir(entry.path))
        except OSError:
            continue

        for child in children:
            if not child.is_dir(follow_symlinks=True):
                continue
            child_energy = os.path.join(child.path, "energy_uj")
            if os.path.isfile(child_energy):
                energy_files.append(child_energy)

    return energy_files


def _linux_get_cpu_temp_celsius():
    import glob

    for hwmon_dir in glob.glob("/sys/class/hwmon/hwmon*"):
        try:
            with open(f"{hwmon_dir}/name") as f:
                chip_name = f.read().strip()
        except OSError:
            continue

        if chip_name in ("coretemp", "k10temp"):
            for label_path in sorted(glob.glob(f"{hwmon_dir}/temp*_label")):
                try:
                    with open(label_path) as f:
                        label = f.read().strip().lower()
                except OSError:
                    continue

                if any(k in label for k in ["package", "tctl", "tdie"]):
                    input_path = label_path.replace("_label", "_input")
                    raw = _read_sysfs(input_path)
                    if raw is not None:
                        return raw // 1000

            raw = _read_sysfs(f"{hwmon_dir}/temp1_input")
            if raw is not None:
                return raw // 1000

    return 0


def _linux_get_cpu_power_watts():
    import glob

    sample_interval = 0.2

    energy_files = _linux_powercap_energy_files()
    preferred = []
    fallback = []
    for energy_file in energy_files:
        sensor_name = _read_text(os.path.join(os.path.dirname(energy_file), "name"))
        sensor_name_l = (sensor_name or "").lower()
        if any(k in sensor_name_l for k in ("package", "cpu", "core", "amd", "intel")):
            preferred.append(energy_file)
        else:
            fallback.append(energy_file)

    energy_candidates = preferred + fallback
    first_reads = {}
    for energy_file in energy_candidates:
        e1 = _read_sysfs(energy_file)
        if e1 is not None:
            first_reads[energy_file] = e1

    if first_reads:
        time.sleep(sample_interval)
        for energy_file, e1 in first_reads.items():
            e2 = _read_sysfs(energy_file)
            if e2 is not None and e2 > e1:
                return int((e2 - e1) / sample_interval / 1_000_000)

    best_guess = 0
    for hwmon in glob.glob("/sys/class/hwmon/hwmon*"):
        name_l = (_read_text(f"{hwmon}/name") or "").lower()
        raw = _read_sysfs(f"{hwmon}/power1_average")
        if raw is None:
            raw = _read_sysfs(f"{hwmon}/power1_input")
        if raw is None or raw <= 0:
            continue

        watts = raw / 1_000_000 if raw > 10_000 else float(raw)
        watts_i = int(watts)
        if watts_i <= 0:
            continue

        if any(k in name_l for k in ("cpu", "k10temp", "fam17h", "zenpower", "coretemp", "rapl", "intel", "amd")):
            return watts_i
        best_guess = max(best_guess, watts_i)

    try:
        sensors_json = subprocess.check_output(
            ["sensors", "-j"],
            encoding="utf-8",
            stderr=subprocess.DEVNULL,
            timeout=2,
        )
        sensors_data = json.loads(sensors_json)
        if isinstance(sensors_data, dict):
            for chip_name, chip_metrics in sensors_data.items():
                if not isinstance(chip_metrics, dict):
                    continue
                chip_l = str(chip_name).lower()
                if not any(k in chip_l for k in ("k10temp", "coretemp", "cpu", "package", "fam17h", "zen")):
                    continue
                for sensor_name, sensor_data in chip_metrics.items():
                    if "power" not in str(sensor_name).lower() or not isinstance(sensor_data, dict):
                        continue
                    for metric_name, metric_value in sensor_data.items():
                        metric_l = str(metric_name).lower()
                        if "input" not in metric_l and "average" not in metric_l:
                            continue
                        value_num = _parse_first_number(metric_value)
                        if value_num is not None and value_num > 0:
                            return int(value_num)
    except (
        FileNotFoundError,
        OSError,
        UnicodeDecodeError,
        subprocess.CalledProcessError,
        subprocess.TimeoutExpired,
        json.JSONDecodeError,
    ):
        pass

    if best_guess > 0:
        return best_guess
    return 0


def _linux_get_gpu_stats():
    import glob

    best_temp = 0
    best_power = 0
    best_load = 0

    for hwmon in glob.glob("/sys/class/drm/card*/device/hwmon/hwmon*"):
        temp = _read_sysfs(f"{hwmon}/temp1_input")
        power = (
            _read_sysfs(f"{hwmon}/power1_average")
            or _read_sysfs(f"{hwmon}/power1_input")
        )
        card_device = hwmon.split("/hwmon/")[0]
        load = _read_sysfs(f"{card_device}/gpu_busy_percent")

        if temp is not None and temp > 0:
            best_temp = max(best_temp, int(temp // 1000))
        if power is not None and power > 0:
            watts = int(power // 1_000_000) if power > 10_000 else int(power)
            best_power = max(best_power, watts)
        if load is not None and load >= 0:
            best_load = max(best_load, int(load))

    if best_temp or best_power or best_load:
        return best_temp, best_power, best_load

    for hwmon in glob.glob("/sys/class/hwmon/hwmon*"):
        name_l = (_read_text(f"{hwmon}/name") or "").lower()
        if not any(k in name_l for k in ("amdgpu", "radeon", "nouveau", "nvidia")):
            continue

        temp = _read_sysfs(f"{hwmon}/temp1_input")
        power = (
            _read_sysfs(f"{hwmon}/power1_average")
            or _read_sysfs(f"{hwmon}/power1_input")
        )
        if temp is not None and temp > 0:
            best_temp = max(best_temp, int(temp // 1000))
        if power is not None and power > 0:
            watts = int(power // 1_000_000) if power > 10_000 else int(power)
            best_power = max(best_power, watts)

    if best_temp or best_power:
        return best_temp, best_power, best_load

    try:
        out = subprocess.check_output(
            ["nvidia-smi", "--query-gpu=temperature.gpu,power.draw,utilization.gpu", "--format=csv,noheader,nounits"],
            encoding="utf-8",
            stderr=subprocess.DEVNULL,
            timeout=2,
        ).strip()
        if out:
            first = out.splitlines()[0]
            parts = [p.strip() for p in first.split(",")]
            temp = _parse_first_number(parts[0] if len(parts) > 0 else None)
            power = _parse_first_number(parts[1] if len(parts) > 1 else None)
            load = _parse_first_number(parts[2] if len(parts) > 2 else None)
            return int(temp or 0), int(power or 0), int(load or 0)
    except (FileNotFoundError, OSError, UnicodeDecodeError, subprocess.CalledProcessError, subprocess.TimeoutExpired):
        pass

    try:
        out = subprocess.check_output(
            ["rocm-smi", "--showtemp", "--showpower", "--showuse", "--json"],
            encoding="utf-8",
            stderr=subprocess.DEVNULL,
            timeout=3,
        )
        data = json.loads(out)
        temp = power = load = 0.0
        for _, metrics in data.items():
            if not isinstance(metrics, dict):
                continue
            for metric_name, metric_value in metrics.items():
                name_l = str(metric_name).lower()
                value_num = _parse_first_number(metric_value)
                if value_num is None or value_num <= 0:
                    continue
                if "temp" in name_l:
                    temp = max(temp, value_num)
                elif "power" in name_l:
                    power = max(power, value_num)
                elif "use" in name_l:
                    load = max(load, value_num)
        if temp or power or load:
            return int(temp), int(power), int(load)
    except (
        FileNotFoundError,
        OSError,
        UnicodeDecodeError,
        subprocess.CalledProcessError,
        subprocess.TimeoutExpired,
        json.JSONDecodeError,
        AttributeError,
        TypeError,
        ValueError,
    ):
        pass

    return 0, 0, 0


def _linux_get_vram_usage_gb():
    import glob

    for card in glob.glob("/sys/class/drm/card*/device"):
        total = _read_sysfs(f"{card}/mem_info_vram_total")
        used = _read_sysfs(f"{card}/mem_info_vram_used")
        if total is not None and used is not None and total > 0:
            return used / (1024 ** 3), total / (1024 ** 3)

    try:
        out = subprocess.check_output(
            ["nvidia-smi", "--query-gpu=memory.used,memory.total", "--format=csv,noheader,nounits"],
            encoding="utf-8",
            stderr=subprocess.DEVNULL,
            timeout=2,
        ).strip()
        if out:
            first = out.splitlines()[0]
            used_str, total_str = [x.strip() for x in first.split(",")[:2]]
            used_mib = float(used_str)
            total_mib = float(total_str)
            return used_mib / 1024, total_mib / 1024
    except (FileNotFoundError, OSError, UnicodeDecodeError, subprocess.CalledProcessError, subprocess.TimeoutExpired, ValueError):
        pass

    return None, None


def diagnose_lhm():
    if sys.platform != "win32":
        print("[DIAGNOSTIC] Running Linux hardware diagnostics...")

        try:
            cpu_t = _linux_get_cpu_temp_celsius()
        except Exception as e:
            cpu_t = 0
            print(f"[DIAGNOSTIC] ✗ CPU Temperature probe failed: {e}")
        if cpu_t:
            print(f"[DIAGNOSTIC] ✓ CPU Temperature: {cpu_t}°C")
        else:
            print("[DIAGNOSTIC] ✗ CPU Temperature: not found")
            print("[DIAGNOSTIC]   FIX: sudo apt install lm-sensors && sudo sensors-detect")

        try:
            cpu_p = _linux_get_cpu_power_watts()
        except Exception as e:
            cpu_p = 0
            print(f"[DIAGNOSTIC] ✗ CPU Power probe failed: {e}")
        if cpu_p:
            print(f"[DIAGNOSTIC] ✓ CPU Power: {cpu_p}W")
        else:
            print("[DIAGNOSTIC] ✗ CPU Power: not found")

        try:
            gpu_t, gpu_p, gpu_l = _linux_get_gpu_stats()
        except Exception as e:
            gpu_t, gpu_p, gpu_l = 0, 0, 0
            print(f"[DIAGNOSTIC] ✗ GPU probe failed: {e}")
        if gpu_t or gpu_p or gpu_l:
            print(f"[DIAGNOSTIC] ✓ GPU: {gpu_t}°C  {gpu_p}W  {gpu_l}% load")
        else:
            print("[DIAGNOSTIC] ✗ GPU stats: not found")

        try:
            net_io = psutil.net_io_counters(pernic=True) or {}
            all_ifaces = list(net_io.keys())
        except Exception as e:
            all_ifaces = []
            print(f"[DIAGNOSTIC] ✗ Network interfaces probe failed: {e}")
        if INTERFACE in all_ifaces:
            print(f"[DIAGNOSTIC] ✓ Network interface '{INTERFACE}' found")
        else:
            print(f"[DIAGNOSTIC] ✗ Interface '{INTERFACE}' not found. Available: {all_ifaces}")
        return True

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
        print("[DIAGNOSTIC] FIX 4: Get LHM at https://github.com/LibreHardwareMonitor/LibreHardwareMonitor/releases")
    except requests.Timeout:
        print("[DIAGNOSTIC] ✗ REST API query timed out")
    except Exception as e:
        print(f"[DIAGNOSTIC] ✗ Error: {e}")
    return False


def get_lhm_data():
    if sys.platform != "win32":
        gpu_temp_value, gpu_power_value, gpu_load_value = _safe_linux_probe(_linux_get_gpu_stats, (0, 0, 0))
        cpu_temp_value = _safe_linux_probe(_linux_get_cpu_temp_celsius, 0)
        cpu_power_value = _safe_linux_probe(_linux_get_cpu_power_watts, 0)
        cpu_load_value = _safe_linux_probe(lambda: int(psutil.cpu_percent(interval=None)), 0)
        return {
            "cpu_temp": cpu_temp_value,
            "cpu_power": cpu_power_value,
            "gpu_temp": gpu_temp_value,
            "gpu_power": gpu_power_value,
            "gpu_load": gpu_load_value,
            "cpu_load": cpu_load_value,
        }

    try:
        response = requests.get(LHM_REST_API, timeout=5)
        if response.status_code == 200:
            return response.json()
    except (requests.ConnectionError, requests.Timeout, json.JSONDecodeError):
        pass
    return None


def _parse_gb(sensor_value) -> float:
    numeric_str = re.sub(r'[^\d.-]', '', str(sensor_value))
    return float(numeric_str)


def _fmt_gb(gb: float) -> str:
    rounded = round(gb)
    for nice in [2, 4, 6, 8, 10, 12, 16, 20, 24, 32, 48, 64, 96, 128]:
        if abs(rounded - nice) <= max(1, int(nice * 0.10)):
            return f"{nice}"
    return f"{rounded}"


# ═════════════════════════════════════════════════════════════════════════════════════════════════════════════════════#
#  CPU SENSORS
# ═════════════════════════════════════════════════════════════════════════════════════════════════════════════════════#

def detect_cpu():
    global cpu_manufacturer

    if sys.platform == "win32":
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

    try:
        with open("/proc/cpuinfo", "r") as f:
            for line in f:
                if line.startswith("model name"):
                    cpu_name = line.split(":", 1)[1].strip()
                    if "intel" in cpu_name.lower():
                        cpu_manufacturer = CPUManufacturer.INTEL
                    elif "amd" in cpu_name.lower():
                        cpu_manufacturer = CPUManufacturer.AMD
                    else:
                        cpu_manufacturer = CPUManufacturer.UNKNOWN
                    return _clean_name(cpu_name)
    except (OSError, IOError):
        pass
    return "CPU Unknown"


def get_cpu_temp_from_lhm(data) -> int:
    if not data:
        return 0

    primary = ("cpu package", "tdie")
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

    primary = ("cpu package",)
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
    if sys.platform != "win32":
        if not data:
            return 0
        return int(data.get("cpu_load", 0))

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

def detect_gpu():
    if sys.platform == "win32":
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

    pci_id = _linux_detect_gpu_pci_id()
    if pci_id in linux_gpu_id_map:
        return linux_gpu_id_map[pci_id]

    gpu_name = _linux_detect_gpu_name()
    if gpu_name:
        return _clean_name(gpu_name)

    return f"Unknown GPU ({pci_id})" if pci_id else "GPU Unknown"


def get_gpu_temp_from_lhm(data) -> int:
    if not data:
        return 0

    primary = ("gpu core", "gpu temperature", "gpu temp")
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
    if sys.platform != "win32":
        if not data:
            return 0
        return int(data.get("gpu_load", 0))

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

def get_dram_total_from_lhm(data) -> str:
    try:
        total_bytes = psutil.virtual_memory().total
        total_gb = total_bytes / (1024 ** 3)
        return f"{_fmt_gb(total_gb)}"
    except (OSError, AttributeError):
        return "error"


def get_dram_used_from_lhm(data) -> float:
    if sys.platform != "win32":
        try:
            used_bytes = psutil.virtual_memory().used
            return round(used_bytes / (1024 ** 3), 1)
        except (OSError, AttributeError):
            return 0.0

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
    if sys.platform != "win32":
        used_gb, _ = _safe_linux_probe(_linux_get_vram_usage_gb, (None, None))
        if used_gb is None:
            return 0.0
        return round(used_gb, 1)

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
    if sys.platform != "win32":
        _, total_gb = _safe_linux_probe(_linux_get_vram_usage_gb, (None, None))
        if total_gb is None:
            return "error"
        return _fmt_gb(total_gb)

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
    if sys.platform != "win32":
        if not data:
            return 0, 0, 0, 0
        return (
            int(data.get("cpu_temp", 0)),
            int(data.get("cpu_power", 0)),
            int(data.get("gpu_temp", 0)),
            int(data.get("gpu_power", 0)),
        )

    return (
        get_cpu_temp_from_lhm(data),
        get_cpu_power_from_lhm(data),
        get_gpu_temp_from_lhm(data),
        get_gpu_power_from_lhm(data),
    )


def detect_dram_type() -> str:
    if sys.platform != "win32":
        return "DDR"

    try:
        result = subprocess.check_output(
            ["powershell", "-Command",
             "(Get-CimInstance Win32_PhysicalMemory | Select-Object -First 1).SMBIOSMemoryType"],
            encoding="utf-8", stderr=subprocess.DEVNULL, timeout=5
        ).strip()
        type_map = {
            "17": "SDRAM", "18": "SDRAM", "19": "SDRAM",
            "20": "DDR", "21": "DDR2", "22": "DDR2",
            "23": "DDR2", "24": "DDR3", "26": "DDR4",
            "34": "DDR5", "35": "DDR5"
        }
        return type_map.get(result, "DDR")
    except (subprocess.CalledProcessError, UnicodeDecodeError, subprocess.TimeoutExpired):
        return "DDR"


def detect_vram_type(gpu_name: str) -> str:
    n = gpu_name.lower()

    if any(x in n for x in ["5090", "5080", "5070 ti", "5070ti", "5070", "5060 ti", "5060ti", "5060"]):
        return "GDDR7"
    if any(x in n for x in ["4090", "4080", "4070 ti", "4070ti", "4070 super", "4070super"]):
        return "GDDR6X"
    if any(x in n for x in ["4070", "4060 ti", "4060ti", "4060"]):
        return "GDDR6"
    if any(x in n for x in ["3090", "3080"]):
        return "GDDR6X"
    if any(x in n for x in ["3070", "3060", "rtx 30", "rtx30"]):
        return "GDDR6"
    if any(x in n for x in ["rtx 20", "rtx20", "2080", "2070", "2060"]):
        return "GDDR6"
    if any(x in n for x in ["1080 ti", "1080ti", "1080"]):
        return "GDDR5X"
    if any(x in n for x in ["gtx 10", "gtx10", "1070", "1060", "1050"]):
        return "GDDR5"
    if any(x in n for x in ["980 ti", "980ti", "980", "970"]):
        return "GDDR5"
    if any(x in n for x in ["rx 9", "rx9"]):
        return "GDDR6"
    if any(x in n for x in ["rx 7", "rx7"]):
        return "GDDR6"
    if any(x in n for x in ["rx 6", "rx6"]):
        return "GDDR6"
    if any(x in n for x in ["rx 5", "rx5"]):
        return "GDDR6"
    if any(x in n for x in ["rx 5", "rx 4", "rx5", "rx4", "rx 580", "rx 570", "rx 480", "rx 470"]):
        return "GDDR5"
    if any(x in n for x in ["radeon", "nvidia", "geforce"]):
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
# WEATHER MONITORING
# ═════════════════════════════════════════════════════════════════════════════════════════════════════════════════════#

WMO_CODES = {
    0: "Clear sky",
    1: "Mainly clear",
    2: "Partly cloudy",
    3: "Overcast",
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


def fetch_weather(lat_lon_str: str):
    global weather_temp, weather_humidity, weather_desc

    try:
        parts = [p.strip() for p in lat_lon_str.split(",")]
        if len(parts) != 2:
            raise ValueError("Expected format: latitude,longitude  (e.g. 51.5,-0.1)")
        lat, lon = float(parts[0]), float(parts[1])
    except (ValueError, AttributeError) as e:
        print(f"[Weather] Bad location format: {e}")
        weather_temp = "?"
        weather_humidity = "?"
        weather_desc = "Bad location"
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

        current = data["current"]
        weather_temp = round(current["temperature_2m"])
        weather_humidity = current["relative_humidity_2m"]
        weather_desc = WMO_CODES.get(current["weather_code"], "Unknown")

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

    weather_temp = "?"
    weather_humidity = "?"
    weather_desc = "Unavailable"
    return False


# ═════════════════════════════════════════════════════════════════════════════════════════════════════════════════════#
# MEDIA MONITORING
# ═════════════════════════════════════════════════════════════════════════════════════════════════════════════════════#

async def get_media_info():
    if sys.platform == "win32" and wmc is not None:
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
                if props is None:
                    return None, None, pos, dur, is_paused
                return props.title, props.artist, pos, dur, is_paused
        except (OSError, AttributeError, RuntimeError):
            pass

    try:
        players = subprocess.check_output(
            ["playerctl", "-l"],
            encoding="utf-8",
            stderr=subprocess.DEVNULL,
            timeout=2
        ).splitlines()

        if not players:
            return None, None, 0, 0, False

        browser_player = next(
            (p for p in players if any(k in p.lower() for k in ["chrome", "chromium", "firefox"])),
            players[0]
        )

        output = subprocess.check_output(
            ["playerctl", "-p", browser_player, "metadata",
             "--format", "{{title}}\n{{artist}}\n{{position}}\n{{mpris:length}}"],
            encoding="utf-8",
            stderr=subprocess.DEVNULL,
            timeout=2
        ).strip().split("\n")

        if len(output) >= 4:
            title = output[0].strip()
            artist = output[1].strip()
            pos = int(output[2]) / 1000
            dur = int(output[3]) / 1000
            is_paused = False

            try:
                status = subprocess.check_output(
                    ["playerctl", "-p", browser_player, "status"],
                    encoding="utf-8",
                    stderr=subprocess.DEVNULL,
                    timeout=2
                ).strip().lower()
                is_paused = status == "paused"
            except subprocess.CalledProcessError:
                pass

            return title, artist, pos, dur, is_paused

    except FileNotFoundError:
        pass
    except subprocess.CalledProcessError:
        pass
    except Exception as e:
        print(f"[MEDIA ERROR] {e}")

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


def create_progress_bar(position_ms, duration_ms, length=15):
    if duration_ms <= 0:
        return "No music playing"
    percent = min(max(position_ms / duration_ms, 0), 1)
    filled_len = int(length * percent)
    if 0 < filled_len < length:
        return (
            progress_filled_char * filled_len
            + progress_border_char
            + progress_empty_char * (length - filled_len - 1)
        )
    return progress_filled_char * filled_len + progress_empty_char * (length - filled_len)


def format_telemetry(value, suffix=""):
    try:
        numeric = float(value)
    except (TypeError, ValueError):
        return "N/A"

    if numeric <= 0:
        return "N/A"
    if numeric.is_integer():
        return f"{int(numeric)}{suffix}"
    return f"{numeric:.1f}{suffix}"


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

    query_cooldown = 0
    weather_cooldown = 0
    weather_interval = 60

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
                    cpu_load = get_cpu_load_from_lhm(lhm_data)
                    gpu_load = get_gpu_load_from_lhm(lhm_data)
                    dram_load = get_dram_used_from_lhm(lhm_data)
                    vram_load = get_vram_used_from_lhm(lhm_data)
                else:
                    cpu_load = gpu_load = 0
                    dram_load = vram_load = 0.0
                query_cooldown = 0

            weather_cooldown += 1
            if weather_cooldown >= weather_interval:
                fetch_weather(location_entry.get().strip())
                weather_cooldown = 0

            if vram_detect == "error":
                vram_detect = get_vram_total_from_lhm(lhm_data)
            if dram_detect == "error":
                dram_detect = get_dram_total_from_lhm(lhm_data)

            prev, up_raw, down_raw, prev_time = get_network_usage(prev, prev_time)

            cur_time_str = time.strftime("%I:%M %p")
            progress_bar = create_progress_bar(pos, dur)
            cpu_power_text = format_telemetry(cpu_wattage, "w")
            cpu_temp_text = format_telemetry(cpu_temp, "℃")
            gpu_power_text = format_telemetry(gpu_wattage, "w")
            gpu_temp_text = format_telemetry(gpu_temp, "℃")

            if clean_song:
                if is_paused:
                    display_artist = f"- {artist}" if artist else ""
                    display_song = f"⏸ {clean_song}" if clean_song else ""
                else:
                    display_artist = f"- {artist}" if artist else ""
                    display_song = f"🎵 {clean_song}" if clean_song else ""
            else:
                display_artist = ""
                display_song = "⏸"

            media_line = display_song
            if display_artist:
                media_line = f"{media_line} {display_artist}" if media_line else display_artist

            enabled_pages = [i for i in range(4) if page_toggles[i].get()]
            if not enabled_pages:
                text = "No pages enabled"
                print(text)
                if client is not None:
                    client.send_message("/chatbox/input", [text, True])
                else:
                    print("Warning: OSC client not initialized")
                time.sleep(5.0)
                continue

            enabled_count = len(enabled_pages)
            switch_interval = max(1, int(SWITCH_INTERVAL))
            page_slot = int((time.time() // switch_interval) % enabled_count)
            page_index = enabled_pages[page_slot]

            if forced_text.get().strip() == "":

                if page_index == 0:
                    text = (
                        f"{page1_line1_text}\n"
                        f"{cur_time_str}\n"
                        f"Download {fmt(down_raw)}\n"
                        f"Upload {fmt(up_raw)}\n"
                        f"{progress_bar}\n"
                        f"{media_line}"
                    )
                elif page_index == 1:
                    text = (
                        f"{page2_line1_text}\n"
                        f"{cur_time_str}\n"
                        f"{cpu_detect} {cpu_load}%\n"
                        f"{cpu_power_text} {cpu_temp_text}\n"
                        f"{gpu_detect} {gpu_load}%\n"
                        f"{gpu_power_text} {gpu_temp_text}\n"
                    )
                elif page_index == 2:
                    text = (
                        f"{page3_line1_text}\n"
                        f"{cur_time_str}\n"
                        f"{dram} {dram_load}GB/{dram_detect}GB\n"
                        f"{vram} {vram_load}GB/{vram_detect}GB\n"
                        f"{progress_bar}\n"
                        f"{media_line}"
                    )
                elif page_index == 3:
                    text = (
                        f"{page4_line1_text}\n"
                        f"{cur_time_str}\n"
                        f"{weather_temp}℃  {weather_humidity}% humidity\n"
                        f"{weather_desc}\n"
                        f"{progress_bar}\n"
                        f"{media_line}"
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
        OSC_IP = ip_entry.get()
        OSC_PORT = int(port_entry.get())
        INTERFACE = iface_entry.get()
        SWITCH_INTERVAL = int(interval_entry.get())
        LHM_REST_API = lhm_entry.get()

        page1_line1_text = page1_entry.get()
        page2_line1_text = page2_entry.get()
        page3_line1_text = page3_entry.get()
        page4_line1_text = page4_entry.get()

        save_config()
        error_text = "Error: Page error value exceeding limit"
        client = SimpleUDPClient(OSC_IP, OSC_PORT)

        running = True
        status_label.config(text="Status: Running", fg=GREEN)

        diagnose_lhm()

        thread = threading.Thread(target=run_osc_loop, daemon=True)
        thread.start()

    except ValueError as e:
        messagebox.showerror("Error", f"Invalid input: {e}")
        status_label.config(text="Status: Error", fg=RED)
    except Exception as e:
        messagebox.showerror("Error", f"Failed to start: {e}")
        status_label.config(text="Status: Error", fg=RED)


def stop_script():
    global running
    running = False
    status_label.config(text="Status: Stopped", fg=RED)


def restart_script():
    stop_script()
    time.sleep(1)
    start_script()


# ═════════════════════════════════════════════════════════════════════════════════════════════════════════════════════#
# CIRCLE TOGGLE WIDGET
# ═════════════════════════════════════════════════════════════════════════════════════════════════════════════════════#

class CircleToggle(tk.Canvas):

    SIZE = 22
    PAD = 3
    COLOR = "#a78bfa"

    def __init__(self, parent, enabled=True, **kwargs):
        super().__init__(
            parent,
            width=self.SIZE, height=self.SIZE,
            bg=BG, highlightthickness=0,
            cursor="hand2",
            **kwargs
        )
        self._enabled = enabled
        self._draw()
        self.bind("<Button-1>", self._on_click)

    def _draw(self):
        self.delete("all")
        p, s = self.PAD, self.SIZE
        if self._enabled:
            self.create_oval(p, p, s - p, s - p,
                             fill=self.COLOR, outline=self.COLOR)
        else:
            self.create_oval(p, p, s - p, s - p,
                             fill="", outline=self.COLOR, width=2)

    def _on_click(self, _event=None):
        self._enabled = not self._enabled
        self._draw()

    def get(self) -> bool:
        return self._enabled

    def set(self, value: bool):
        self._enabled = bool(value)
        self._draw()


# ═════════════════════════════════════════════════════════════════════════════════════════════════════════════════════#
# GUI
# ═════════════════════════════════════════════════════════════════════════════════════════════════════════════════════#

cfg = load_config()
progress_filled_char = normalize_progress_char(
    cfg.get("progress_filled_char"), DEFAULT_PROGRESS_FILLED_CHAR
)
progress_border_char = normalize_progress_char(
    cfg.get("progress_border_char"), DEFAULT_PROGRESS_BORDER_CHAR
)
progress_empty_char = normalize_progress_char(
    cfg.get("progress_empty_char"), DEFAULT_PROGRESS_EMPTY_CHAR
)

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
root.title("OSC Chatbox")
root.geometry("560x620")
root.minsize(520, 560)
root.configure(bg=BG)
root.resizable(True, True)

title_bar = tk.Frame(root, bg=PANEL, pady=10)
title_bar.pack(fill="x")

header_title_label = tk.Label(
    title_bar,
    text="◈  OSC CHATBOX",
    bg=PANEL,
    fg=ACCENT2,
    font=(UI_FONT, 13, "bold")
)
header_title_label.pack(side="left", padx=16)

status_label = tk.Label(
    title_bar,
    text="Status: Stopped",
    bg=PANEL,
    fg=RED,
    font=(UI_FONT, 9)
)
status_label.pack(side="right", padx=16)

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


# ── Helpers ────────────────────────────────────────────────────────────────

def open_settings():
    set_win = tk.Toplevel(root)
    set_win.title("Settings")
    set_win.configure(bg=BG)
    set_win.resizable(False, False)

    root.update_idletasks()
    sw = root.winfo_width()
    sh = root.winfo_height()
    sx = root.winfo_x()
    sy = root.winfo_y()
    set_win.geometry(f"{sw}x{sh}+{sx}+{sy}")

    pages = [
        {
            "title": "Settings",
            "content_type": "scale",
        },
    ]

    current_page = [0]

    header = tk.Frame(set_win, bg=PANEL, pady=10)
    header.pack(fill="x")

    title_label = tk.Label(
        header,
        text="",
        bg=PANEL,
        fg=ACCENT2,
        font=(UI_FONT, 12, "bold")
    )
    title_label.pack(side="left", padx=16)

    page_indicator = tk.Label(
        header,
        text="",
        bg=PANEL,
        fg=SUBTEXT,
        font=(UI_FONT, 8)
    )
    page_indicator.pack(side="right", padx=16)

    tk.Frame(set_win, bg=BORDER, height=1).pack(fill="x")

    content_frame = tk.Frame(set_win, bg=PANEL, highlightthickness=1, highlightbackground=BORDER)
    content_frame.pack(padx=20, pady=(14, 0), fill="both", expand=True)

    def confirm_reset():
        if messagebox.askyesno("Reset", "Reset all settings to defaults?"):
            reset_to_defaults()

    def build_scale_page():
        for w in content_frame.winfo_children():
            w.destroy()

        tk.Label(content_frame, text="UI Scale",
                 bg=PANEL, fg=ACCENT2, font=(UI_FONT, 10, "bold")).pack(pady=(20, 8))

        scale_var = tk.DoubleVar(value=ui_scale)

        slider = tk.Scale(
            content_frame,
            from_=0.7, to=2.0,
            resolution=0.05,
            orient="horizontal",
            variable=scale_var,
            bg=PANEL, fg=TEXT,
            troughcolor=BORDER,
            activebackground=ACCENT2,
            highlightthickness=0,
            sliderrelief="flat",
            length=300,
            command=lambda v: apply_scale(float(v)),
        )
        slider.pack(pady=4)

        pct_label = tk.Label(content_frame, text="", bg=PANEL, fg=SUBTEXT,
                             font=(UI_FONT, 9))
        pct_label.pack()

        tk.Label(content_frame, text="Config",
                 bg=PANEL, fg=ACCENT2, font=(UI_FONT, 10, "bold")).pack(pady=(20, 8))

        tk.Button(
            content_frame,
            text="Reset to Defaults",
            bg=BTN_BG,
            fg=SUBTEXT,
            relief="flat",
            activebackground=BORDER,
            activeforeground=TEXT,
            cursor="hand2",
            font=(UI_FONT, 9, "bold"),
            command=confirm_reset
        ).pack(pady=(20, 5))

        tk.Label(content_frame, text="Progress Bar",
                 bg=PANEL, fg=ACCENT2, font=(UI_FONT, 10, "bold")).pack(pady=(20, 8))

        chars_frame = tk.Frame(content_frame, bg=PANEL)
        chars_frame.pack(pady=(0, 6))

        tk.Label(chars_frame, text="Filled", bg=PANEL, fg=SUBTEXT, font=(UI_FONT, 8)).grid(row=0, column=0, padx=4)
        tk.Label(chars_frame, text="Border", bg=PANEL, fg=SUBTEXT, font=(UI_FONT, 8)).grid(row=0, column=1, padx=4)
        tk.Label(chars_frame, text="Empty",  bg=PANEL, fg=SUBTEXT, font=(UI_FONT, 8)).grid(row=0, column=2, padx=4)

        filled_char_entry = tk.Entry(
            chars_frame, width=4, justify="center",
            bg=ENTRY_BG, fg=TEXT, insertbackground=ACCENT, relief="flat",
            font=(UI_FONT, 9), highlightthickness=1,
            highlightbackground=BORDER, highlightcolor=ACCENT
        )
        border_char_entry = tk.Entry(
            chars_frame, width=4, justify="center",
            bg=ENTRY_BG, fg=TEXT, insertbackground=ACCENT, relief="flat",
            font=(UI_FONT, 9), highlightthickness=1,
            highlightbackground=BORDER, highlightcolor=ACCENT
        )
        empty_char_entry = tk.Entry(
            chars_frame, width=4, justify="center",
            bg=ENTRY_BG, fg=TEXT, insertbackground=ACCENT, relief="flat",
            font=(UI_FONT, 9), highlightthickness=1,
            highlightbackground=BORDER, highlightcolor=ACCENT
        )

        filled_char_entry.grid(row=1, column=0, padx=4, pady=(2, 0))
        border_char_entry.grid(row=1, column=1, padx=4, pady=(2, 0))
        empty_char_entry.grid(row=1, column=2, padx=4, pady=(2, 0))

        filled_char_entry.insert(0, progress_filled_char)
        border_char_entry.insert(0, progress_border_char)
        empty_char_entry.insert(0, progress_empty_char)

        preview_frame = tk.Frame(content_frame, bg=PANEL)
        preview_frame.pack(pady=(2, 8))

        tk.Label(preview_frame, text="Preview", bg=PANEL, fg=SUBTEXT, font=(UI_FONT, 8)).grid(
            row=0, column=0, columnspan=3, pady=(0, 3)
        )

        filled_preview = tk.Label(
            preview_frame,
            text=progress_filled_char * 6,
            bg=BORDER, fg=TEXT,
            font=(UI_FONT, 10), width=8, padx=4, pady=2,
        )
        filled_preview.grid(row=1, column=0, padx=4)

        border_preview = tk.Label(
            preview_frame,
            text=progress_border_char * 6,
            bg=BORDER, fg=TEXT,
            font=(UI_FONT, 10), width=8, padx=4, pady=2,
        )
        border_preview.grid(row=1, column=1, padx=4)

        empty_preview = tk.Label(
            preview_frame,
            text=progress_empty_char * 6,
            bg=BORDER, fg=ACCENT2,
            font=(UI_FONT, 10), width=8, padx=4, pady=2,
        )
        empty_preview.grid(row=1, column=2, padx=4)

        def set_entry_char(entry, value):
            if entry.get() != value:
                entry.delete(0, tk.END)
                entry.insert(0, value)

        def refresh_progress_previews():
            filled_preview.config(text=progress_filled_char * 6)
            border_preview.config(text=progress_border_char * 6)
            empty_preview.config(text=progress_empty_char * 6)

        def apply_progress_char_settings(_event=None):
            global progress_filled_char, progress_border_char, progress_empty_char

            progress_filled_char = normalize_progress_char(
                filled_char_entry.get(), DEFAULT_PROGRESS_FILLED_CHAR
            )
            progress_border_char = normalize_progress_char(
                border_char_entry.get(), DEFAULT_PROGRESS_BORDER_CHAR
            )
            progress_empty_char = normalize_progress_char(
                empty_char_entry.get(), DEFAULT_PROGRESS_EMPTY_CHAR
            )

            set_entry_char(filled_char_entry, progress_filled_char)
            set_entry_char(border_char_entry, progress_border_char)
            set_entry_char(empty_char_entry, progress_empty_char)
            refresh_progress_previews()
            save_config()

        for progress_entry in (filled_char_entry, border_char_entry, empty_char_entry):
            progress_entry.bind("<KeyRelease>", apply_progress_char_settings)
            progress_entry.bind("<FocusOut>",   apply_progress_char_settings)

        def update_pct(*_):
            pct_label.config(text=f"{int(scale_var.get() * 100)}%")
            set_win.update_idletasks()
            set_win.geometry(f"{root.winfo_width()}x{root.winfo_height()}"
                             f"+{root.winfo_x()}+{root.winfo_y()}")

        scale_var.trace_add("write", update_pct)
        update_pct()

    def show_page(idx):
        p = pages[idx]
        title_label.config(text=p["title"])
        page_indicator.config(text=f"Page {idx + 1} of {len(pages)}")
        prev_btn.config(state="normal" if idx > 0 else "disabled")
        is_last = idx == len(pages) - 1
        next_btn.config(text="Finish" if is_last else "Next →")
        if p["content_type"] == "scale":
            build_scale_page()

    nav_frame = tk.Frame(set_win, bg=BG)
    nav_frame.pack(fill="x", padx=20, pady=(0, 14))
    nav_frame.columnconfigure(1, weight=1)

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

    def next_or_finish():
        if current_page[0] < len(pages) - 1:
            current_page[0] += 1
            show_page(current_page[0])
        else:
            set_win.destroy()

    next_btn = tk.Button(
        nav_frame, text="Next →", bg=BTN_BG, fg=BTN_FG, relief="flat", width=10,
        command=next_or_finish
    )
    next_btn.grid(row=0, column=2, sticky="e")
    next_btn.configure(
        bg=ACCENT, fg="#FFFFFF", activebackground=ACCENT2, activeforeground="#FFFFFF",
        cursor="hand2", font=(UI_FONT, 9, "bold"),
    )

    show_page(0)


def open_help():
    help_win = tk.Toplevel(root)
    help_win.title("OSC Chatbox Tutorial")
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
            "title": "Page Toggles",
            "content": (
                "Page Toggles — The four circle buttons below\n"
                "the Start/Stop/Restart buttons control which\n"
                "pages appear in the rotation.\n\n"
                "● Filled circle  = page is ENABLED\n"
                "○ Hollow circle = page is DISABLED\n\n"
                "Each toggle is labelled with the page name\n"
                "above the circle and P1–P4 below it.\n\n"
                "Click a circle to toggle it on or off.\n"
                "Disabled pages are skipped entirely.\n\n"
                "If ALL pages are disabled, the chatbox will\n"
                "display 'No pages enabled' until at least\n"
                "one page is turned back on."
            )
        },
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
                "1. Get LHM at https://github.com/LibreHardwareMonitor/LibreHardwareMonitor/releases\n"
                "2. Open LibreHardwareMonitor.exe\n"
                "3. Go to Options → Web Server\n"
                "4. Click 'Run' and make sure port is 8085\n"
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

    header = tk.Frame(help_win, bg=PANEL, pady=10)
    header.pack(fill="x")

    title_label = tk.Label(
        header, text="", bg=PANEL, fg=ACCENT2, font=(UI_FONT, 12, "bold")
    )
    title_label.pack(side="left", padx=16)

    page_indicator = tk.Label(
        header, text="", bg=PANEL, fg=SUBTEXT, font=(UI_FONT, 8)
    )
    page_indicator.pack(side="right", padx=16)

    tk.Frame(help_win, bg=BORDER, height=1).pack(fill="x")

    content_panel = tk.Frame(help_win, bg=PANEL, highlightthickness=1, highlightbackground=BORDER)
    content_panel.pack(padx=20, pady=(14, 0), fill="both", expand=True)

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

    def next_or_finish():
        if current_page[0] < len(pages) - 1:
            current_page[0] += 1
            show_page(current_page[0])
        else:
            help_win.destroy()

    next_btn = tk.Button(
        nav_frame, text="Next →", bg=BTN_BG, fg=BTN_FG, relief="flat", width=10,
        command=next_or_finish
    )
    next_btn.grid(row=0, column=2, sticky="e")
    next_btn.configure(
        bg=ACCENT, fg="#FFFFFF", activebackground=ACCENT2, activeforeground="#FFFFFF",
        cursor="hand2", font=(UI_FONT, 9, "bold"),
    )

    show_page(0)


frame.columnconfigure(1, weight=1)


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

# ── Data Config ────────────────────────────────────────────────────────────
tk.Label(
    frame,
    text="Data Config",
    bg=BG,
    fg=ACCENT2,
    font=(UI_FONT, 11, "bold"),
).grid(row=0, column=0, columnspan=2, pady=(2, 6))

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

# ── Page Text ──────────────────────────────────────────────────────────────
tk.Label(
    frame,
    text="Page Text",
    bg=BG,
    fg=ACCENT2,
    font=(UI_FONT, 11, "bold"),
).grid(row=7, column=0, columnspan=2, pady=(8, 6))

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

# ── Toggles ────────────────────────────────────────────────────────────────
toggle_outer = tk.Frame(frame, bg=BG)
toggle_outer.grid(row=13, column=0, columnspan=2, pady=(6, 2), sticky="ew")

toggle_inner = tk.Frame(toggle_outer, bg=BG)
toggle_inner.pack(anchor="center")

PAGE_NAMES = ["Network", "CPU/GPU", "RAM", "Weather"]
PAGE_NUMBERS = ["P1", "P2", "P3", "P4"]
PAGE_ENABLED_KEYS = ["page1_enabled", "page2_enabled", "page3_enabled", "page4_enabled"]

for col, (name, num, key) in enumerate(zip(PAGE_NAMES, PAGE_NUMBERS, PAGE_ENABLED_KEYS)):
    cell = tk.Frame(toggle_inner, bg=BG)
    cell.grid(row=0, column=col, padx=18)

    name_lbl = tk.Label(cell, text=name, bg=BG, fg=SUBTEXT, font=(UI_FONT, 8))
    name_lbl.pack()
    scalable_widgets.append((name_lbl, 8, ()))

    tog = CircleToggle(cell, enabled=bool(cfg.get(key, True)))
    tog.pack()

    num_lbl = tk.Label(cell, text=num, bg=BG, fg=ACCENT2, font=(UI_FONT, 8, "bold"))
    num_lbl.pack()
    scalable_widgets.append((num_lbl, 8, ()))

    page_toggles.append(tog)

# ── Main Buttons ───────────────────────────────────────────────────────────
button_frame = tk.Frame(frame, bg=BG)
button_frame.grid(row=14, column=0, columnspan=2, pady=15, sticky="ew")

button_frame.columnconfigure(0, weight=1)
button_frame.columnconfigure(1, weight=1)
button_frame.columnconfigure(2, weight=1)

tk.Button(
    button_frame,
    text="Start",
    command=start_script,
    bg=ACCENT,
    fg="#FFFFFF",
    relief="flat",
    activebackground=ACCENT2,
    activeforeground="#FFFFFF",
    cursor="hand2",
    font=(UI_FONT, 9, "bold"),
).grid(row=0, column=0, sticky="ew", padx=2)

tk.Button(
    button_frame,
    text="Stop",
    command=stop_script,
    bg=BTN_BG,
    fg=SUBTEXT,
    relief="flat",
    activebackground=BORDER,
    activeforeground=TEXT,
    cursor="hand2",
    font=(UI_FONT, 9, "bold"),
).grid(row=0, column=1, sticky="ew", padx=2)

tk.Button(
    button_frame,
    text="Restart",
    command=restart_script,
    bg=BTN_BG,
    fg=SUBTEXT,
    relief="flat",
    activebackground=BORDER,
    activeforeground=TEXT,
    cursor="hand2",
    font=(UI_FONT, 9, "bold"),
).grid(row=0, column=2, sticky="ew", padx=2)

# ── Bottom Bar: Help | Settings ───────────────────────────────────────────
bottom_bar = tk.Frame(frame, bg=BG)
bottom_bar.grid(row=15, column=0, columnspan=2, pady=6, sticky="ew")

bottom_bar.columnconfigure(0, weight=1)
bottom_bar.columnconfigure(1, weight=1)

help_btn = square_button(bottom_bar, "？", open_help, base_size=32)
help_btn.grid(row=0, column=0, sticky="w", padx=6)

settings_btn = square_button(bottom_bar, "⚙", open_settings, base_size=32)
settings_btn.grid(row=0, column=1, sticky="e", padx=6)

root.mainloop()