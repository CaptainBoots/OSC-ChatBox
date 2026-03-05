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

cpu_wattage = "error"
cpu_temp = "error"
gpu_wattage = "error"
gpu_temp = "error"
dram_load = "error"
vram_load = "error"

cpu_manufacturer = CPUManufacturer.UNKNOWN

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
        return _clean_name(gpu_name)
    except (subprocess.CalledProcessError, UnicodeDecodeError, subprocess.TimeoutExpired):
        return "GPU Unknown"


def get_lhm_data():
    try:
        response = requests.get(LHM_REST_API, timeout=5)
        if response.status_code == 200:
            return response.json()
    except (requests.ConnectionError, requests.Timeout, json.JSONDecodeError):
        pass
    return None


def parse_lhm_data(data):
    cpu_temp_val = 0
    cpu_power_val = 0
    gpu_temp_val = 0
    gpu_power_val = 0

    if not data or "Children" not in data:
        return cpu_temp_val, cpu_power_val, gpu_temp_val, gpu_power_val

    try:
        for top_level in data.get("Children", []):
            for hardware in top_level.get("Children", []):
                hardware_text = hardware.get("Text", "").lower()

                if "intel" in hardware_text:
                    for category in hardware.get("Children", []):
                        category_text = category.get("Text", "").lower()

                        if "temperature" in category_text:
                            for sensor in category.get("Children", []):
                                sensor_text = sensor.get("Text", "").lower()
                                if "cpu package" in sensor_text:
                                    try:
                                        sensor_value = sensor.get("Value", 0)
                                        numeric_str = re.sub(r'[^\d.-]', '', str(sensor_value))
                                        cpu_temp_val = int(float(numeric_str))
                                    except (ValueError, TypeError):
                                        pass

                        if "power" in category_text:
                            for sensor in category.get("Children", []):
                                sensor_text = sensor.get("Text", "").lower()
                                if "cpu package" in sensor_text:
                                    try:
                                        sensor_value = sensor.get("Value", 0)
                                        numeric_str = re.sub(r'[^\d.-]', '', str(sensor_value))
                                        cpu_power_val = int(float(numeric_str))
                                    except (ValueError, TypeError):
                                        pass

                elif "amd radeon" in hardware_text:
                    for category in hardware.get("Children", []):
                        category_text = category.get("Text", "").lower()

                        if "temperature" in category_text:
                            for sensor in category.get("Children", []):
                                sensor_text = sensor.get("Text", "").lower()
                                if "gpu core" in sensor_text and "distance" not in sensor_text:
                                    try:
                                        sensor_value = sensor.get("Value", 0)
                                        numeric_str = re.sub(r'[^\d.-]', '', str(sensor_value))
                                        gpu_temp_val = int(float(numeric_str))
                                    except (ValueError, TypeError):
                                        pass

                        if "power" in category_text:
                            for sensor in category.get("Children", []):
                                sensor_text = sensor.get("Text", "").lower()
                                if "gpu package" in sensor_text:
                                    try:
                                        sensor_value = sensor.get("Value", 0)
                                        numeric_str = re.sub(r'[^\d.-]', '', str(sensor_value))
                                        gpu_power_val = int(float(numeric_str))
                                    except (ValueError, TypeError):
                                        pass

    except (KeyError, TypeError, AttributeError, ValueError):
        pass

    return cpu_temp_val, cpu_power_val, gpu_temp_val, gpu_power_val


def get_gpu_load_from_lhm(data):
    if not data or "Children" not in data:
        return 0

    try:
        for top_level in data.get("Children", []):
            for hardware in top_level.get("Children", []):
                hardware_text = hardware.get("Text", "").lower()

                if "amd radeon" in hardware_text:
                    for category in hardware.get("Children", []):
                        category_text = category.get("Text", "").lower()

                        if "load" in category_text:
                            for sensor in category.get("Children", []):
                                sensor_text = sensor.get("Text", "").lower()
                                if "gpu core" in sensor_text:
                                    try:
                                        sensor_value = sensor.get("Value", 0)
                                        numeric_str = re.sub(r'[^\d.-]', '', str(sensor_value))
                                        return int(float(numeric_str))
                                    except (ValueError, TypeError):
                                        pass
    except (KeyError, TypeError, AttributeError, ValueError):
        pass

    return 0


def get_cpu_load_from_lhm(data):
    if not data or "Children" not in data:
        return 0

    try:
        for top_level in data.get("Children", []):
            for hardware in top_level.get("Children", []):
                hardware_text = hardware.get("Text", "").lower()

                if "intel" in hardware_text:
                    for category in hardware.get("Children", []):
                        category_text = category.get("Text", "").lower()

                        if "load" in category_text:
                            for sensor in category.get("Children", []):
                                sensor_text = sensor.get("Text", "").lower()
                                if "cpu total" in sensor_text:
                                    try:
                                        sensor_value = sensor.get("Value", 0)
                                        numeric_str = re.sub(r'[^\d.-]', '', str(sensor_value))
                                        return int(float(numeric_str))
                                    except (ValueError, TypeError):
                                        pass

                elif "amd" in hardware_text and "radeon" not in hardware_text:
                    for category in hardware.get("Children", []):
                        category_text = category.get("Text", "").lower()

                        if "load" in category_text:
                            for sensor in category.get("Children", []):
                                sensor_text = sensor.get("Text", "").lower()
                                if "cpu total" in sensor_text:
                                    try:
                                        sensor_value = sensor.get("Value", 0)
                                        numeric_str = re.sub(r'[^\d.-]', '', str(sensor_value))
                                        return int(float(numeric_str))
                                    except (ValueError, TypeError):
                                        pass
    except (KeyError, TypeError, AttributeError, ValueError):
        pass

    return 0


def get_dram_used_from_lhm(data):
    """Return physical RAM in use as a float GB, e.g. 13.5"""
    if not data or "Children" not in data:
        return 0.0

    try:
        for top_level in data.get("Children", []):
            for hardware in top_level.get("Children", []):
                hardware_text = hardware.get("Text", "").lower()

                if hardware_text == "total memory":
                    for category in hardware.get("Children", []):
                        if "data" in category.get("Text", "").lower():
                            for sensor in category.get("Children", []):
                                sensor_text = sensor.get("Text", "").lower()
                                if sensor_text == "memory used":
                                    try:
                                        sensor_value = sensor.get("Value", 0)
                                        numeric_str = re.sub(r'[^\d.-]', '', str(sensor_value))
                                        return round(float(numeric_str), 1)
                                    except (ValueError, TypeError):
                                        pass
    except (KeyError, TypeError, AttributeError, ValueError):
        pass

    return 0.0


def get_vram_used_from_lhm(data):
    """Return VRAM in use as a float GB at 1dp, e.g. 2.4"""
    if not data or "Children" not in data:
        return 0.0

    try:
        for top_level in data.get("Children", []):
            for hardware in top_level.get("Children", []):
                hardware_text = hardware.get("Text", "").lower()

                if "amd radeon" in hardware_text:
                    for category in hardware.get("Children", []):
                        for sensor in category.get("Children", []):
                            sensor_text = sensor.get("Text", "").lower()
                            if "gpu memory used" in sensor_text:
                                try:
                                    sensor_value = sensor.get("Value", 0)
                                    numeric_str = re.sub(r'[^\d.-]', '', str(sensor_value))
                                    # LHM reports GPU memory in MB — convert to GB
                                    return round(float(numeric_str) / 1024, 1)
                                except (ValueError, TypeError):
                                    pass
    except (KeyError, TypeError, AttributeError, ValueError):
        pass

    return 0.0


def _parse_gb(sensor_value) -> float:
    """Extract a float GB value from an LHM sensor value string."""
    numeric_str = re.sub(r'[^\d.-]', '', str(sensor_value))
    return float(numeric_str)


def _fmt_gb(gb: float) -> str:
    """Round a raw GB float to the nearest standard capacity, e.g. 31.9 → '32GB'."""
    rounded = round(gb)
    for nice in [2, 4, 6, 8, 10, 12, 16, 20, 24, 32, 48, 64, 96, 128]:
        if abs(rounded - nice) <= max(1, int(nice * 0.10)):
            return f"{nice}GB"
    return f"{rounded}GB"


# ─────────────────────────────────────────────────────────────────────────────
# DRAM total  (Generic Memory → Data → Memory Used + Memory Available)
# LHM has no dedicated "total" sensor so we add used + available.
# ─────────────────────────────────────────────────────────────────────────────
def get_dram_total_from_lhm(data) -> str:
    """Return total physical RAM as a string like '32GB RAM' using psutil."""
    try:
        total_bytes = psutil.virtual_memory().total
        total_gb = total_bytes / (1024 ** 3)
        return f"{_fmt_gb(total_gb)} DRAM"
    except Exception:
        return "DRAM"


# ─────────────────────────────────────────────────────────────────────────────
# VRAM total  (AMD Radeon → Data → GPU Memory Total, or Used + Free)
# ─────────────────────────────────────────────────────────────────────────────
def get_vram_total_from_lhm(data) -> str:
    """Return a string like '8GB VRAM' sourced from LHM. Falls back to 'VRAM'."""
    if not data or "Children" not in data:
        return "VRAM"

    try:
        for top_level in data.get("Children", []):
            for hardware in top_level.get("Children", []):
                hardware_text = hardware.get("Text", "").lower()

                if "amd radeon" in hardware_text:
                    total_gb = None
                    used_gb = None
                    free_gb = None

                    for category in hardware.get("Children", []):
                        for sensor in category.get("Children", []):
                            sensor_text = sensor.get("Text", "").lower()
                            try:
                                # LHM reports GPU memory in MB — divide by 1024 to get GB
                                if "gpu memory total" in sensor_text:
                                    total_gb = _parse_gb(sensor.get("Value", 0)) / 1024
                                elif "gpu memory used" in sensor_text:
                                    used_gb = _parse_gb(sensor.get("Value", 0)) / 1024
                                elif "gpu memory free" in sensor_text:
                                    free_gb = _parse_gb(sensor.get("Value", 0)) / 1024
                            except (ValueError, TypeError):
                                pass

                    if total_gb is not None and total_gb > 0:
                        return f"{_fmt_gb(total_gb)} VRAM"
                    if used_gb is not None and free_gb is not None:
                        return f"{_fmt_gb(used_gb + free_gb)} VRAM"
    except (KeyError, TypeError, AttributeError, ValueError):
        pass

    return "VRAM"


def diagnose_lhm():
    try:
        response = requests.get(LHM_REST_API, timeout=5)
        if response.status_code == 200:
            print("[DIAGNOSTIC] ✓ API Connection: SUCCESS")
            data = response.json()
            sensor_count = len(data.get("Children", []))   # type: ignore
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
# MEDIA MONITORING
# ═════════════════════════════════════════════════════════════════════════════════════════════════════════════════════#

async def get_media_info():
    try:
        manager = await wmc.GlobalSystemMediaTransportControlsSessionManager.request_async()
        session = manager.get_current_session()
        if session:
            props = await session.try_get_media_properties_async()
            timeline = session.get_timeline_properties()
            pos = timeline.position.total_seconds() * 1000
            dur = timeline.end_time.total_seconds() * 1000
            return props.title, props.artist, pos, dur
    except (OSError, AttributeError, RuntimeError):
        pass
    return None, None, 0, 0


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
        return "─" * length
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

    # Query LHM once at startup to get stable total capacity labels.
    startup_lhm = get_lhm_data()
    dram_detect = get_dram_total_from_lhm(startup_lhm)
    vram_detect = get_vram_total_from_lhm(startup_lhm)

    print(f"\n{'=' * 60}")
    print(f"CPU:  {cpu_detect} ({cpu_manufacturer.value})")
    print(f"GPU:  {gpu_detect}")
    print(f"DRAM: {dram_detect}")
    print(f"VRAM: {vram_detect}")
    print(f"{'=' * 60}")

    query_cooldown = 0
    cpu_load = 0
    gpu_load = 0
    dram_load = 0
    vram_load = 0

    while running:
        try:
            song, artist, pos, dur = asyncio.run(get_media_info())
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
                    cpu_load  = 0
                    gpu_load  = 0
                    dram_load = 0.0
                    vram_load = 0.0
                query_cooldown = 0

            prev, up_raw, down_raw, prev_time = get_network_usage(prev, prev_time)

            cur_time_str = time.strftime("%I:%M %p")
            progress_bar = create_progress_bar(pos, dur)

            display_artist = f"- {artist}" if artist else ""
            display_song = f"🎵 {clean_song}" if clean_song else ""

            page_index = int((time.time() // SWITCH_INTERVAL) % 3)

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
            else:
                text = (
                    f"{page3_line1_text}\n"
                    f"{cur_time_str}\n"
                    f"{dram_detect} {dram_load}GB\n"
                    f"{vram_detect} {vram_load}GB\n"
                    f"{progress_bar}\n"
                    f"{display_song} {display_artist}"
                )

            print(text)

            if client is not None:
                client.send_message("/chatbox/input", [text, True])  # type: ignore
            else:
                print("Warning: OSC client not initialized")

            time.sleep(5.0)

        except Exception as e:
            print(f"Error in OSC loop: {e}")
            time.sleep(1)


# ═════════════════════════════════════════════════════════════════════════════════════════════════════════════════════#
# SCRIPT CONTROL
# ═════════════════════════════════════════════════════════════════════════════════════════════════════════════════════#

def start_script():
    global running, client, OSC_IP, OSC_PORT, INTERFACE, SWITCH_INTERVAL, LHM_REST_API
    global page1_line1_text, page2_line1_text, page3_line1_text

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

BG = "#121212"
FG = "#E0E0E0"
ENTRY_BG = "#1E1E1E"
BTN_BG = "#2A2A2A"
BTN_FG = "#FFFFFF"

root = tk.Tk()
root.title("OSC Chatbox")
root.geometry("450x350")
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


frame.columnconfigure(1, weight=1)

dark_label("OSC IP", 0)
ip_entry = dark_entry(0, OSC_IP)

dark_label("OSC Port", 1)
port_entry = dark_entry(1, str(OSC_PORT))

dark_label("Network Interface", 2)
iface_entry = dark_entry(2, INTERFACE)

dark_label("Switch Interval", 3)
interval_entry = dark_entry(3, str(SWITCH_INTERVAL))

dark_label("LHM Interface", 4)
lhm_entry = dark_entry(4, LHM_REST_API)

dark_label("Page 1 Text", 5)
page1_entry = dark_entry(5, "Thx for using Boots's osc code")

dark_label("Page 2 Text", 6)
page2_entry = dark_entry(6, "Join the discord server at https://discord.gg/XdfKAWu6Ph")

dark_label("Page 3 Text", 7)
page3_entry = dark_entry(7, "hi put your text here :3")

button_frame = tk.Frame(frame, bg=BG)
button_frame.grid(row=8, column=0, columnspan=2, pady=15, sticky="ew")
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
status_label.grid(row=9, column=0, columnspan=2)

root.mainloop()