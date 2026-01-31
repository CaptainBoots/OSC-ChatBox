# ═════════════════════════════════════════════════════════════════════════════════════════════════════════════════════#
#                                    OSC Python Script - LibreHardwareMonitor REST API Edition                          #
# ═════════════════════════════════════════════════════════════════════════════════════════════════════════════════════#
# Uses REST API instead of WMI - NO ADMIN REQUIRED
# LibreHardwareMonitor must be running with web server enabled

import asyncio
import re
import subprocess
import threading
import time
import tkinter as tk
from tkinter import messagebox
import psutil
import winrt.windows.media.control as wmc
from pythonosc.udp_client import SimpleUDPClient
from enum import Enum
import requests
import json


class CPUManufacturer(Enum):
    INTEL = "Intel"
    AMD = "AMD"
    UNKNOWN = "Unknown"


OSC_IP = "127.0.0.1"
OSC_PORT = 9000
INTERFACE = "Ethernet"
SWITCH_INTERVAL = 30
LHM_REST_API = "http://localhost:8888/data.json"  # Default LHM web server port

print("OSC Script - LibreHardwareMonitor REST API Edition")
print("=" * 70)

client = None
running = False
page1_line1_text = "-enter text-"
page2_line1_text = "-enter text-"

cpu_wattage = "-error-"
cpu_temp = "-error-"
gpu_wattage = "-error-"
gpu_temp = "-error-"

cpu_manufacturer = CPUManufacturer.UNKNOWN


# ═════════════════════════════════════════════════════════════════════════════════════════════════════════════════════#
# REST API DATA RETRIEVAL
# ═════════════════════════════════════════════════════════════════════════════════════════════════════════════════════#

def get_lhm_data():
    """
    Get sensor data from LibreHardwareMonitor REST API.
    No admin required - just need LHM running with web server enabled.
    """
    try:
        response = requests.get(LHM_REST_API, timeout=5)
        if response.status_code == 200:
            return response.json()
    except (requests.ConnectionError, requests.Timeout, json.JSONDecodeError) as e:
        print(f"[WARNING] REST API query failed: {e}")
    return None


def parse_lhm_data(data):
    """Parse LibreHardwareMonitor JSON data and extract sensor values"""
    cpu_temp_val = 0
    cpu_power_val = 0
    gpu_temp_val = 0
    gpu_power_val = 0

    if not data or "Children" not in data:
        return cpu_temp_val, cpu_power_val, gpu_temp_val, gpu_power_val

    try:
        for component in data.get("Children", []):
            component_name = component.get("Text", "").lower()

            # Check if this is CPU
            if "cpu" in component_name:
                for child in component.get("Children", []):
                    child_name = child.get("Text", "").lower()

                    # Look for CPU temperature
                    if "core" in child_name or "package" in child_name:
                        for sensor in child.get("Children", []):
                            sensor_name = sensor.get("Text", "").lower()
                            if "temperature" in sensor_name:
                                try:
                                    cpu_temp_val = int(float(sensor.get("Value", 0)))
                                except (ValueError, TypeError):
                                    pass

                    # Look for CPU power
                    if "power" in child_name or "package" in child_name:
                        try:
                            cpu_power_val = int(float(child.get("Value", 0)))
                        except (ValueError, TypeError):
                            pass

            # Check if this is GPU
            if "gpu" in component_name or "amd" in component_name or "nvidia" in component_name:
                for child in component.get("Children", []):
                    child_name = child.get("Text", "").lower()

                    # Look for GPU temperature
                    if "temperature" in child_name or "temp" in child_name:
                        try:
                            gpu_temp_val = int(float(child.get("Value", 0)))
                        except (ValueError, TypeError):
                            pass

                    # Look for GPU power
                    if "power" in child_name:
                        try:
                            gpu_power_val = int(float(child.get("Value", 0)))
                        except (ValueError, TypeError):
                            pass

    except (KeyError, TypeError, AttributeError):
        pass

    return cpu_temp_val, cpu_power_val, gpu_temp_val, gpu_power_val


# ═════════════════════════════════════════════════════════════════════════════════════════════════════════════════════#
# DIAGNOSTICS
# ═════════════════════════════════════════════════════════════════════════════════════════════════════════════════════#

def diagnose_lhm():
    """Run LibreHardwareMonitor connection diagnostics"""
    print("\n[DIAGNOSTIC] Testing LibreHardwareMonitor REST API...")
    print("[DIAGNOSTIC] " + "=" * 66)

    try:
        response = requests.get(LHM_REST_API, timeout=5)
        if response.status_code == 200:
            print("[DIAGNOSTIC] ✓ REST API Connection: SUCCESS")
            data = response.json()
            sensor_count = len(data.get("Children", []))
            print(f"[DIAGNOSTIC] ✓ Sensors Found: {sensor_count} components")
            print("[DIAGNOSTIC] ✓ No admin privileges required!")
            print("[DIAGNOSTIC] " + "=" * 66)
            return True
        else:
            print(f"[DIAGNOSTIC] ✗ REST API returned status: {response.status_code}")
    except requests.ConnectionError:
        print("[DIAGNOSTIC] ✗ Cannot connect to LibreHardwareMonitor REST API")
        print("[DIAGNOSTIC] FIX 1: Make sure LibreHardwareMonitor.exe is RUNNING")
        print("[DIAGNOSTIC] FIX 2: Enable web server in LHM (Options → Web server)")
        print("[DIAGNOSTIC] FIX 3: Check port is 8888 (default)")
    except requests.Timeout:
        print("[DIAGNOSTIC] ✗ REST API query timed out")
    except Exception as e:
        print(f"[DIAGNOSTIC] ✗ Error: {e}")

    print("[DIAGNOSTIC] " + "=" * 66)
    return False


# ═════════════════════════════════════════════════════════════════════════════════════════════════════════════════════#
# HARDWARE MONITORING
# ═════════════════════════════════════════════════════════════════════════════════════════════════════════════════════#

def _clean_name(name: str):
    """Clean hardware names"""
    name = re.sub(r"\(.*?\)|\[.*?]|\{.*?}", "", name)
    name = name.split("@")[0]
    name = re.sub(r"\s+", " ", name).strip()
    return name


def detect_cpu():
    """Detect CPU name and manufacturer"""
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
    """Detect GPU name"""
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


def get_gpu_load():
    """Get GPU utilization percentage"""
    try:
        cmd = (
            'Get-Counter "\\GPU Engine(*3D*)\\Utilization Percentage" | '
            'Select-Object -ExpandProperty CounterSamples | '
            'Select-Object -ExpandProperty CookedValue'
        )
        result = subprocess.check_output(
            ["powershell", "-Command", cmd], encoding='utf-8', stderr=subprocess.DEVNULL, timeout=5
        )
        values = [float(v) for v in result.strip().split('\n') if v.strip()]
        return int(max(values)) if values else 0
    except (subprocess.CalledProcessError, ValueError, IndexError, subprocess.TimeoutExpired):
        return 0


# ═════════════════════════════════════════════════════════════════════════════════════════════════════════════════════#
# NETWORK MONITORING
# ═════════════════════════════════════════════════════════════════════════════════════════════════════════════════════#

def fmt(bps):
    """Format bytes per second"""
    if bps > 1024 * 1024:
        return f"{bps / (1024 * 1024):.2f} MB/s"
    return f"{bps / 1024:.1f} KB/s"


def get_network_usage(prev, prev_time):
    """Get network speeds"""
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
    """Get current media information"""
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
    """Clean media title"""
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
    """Create progress bar"""
    if duration_ms <= 0:
        return "─" * length
    percent = min(max(position_ms / duration_ms, 0), 1)
    filled_len = int(length * percent)
    return "■" * filled_len + "□" * (length - filled_len)


# ═════════════════════════════════════════════════════════════════════════════════════════════════════════════════════#
# MAIN OSC LOOP
# ═════════════════════════════════════════════════════════════════════════════════════════════════════════════════════#

def run_osc_loop():
    """Main OSC loop"""
    global running, cpu_wattage, cpu_temp, gpu_wattage, gpu_temp

    all_stats = psutil.net_io_counters(pernic=True)
    if INTERFACE not in all_stats:
        print(f"Error: {INTERFACE} not found. Available: {list(all_stats.keys())}")
        running = False
        return

    prev = all_stats[INTERFACE]
    prev_time = time.time()

    cpu_detect = detect_cpu()
    gpu_detect = detect_gpu()

    print(f"\n{'=' * 60}")
    print(f"CPU: {cpu_detect} ({cpu_manufacturer.value})")
    print(f"GPU: {gpu_detect}")
    print(f"Using: LibreHardwareMonitor REST API")
    print(f"{'=' * 60}")
    print("Sending live data to OSC port..\n")

    query_cooldown = 0

    while running:
        try:
            song, artist, pos, dur = asyncio.run(get_media_info())
            clean_song = clean_title(song)

            cpu = psutil.cpu_percent()
            gpu = get_gpu_load()

            # Query sensors every 3 iterations
            query_cooldown += 1
            if query_cooldown >= 3:
                lhm_data = get_lhm_data()
                if lhm_data:
                    cpu_temp, cpu_wattage, gpu_temp, gpu_wattage = parse_lhm_data(lhm_data)
                query_cooldown = 0

            prev, up_raw, down_raw, prev_time = get_network_usage(prev, prev_time)

            cur_time_str = time.strftime("%I:%M %p")
            progress_bar = create_progress_bar(pos, dur)

            display_artist = f"-{artist}" if artist else ""
            display_song = f"🎵 {clean_song}" if clean_song else ""

            page_index = int((time.time() // SWITCH_INTERVAL) % 2)

            if page_index == 0:
                text = (
                    f"{page1_line1_text}\n"
                    f"{cur_time_str}\n"
                    f"Download {fmt(down_raw)}\n"
                    f"Upload {fmt(up_raw)}\n"
                    f"{progress_bar}\n"
                    f"{display_song} {display_artist}"
                )
            else:
                text = (
                    f"{page2_line1_text}\n"
                    f"{cur_time_str}\n"
                    f"{cpu_detect} {cpu}%\n"
                    f"{cpu_wattage}w {cpu_temp}℃\n"
                    f"{gpu_detect} {gpu}%\n"
                    f"{gpu_wattage}w {gpu_temp}℃\n"
                )

            client.send_message("/chatbox/input", [text, True])
            time.sleep(5.0)

        except Exception as e:
            print(f"Error in OSC loop: {e}")
            time.sleep(1)


# ═════════════════════════════════════════════════════════════════════════════════════════════════════════════════════#
# SCRIPT CONTROL
# ═════════════════════════════════════════════════════════════════════════════════════════════════════════════════════#

def start_script():
    """Start the script"""
    global running, client, OSC_IP, OSC_PORT, INTERFACE, SWITCH_INTERVAL, LHM_REST_API
    global page1_line1_text, page2_line1_text

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

        client = SimpleUDPClient(OSC_IP, OSC_PORT)

        running = True
        status_label.config(text="Status: Running", fg="#4CFF4C")
        info_label.config(text="LibreHardwareMonitor REST API", fg="#FFB84D")

        print("\n[INFO] Starting diagnostics...")
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
    """Stop the script"""
    global running
    running = False
    status_label.config(text="Status: Stopped", fg="#FF4C4C")
    info_label.config(text="")


# ═════════════════════════════════════════════════════════════════════════════════════════════════════════════════════#
# GUI
# ═════════════════════════════════════════════════════════════════════════════════════════════════════════════════════#

BG = "#121212"
FG = "#E0E0E0"
ENTRY_BG = "#1E1E1E"
BTN_BG = "#2A2A2A"
BTN_FG = "#FFFFFF"

root = tk.Tk()
root.title("OSC Chatbox - LibreHardwareMonitor REST API")
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

frame.columnconfigure(1, weight=1)

dark_label("OSC IP", 0)
ip_entry = dark_entry(0, OSC_IP)

dark_label("OSC Port", 1)
port_entry = dark_entry(1, str(OSC_PORT))

dark_label("Network Interface", 2)
iface_entry = dark_entry(2, INTERFACE)

dark_label("Switch Interval", 3)
interval_entry = dark_entry(3, str(SWITCH_INTERVAL))

dark_label("LHM REST API", 4)
lhm_entry = dark_entry(4, LHM_REST_API)

dark_label("Page 1 Text", 5)
page1_entry = dark_entry(5, "Thx for using boot's osc code")

dark_label("Page 2 Text", 6)
page2_entry = dark_entry(6, "hi put your text here :3")

start_btn = tk.Button(frame, text="Start", command=start_script,
                      bg=BTN_BG, fg=BTN_FG, relief="flat")
start_btn.grid(row=7, column=0, pady=15, sticky="ew")

stop_btn = tk.Button(frame, text="Stop", command=stop_script,
                     bg=BTN_BG, fg=BTN_FG, relief="flat")
stop_btn.grid(row=7, column=1, pady=15, sticky="ew")

status_label = tk.Label(frame, text="Status: Stopped", bg=BG, fg="#FF4C4C")
status_label.grid(row=8, column=0, columnspan=2)

info_label = tk.Label(frame, text="", bg=BG, fg="#FFB84D", wraplength=400)
info_label.grid(row=9, column=0, columnspan=2, pady=5)

help_text = tk.Label(
    frame,
    text="Note: Enable web server in LibreHardwareMonitor\n(Options → Web server)",
    bg=BG, fg="#888888", font=("Arial", 8)
)
help_text.grid(row=10, column=0, columnspan=2, pady=5)

root.mainloop()