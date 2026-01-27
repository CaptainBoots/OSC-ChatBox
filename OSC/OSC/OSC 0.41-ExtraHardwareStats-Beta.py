#═════════════════════════════════════════════════════════════════════════════════════════════════════════════════════#
#                                              OSC Python Script                                                      #
#═════════════════════════════════════════════════════════════════════════════════════════════════════════════════════#
#Hi :3
#Wellcome to my code

#═════════════════════════════════════════════════════════════════════════════════════════════════════════════════════#
# Imports
#═════════════════════════════════════════════════════════════════════════════════════════════════════════════════════#

import asyncio
import re
import subprocess
import threading
import time
import tkinter as tk
import psutil
import winrt.windows.media.control as wmc
from pythonosc.udp_client import SimpleUDPClient

#═════════════════════════════════════════════════════════════════════════════════════════════════════════════════════#
# CONFIGURATION & GLOBAL VARIABLES
#═════════════════════════════════════════════════════════════════════════════════════════════════════════════════════#

OSC_IP = "127.0.0.1"
OSC_PORT = 9000
INTERFACE = "Ethernet"
SWITCH_INTERVAL = 30

print("This a beta version use at your own risk")

client = None
running = False
page1_line1_text = "-enter text-"
page2_line1_text = "-enter text-"

# Add these new global variables
cpu_wattage = 0
cpu_temp = 0
gpu_wattage = 0
gpu_temp = 0

#═════════════════════════════════════════════════════════════════════════════════════════════════════════════════════#
# UTILITY FUNCTIONS
#═════════════════════════════════════════════════════════════════════════════════════════════════════════════════════#

def fmt(bps):
    if bps > 1024 * 1024:
        return f"{bps / (1024 * 1024):.2f} MB/s"
    return f"{bps / 1024:.1f} KB/s"


def _clean_name(name: str):
    name = re.sub(r"\(.*?\)|\[.*?]|\{.*?}", "", name)
    name = name.split("@")[0]
    name = re.sub(r"\s+", " ", name).strip()
    return name
#═════════════════════════════════════════════════════════════════════════════════════════════════════════════════════#
# HARDWARE MONITORING
#═════════════════════════════════════════════════════════════════════════════════════════════════════════════════════#

#════════════════════════════════════════════════════Cpu══════════════════════════════════════════════════════════════#

def detect_cpu():
    try:
        cpu_name = subprocess.check_output(
            ["powershell", "-Command", "(Get-CimInstance Win32_Processor).Name"],
            encoding="utf-8",
            stderr=subprocess.DEVNULL
        ).strip()
        return _clean_name(cpu_name)
    except (subprocess.CalledProcessError, UnicodeDecodeError):
        return "CPU Unknown"

def get_cpu_wattage():
    """Get CPU power consumption from OpenHardwareMonitor or WMI"""
    # Method 1: Try OpenHardwareMonitor (most reliable)
    try:
        cmd = (
            'Get-WmiObject -Namespace "root/OpenHardwareMonitor" -Class Sensor | '
            'Where-Object {$_.SensorType -eq "Power" -and $_.Name -like "*CPU*"} | '
            'Select-Object -ExpandProperty Value | Select-Object -First 1'
        )
        result = subprocess.check_output(["powershell", "-Command", cmd], encoding='utf-8', stderr=subprocess.DEVNULL)
        if result.strip():
            wattage = float(result.strip())
            return int(wattage)
    except (subprocess.CalledProcessError, ValueError):
        pass

    # Method 2: Try WMI Power consumption
    try:
        cmd = (
            'Get-WmiObject -Namespace "root/cimv2" -Class Win32_PerfFormattedData_Counters_ProcessorInformation | '
            'Select-Object -First 1 -ExpandProperty "C2 Transitions/sec"'
        )
        result = subprocess.check_output(["powershell", "-Command", cmd], encoding='utf-8', stderr=subprocess.DEVNULL)
        if result.strip():
            value = float(result.strip())
            # Rough conversion based on C-state transitions
            return int(max(30, min(150, value / 10)))
    except (subprocess.CalledProcessError, ValueError):
        pass

    # Fallback: estimation based on load
    try:
        cpu_percent = psutil.cpu_percent(interval=0.05)
        estimated_watts = int((cpu_percent / 100) * 150)
        return estimated_watts
    except:
        pass

    return 0

def get_cpu_temp():
    try:
        cmd = (
            'Get-WmiObject -Namespace "root/cimv2" -Class Win32_PerfFormattedData_Counters_ThermalZoneInformation | '
            'Select-Object -First 1 -ExpandProperty HighPrecisionTemperature'
        )
        result = subprocess.check_output(["powershell", "-Command", cmd], encoding='utf-8', stderr=subprocess.DEVNULL)
        if result.strip():
            # Temperature is in tenths of Kelvin
            temp_kelvin = float(result.strip()) / 10.0
            temp_celsius = temp_kelvin - 273.15
            return int(temp_celsius)
    except (subprocess.CalledProcessError, ValueError):
        pass

    return 0

#════════════════════════════════════════════════════Gpu══════════════════════════════════════════════════════════════#

def detect_gpu():
    try:
        gpu_name = subprocess.check_output(
            ["powershell", "-Command", "(Get-CimInstance Win32_VideoController).Name"],
            encoding="utf-8",
            stderr=subprocess.DEVNULL
        ).strip()
        return _clean_name(gpu_name)
    except (subprocess.CalledProcessError, UnicodeDecodeError):
        return "GPU Unknown"

def get_gpu_load():
    try:
        cmd = (
            'Get-Counter "\\GPU Engine(*3D*)\\Utilization Percentage" | '
            'Select-Object -ExpandProperty CounterSamples | '
            'Select-Object -ExpandProperty CookedValue'
        )
        result = subprocess.check_output(["powershell", "-Command", cmd], encoding='utf-8', stderr=subprocess.DEVNULL)
        values = [float(v) for v in result.strip().split('\n') if v.strip()]
        return int(max(values)) if values else 0
    except (subprocess.CalledProcessError, ValueError, IndexError):
        return 0

def get_gpu_wattage():
    """Get GPU power consumption from OpenHardwareMonitor or estimation"""
    # Method 1: Try OpenHardwareMonitor (most reliable)
    try:
        cmd = (
            'Get-WmiObject -Namespace "root/OpenHardwareMonitor" -Class Sensor | '
            'Where-Object {$_.SensorType -eq "Power" -and ($_.Name -like "*GPU*" -or $_.Name -like "*3D*")} | '
            'Select-Object -ExpandProperty Value | Select-Object -First 1'
        )
        result = subprocess.check_output(["powershell", "-Command", cmd], encoding='utf-8', stderr=subprocess.DEVNULL)
        if result.strip():
            wattage = float(result.strip())
            return int(wattage)
    except (subprocess.CalledProcessError, ValueError):
        pass

    # Method 2: Try WMI GPU performance data
    try:
        cmd = (
            'Get-WmiObject -Namespace "root/cimv2" -Class Win32_PerfFormattedData_Counters_GPUEngine | '
            'Select-Object -First 1 -ExpandProperty "Utilization Percentage"'
        )
        result = subprocess.check_output(["powershell", "-Command", cmd], encoding='utf-8', stderr=subprocess.DEVNULL)
        if result.strip():
            gpu_util = float(result.strip())
            estimated_watts = int((gpu_util / 100) * 300)
            return estimated_watts
    except (subprocess.CalledProcessError, ValueError):
        pass

    # Fallback: estimation based on GPU load
    try:
        gpu_load = get_gpu_load()
        estimated_watts = int((gpu_load / 100) * 300)
        return estimated_watts
    except:
        pass

    return 0

def get_gpu_temp():
    try:
        cmd = (
            'Get-WmiObject -Namespace "root/OpenHardwareMonitor" -Class Sensor | '
            'Where-Object {$_.SensorType -eq "Temperature" -and $_.Name -like "*GPU Core*"} | '
            'Select-Object -ExpandProperty Value'
        )
        result = subprocess.check_output(["powershell", "-Command", cmd], encoding='utf-8', stderr=subprocess.DEVNULL)
        temp = float(result.strip())
        return int(temp)
    except (subprocess.CalledProcessError, ValueError):
        return 0

#═════════════════════════════════════════════════════════════════════════════════════════════════════════════════════#
# NETWORK MONITORING
#═════════════════════════════════════════════════════════════════════════════════════════════════════════════════════#

def get_network_usage(prev, prev_time):
    now = time.time()
    try:
        cur = psutil.net_io_counters(pernic=True)[INTERFACE]
        elapsed = now - prev_time
        up = (cur.bytes_sent - prev.bytes_sent) / elapsed
        down = (cur.bytes_recv - prev.bytes_recv) / elapsed
        return cur, up, down, now
    except (KeyError, ZeroDivisionError):
        return prev, 0, 0, now

#═════════════════════════════════════════════════════════════════════════════════════════════════════════════════════#
# MEDIA MONITORING
#═════════════════════════════════════════════════════════════════════════════════════════════════════════════════════#

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

#═════════════════════════════════════════════════════════════════════════════════════════════════════════════════════#
# MAIN OSC LOOP
#═════════════════════════════════════════════════════════════════════════════════════════════════════════════════════#

def run_osc_loop():
    global running, cpu_wattage, cpu_temp, gpu_wattage, gpu_temp

    all_stats = psutil.net_io_counters(pernic=True)
    if INTERFACE not in all_stats:
        print(f"Interface not found: {INTERFACE}")
        running = False
        return

    prev = all_stats[INTERFACE]
    prev_time = time.time()

    cpu_detect = detect_cpu()
    gpu_detect = detect_gpu()

    print(cpu_detect)
    print(gpu_detect)
    print("Sending live data to VRChat...")

    while running:
        song, artist, pos, dur = asyncio.run(get_media_info())
        clean_song = clean_title(song)

        cpu = psutil.cpu_percent()
        gpu = get_gpu_load()

        cpu_temp = get_cpu_temp()
        cpu_wattage = get_cpu_wattage()
        gpu_temp = get_gpu_temp()
        gpu_wattage = get_gpu_wattage()

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

        client.send_message("/chatbox/input", [text, True])  # type: ignore
        time.sleep(1.6)

#═════════════════════════════════════════════════════════════════════════════════════════════════════════════════════#
# SCRIPT CONTROL FUNCTIONS
#═════════════════════════════════════════════════════════════════════════════════════════════════════════════════════#

def start_script():
    global running, client, OSC_IP, OSC_PORT, INTERFACE, SWITCH_INTERVAL
    global page1_line1_text, page2_line1_text

    if running:
        return

    OSC_IP = ip_entry.get()
    OSC_PORT = int(port_entry.get())
    INTERFACE = iface_entry.get()
    SWITCH_INTERVAL = int(interval_entry.get())

    page1_line1_text = page1_entry.get()
    page2_line1_text = page2_entry.get()

    client = SimpleUDPClient(OSC_IP, OSC_PORT)

    running = True
    status_label.config(text="Status: Running", fg="#4CFF4C")

    thread = threading.Thread(target=run_osc_loop, daemon=True)
    thread.start()


def stop_script():
    global running
    running = False
    status_label.config(text="Status: Stopped", fg="#FF4C4C")

#═════════════════════════════════════════════════════════════════════════════════════════════════════════════════════#
# GUI SETUP
#═════════════════════════════════════════════════════════════════════════════════════════════════════════════════════#

BG = "#121212"
FG = "#E0E0E0"
ENTRY_BG = "#1E1E1E"
BTN_BG = "#2A2A2A"
BTN_FG = "#FFFFFF"

root = tk.Tk()
root.title("VRChat OSC Monitor")
root.geometry("400x360")
root.configure(bg=BG)
root.resizable(False, False)

frame = tk.Frame(root, bg=BG)
frame.pack(fill="both", expand=True, padx=10, pady=10)

#═════════════════════════════════════════════════════════════════════════════════════════════════════════════════════#
# GUI Helper Functions
#═════════════════════════════════════════════════════════════════════════════════════════════════════════════════════#

def dark_label(text, r):
    lbl = tk.Label(frame, text=text, bg=BG, fg=FG, anchor="w")
    lbl.grid(row=r, column=0, sticky="w", pady=4)
    return lbl


def dark_entry(r, default=""):
    e = tk.Entry(frame, bg=ENTRY_BG, fg=FG, insertbackground=FG, relief="flat")
    e.insert(0, default)
    e.grid(row=r, column=1, pady=4, sticky="ew")
    return e

#─────────────────────────────────────────────────────────────────────────────────────────────────────────────────────#
# GUI Layout
#─────────────────────────────────────────────────────────────────────────────────────────────────────────────────────#

frame.columnconfigure(1, weight=1)

dark_label("OSC IP", 0)
ip_entry = dark_entry(0, OSC_IP)

dark_label("OSC Port", 1)
port_entry = dark_entry(1, str(OSC_PORT))

dark_label("Network Interface", 2)
iface_entry = dark_entry(2, INTERFACE)

dark_label("Switch Interval", 3)
interval_entry = dark_entry(3, str(SWITCH_INTERVAL))

dark_label("Page 1 - Line 1 Text", 4)
page1_entry = dark_entry(4, "-enter text-")

dark_label("Page 2 - Line 1 Text", 5)
page2_entry = dark_entry(5, "-enter text-")

start_btn = tk.Button(frame, text="Start", command=start_script,
                      bg=BTN_BG, fg=BTN_FG, relief="flat")
start_btn.grid(row=6, column=0, pady=15, sticky="ew")

stop_btn = tk.Button(frame, text="Stop", command=stop_script,
                     bg=BTN_BG, fg=BTN_FG, relief="flat")
stop_btn.grid(row=6, column=1, pady=15, sticky="ew")

status_label = tk.Label(frame, text="Status: Stopped", bg=BG, fg="#FF4C4C")
status_label.grid(row=7, column=0, columnspan=2)

#═════════════════════════════════════════════════════════════════════════════════════════════════════════════════════#
# START APPLICATION
#═════════════════════════════════════════════════════════════════════════════════════════════════════════════════════#

root.mainloop()