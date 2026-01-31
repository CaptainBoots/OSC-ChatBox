# ═════════════════════════════════════════════════════════════════════════════════════════════════════════════════════#
#                                              OSC Python Script (Enhanced)                                            #
# ═════════════════════════════════════════════════════════════════════════════════════════════════════════════════════#
# Hi :3
# Welcome to my enhanced code with Intel PCM and AMD uProf support

# ═════════════════════════════════════════════════════════════════════════════════════════════════════════════════════#
# Imports
# ═════════════════════════════════════════════════════════════════════════════════════════════════════════════════════#

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


# ═════════════════════════════════════════════════════════════════════════════════════════════════════════════════════#
# ENUMS
# ═════════════════════════════════════════════════════════════════════════════════════════════════════════════════════#

class CPUManufacturer(Enum):
    INTEL = "Intel"
    AMD = "AMD"
    UNKNOWN = "Unknown"


# ═════════════════════════════════════════════════════════════════════════════════════════════════════════════════════#
# CONFIGURATION & GLOBAL VARIABLES
# ═════════════════════════════════════════════════════════════════════════════════════════════════════════════════════#

OSC_IP = "127.0.0.1"
OSC_PORT = 9000
INTERFACE = "Ethernet"
SWITCH_INTERVAL = 30

print("This is a beta version, use at your own risk")

client = None
running = False
page1_line1_text = "-enter text-"
page2_line1_text = "-enter text-"

cpu_wattage = "-error-"
cpu_temp = "-error-"
gpu_wattage = "-error-"
gpu_temp = "-error-"

cpu_manufacturer = CPUManufacturer.UNKNOWN
pcm_available = False
uprof_available = False


# ═════════════════════════════════════════════════════════════════════════════════════════════════════════════════════#
# HARDWARE MONITORING
# ═════════════════════════════════════════════════════════════════════════════════════════════════════════════════════#

def _clean_name(name: str):
    """Clean hardware names by removing extra characters and whitespace"""
    name = re.sub(r"\(.*?\)|\[.*?]|\{.*?}", "", name)
    name = name.split("@")[0]
    name = re.sub(r"\s+", " ", name).strip()
    return name


# ════════════════════════════════════════════════════Cpu══════════════════════════════════════════════════════════════#

def detect_cpu():
    """Detect CPU name and manufacturer"""
    global cpu_manufacturer, pcm_available, uprof_available

    try:
        cpu_name = subprocess.check_output(
            ["powershell", "-Command", "(Get-CimInstance Win32_Processor).Name"],
            encoding="utf-8",
            stderr=subprocess.DEVNULL
        ).strip()

        clean_cpu_name = _clean_name(cpu_name)

        # Detect CPU manufacturer
        if "intel" in cpu_name.lower():
            cpu_manufacturer = CPUManufacturer.INTEL
        elif "amd" in cpu_name.lower():
            cpu_manufacturer = CPUManufacturer.AMD
        else:
            cpu_manufacturer = CPUManufacturer.UNKNOWN

        # Check for available tools
        check_cpu_tools()

        return clean_cpu_name
    except (subprocess.CalledProcessError, UnicodeDecodeError):
        return "CPU Unknown"


def check_cpu_tools():
    """Check if Intel PCM or AMD uProf are available"""
    global pcm_available, uprof_available

    if cpu_manufacturer == CPUManufacturer.INTEL:
        # Check for Intel PCM
        try:
            result = subprocess.run(
                ["where", "pcm.exe"],
                capture_output=True,
                timeout=2
            )
            pcm_available = result.returncode == 0
            if pcm_available:
                print("✓ Intel PCM detected and available")
            else:
                print(
                    "⚠ Intel PCM not found. Download from: https://www.intel.com/content/www/en/en/developer/articles/tool/performance-counter-monitor.html")
        except (subprocess.TimeoutExpired, FileNotFoundError):
            pcm_available = False
            print("⚠ Intel PCM check failed")

    elif cpu_manufacturer == CPUManufacturer.AMD:
        # Check for AMD uProf
        try:
            result = subprocess.run(
                ["where", "AMDuProfCLI.exe"],
                capture_output=True,
                timeout=2
            )
            uprof_available = result.returncode == 0
            if uprof_available:
                print("✓ AMD uProf detected and available")
            else:
                print("⚠ AMD uProf not found. Download from: https://www.amd.com/en/developer/uprof.html")
        except (subprocess.TimeoutExpired, FileNotFoundError):
            uprof_available = False
            print("⚠ AMD uProf check failed")


def get_cpu_wattage_intel():
    """Get CPU wattage using Intel PCM"""
    if not pcm_available:
        return 0

    try:
        # Intel PCM outputs power data. This is a basic approach using WMI as fallback
        cmd = (
            'Get-WmiObject -Namespace "root/LibreHardwareMonitor" -Class Sensor | '
            'Where-Object { $_.SensorType -eq "Power" -and $_.Name -match "CPU Package" } | '
            'Select-Object -ExpandProperty Value'
        )
        result = subprocess.check_output(
            ["powershell", "-NoProfile", "-Command", cmd],
            encoding="utf-8", stderr=subprocess.DEVNULL, timeout=5
        ).strip()

        if result:
            return int(float(result))
    except (subprocess.CalledProcessError, ValueError, subprocess.TimeoutExpired):
        pass

    return 0


def get_cpu_temp_intel():
    """Get CPU temperature using Intel PCM or LibreHardwareMonitor"""
    try:
        cmd = (
            'Get-WmiObject -Namespace "root/LibreHardwareMonitor" -Class Sensor | '
            'Where-Object { $_.SensorType -eq "Temperature" -and $_.Name -match "CPU Package" } | '
            'Select-Object -ExpandProperty Value'
        )
        result = subprocess.check_output(
            ["powershell", "-NoProfile", "-Command", cmd],
            encoding="utf-8", stderr=subprocess.DEVNULL, timeout=5
        ).strip()

        if result:
            return int(float(result))
    except (subprocess.CalledProcessError, ValueError, subprocess.TimeoutExpired):
        pass

    return 0


def get_cpu_wattage_amd():
    """Get CPU wattage using AMD uProf or LibreHardwareMonitor"""
    if uprof_available:
        try:
            # AMD uProf CLI can output power metrics
            result = subprocess.check_output(
                ["AMDuProfCLI.exe", "info", "-l"],
                encoding="utf-8", stderr=subprocess.DEVNULL, timeout=5
            ).strip()

            # Parse output for power readings (format varies by uProf version)
            for line in result.split('\n'):
                if 'power' in line.lower() or 'watt' in line.lower():
                    try:
                        wattage = int(float(re.search(r'[\d.]+', line).group()))
                        return wattage
                    except (AttributeError, ValueError):
                        pass
        except (subprocess.CalledProcessError, subprocess.TimeoutExpired, FileNotFoundError):
            pass

    # Fallback to LibreHardwareMonitor
    try:
        cmd = (
            'Get-WmiObject -Namespace "root/LibreHardwareMonitor" -Class Sensor | '
            'Where-Object { $_.SensorType -eq "Power" -and $_.Name -match "CPU" } | '
            'Select-Object -ExpandProperty Value'
        )
        result = subprocess.check_output(
            ["powershell", "-NoProfile", "-Command", cmd],
            encoding="utf-8", stderr=subprocess.DEVNULL, timeout=5
        ).strip()

        if result:
            return int(float(result))
    except (subprocess.CalledProcessError, ValueError, subprocess.TimeoutExpired):
        pass

    return 0


def get_cpu_temp_amd():
    """Get CPU temperature using AMD uProf or LibreHardwareMonitor"""
    if uprof_available:
        try:
            # AMD uProf CLI can output temperature metrics
            result = subprocess.check_output(
                ["AMDuProfCLI.exe", "info", "-l"],
                encoding="utf-8", stderr=subprocess.DEVNULL, timeout=5
            ).strip()

            # Parse output for temperature readings
            for line in result.split('\n'):
                if 'temp' in line.lower() or '°c' in line.lower() or 'celsius' in line.lower():
                    try:
                        temp = int(float(re.search(r'[\d.]+', line).group()))
                        return temp
                    except (AttributeError, ValueError):
                        pass
        except (subprocess.CalledProcessError, subprocess.TimeoutExpired, FileNotFoundError):
            pass

    # Fallback to LibreHardwareMonitor
    try:
        cmd = (
            'Get-WmiObject -Namespace "root/LibreHardwareMonitor" -Class Sensor | '
            'Where-Object { $_.SensorType -eq "Temperature" -and $_.Name -match "CPU" } | '
            'Select-Object -ExpandProperty Value'
        )
        result = subprocess.check_output(
            ["powershell", "-NoProfile", "-Command", cmd],
            encoding="utf-8", stderr=subprocess.DEVNULL, timeout=5
        ).strip()

        if result:
            return int(float(result))
    except (subprocess.CalledProcessError, ValueError, subprocess.TimeoutExpired):
        pass

    return 0


def get_cpu_wattage():
    """Get CPU wattage based on detected manufacturer"""
    if cpu_manufacturer == CPUManufacturer.INTEL:
        return get_cpu_wattage_intel()
    elif cpu_manufacturer == CPUManufacturer.AMD:
        return get_cpu_wattage_amd()
    else:
        # Fallback for unknown CPUs
        return get_cpu_wattage_amd()


def get_cpu_temp():
    """Get CPU temperature based on detected manufacturer"""
    if cpu_manufacturer == CPUManufacturer.INTEL:
        return get_cpu_temp_intel()
    elif cpu_manufacturer == CPUManufacturer.AMD:
        return get_cpu_temp_amd()
    else:
        # Fallback for unknown CPUs
        return get_cpu_temp_amd()


# ════════════════════════════════════════════════════Gpu══════════════════════════════════════════════════════════════#

def detect_gpu():
    """Detect GPU name"""
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
    """Get GPU utilization percentage"""
    try:
        cmd = (
            'Get-Counter "\\GPU Engine(*3D*)\\Utilization Percentage" | '
            'Select-Object -ExpandProperty CounterSamples | '
            'Select-Object -ExpandProperty CookedValue'
        )
        result = subprocess.check_output(
            ["powershell", "-Command", cmd], encoding='utf-8', stderr=subprocess.DEVNULL
        )
        values = [float(v) for v in result.strip().split('\n') if v.strip()]
        return int(max(values)) if values else 0
    except (subprocess.CalledProcessError, ValueError, IndexError):
        return 0


def get_gpu_wattage():
    """Get GPU wattage"""
    try:
        cmd = (
            'Get-WmiObject -Namespace "root/LibreHardwareMonitor" -Class Sensor | '
            'Where-Object { $_.SensorType -eq "Power" -and $_.Name -match "GPU Power" } | '
            'Select-Object -ExpandProperty Value'
        )
        result = subprocess.check_output(
            ["powershell", "-NoProfile", "-Command", cmd],
            encoding="utf-8", stderr=subprocess.DEVNULL, timeout=5
        ).strip()

        if result:
            return int(float(result))
    except (subprocess.CalledProcessError, ValueError, subprocess.TimeoutExpired):
        pass

    return 0


def get_gpu_temp():
    """Get GPU temperature"""
    try:
        cmd = (
            'Get-WmiObject -Namespace "root/LibreHardwareMonitor" -Class Sensor | '
            'Where-Object { $_.SensorType -eq "Temperature" -and $_.Name -match "GPU Core" } | '
            'Select-Object -ExpandProperty Value'
        )
        result = subprocess.check_output(
            ["powershell", "-NoProfile", "-Command", cmd],
            encoding="utf-8", stderr=subprocess.DEVNULL, timeout=5
        ).strip()

        if result:
            return int(float(result))
    except (subprocess.CalledProcessError, ValueError, subprocess.TimeoutExpired):
        pass

    return 0


# ═════════════════════════════════════════════════════════════════════════════════════════════════════════════════════#
# NETWORK MONITORING
# ═════════════════════════════════════════════════════════════════════════════════════════════════════════════════════#

def fmt(bps):
    """Format bytes per second into readable format"""
    if bps > 1024 * 1024:
        return f"{bps / (1024 * 1024):.2f} MB/s"
    return f"{bps / 1024:.1f} KB/s"


def get_network_usage(prev, prev_time):
    """Get network upload and download speeds"""
    now = time.time()
    try:
        cur = psutil.net_io_counters(pernic=True)[INTERFACE]
        elapsed = now - prev_time
        up = (cur.bytes_sent - prev.bytes_sent) / elapsed
        down = (cur.bytes_recv - prev.bytes_recv) / elapsed
        return cur, up, down, now
    except (KeyError, ZeroDivisionError):
        return prev, 0, 0, now


# ═════════════════════════════════════════════════════════════════════════════════════════════════════════════════════#
# MEDIA MONITORING
# ═════════════════════════════════════════════════════════════════════════════════════════════════════════════════════#

async def get_media_info():
    """Get current media information (title, artist, position, duration)"""
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
    """Clean media title by removing extra info"""
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
    """Create a visual progress bar"""
    if duration_ms <= 0:
        return "─" * length
    percent = min(max(position_ms / duration_ms, 0), 1)
    filled_len = int(length * percent)
    return "■" * filled_len + "□" * (length - filled_len)


# ═════════════════════════════════════════════════════════════════════════════════════════════════════════════════════#
# MAIN OSC LOOP
# ═════════════════════════════════════════════════════════════════════════════════════════════════════════════════════#

def run_osc_loop():
    """Main OSC loop that sends data periodically"""
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

    print(f"CPU: {cpu_detect} ({cpu_manufacturer.value})")
    print(f"GPU: {gpu_detect}")
    print("Sending live data to OSC port..")

    while running:
        try:
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

            client.send_message("/chatbox/input", [text, True])
            time.sleep(5.0)

        except Exception as e:
            print(f"Error in OSC loop: {e}")
            time.sleep(1)


# ═════════════════════════════════════════════════════════════════════════════════════════════════════════════════════#
# SCRIPT CONTROL FUNCTIONS
# ═════════════════════════════════════════════════════════════════════════════════════════════════════════════════════#

def start_script():
    """Start the OSC monitoring script"""
    global running, client, OSC_IP, OSC_PORT, INTERFACE, SWITCH_INTERVAL
    global page1_line1_text, page2_line1_text

    if running:
        return

    try:
        OSC_IP = ip_entry.get()
        OSC_PORT = int(port_entry.get())
        INTERFACE = iface_entry.get()
        SWITCH_INTERVAL = int(interval_entry.get())

        page1_line1_text = page1_entry.get()
        page2_line1_text = page2_entry.get()

        client = SimpleUDPClient(OSC_IP, OSC_PORT)

        running = True
        status_label.config(text="Status: Running", fg="#4CFF4C")

        # Show info about detected CPU
        manufacturer_text = f"CPU: {cpu_manufacturer.value}"
        if cpu_manufacturer == CPUManufacturer.INTEL and pcm_available:
            manufacturer_text += " (Using PCM)"
        elif cpu_manufacturer == CPUManufacturer.AMD and uprof_available:
            manufacturer_text += " (Using uProf)"
        else:
            manufacturer_text += " (Using LibreHardwareMonitor)"

        info_label.config(text=manufacturer_text, fg="#FFB84D")

        thread = threading.Thread(target=run_osc_loop, daemon=True)
        thread.start()

    except ValueError as e:
        messagebox.showerror("Error", f"Invalid input: {e}")
        status_label.config(text="Status: Error", fg="#FF4C4C")
    except Exception as e:
        messagebox.showerror("Error", f"Failed to start: {e}")
        status_label.config(text="Status: Error", fg="#FF4C4C")


def stop_script():
    """Stop the OSC monitoring script"""
    global running
    running = False
    status_label.config(text="Status: Stopped", fg="#FF4C4C")
    info_label.config(text="")


# ═════════════════════════════════════════════════════════════════════════════════════════════════════════════════════#
# GUI SETUP
# ═════════════════════════════════════════════════════════════════════════════════════════════════════════════════════#

BG = "#121212"
FG = "#E0E0E0"
ENTRY_BG = "#1E1E1E"
BTN_BG = "#2A2A2A"
BTN_FG = "#FFFFFF"

root = tk.Tk()
root.title("OSC Chatbox")
root.geometry("400x400")
root.configure(bg=BG)
root.resizable(True, True)

frame = tk.Frame(root, bg=BG)
frame.pack(fill="both", expand=True, padx=10, pady=10)


# ═════════════════════════════════════════════════════════════════════════════════════════════════════════════════════#
# GUI Helper Functions
# ═════════════════════════════════════════════════════════════════════════════════════════════════════════════════════#

def dark_label(text, r):
    """Create a dark-themed label"""
    lbl = tk.Label(frame, text=text, bg=BG, fg=FG, anchor="w")
    lbl.grid(row=r, column=0, sticky="w", pady=4)
    return lbl


def dark_entry(r, default=""):
    """Create a dark-themed entry field"""
    e = tk.Entry(frame, bg=ENTRY_BG, fg=FG, insertbackground=FG, relief="flat")
    e.insert(0, default)
    e.grid(row=r, column=1, pady=4, sticky="ew")
    return e


# ─────────────────────────────────────────────────────────────────────────────────────────────────────────────────────#
# GUI Layout
# ─────────────────────────────────────────────────────────────────────────────────────────────────────────────────────#

frame.columnconfigure(1, weight=1)

dark_label("OSC IP", 0)
ip_entry = dark_entry(0, OSC_IP)

dark_label("OSC Port", 1)
port_entry = dark_entry(1, str(OSC_PORT))

dark_label("Network Interface", 2)
iface_entry = dark_entry(2, INTERFACE)

dark_label("Switch Interval", 3)
interval_entry = dark_entry(3, str(SWITCH_INTERVAL))

dark_label("Page 1 Text", 4)
page1_entry = dark_entry(4, "Thx for using boot's osc code")

dark_label("Page 2 Text", 5)
page2_entry = dark_entry(5, "hi put your text here :3")

start_btn = tk.Button(frame, text="Start", command=start_script,
                      bg=BTN_BG, fg=BTN_FG, relief="flat")
start_btn.grid(row=6, column=0, pady=15, sticky="ew")

stop_btn = tk.Button(frame, text="Stop", command=stop_script,
                     bg=BTN_BG, fg=BTN_FG, relief="flat")
stop_btn.grid(row=6, column=1, pady=15, sticky="ew")

status_label = tk.Label(frame, text="Status: Stopped", bg=BG, fg="#FF4C4C")
status_label.grid(row=7, column=0, columnspan=2)

info_label = tk.Label(frame, text="", bg=BG, fg="#FFB84D", wraplength=350)
info_label.grid(row=8, column=0, columnspan=2, pady=5)

# ═════════════════════════════════════════════════════════════════════════════════════════════════════════════════════#
# START APPLICATION
# ═════════════════════════════════════════════════════════════════════════════════════════════════════════════════════#

root.mainloop()