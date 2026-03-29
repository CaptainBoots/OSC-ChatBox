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
import re
import threading
import time
import site
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
        subprocess.check_call([
            sys.executable, "-m", "pip", "install",
            package, "--break-system-packages"
        ])
        # Make sure newly installed packages are findable
        user_site = site.getusersitepackages()
        if user_site not in sys.path:
            sys.path.insert(0, user_site)

install_if_missing("python-osc", "pythonosc")
install_if_missing("psutil", "psutil")
install_if_missing("requests", "requests")

import psutil
from pythonosc.udp_client import SimpleUDPClient

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

INTERFACE = "eno1"

SWITCH_INTERVAL = 30

LHM_REST_API = "http://localhost:8085/data.json"

client = None
running = False

page1_line1_text = "-enter text-"
page2_line1_text = "-enter text-"

cpu_wattage = "error"
cpu_temp = "error"
gpu_wattage = "error"
gpu_temp = "error"

cpu_manufacturer = CPUManufacturer.UNKNOWN

# ═════════════════════════════════════════════════════════════════════════════════════════════════════════════════════#
# HARDWARE MONITORING
# ═════════════════════════════════════════════════════════════════════════════════════════════════════════════════════#

gpu_id_map = {
    # AMD RDNA3
    "1002:7590": "Radeon RX 9060 XT",
    "1002:7340": "Radeon RX 7900 XTX",
    "1002:7341": "Radeon RX 7900 XT",
    "1002:71b0": "Radeon RX 7800 XT",
    "1002:7438": "Radeon RX 7700 XT",
    "1002:7439": "Radeon RX 7700",
    "1002:7540": "Radeon RX 7600",
    "1002:7541": "Radeon RX 7600 XT",

    # AMD RDNA2
    "1002:73bf": "Radeon RX 6950 XT",
    "1002:73b7": "Radeon RX 6900 XT",
    "1002:73b8": "Radeon RX 6800",
    "1002:73ab": "Radeon RX 6600 XT",
    "1002:73a0": "Radeon RX 6600",
    "1002:73a4": "Radeon RX 6500 XT",

    # NVIDIA RTX 40 Series
    "10de:2498": "GeForce RTX 4090",
    "10de:2487": "GeForce RTX 4080",
    "10de:2490": "GeForce RTX 4070 Ti",
    "10de:248e": "GeForce RTX 4070",
    "10de:2481": "GeForce RTX 4060 Ti",
    "10de:2480": "GeForce RTX 4060",

    # NVIDIA RTX 30 Series
    "10de:2204": "GeForce RTX 3090",
    "10de:2206": "GeForce RTX 3080 Ti",
    "10de:1e87": "GeForce RTX 3080",
    "10de:1e84": "GeForce RTX 3070",
    "10de:2485": "GeForce RTX 3060 Ti",
    "10de:2483": "GeForce RTX 3060",

    # Intel Integrated
    "8086:3e92": "Intel UHD Graphics 630",
    "8086:9bc5": "Intel Iris Xe Graphics",
}

def _clean_name(name: str):
    name = re.sub(r"\(.*?\)|\[.*?]|\{.*?}", "", name)
    name = name.split("@")[0]
    name = re.sub(r"\s+", " ", name).strip()
    return name


def detect_cpu():
    global cpu_manufacturer
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


def detect_gpu_pci_id():
    try:
        output = subprocess.check_output(
            ["lspci", "-nn"], encoding="utf-8", stderr=subprocess.DEVNULL
        )
        for line in output.splitlines():
            if "VGA" in line or "3D controller" in line:
                match = re.search(r"\[(\w{4}:\w{4})\]", line)
                if match:
                    return match.group(1).lower()
    except Exception:
        pass
    return None


def detect_gpu():
    pci_id = detect_gpu_pci_id()
    return gpu_id_map.get(pci_id, f"Unknown GPU ({pci_id})" if pci_id else "GPU Unknown")


def _read_sysfs(path):
    """Read a single integer value from a sysfs file, return None on failure."""
    try:
        with open(path, "r") as f:
            return int(f.read().strip())
    except (OSError, ValueError):
        return None


def _get_cpu_temp_celsius():
    """
    Try multiple sysfs sources for CPU package temperature.
    Works for both Intel (coretemp) and AMD (k10temp) CPUs.
    """
    import glob

    # Look for a hwmon device with 'coretemp' or 'k10temp' as the driver name
    for hwmon_dir in glob.glob("/sys/class/hwmon/hwmon*"):
        try:
            with open(f"{hwmon_dir}/name") as f:
                name = f.read().strip()
        except OSError:
            continue

        if name in ("coretemp", "k10temp"):
            # coretemp: temp1 is usually "Package id 0"
            # k10temp:  Tctl/Tdie is temp1
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
                        return raw // 1000  # millidegrees → degrees

            # Fallback: just use temp1_input if no label matched
            raw = _read_sysfs(f"{hwmon_dir}/temp1_input")
            if raw is not None:
                return raw // 1000

    return 0


def _get_cpu_power_watts():
    """
    Return the CPU power in watts.
    Works for Intel (RAPL) and AMD (hwmon) CPUs.
    """
    import glob
    import os
    import time

    # Intel RAPL
    rapl = "/sys/class/powercap/intel-rapl/intel-rapl:0/energy_uj"
    if os.path.exists(rapl):
        e1 = None
        e2 = None
        try:
            e1 = int(open(rapl).read().strip())
            time.sleep(0.2)
            e2 = int(open(rapl).read().strip())
        except:
            pass
        if e1 is not None and e2 is not None and e2 > e1:
            return int((e2 - e1) / 0.2 / 1_000_000)  # µJ → W

    # AMD hwmon
    for hwmon in glob.glob("/sys/class/hwmon/hwmon*"):
        try:
            name = open(f"{hwmon}/name").read().strip()
            if name.lower() in ("k10temp", "fam17h", "zenpower"):
                raw = _read_sysfs(f"{hwmon}/power1_input")
                if raw:
                    return raw // 1_000_000  # µW → W
        except:
            continue

    return 0


def _get_gpu_stats():
    import glob

    for hwmon in glob.glob("/sys/class/drm/card*/device/hwmon/hwmon*"):
        temp = _read_sysfs(f"{hwmon}/temp1_input")
        power = (
            _read_sysfs(f"{hwmon}/power1_average")
            or _read_sysfs(f"{hwmon}/power1_input")
        )

        if temp is not None:
            temp = temp // 1000

        if power is not None:
            power = power // 1_000_000

        return temp or 0, power or 0, 0

    return 0, 0, 0


def get_lhm_data():
    """
    Collect CPU/GPU stats for OSC display.
    """
    gpu_temp, gpu_power, gpu_load = _get_gpu_stats()  # GPU function you already have

    return {
        "cpu_temp":  _get_cpu_temp_celsius(),
        "cpu_power": _get_cpu_power_watts(),  # uses the function we just defined
        "gpu_temp":  gpu_temp,
        "gpu_power": gpu_power,
        "gpu_load":  gpu_load,
        "cpu_load":  int(psutil.cpu_percent(interval=None)),
    }


def parse_lhm_data(data):
    if not data:
        return 0, 0, 0, 0
    return (
        data.get("cpu_temp",  0),
        data.get("cpu_power", 0),
        data.get("gpu_temp",  0),
        data.get("gpu_power", 0),
    )


def get_cpu_load_from_lhm(data):
    if not data:
        return 0
    return data.get("cpu_load", 0)


def get_gpu_load_from_lhm(data):
    if not data:
        return 0
    return data.get("gpu_load", 0)


def diagnose_lhm():
    import glob
    print("[DIAGNOSTIC] Running Linux hardware diagnostics...")

    # CPU temp
    cpu_t = _get_cpu_temp_celsius()
    if cpu_t:
        print(f"[DIAGNOSTIC] ✓ CPU Temperature: {cpu_t}°C")
    else:
        print("[DIAGNOSTIC] ✗ CPU Temperature: not found")
        print("[DIAGNOSTIC]   FIX: sudo apt install lm-sensors && sudo sensors-detect")

    # CPU power
    cpu_p = _get_cpu_power_watts()
    if cpu_p:
        print(f"[DIAGNOSTIC] ✓ CPU Power: {cpu_p}W")
    else:
        print("[DIAGNOSTIC] ✗ CPU Power: not found")
        print("[DIAGNOSTIC]   FIX (Intel): sudo chmod a+r /sys/class/powercap/intel-rapl/intel-rapl:0/energy_uj")
        print("[DIAGNOSTIC]   FIX (AMD):   check /sys/class/hwmon/hwmon*/name for k10temp")

    # rocm-smi / GPU
    gpu_t, gpu_p, gpu_l = _get_gpu_stats_rocm()
    if gpu_t or gpu_p or gpu_l:
        print(f"[DIAGNOSTIC] ✓ GPU: {gpu_t}°C  {gpu_p}W  {gpu_l}% load")
    else:
        print("[DIAGNOSTIC] ✗ GPU stats: not found")
        print("[DIAGNOSTIC]   FIX: sudo apt install rocm-smi  (AMD only)")

    # Network interface
    all_ifaces = list(psutil.net_io_counters(pernic=True).keys())
    if INTERFACE in all_ifaces:
        print(f"[DIAGNOSTIC] ✓ Network interface '{INTERFACE}' found")
    else:
        print(f"[DIAGNOSTIC] ✗ Interface '{INTERFACE}' not found. Available: {all_ifaces}")

    return True

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
# MEDIA MONITORING - Browser & YouTube Music compatible
# ═════════════════════════════════════════════════════════════════════════════════════════════════════════════════════#

def get_media_info():
    try:
        players = subprocess.check_output(
            ["playerctl", "-l"],
            encoding="utf-8",
            stderr=subprocess.DEVNULL,
            timeout=2
        ).splitlines()

        if not players:
            return None, None, 0, 0

        browser_player = next(
            (p for p in players if any(k in p.lower() for k in ["chrome", "chromium", "firefox"])),
            players[0]  # fallback to first player
        )

        output = subprocess.check_output(
            ["playerctl", "-p", browser_player, "metadata",
             "--format", "{{title}}\n{{artist}}\n{{position}}\n{{mpris:length}}"],
            encoding="utf-8",
            stderr=subprocess.DEVNULL,
            timeout=2
        ).strip().split("\n")

        if len(output) >= 4:
            title  = output[0].strip()
            artist = output[1].strip()
            pos    = int(output[2]) / 1000        # µs → ms
            dur    = int(output[3]) / 1000        # µs → ms
            return title, artist, pos, dur

    except subprocess.CalledProcessError:
        pass
    except Exception as e:
        print(f"[MEDIA ERROR] {e}")

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
    return re.sub(r"\s+", " ", parts[0] if parts else title).strip()


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
    global running, cpu_wattage, cpu_temp, gpu_wattage, gpu_temp, client

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
    print(f"{'=' * 60}")

    query_cooldown = 0
    cpu = 0
    gpu = 0

    while running:
        try:
            song, artist, pos, dur = get_media_info()
            clean_song = clean_title(song)

            query_cooldown += 1
            if query_cooldown >= 3:
                lhm_data = get_lhm_data()
                if lhm_data:
                    cpu_temp, cpu_wattage, gpu_temp, gpu_wattage = parse_lhm_data(lhm_data)
                    cpu = get_cpu_load_from_lhm(lhm_data)
                    gpu = get_gpu_load_from_lhm(lhm_data)
                else:
                    cpu = 0
                    gpu = 0
                query_cooldown = 0

            prev, up_raw, down_raw, prev_time = get_network_usage(prev, prev_time)

            cur_time_str = time.strftime("%I:%M %p")
            progress_bar = create_progress_bar(pos, dur)

            display_artist = f"- {artist}" if artist else ""
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

dark_label("Page 1 Text", 5)
page1_entry = dark_entry(5, "Thx for using boot's osc code")

dark_label("Page 2 Text", 6)
page2_entry = dark_entry(6, "hi put your text here :3")

button_frame = tk.Frame(frame, bg=BG)
button_frame.grid(row=7, column=0, columnspan=2, pady=15, sticky="ew")
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
status_label.grid(row=8, column=0, columnspan=2)

root.mainloop()