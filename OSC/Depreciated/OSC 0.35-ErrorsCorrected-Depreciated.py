import asyncio
import re
import subprocess
import threading
import time
import tkinter as tk
import psutil
import winrt.windows.media.control as wmc
from pythonosc.udp_client import SimpleUDPClient

OSC_IP = "127.0.0.1"
OSC_PORT = 9000
INTERFACE = "Ethernet"
SWITCH_INTERVAL = 30

print("This a beta version use at your own risk")

client = None
running = False
page1_line1_text = "-enter text-"
page2_line1_text = "-enter text-"


def fmt(bps):
    if bps > 1024 * 1024:
        return f"{bps / (1024 * 1024):.2f} MB/s"
    return f"{bps / 1024:.1f} KB/s"


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


def _clean_name(name: str):
    name = re.sub(r"\(.*?\)|\[.*?]|\{.*?}", "", name)
    name = name.split("@")[0]
    name = re.sub(r"\s+", " ", name).strip()
    return name


def detect_cpu():
    try:
        cpu_name = subprocess.check_output(
            ["powershell", "-Command", "(Get-CimInstance Win32_Processor).Name"],
            encoding="utf-8",
            stderr=subprocess.DEVNULL
        ).strip()
        return _clean_name(cpu_name)
    except (subprocess.CalledProcessError, ValueError, IndexError):
        return 0


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


def create_progress_bar(position_ms, duration_ms, length=13):
    if duration_ms <= 0:
        return "─" * length
    percent = min(max(position_ms / duration_ms, 0), 1)
    filled_len = int(length * percent)
    return "■" * filled_len + "□" * (length - filled_len)


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


def clean_title(raw_title,):
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


def run_osc_loop():
    global running

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
        clean_song = clean_title(song)  # Remove the artist parameter

        cpu = psutil.cpu_percent()
        gpu = get_gpu_load()
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
                f"{gpu_detect} {gpu}%\n"
                f"{progress_bar}\n"
                f"{display_song} {display_artist}"
            )

        client.send_message("/chatbox/input", [text, True])  # type: ignore
        time.sleep(1.6)


#──────────────────────────────────────────────────────────UI──────────────────────────────────────────────────────────#

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


# Dark theme colors
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

# New controls for your first lines
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

root.mainloop()