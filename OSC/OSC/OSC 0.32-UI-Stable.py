import time
import asyncio
import psutil
import subprocess
import re
import threading
import tkinter as tk
from tkinter import ttk
from pythonosc.udp_client import SimpleUDPClient
import winrt.windows.media.control as wmc

OSC_IP = "127.0.0.1"
OSC_PORT = 9000
INTERFACE = "Ethernet"
SWITCH_INTERVAL = 30

client = None
running = False


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
    except:
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
    except:
        return "CPU Unknown"


def detect_gpu():
    try:
        gpu_name = subprocess.check_output(
            ["powershell", "-Command", "(Get-CimInstance Win32_VideoController).Name"],
            encoding="utf-8",
            stderr=subprocess.DEVNULL
        ).strip()
        return _clean_name(gpu_name)
    except:
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
    except:
        pass
    return None, None, 0, 0


def clean_title(raw_title, artist=None):
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
    except:
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
        clean_song = clean_title(song, artist)

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
                f"Im running this shit with python\n"
                f"{cur_time_str}\n"
                f"Download {fmt(down_raw)}\n"
                f"Upload {fmt(up_raw)}\n"
                f"{progress_bar}\n"
                f"{display_song} {display_artist}"
            )
        else:
            text = (
                f"Blasting Music\n"
                f"{cur_time_str}\n"
                f"{cpu_detect} {cpu}%\n"
                f"{gpu_detect} {gpu}%\n"
                f"{progress_bar}\n"
                f"{display_song} {display_artist}"
            )

        client.send_message("/chatbox/input", [text, True])
        time.sleep(1.6)


# ---------------- UI ---------------- #

def start_script():
    global running, client, OSC_IP, OSC_PORT, INTERFACE, SWITCH_INTERVAL

    if running:
        return

    OSC_IP = ip_entry.get()
    OSC_PORT = int(port_entry.get())
    INTERFACE = iface_entry.get()
    SWITCH_INTERVAL = int(interval_entry.get())

    client = SimpleUDPClient(OSC_IP, OSC_PORT)

    running = True
    status_label.config(text="Status: Running", foreground="green")

    thread = threading.Thread(target=run_osc_loop, daemon=True)
    thread.start()


def stop_script():
    global running
    running = False
    status_label.config(text="Status: Stopped", foreground="red")


root = tk.Tk()
root.title("VRChat OSC Monitor")
root.geometry("360x260")
root.resizable(False, False)

frame = ttk.Frame(root, padding=10)
frame.pack(fill="both", expand=True)

ttk.Label(frame, text="OSC IP").grid(row=0, column=0, sticky="w")
ip_entry = ttk.Entry(frame)
ip_entry.insert(0, OSC_IP)
ip_entry.grid(row=0, column=1)

ttk.Label(frame, text="OSC Port").grid(row=1, column=0, sticky="w")
port_entry = ttk.Entry(frame)
port_entry.insert(0, str(OSC_PORT))
port_entry.grid(row=1, column=1)

ttk.Label(frame, text="Network Interface").grid(row=2, column=0, sticky="w")
iface_entry = ttk.Entry(frame)
iface_entry.insert(0, INTERFACE)
iface_entry.grid(row=2, column=1)

ttk.Label(frame, text="Switch Interval").grid(row=3, column=0, sticky="w")
interval_entry = ttk.Entry(frame)
interval_entry.insert(0, str(SWITCH_INTERVAL))
interval_entry.grid(row=3, column=1)

start_btn = ttk.Button(frame, text="Start", command=start_script)
start_btn.grid(row=4, column=0, pady=10)

stop_btn = ttk.Button(frame, text="Stop", command=stop_script)
stop_btn.grid(row=4, column=1, pady=10)

status_label = ttk.Label(frame, text="Status: Stopped", foreground="red")
status_label.grid(row=5, column=0, columnspan=2)

root.mainloop()
