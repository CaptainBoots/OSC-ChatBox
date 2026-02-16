import time
import psutil
import subprocess
from pythonosc.udp_client import SimpleUDPClient
import ctypes

# ---------------- CONFIG ----------------
OSC_IP = "127.0.0.1"
OSC_PORT = 9000
INTERFACE = "Ethernet"
SWITCH_INTERVAL = 30

print("This code is depreciated use at your own risk")

client = SimpleUDPClient(OSC_IP, OSC_PORT)

# Windows API functions
GetForegroundWindow = ctypes.windll.user32.GetForegroundWindow
GetWindowTextLengthW = ctypes.windll.user32.GetWindowTextLengthW
GetWindowTextW = ctypes.windll.user32.GetWindowTextW

def fmt(bps):
    if bps > 1024 * 1024:
        return f"{bps / (1024 * 1024):.2f} MB/s"
    return f"{bps / 1024:.1f} KB/s"

def get_amd_gpu_load():
    try:
        ps_cmd = 'Get-Counter "\\GPU Engine(*3D*)\\Utilization Percentage" | Select-Object -ExpandProperty CounterSamples | Select-Object -ExpandProperty CookedValue'
        result = subprocess.check_output(["powershell", "-Command", ps_cmd], encoding='utf-8', stderr=subprocess.DEVNULL)
        values = [float(val) for val in result.strip().split('\n') if val.strip()]
        if values:
            return int(max(values))
    except Exception:
        pass
    return 0

def create_progress_bar(position_ms, duration_ms, length=13):
    if duration_ms <= 0:
        return "─" * length
    percent = min(max(position_ms / duration_ms, 0), 1)
    filled_len = int(length * percent)
    bar = "■" * filled_len + "□" * (length - filled_len)
    return bar

def get_media_info():
    """
    Detects media from the currently active window title.
    Returns: song, artist, pos_ms, dur_ms
    """
    hwnd = GetForegroundWindow()
    length = GetWindowTextLengthW(hwnd)
    buff = ctypes.create_unicode_buffer(length + 1)
    GetWindowTextW(hwnd, buff, length + 1)
    title = buff.value

    # Split title like "Song - Artist" or browser tab like "Song name - YouTube"
    if " - " in title:
        parts = title.split(" - ")
        song = parts[0].strip()
        artist = parts[1].strip() if len(parts) > 1 else "Unknown Artist"
        pos = 0
        dur = 180000  # 3 min default
        return song, artist, pos, dur
    return None, None, 0, 0

def run_osc_loop():
    try:
        prev = psutil.net_io_counters(pernic=True)[INTERFACE]
    except KeyError:
        print(f"Error: {INTERFACE} not found.")
        return

    prev_time = time.time()
    print("Sending live data to VRChat...")

    while True:
        now = time.time()

        # ---- System stats ----
        cpu = psutil.cpu_percent()
        gpu = get_amd_gpu_load()

        # ---- Network speeds ----
        cur = psutil.net_io_counters(pernic=True)[INTERFACE]
        elapsed = now - prev_time
        up_raw = (cur.bytes_sent - prev.bytes_sent) / elapsed
        down_raw = (cur.bytes_recv - prev.bytes_recv) / elapsed
        up_speed = fmt(up_raw)
        down_speed = fmt(down_raw)

        # ---- Media info ----
        song, artist, pos, dur = get_media_info()
        display_artist = f"{artist}" if artist else "Unknown Artist"
        display_song = f"🎵 {song}" if song else "🎵 No Music Playing"
        progress_bar = create_progress_bar(pos, dur)

        # ---- Page switching ----
        page_index = int((now // SWITCH_INTERVAL) % 2)

        if page_index == 0:  # Page 1
            text = (
                "Im running this shit with python get gud bro\n"
                f"{time.strftime('%I:%M %p')}\n"
                f"Download: {down_speed}\n"
                f"Upload: {up_speed}\n"
                f"{progress_bar}\n"
                f"{display_song} - {display_artist}"
            )
        else:  # Page 2
            text = (
                "Blasting Music\n"
                f"CPU {cpu}%\n"
                f"GPU {gpu}%\n"
                f"{progress_bar}\n"
                f"{display_song} - {display_artist}"
            )

        client.send_message("/chatbox/input", [text, True])

        prev = cur
        prev_time = now
        time.sleep(1.6)

if __name__ == "__main__":
    run_osc_loop()

