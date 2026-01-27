import time
import psutil
import subprocess
import win32gui
from pythonosc.udp_client import SimpleUDPClient

# ---------------- CONFIG ----------------
OSC_IP = "127.0.0.1"
OSC_PORT = 9000
INTERFACE = "Ethernet"
SWITCH_INTERVAL = 30

client = SimpleUDPClient(OSC_IP, OSC_PORT)

# ---------------- HELPERS ----------------
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


def create_progress_bar(position_percent, length=13):
    """position_percent: 0.0 -> 1.0"""
    position_percent = min(max(position_percent, 0), 1)
    filled_len = int(length * position_percent)
    bar = "■" * filled_len + "□" * (length - filled_len)
    return bar


def get_current_media():
    """Detect media from window titles"""
    media_title = None
    artist = None
    album = None
    pos = 0
    dur = 180  # simulate 3 min tracks

    def enum_windows_callback(hwnd, results):
        if win32gui.IsWindowVisible(hwnd):
            title = win32gui.GetWindowText(hwnd)
            if title:
                results.append(title)

    windows = []
    win32gui.EnumWindows(enum_windows_callback, windows)

    for title in windows:
        if " - " in title:
            parts = title.split(" - ")
            media_title = parts[0].strip()
            artist = parts[1].strip() if len(parts) > 1 else "Unknown"
            break

    return media_title, artist, album, pos, dur


# ---------------- MAIN LOOP ----------------
def run_osc_loop():
    try:
        prev = psutil.net_io_counters(pernic=True)[INTERFACE]
    except KeyError:
        print(f"Error: {INTERFACE} not found.")
        return

    prev_time = time.time()
    print("Sending live data to VRChat...")

    progress_counter = 0  # used to animate progress bar

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
        song, artist, album, pos_ms, dur_ms = get_current_media()
        display_artist = f" {artist}" if artist else " Unknown Artist"
        display_song = f"🎵 {song}" if song else "🎵 No Music Playing"

        # ---- Animate progress bar ----
        progress_counter = (progress_counter + 1) % dur_ms
        progress_percent = progress_counter / dur_ms
        progress_bar = create_progress_bar(progress_percent)

        # ---- Page switching ----
        page_index = int((now // SWITCH_INTERVAL) % 2)

        if page_index == 0:  # Page 1
            text = (
                "Im running this shit with python get gud bro\n"
                f"{time.strftime('%I:%M %p')}\n"
                f"Download: {down_speed}\n"
                f"Upload: {up_speed}\n"
                f"{progress_bar}\n"
                f"{display_song}-{display_artist}"
            )
        else:  # Page 2
            text = (
                "Blasting Music\n"
                f"Intel Core I7 8700 {cpu}%\n"
                f"AMD Radeon 9060 XT {gpu}%\n"
                f"{progress_bar}\n"
                f"{display_song}-{display_artist}"
            )

        # ---- Send OSC ----
        client.send_message("/chatbox/input", [text, True])

        prev = cur
        prev_time = now
        time.sleep(1.6)


# ---------------- RUN ----------------
if __name__ == "__main__":
    run_osc_loop()
