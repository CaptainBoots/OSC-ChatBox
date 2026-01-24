import time
import psutil
import subprocess
from pythonosc.udp_client import SimpleUDPClient
import pygetwindow as gw  # pip install PyGetWindow

# ---------------- CONFIG ----------------
OSC_IP = "127.0.0.1"
OSC_PORT = 9000
INTERFACE = "Ethernet"
SWITCH_INTERVAL = 30

client = SimpleUDPClient(OSC_IP, OSC_PORT)

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
    Detects media from open windows. Returns:
    song, artist, progress_position_ms, duration_ms
    """
    windows = [w.title for w in gw.getWindowsWithTitle('') if w.title.strip()]
    for title in windows:
        # Spotify / VLC style "Song - Artist"
        if " - " in title:
            parts = title.split(" - ")
            song = parts[0].strip()
            artist = parts[1].strip() if len(parts) > 1 else "Unknown"
            # Fake progress for demonstration (as we can't read VLC/Spotify timing here)
            pos = 0
            dur = 180000  # 3 minutes in ms default
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
