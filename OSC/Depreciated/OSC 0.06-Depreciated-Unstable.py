import time
import asyncio
import psutil
import subprocess
import win32gui
import re
from pythonosc.udp_client import SimpleUDPClient
import winrt.windows.media.control as wmc

# -------------------- CONFIG --------------------
OSC_IP = "127.0.0.1"
OSC_PORT = 9000
INTERFACE = "Ethernet"
SWITCH_INTERVAL = 30

client = SimpleUDPClient(OSC_IP, OSC_PORT)

# -------------------- VRChat FPS MONITOR --------------------
class VRCFPSMonitor:
    def __init__(self):
        self.last_time = None
        self.last_title = ""
        self.fps = 0

    def get_vrc_fps(self):
        try:
            # Find VRChat window
            def enum_windows(hwnd, result):
                title = win32gui.GetWindowText(hwnd)
                if "VRChat" in title:
                    result.append(title)

            windows = []
            win32gui.EnumWindows(enum_windows, windows)
            if not windows:
                return 0  # VRChat not running

            title = windows[0]
            now = time.time()

            # Try to extract FPS from title "VRChat (FPS: 72)"
            m = re.search(r'FPS[: ]+(\d+)', title)
            if m:
                fps = int(m.group(1))
            else:
                # fallback estimate based on loop time
                if self.last_time is None:
                    fps = 0
                else:
                    dt = now - self.last_time
                    fps = int(1 / max(dt, 0.001))

            self.last_time = now
            self.last_title = title
            self.fps = fps
            return fps

        except Exception:
            return 0


vrc_fps_monitor = VRCFPSMonitor()

# -------------------- HELPERS --------------------
def fmt(bps):
    if bps > 1024 * 1024:
        return f"{bps / (1024 * 1024):.2f} MB/s"
    return f"{bps / 1024:.1f} KB/s"


def get_amd_gpu_load():
    try:
        ps_cmd = (
            'Get-Counter "\\GPU Engine(*3D*)\\Utilization Percentage" | '
            'Select-Object -ExpandProperty CounterSamples | '
            'Select-Object -ExpandProperty CookedValue'
        )
        result = subprocess.check_output(
            ["powershell", "-Command", ps_cmd],
            encoding="utf-8",
            stderr=subprocess.DEVNULL
        )
        values = [float(v) for v in result.strip().split("\n") if v.strip()]
        if values:
            return int(max(values))
    except Exception:
        pass
    return 0


def create_progress_bar(position_ms, duration_ms, length=13):
    if duration_ms <= 0:
        return "──" * length
    percent = min(max(position_ms / duration_ms, 0), 1)
    filled_len = int(length * percent)
    bar = "■" * filled_len + "□" * (length - filled_len)
    return bar


async def get_media_info():
    try:
        manager = await wmc.GlobalSystemMediaTransportControlsSessionManager.request_async()
        session = manager.get_current_session()
        if session:
            props = await session.try_get_media_properties_async()
            timeline = session.get_timeline_properties()
            pos = timeline.position.total_seconds() * 1000
            dur = timeline.end_time.total_seconds() * 1000
            return props.title, props.artist, props.album_title, pos, dur
    except Exception:
        pass
    return None, None, None, 0, 0

# -------------------- MAIN OSC LOOP --------------------
def run_osc_loop():
    try:
        prev = psutil.net_io_counters(pernic=True)[INTERFACE]
    except KeyError:
        print(f"Error: {INTERFACE} not found.")
        return

    prev_time = time.time()
    print("Sending live data to VRChat...")

    while True:
        song, artist, album, pos, dur = asyncio.run(get_media_info())
        cpu = psutil.cpu_percent()
        gpu = get_amd_gpu_load()
        vrc_fps = vrc_fps_monitor.get_vrc_fps()

        now = time.time()
        cur_time_str = time.strftime("%I:%M %p")

        cur_net = psutil.net_io_counters(pernic=True)[INTERFACE]
        elapsed = now - prev_time
        up_raw = (cur_net.bytes_sent - prev.bytes_sent) / elapsed
        down_raw = (cur_net.bytes_recv - prev.bytes_recv) / elapsed
        down_speed = fmt(down_raw)
        up_speed = fmt(up_raw)

        display_artist = f" {artist}" if artist else " Unknown Artist"
        display_song = f"🎵 {song}" if song else "🎵 No Music Playing"
        progress_bar = create_progress_bar(pos, dur)

        page_index = int((now // SWITCH_INTERVAL) % 2)

        if page_index == 0:
            text = (
                "Im running this shit with python get gud bro\n"
                f"{cur_time_str}\n"
                f"Download: {down_speed}\n"
                f"Upload: {up_speed}\n"
                f"{progress_bar}\n"
                f"{display_song}-{display_artist}"
            )
        else:
            text = (
                "Blasting Music\n"
                f"CPU {cpu}%\n"
                f"GPU {gpu}%\n"
                f"VRC FPS {vrc_fps}\n"
                f"{progress_bar}\n"
                f"{display_song}-{display_artist}"
            )

        client.send_message("/chatbox/input", [text, True])

        prev = cur_net
        prev_time = now
        time.sleep(1.6)


if __name__ == "__main__":
    run_osc_loop()
