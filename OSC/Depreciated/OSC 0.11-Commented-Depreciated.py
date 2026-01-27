import time
import asyncio
import psutil
import subprocess
import GPUtil
from pythonosc.udp_client import SimpleUDPClient
import winrt.windows.media.control as wmc

# -------------------- CONFIG --------------------
OSC_IP = "127.0.0.1"
OSC_PORT = 9000
INTERFACE = "Ethernet"
SWITCH_INTERVAL = 30

client = SimpleUDPClient(OSC_IP, OSC_PORT)

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
            ["powershell", "-Command", ps_cmd], encoding="utf-8", stderr=subprocess.DEVNULL
        )
        values = [float(val) for val in result.strip().split('\n') if val.strip()]
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
        # ---- Media info ----
        song, artist, album, pos, dur = asyncio.run(get_media_info())

        # ---- System stats ----
        cpu = psutil.cpu_percent()
        gpu = get_amd_gpu_load()
        ram = psutil.virtual_memory()
        ram_text = f"RAM {ram.used // 1024**2}/{ram.total // 1024**2} MB"

        # VRAM usage using GPUtil
        try:
            gpus = GPUtil.getGPUs()
            if gpus:
                gpu0 = gpus[0]
                vram_text = f"VRAM {int(gpu0.memoryUsed)}/{int(gpu0.memoryTotal)} MB"
            else:
                vram_text = "VRAM N/A"
        except Exception:
            vram_text = "VRAM N/A"

        # ---- Network ----
        now = time.time()
        cur = psutil.net_io_counters(pernic=True)[INTERFACE]
        elapsed = now - prev_time
        up_raw = (cur.bytes_sent - prev.bytes_sent) / elapsed
        down_raw = (cur.bytes_recv - prev.bytes_recv) / elapsed
        down_speed = fmt(down_raw)
        up_speed = fmt(up_raw)
        cur_time_str = time.strftime("%I:%M %p")

        display_artist = f" {artist}" if artist else " Unknown Artist"
        display_song = f"🎵 {song}" if song else "🎵 No Music Playing"
        progress_bar = create_progress_bar(pos, dur)

        # ---- Page switching ----
        page_index = int((now // SWITCH_INTERVAL) % 2)

        if page_index == 0:  # Page 1
            text = (
                "Im running this shit with python get gud bro\n"
                f"{cur_time_str}\n"
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
                f"{ram_text} | {vram_text}\n"
                f"{progress_bar}\n"
                f"{display_song}-{display_artist}"
            )

        client.send_message("/chatbox/input", [text, True])

        prev = cur
        prev_time = now
        time.sleep(1.6)


if __name__ == "__main__":
    run_osc_loop()

