import time
import asyncio
import psutil
import subprocess
from pythonosc.udp_client import SimpleUDPClient
import winrt.windows.media.control as wmc

# Configuration
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
        # Pulls the 3D Engine load via PowerShell
        ps_cmd = 'Get-Counter "\\GPU Engine(*3D*)\\Utilization Percentage" | Select-Object -ExpandProperty CounterSamples | Select-Object -ExpandProperty CookedValue'
        result = subprocess.check_output(["powershell", "-Command", ps_cmd], encoding='utf-8',
                                         stderr=subprocess.DEVNULL)
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
    return f"{bar}"


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


def run_osc_loop():
    try:
        # Verify interface exists
        all_stats = psutil.net_io_counters(pernic=True)
        if INTERFACE not in all_stats:
            print(f"Error: {INTERFACE} not found. Available: {list(all_stats.keys())}")
            return

        prev = all_stats[INTERFACE]
    except Exception as e:
        print(f"Startup error: {e}")
        return

    prev_time = time.time()
    print("Sending live data to VRChat...")

    while True:
        # Handle async media gathering
        song, artist, album, pos, dur = asyncio.run(get_media_info())

        # System Stats
        cpu = psutil.cpu_percent()
        gpu = get_amd_gpu_load()
        ram = psutil.virtual_memory()

        # Networking
        now = time.time()
        try:
            cur = psutil.net_io_counters(pernic=True)[INTERFACE]
            elapsed = now - prev_time
            up_raw = (cur.bytes_sent - prev.bytes_sent) / elapsed
            down_raw = (cur.bytes_recv - prev.bytes_recv) / elapsed
        except Exception:
            up_raw, down_raw = 0, 0
            cur = prev

        cur_time_str = time.strftime("%I:%M %p")
        display_artist = f" -{artist}" if artist else "aaaaaaaaaaaaaaaaaaah"
        display_song = f"🎵 {song}" if song else ""
        progress_bar = create_progress_bar(pos, dur)

        page_index = int((now // SWITCH_INTERVAL) % 2)

        if page_index == 0:  # Page 1: Network & Clock
            text = (
                f"Im running this shit with python\n"
                f"{cur_time_str}\n"
                f"Download {fmt(down_raw)}\n"
                f"Upload {fmt(up_raw)}\n"
                f"{progress_bar}\n"
                f"{display_song}{display_artist}"
            )
        else:  # Page 2: Hardware Stats
            text = (
                f"Blasting Music\n"
                f"{cur_time_str}\n"
                f"Intel Core I7 8700 {cpu}%\n"
                f"AMD Radeon 9060 XT {gpu}%\n"
                f"{progress_bar}\n"
                f"{display_song}-{display_artist}"
            )

        # Send to VRChat
        client.send_message("/chatbox/input", [text, True])

        prev = cur
        prev_time = now
        time.sleep(1.6)


if __name__ == "__main__":
    run_osc_loop()