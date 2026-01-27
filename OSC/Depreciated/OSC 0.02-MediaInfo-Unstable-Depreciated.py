import time
import asyncio
import psutil
from pythonosc.udp_client import SimpleUDPClient

# The new modular imports
import winrt.windows.media.control as wmc

# --- Configuration ---
OSC_IP = "127.0.0.1"
OSC_PORT = 9000
INTERFACE = "Ethernet"

print("This code is depreciated use at your own risk")

client = SimpleUDPClient(OSC_IP, OSC_PORT)


def fmt(bps):
    if bps > 1024 * 1024:
        return f"{bps / (1024 * 1024):.2f} MB/s"
    return f"{bps / 1024:.1f} KB/s"


def create_progress_bar(position_ms, duration_ms, length=15):
    if duration_ms <= 0:
        return "──" * length

    percent = min(max(position_ms / duration_ms, 0), 1)
    filled_len = int(length * percent)
    bar = "■" * filled_len + " " * (length - filled_len)
    return f"[{bar}]"


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
    except Exception:
        pass
    return None, None, 0, 0


def run_osc_loop():
    try:
        prev = psutil.net_io_counters(pernic=True)[INTERFACE]
    except KeyError:
        print(f"Available interfaces: {list(psutil.net_io_counters(pernic=True).keys())}")
        return

    prev_time = time.time()
    print("Script started! Check VRChat chatbox...")

    while True:
        # 1. Get Metadata (using the async function)
        song, artist = asyncio.run(get_media_info())

        # 2. Get Network Stats
        now = time.time()
        cur = psutil.net_io_counters(pernic=True)[INTERFACE]
        elapsed = now - prev_time

        up = (cur.bytes_sent - prev.bytes_sent) / elapsed
        down = (cur.bytes_recv - prev.bytes_recv) / elapsed

        # 3. Format the lines
        artist = f" {artist}" if artist else "🎤 Artist: Unknown"
        down = f"Download: {fmt(down)}"
        up = f"Upload: {fmt(up)}"
        song = f"🎵 {song}" if song else "🎵 No Song Playing"

        # 4. Construct the block
        # VRChat Chatbox layout (6 lines total)
        text = (
            "im running this shit with python get gud bro\n"
            f"{down}\n"
            f"{up}\n"
            f"{progress_bar}\n"  # Empty line 4
            f"\n"  # Empty line 5
            f"{song}-{artist}"
        )

        # 5. Send it
        client.send_message("/chatbox/input", [text, True])

        # Reset for next interval
        prev = cur
        prev_time = now
        time.sleep(1.6)


if __name__ == "__main__":
    run_osc_loop()
