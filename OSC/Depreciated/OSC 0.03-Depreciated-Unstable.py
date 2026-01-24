import time
import asyncio
import psutil
from pythonosc.udp_client import SimpleUDPClient
import winrt.windows.media.control as wmc

# --- Configuration ---
OSC_IP = "127.0.0.1"
OSC_PORT = 9000
INTERFACE = "Ethernet"
SWITCH_INTERVAL = 30

client = SimpleUDPClient(OSC_IP, OSC_PORT)


def fmt(bps):
    if bps > 1024 * 1024:
        return f"{bps / (1024 * 1024):.2f} MB/s"
    return f"{bps / 1024:.1f} KB/s"


def create_progress_bar(position_ms, duration_ms, length=13):
    """Generates a visual progress bar string"""
    if duration_ms <= 0:
        return "──" * length

    percent = min(max(position_ms / duration_ms, 0), 1)
    filled_len = int(length * percent)
    # Using '■' and ' ' for a clean look
    bar = "■" * filled_len + "□" * (length - filled_len)
    return f"{bar}"


async def get_media_info():
    """Grabs Title, Artist, and Progress from Windows Media"""
    try:
        manager = await wmc.GlobalSystemMediaTransportControlsSessionManager.request_async()
        session = manager.get_current_session()
        if session:
            props = await session.try_get_media_properties_async()
            timeline = session.get_timeline_properties()

            # Convert Windows times (100-nanosecond intervals) to milliseconds
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
        print(f"Error: {INTERFACE} not found.")
        return

    prev_time = time.time()
    print("Sending live data to VRChat...")

    while True:
        # 1. Get Metadata and Timeline
        song, artist, pos, dur = asyncio.run(get_media_info())

        # 2. Get Network Stats
        now = time.time()
        cur = psutil.net_io_counters(pernic=True)[INTERFACE]
        elapsed = now - prev_time
        up_speed = (cur.bytes_sent - prev.bytes_sent) / elapsed
        down_speed = (cur.bytes_recv - prev.bytes_recv) / elapsed

        # 3. Format the Strings
        display_artist = f" {artist}" if artist else " Unknown Artist"
        display_song = f"🎵 {song}" if song else "🎵 No Music Playing"
        down_speed = f"{fmt(down_speed)}" if song else "🎵 No Music Playing"
        up_speed = f"{fmt(up_speed)}" if song else "🎵 No Music Playing"
        progress_bar = create_progress_bar(pos, dur)

        page_index = int((now // SWITCH_INTERVAL) % 2)

        # 4. Construct the Final Chatbox Block
        text = (
            "im running this shit with python get gud bro\n"  # Line 1
            f"Download {down_speed}\n"  # Line 2
            f"Upload {up_speed}\n"  # Line 3
            f"\n"  # Line 4
            f"{progress_bar}\n"  # Line 5 
            f"{display_song}-{display_artist}"  # Line 6
        )

        if
        page_index == 0:
        text = (
            "im running this shit with python get gud bro\n"  # Line 1
            f"\n"  # Line 2
            f"Download {down_speed}\n"  # Line 3
            f"Upload {up_speed}\n"  # Line 4
            f"{progress_bar}\n"  # Line 5 
            f"{display_song}-{display_artist}"  # Line 6
        )
    else:
        text = (
            "im running this shit with python get gud bro\n"  # Line 1
            f"\n"  # Line 2
            f"System Load: {gpu}%\n"  # Line 3
            f""System Load: {cpu}%\n"  # Line 4
            f"{progress_bar}\n"  # Line 5 
            f"{display_song}-{display_artist}"  # Line 6
        )

        # 5. Send to OSC
        client.send_message("/chatbox/input", [text, True])

        # Reset loop
        prev = cur
        prev_time = now
        time.sleep(1.6)


if __name__ == "__main__":
    run_osc_loop()