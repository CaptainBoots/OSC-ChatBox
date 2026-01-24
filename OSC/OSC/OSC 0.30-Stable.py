import time
import asyncio
import psutil
import subprocess
from pythonosc.udp_client import SimpleUDPClient
import winrt.windows.media.control as wmc

OSC_IP = "127.0.0.1" #Change to ip of device VRC is running on
OSC_PORT = 9000 #Change to the OSC port you want to use
INTERFACE = "Ethernet" #Chamge to interface you want to measure: Ethernet, "the name of your WI-FI" or any other interface
SWITCH_INTERVAL = 30 #Change to decide how often the page number changes

client = SimpleUDPClient(OSC_IP, OSC_PORT)

def fmt(bps): #Formats the Bits per second for the network reader
    if bps > 1024 * 1024:
        return f"{bps / (1024 * 1024):.2f} MB/s"
    return f"{bps / 1024:.1f} KB/s"

def get_gpu_load(): #Gets the usage of the gpu
    try:
        cmd = (
            'Get-Counter "\\GPU Engine(*3D*)\\Utilization Percentage" | '
            'Select-Object -ExpandProperty CounterSamples | '
            'Select-Object -ExpandProperty CookedValue'
        )
        result = subprocess.check_output(["powershell", "-Command", cmd], encoding='utf-8', stderr=subprocess.DEVNULL)
        values = [float(v) for v in result.strip().split('\n') if v.strip()]
        return int(max(values)) if values else 0
    except subprocess.CalledProcessError:
        return 0
    except ValueError:
        return 0

def create_progress_bar(position_ms, duration_ms, length=13): #Creates the media progress bar
    if duration_ms <= 0:
        return "─" * length
    percent = min(max(position_ms / duration_ms, 0), 1)
    filled_len = int(length * percent)
    return "■" * filled_len + "□" * (length - filled_len)

async def get_media_info(): #Gets the Media info for the media title and artist
    try:
        manager = await wmc.GlobalSystemMediaTransportControlsSessionManager.request_async()
        session = manager.get_current_session()
        if session:
            props = await session.try_get_media_properties_async()
            timeline = session.get_timeline_properties()
            pos = timeline.position.total_seconds() * 1000
            dur = timeline.end_time.total_seconds() * 1000
            return props.title, props.artist, props.album_title, pos, dur
    except (AttributeError, TypeError):
        pass
    return None, None, None, 0, 0

import re

def clean_title(raw_title, artist=None): #Cleans the media title
    if not raw_title:
        return ""

    title = raw_title

    title = re.sub(r"\(.*?\)|\[.*?]|\{.*?}", "", title)

    junk_words = [     #Add words you want to be removed from the titles
        "official", "video", "lyrics", "audio", "hd", "4k", "remastered",
        "live", "visualizer", "explicit", "clean", "version", "mix"
    ]
    pattern = r"\b(" + "|".join(junk_words) + r")\b"
    title = re.sub(pattern, "", title, flags=re.IGNORECASE)

    title = re.sub(r"\b(ft\.|feat\.|featuring).*", "", title, flags=re.IGNORECASE)

    parts = [p.strip() for p in re.split(r"[-–|•]", title) if len(p.strip()) > 2]

    if len(parts) >= 2:
        if artist and artist.lower() in parts[0].lower():
            title = parts[1]
        else:
            title = parts[0]
    elif parts:
        title = parts[0]

    title = re.sub(r"\s+", " ", title).strip()

    return title


def get_network_usage(prev, prev_time): #gets the interface network data
    now = time.time()
    try:
        cur = psutil.net_io_counters(pernic=True)[INTERFACE]
        elapsed = now - prev_time
        up = (cur.bytes_sent - prev.bytes_sent) / elapsed
        down = (cur.bytes_recv - prev.bytes_recv) / elapsed
        return cur, up, down, now
    except KeyError:
        return prev, 0, 0, now

def run_osc_loop():
    all_stats = psutil.net_io_counters(pernic=True)
    if INTERFACE not in all_stats:
        print(f"Error: {INTERFACE} not found. Available: {list(all_stats.keys())}")
        return

    prev = all_stats[INTERFACE]
    prev_time = time.time()
    print("Sending live data to VRChat...")

    while True:
        song, artist, _, pos, dur = asyncio.run(get_media_info())
        clean_song = clean_title(song, artist)

        cpu = psutil.cpu_percent()
        gpu = get_gpu_load()
        prev, up_raw, down_raw, prev_time = get_network_usage(prev, prev_time)

        cur_time_str = time.strftime("%I:%M %p")
        progress_bar = create_progress_bar(pos, dur)
        display_artist = f"-{artist}" if artist else ""
        display_song = f"🎵 {clean_song}" if clean_song else ""

        page_index = int((prev_time // SWITCH_INTERVAL) % 2)

        if page_index == 0: #Page 1
            text = (
                f"Im running this shit with python\n" #Edit this text for the top status on page
                f"{cur_time_str}\n" #The time
                f"Download {fmt(down_raw)}\n" #Download usage
                f"Upload {fmt(up_raw)}\n" #Upload usage
                f"{progress_bar}\n" #Media progress bar
                f"{display_song} {display_artist}" #Song and artist
            )
        else:
            text = ( #Page 2
                f"Blasting Music\n" #Edit this text for the top status on page 2
                f"{cur_time_str}\n" #The time
                f"Intel Core I7 8700 {cpu}%\n" #Cpu usage
                f"AMD Radeon 9060 XT {gpu}%\n" #Gpu usage
                f"{progress_bar}\n" #Media progress bar
                f"{display_song} {display_artist}" #Song and artist
            )

        client.send_message("/chatbox/input", [text, True])
        time.sleep(1.6)

if __name__ == "__main__":
    run_osc_loop()