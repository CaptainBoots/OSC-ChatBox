import time
import asyncio
import psutil
import subprocess
import openvr
import GPUtil
from pythonosc.udp_client import SimpleUDPClient
import winrt.windows.media.control as wmc

OSC_IP = "127.0.0.1"
OSC_PORT = 9000
INTERFACE = "Ethernet"
SWITCH_INTERVAL = 30

client = SimpleUDPClient(OSC_IP, OSC_PORT)

print("This code is depreciated use at your own risk")

def fmt(bps):
    if bps > 1024 * 1024:
        return f"{bps / (1024 * 1024):.2f} MB/s"
    return f"{bps / 1024:.1f} KB/s"


def get_amd_gpu_load():
    try:
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
        now = time.time()
        cur = psutil.net_io_counters(pernic=True)[INTERFACE]
        elapsed = now - prev_time
        up_raw = (cur.bytes_sent - prev.bytes_sent) / elapsed
        down_raw = (cur.bytes_recv - prev.bytes_recv) / elapsed
        cur_time_str = time.strftime("%I:%M %p")

        display_artist = f" {artist}" if artist else " Unknown Artist"
        display_song = f"🎵 {song}" if song else "🎵 No Music Playing"
        down_speed = fmt(down_raw)
        up_speed = fmt(up_raw)
        progress_bar = create_progress_bar(pos, dur)

        # RAM info
        ram = psutil.virtual_memory()
        ram_text = f"RAM {ram.used // 1024**2}/{ram.total // 1024**2} MB"

        # VRAM info (Windows only, via dxdiag)
        try:
            import wmi
            w = wmi.WMI()
            vram_text = ""
            for gpu_info in w.Win32_VideoController():
                if gpu_info.AdapterRAM:
                    vram_used = gpu_info.AdapterRAM // 1024**2
                    vram_total = gpu_info.AdapterRAM // 1024**2  # approximate total = max RAM reported
                    vram_text = f"VRAM {vram_used}/{vram_total} MB"
        except Exception:
            vram_text = "VRAM N/A"

        page_index = int((now // SWITCH_INTERVAL) % 2)

        if page_index == 0:  # page 1
            text = (
                "Im running this shit with python get gud bro\n"
                f"{cur_time_str}\n" 
                f"Download: {down_speed}\n"  
                f"Upload: {up_speed}\n"
                f"{progress_bar}\n"
                f"{display_song}-{display_artist}"
            )
        else:  # page 2
            text = (
                "Blasting Music\n"
                f"Intel Core I7 8700 {cpu}%\n"
                f"AMD Radeon 9060 XT {gpu}%\n"
                f"{ram_text}\n"
                f"{progress_bar}\n"
                f"{display_song}-{display_artist}"
            )

        client.send_message("/chatbox/input", [text, True])

        prev = cur
        prev_time = now
        time.sleep(1.6)


if __name__ == "__main__":
    run_osc_loop()