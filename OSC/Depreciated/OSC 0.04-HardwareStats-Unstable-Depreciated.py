import time
import asyncio
import psutil
import subprocess
import openvr
from pythonosc.udp_client import SimpleUDPClient
import winrt.windows.media.control as wmc

OSC_IP = "127.0.0.1"
OSC_PORT = 9000
INTERFACE = "Ethernet"
SWITCH_INTERVAL = 30

print("This code is depreciated use at your own risk")

client = SimpleUDPClient(OSC_IP, OSC_PORT)


class SteamVRTracker:
    def __init__(self):
        self.vr_system = None
        self.reconnect_timer = 0
        self.last_frame_count = None
        self.last_time = None
        self.debug_printed = False  # only print once

    def get_steamvr_fps(self):
        try:
            if not self.vr_system:
                if time.time() > self.reconnect_timer:
                    openvr.init(openvr.VRApplication_Other)
                    self.vr_system = True
                    print("Connected to SteamVR!")
                    self.last_frame_count = None
                    self.last_time = None
                    time.sleep(0.2)
                else:
                    return 0

            result = openvr.VRCompositor().getFrameTiming()

            # Unpack tuple if needed
            if isinstance(result, tuple):
                if len(result) == 2:
                    success, timing = result
                    if not success:
                        return 0
                else:
                    timing = result[0]
            else:
                timing = result

            # DEBUG: print available fields once
            if not self.debug_printed:
                print("Timing fields:", [a for a in dir(timing) if not a.startswith("_")])
                self.debug_printed = True

            # Try to find a frame counter field
            frame_count = None
            for attr in dir(timing):
                if "Frame" in attr and "Num" in attr:
                    frame_count = getattr(timing, attr)
                    break

            if frame_count is None:
                return 0

            now = time.time()

            if self.last_frame_count is None:
                self.last_frame_count = frame_count
                self.last_time = now
                return 0

            frames_passed = frame_count - self.last_frame_count
            time_passed = now - self.last_time

            self.last_frame_count = frame_count
            self.last_time = now

            if time_passed > 0 and frames_passed >= 0:
                fps = int(frames_passed / time_passed)
                return min(max(fps, 1), 300)

            return 0

        except Exception as e:
            print("SteamVR error:", e)
            self.vr_system = None
            self.reconnect_timer = time.time() + 5
            return 0

vr_stats = SteamVRTracker()


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
    print("Timing fields:", [a for a in dir(timing) if not a.startswith("_")])

    while True:
        song, artist, album, pos, dur = asyncio.run(get_media_info())
        cpu = psutil.cpu_percent()
        gpu = get_amd_gpu_load()
        vr_fps = vr_stats.get_steamvr_fps()

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
                f"Fps {vr_fps}\n"
                f"{progress_bar}\n"  
                f"{display_song}-{display_artist}"
            )

        client.send_message("/chatbox/input", [text, True])

        prev = cur_net
        prev_time = now
        time.sleep(1.6)


if __name__ == "__main__":
    run_osc_loop()