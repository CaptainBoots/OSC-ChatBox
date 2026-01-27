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
        self.refresh_rate = 90  # default fallback
        self.connected = False

    def get_steamvr_fps(self):
        try:
            if not self.connected:
                if time.time() > 0:
                    # Initialize VR
                    openvr.init(openvr.VRApplication_Other)
                    self.vr_system = openvr.VRSystem()
                    self.connected = True
                    print("Connected to SteamVR!")

                    # Try to read recommended refresh rate from HMD
                    try:
                        self.refresh_rate = int(self.vr_system.getFloatTrackedDeviceProperty(
                            openvr.k_unTrackedDeviceIndex_Hmd,
                            openvr.Prop_DisplayFrequency_Float
                        ))
                    except Exception:
                        pass


            # This field exists on ALL OpenVR builds
            frame_count = timing.m_nNumFramePresents

            now = time.time()

            if self.last_frame is None:
                self.last_frame = frame_count
                self.last_time = now
                return 0

            frames = frame_count - self.last_frame
            dt = now - self.last_time

            self.last_frame = frame_count
            self.last_time = now

            if dt > 0 and frames >= 0:
                fps = int(frames / dt)
                return min(max(fps, 1), 300)

            return 0

        except Exception as e:
            print("SteamVR error:", e)
            self.connected = False
            self.reconnect_timer = time.time() + 5
            return 0


vr_stats = SteamVRTracker()


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
                f"FPS {vr_fps}\n"
                f"{progress_bar}\n"
                f"{display_song}-{display_artist}"
            )

        client.send_message("/chatbox/input", [text, True])

        prev = cur_net
        prev_time = now
        time.sleep(1.6)


if __name__ == "__main__":
    run_osc_loop()
