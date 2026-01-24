import time
import psutil
from pythonosc.udp_client import SimpleUDPClient

OSC_IP = "127.0.0.1"
OSC_PORT = 9005

STEAMVR_PROCESSES = [
    "vrserver.exe",
    "vrcompositor.exe"
]

client = SimpleUDPClient(OSC_IP, OSC_PORT)

def get_steamvr_bytes():
    sent = 0
    recv = 0

    for proc in psutil.process_iter(['name']):
        try:
            if proc.info['name'] in STEAMVR_PROCESSES:
                conns = proc.net_io_counters()
                sent += conns.bytes_sent
                recv += conns.bytes_recv
        except (psutil.NoSuchProcess, psutil.AccessDenied):
            pass

    return sent, recv


prev_sent, prev_recv = get_steamvr_bytes()
prev_time = time.time()

def fmt(bps):
    if bps > 1024 * 1024:
        return f"{bps / (1024 * 1024):.2f} MB/s"
    return f"{bps / 1024:.1f} KB/s"

while True:
    time.sleep(1)

    now = time.time()
    sent, recv = get_steamvr_bytes()
    elapsed = now - prev_time

    up = (sent - prev_sent) / elapsed
    down = (recv - prev_recv) / elapsed

    text = f"Download {fmt(down)}           Upload {fmt(up)}"

    client.send_message("/chatbox/input", [text, True])

    prev_sent, prev_recv = sent, recv
    prev_time = now

