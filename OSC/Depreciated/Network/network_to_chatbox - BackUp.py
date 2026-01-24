import time
import psutil
from pythonosc.udp_client import SimpleUDPClient

OSC_IP = "127.0.0.1"
OSC_PORT = 9005

client = SimpleUDPClient(OSC_IP, OSC_PORT)

prev = psutil.net_io_counters()
prev_time = time.time()

def format_speed(bps):
    if bps > 1024 * 1024:
        return f"{bps / (1024 * 1024):.2f} MB/s"
    return f"{bps / 1024:.1f} KB/s"

while True:
    time.sleep(1)

    now = time.time()
    current = psutil.net_io_counters()

    elapsed = now - prev_time
    up = (current.bytes_sent - prev.bytes_sent) / elapsed
    down = (current.bytes_recv - prev.bytes_recv) / elapsed

    text = f"  Upload {format_speed(up)}      Download {format_speed(down)}"

    client.send_message("/chatbox/input", [text, True])

    prev = current
    prev_time = now
