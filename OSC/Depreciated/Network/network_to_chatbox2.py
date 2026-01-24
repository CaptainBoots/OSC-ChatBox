import time
import psutil
from pythonosc.udp_client import SimpleUDPClient

OSC_IP = "127.0.0.1"
OSC_PORT = 9005
INTERFACE = "Ethernet"   # <<< CHANGE THIS if needed

client = SimpleUDPClient(OSC_IP, OSC_PORT)

prev = psutil.net_io_counters(pernic=True)[INTERFACE]
prev_time = time.time()

def format_speed(bps):
    if bps > 1024 * 1024:
        return f"{bps / (1024 * 1024):.2f} MB/s"
    return f"{bps / 1024:.1f} KB/s"

while True:
    time.sleep(1)

    now = time.time()
    current = psutil.net_io_counters(pernic=True)[INTERFACE]

    elapsed = now - prev_time
    up = (current.bytes_sent - prev.bytes_sent) / elapsed
    down = (current.bytes_recv - prev.bytes_recv) / elapsed

    text = f"Download {format_speed(down)}           Upload {format_speed(up)}"

    client.send_message("/chatbox/input", [text, True])

    prev = current
    prev_time = now
