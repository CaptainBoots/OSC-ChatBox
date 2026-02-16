import time
import psutil
from pythonosc.udp_client import SimpleUDPClient

OSC_IP = "127.0.0.1"
OSC_PORT = 9000
INTERFACE = "Ethernet" 

client = SimpleUDPClient(OSC_IP, OSC_PORT)

def fmt(bps):
    if bps > 1024 * 1024:
        return f"{bps / (1024 * 1024):.2f} MB/s"
    return f"{bps / 1024:.1f} KB/s"

prev = psutil.net_io_counters(pernic=True)[INTERFACE]
prev_time = time.time()

while True:
    time.sleep(1)

    now = time.time()
    cur = psutil.net_io_counters(pernic=True)[INTERFACE]

    elapsed = now - prev_time
    up = (cur.bytes_sent - prev.bytes_sent) / elapsed
    down = (cur.bytes_recv - prev.bytes_recv) / elapsed

    text = f"Download {fmt(down)}           Upload {fmt(up)}"
    client.send_message("/chatbox/input", [text, True])

    prev = cur
    prev_time = now

