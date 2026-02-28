# ═════════════════════════════════════════════════════════════════════════════════════════════════════════════════════#
#                                              OSC Python Script                                                      #
# ═════════════════════════════════════════════════════════════════════════════════════════════════════════════════════#
# Hi :3
# Wellcome to my code

# ═════════════════════════════════════════════════════════════════════════════════════════════════════════════════════#
# Imports
# ═════════════════════════════════════════════════════════════════════════════════════════════════════════════════════#

import time
import threading
import psutil
from pythonosc.udp_client import SimpleUDPClient

from kivy.app import App
from kivy.clock import Clock
from kivy.uix.boxlayout import BoxLayout
from kivy.uix.button import Button
from kivy.uix.label import Label
from kivy.uix.textinput import TextInput

# ═════════════════════════════════════════════════════════════════════════════════════════════════════════════════════#
# CONFIGURATION & GLOBAL VARIABLES
# ═════════════════════════════════════════════════════════════════════════════════════════════════════════════════════#

SWITCH_INTERVAL = 30
INTERFACE = "wlan0"  # Android default
running = False
client = None

page1_text = ""
page2_text = ""

# ═════════════════════════════════════════════════════════════════════════════════════════════════════════════════════#
# NETWORK MONITORING
# ═════════════════════════════════════════════════════════════════════════════════════════════════════════════════════#

def fmt(bps):
    if bps >= 1_000_000:
        return f"{bps / 1_000_000:.2f} Mb/s"
    return f"{bps / 1_000:.1f} Kb/s"


def get_network(prev, prev_time):
    now = time.time()
    cur = psutil.net_io_counters(pernic=True).get(INTERFACE)
    if not cur:
        return prev, 0, 0, now

    dt = now - prev_time
    up = (cur.bytes_sent - prev.bytes_sent) / dt if dt else 0
    down = (cur.bytes_recv - prev.bytes_recv) / dt if dt else 0
    return cur, up, down, now

# ═════════════════════════════════════════════════════════════════════════════════════════════════════════════════════#
# MEDIA MONITORING
# ═════════════════════════════════════════════════════════════════════════════════════════════════════════════════════#

def get_media_info():
    # ANDROID PLACEHOLDER
    return "No Media", "", 0, 0


# ═════════════════════════════════════════════════════════════════════════════════════════════════════════════════════#
# MAIN OSC LOOP
# ═════════════════════════════════════════════════════════════════════════════════════════════════════════════════════#

def osc_loop():
    global running

    stats = psutil.net_io_counters(pernic=True)
    prev = stats.get(INTERFACE)
    prev_time = time.time()

    while running:
        song, artist, pos, dur = get_media_info()
        prev, up, down, prev_time = get_network(prev, prev_time)

        page = int((time.time() // SWITCH_INTERVAL) % 2)
        text = page1_text if page == 0 else page2_text

        msg = (
            f"{text}\n"
            f"Download {fmt(down)}\n"
            f"Upload {fmt(up)}\n"
            f"{song}"
        )

        try:
            client.send_message("/chatbox/input", [msg, True])
        except Exception as e:
            print("OSC error:", e)

        time.sleep(5)


# ═════════════════════════════════════════════════════════════════════════════════════════════════════════════════════#
# UI
# ═════════════════════════════════════════════════════════════════════════════════════════════════════════════════════#

class MainUI(BoxLayout):
    def __init__(self, **kwargs):
        super().__init__(orientation="vertical", spacing=10, padding=10)

        self.ip = TextInput(text="192.168.0.10", hint_text="OSC IP")
        self.port = TextInput(text="9000", hint_text="OSC Port", input_filter="int")
        self.page1 = TextInput(text="Page 1 Text")
        self.page2 = TextInput(text="Page 2 Text")

        self.status = Label(text="Stopped")

        start = Button(text="Start")
        stop = Button(text="Stop")

        start.bind(on_press=self.start)
        stop.bind(on_press=self.stop)

        for w in (self.ip, self.port, self.page1, self.page2, start, stop, self.status):
            self.add_widget(w)

    def start(self, *_):
        global running, client, page1_text, page2_text

        if running:
            return

        page1_text = self.page1.text
        page2_text = self.page2.text

        client = SimpleUDPClient(self.ip.text, int(self.port.text))
        running = True

        threading.Thread(target=osc_loop, daemon=True).start()
        self.status.text = "Running"

    def stop(self, *_):
        global running
        running = False
        self.status.text = "Stopped"


class OSCApp(App):
    def build(self):
        return MainUI()


if __name__ == "__main__":
    OSCApp().run()
