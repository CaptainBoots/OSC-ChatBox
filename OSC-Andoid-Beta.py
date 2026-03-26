# ════════════════════════#
# OSC Python Script #
# ════════════════════════#
# Hi :3
# Wellcome to my code


# ════════════════════════#
# Imports
# ════════════════════════#

import sys
import subprocess
import time
import threading
import importlib

def install_if_missing(package, import_name=None):
    if import_name is None:
        import_name = package.split("==")[0].replace("-", "_")

    try:
        importlib.import_module(import_name)
    except ImportError:
        print(f"Installing {package}...")
        subprocess.check_call([sys.executable, "-m", "pip", "install", package])

install_if_missing("python-osc==1.9.3", "pythonosc")
install_if_missing("psutil==7.2.2", "psutil")
install_if_missing("kivy==2.3.1", "kivy")

import psutil
from pythonosc.udp_client import SimpleUDPClient
from kivy.app import App
from kivy.uix.boxlayout import BoxLayout
from kivy.uix.button import Button
from kivy.uix.label import Label
from kivy.uix.textinput import TextInput

# ════════════════════════#
# CONFIGURATION & GLOBAL VARIABLES
# ════════════════════════#

SWITCH_INTERVAL = 20
INTERFACE = "wlan0"
running = False
client = None

page1_text = ""
page2_text = ""

# ════════════════════════#
# MAIN OSC LOOP
# ════════════════════════#

def osc_loop():
    global running

    try:
        stats = psutil.net_io_counters(pernic=True)
        prev = stats.get(INTERFACE)
    except Exception:
        prev = None  # Android fallback

    prev_time = time.time()

    while running:
        song, artist, pos, dur = get_media_info()
        prev, up, down, prev_time = get_network(prev, prev_time)

        page = int((time.time() // SWITCH_INTERVAL) % 2)
        text = page1_text if page == 0 else page2_text

        msg = {text}

        try:
            client.send_message("/chatbox/input", [msg, True])
            print("Sent:", msg)
        except Exception as e:
            print("OSC error:", e)

        time.sleep(5)

# ════════════════════════#
# UI
# ════════════════════════#

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