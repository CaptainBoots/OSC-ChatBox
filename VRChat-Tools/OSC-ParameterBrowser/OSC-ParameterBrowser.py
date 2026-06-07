import tkinter as tk
from tkinter import ttk, messagebox
import threading
import queue
import socket
import time
import struct

try:
    from pythonosc import udp_client

    PYTHON_OSC = True
except ImportError:
    PYTHON_OSC = False

VERSION = "0.0.1"

print("OSC Parameter Browser")
print("Made By Boots")
print(f"Version {VERSION}")

BG = "#0f0f13"
PANEL = "#17171f"
PANEL2 = "#1f1f2b"
BORDER = "#2a2a3d"
TEXT = "#e8e4ff"
SUBTEXT = "#8878cc"
ACCENT = "#4cf5ff"
ACCENT2 = "#5c3fd4"
ACCENT3 = "#a87fff"
GREEN = "#4cff91"
RED = "#ff4c6a"
YELLOW = "#ffd166"
CYAN = "#4cf5ff"
BTN_BG = "#252535"

UI_FONT = "Segoe UI"
FONT_MONO = ("Consolas", 10)
FONT_SMALL = (UI_FONT, 9)
FONT_LABEL = (UI_FONT, 10, "bold")
FONT_BIG = (UI_FONT, 14, "bold")


def make_client(ip, port):
    if not PYTHON_OSC: return None
    try:
        return udp_client.SimpleUDPClient(ip, port)
    except:
        return None


def send_osc(c, addr, val):
    if c is None: return False
    try:
        c.send_message(addr, val); return True
    except:
        return False


def parse_osc(data):
    try:
        end = data.index(b'\x00')
        addr = data[:end].decode("utf-8", "replace")
        ap = (end + 4) & ~3
        rest = data[ap:]
        if not rest or rest[0:1] != b',': return None
        te = rest.index(b'\x00', 1)
        tags = rest[1:te].decode("ascii", "replace")
        tp = (te + 4) & ~3
        rest = rest[tp:]
        tag = tags[0] if tags else "?"
        if tag == 'f':
            val = round(struct.unpack_from(">f", rest)[0], 5)
        elif tag == 'i':
            val = struct.unpack_from(">i", rest)[0]
        elif tag == 'T':
            val = True
        elif tag == 'F':
            val = False
        elif tag == 's':
            se = rest.index(b'\x00');
            val = rest[:se].decode("utf-8", "replace")
        else:
            val = None
        tn = {"f": "float", "i": "int", "T": "bool", "F": "bool", "s": "string"}.get(tag, "?")
        return addr, tn, val
    except:
        return None


class Dot(tk.Canvas):
    def __init__(self, parent, **kw):
        super().__init__(parent, width=10, height=10, bg=parent["bg"], highlightthickness=0, **kw)
        self._s = "off";
        self._draw()

    def _draw(self):
        self.delete("all")
        c = {"on": GREEN, "off": RED, "warn": YELLOW}.get(self._s, SUBTEXT)
        self.create_oval(1, 1, 9, 9, fill=c, outline="")

    def set(self, s): self._s = s; self._draw()


class OSCBrowserApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("VRChat OSC Parameter Browser")
        self.configure(bg=BG)
        self.geometry("850x680")

        self._run = False
        self._params = {}
        self._q = queue.Queue()

        self._build_ui()
        self._poll()

        if not PYTHON_OSC:
            messagebox.showwarning("Missing Dependency",
                                   "python-osc is not installed.\nSend functions will be unavailable.\nRun: pip install python-osc")

    def _build_ui(self):
        # ─── TOP BAR ───────────────────────────────────────────────────────
        top_bar = tk.Frame(self, bg=PANEL, height=48)
        top_bar.pack(fill="x", side="top")
        top_bar.pack_propagate(False)

        lbl_title = tk.Label(top_bar, text="🔍 OSC Parameter Browser", font=FONT_BIG, fg=TEXT, bg=PANEL)
        lbl_title.pack(side="left", padx=12, pady=8)

        lbl_ver = tk.Label(top_bar, text="v0.0.1", font=(UI_FONT, 9, "italic"), fg=SUBTEXT, bg=PANEL)
        lbl_ver.pack(side="left", padx=(0, 12), pady=(14, 0))

        # ─── CONFIGURATION NETWORK ROW ─────────────────────────────────────
        cfg_row = tk.Frame(self, bg=BG)
        cfg_row.pack(fill="x", padx=12, pady=(12, 4))

        tk.Label(cfg_row, text="Target IP:", font=FONT_SMALL, fg=SUBTEXT, bg=BG).pack(side="left")
        self._ip_var = tk.StringVar(value="127.0.0.1")
        tk.Entry(cfg_row, textvariable=self._ip_var, font=FONT_MONO, bg=BG, fg=TEXT, width=12, relief="flat",
                 highlightthickness=1, highlightbackground=BORDER, highlightcolor=ACCENT).pack(side="left", padx=4,
                                                                                               ipady=1)

        tk.Label(cfg_row, text="Send Port:", font=FONT_SMALL, fg=SUBTEXT, bg=BG).pack(side="left", padx=(6, 0))
        self._sport_var = tk.StringVar(value="9000")
        tk.Entry(cfg_row, textvariable=self._sport_var, font=FONT_MONO, bg=BG, fg=TEXT, width=6, relief="flat",
                 highlightthickness=1, highlightbackground=BORDER, highlightcolor=ACCENT).pack(side="left", padx=4,
                                                                                               ipady=1)

        tk.Label(cfg_row, text="Listen Port:", font=FONT_SMALL, fg=SUBTEXT, bg=BG).pack(side="left", padx=(6, 0))
        self._lport_var = tk.StringVar(value="9001")
        tk.Entry(cfg_row, textvariable=self._lport_var, font=FONT_MONO, bg=BG, fg=TEXT, width=6, relief="flat",
                 highlightthickness=1, highlightbackground=BORDER, highlightcolor=ACCENT).pack(side="left", padx=4,
                                                                                               ipady=1)

        self._dot = Dot(cfg_row)
        self._dot.pack(side="left", padx=(12, 6))

        tk.Button(cfg_row, text="Listen", command=self._start, bg=GREEN, fg=BG, activebackground=GREEN, relief="flat",
                  font=FONT_SMALL, cursor="hand2", padx=8).pack(side="left", padx=2)
        tk.Button(cfg_row, text="Stop", command=self._stop, bg=BTN_BG, fg=TEXT, activebackground=BTN_BG, relief="flat",
                  font=FONT_SMALL, cursor="hand2", padx=8).pack(side="left", padx=2)
        tk.Button(cfg_row, text="Clear Data", command=self._clear, bg=PANEL, fg=SUBTEXT, activebackground=PANEL,
                  relief="flat", font=FONT_SMALL, cursor="hand2", padx=8).pack(side="left", padx=2)

        # ─── INJECTOR & FILTER CONTROLS BAR ───────────────────────────────
        controls_frame = tk.Frame(self, bg=PANEL, bd=1, relief="solid", highlightbackground=BORDER, pady=8, padx=8)
        controls_frame.pack(fill="x", padx=12, pady=6)

        # Filter Sub-Row
        fb = tk.Frame(controls_frame, bg=PANEL)
        fb.pack(fill="x", pady=2)
        tk.Label(fb, text="Filter Address: ", font=FONT_SMALL, fg=SUBTEXT, bg=PANEL).pack(side="left")
        self._fv = tk.StringVar();
        self._fv.trace_add("write", lambda *_: self._refresh())
        tk.Entry(fb, textvariable=self._fv, font=FONT_MONO, bg=BG, fg=TEXT, insertbackground=ACCENT, relief="flat",
                 highlightthickness=1, highlightbackground=BORDER, highlightcolor=ACCENT, width=40).pack(side="left",
                                                                                                         padx=4,
                                                                                                         ipady=2)

        # Injector Sub-Row
        sb = tk.Frame(controls_frame, bg=PANEL)
        sb.pack(fill="x", pady=(6, 2))
        tk.Label(sb, text="Inject Address: ", font=FONT_SMALL, fg=SUBTEXT, bg=PANEL).pack(side="left")
        self._av = tk.StringVar(value="/avatar/parameters/")
        tk.Entry(sb, textvariable=self._av, font=FONT_MONO, bg=BG, fg=TEXT, insertbackground=ACCENT, relief="flat",
                 highlightthickness=1, highlightbackground=BORDER, highlightcolor=ACCENT, width=32).pack(side="left",
                                                                                                         padx=2,
                                                                                                         ipady=2)

        tk.Label(sb, text="Val:", font=FONT_SMALL, fg=SUBTEXT, bg=PANEL).pack(side="left", padx=(4, 0))
        self._vv = tk.StringVar(value="0")
        tk.Entry(sb, textvariable=self._vv, font=FONT_MONO, bg=BG, fg=TEXT, insertbackground=ACCENT, relief="flat",
                 highlightthickness=1, highlightbackground=BORDER, highlightcolor=ACCENT, width=8).pack(side="left",
                                                                                                        padx=2, ipady=2)

        self._tc = ttk.Combobox(sb, values=["float", "int", "bool", "string"], width=7, font=FONT_SMALL)
        self._tc.set("float");
        self._tc.pack(side="left", padx=4)
        tk.Button(sb, text="Send Packet", command=self._send, bg=ACCENT3, fg=TEXT, relief="flat", font=FONT_SMALL,
                  cursor="hand2", padx=8).pack(side="left", padx=4)

        # ─── DATA PARAMETERS TREEVIEW ──────────────────────────────────────
        style = ttk.Style();
        style.theme_use("default")
        style.configure("P.Treeview", background=PANEL, foreground=TEXT, fieldbackground=PANEL, rowheight=22,
                        font=FONT_SMALL)
        style.configure("P.Treeview.Heading", background=PANEL2, foreground=ACCENT, font=FONT_SMALL, borderwidth=0)
        style.map("P.Treeview", background=[("selected", ACCENT2)])

        tf = tk.Frame(self, bg=BG)
        tf.pack(fill="both", expand=True, padx=12, pady=6)

        self._tree = ttk.Treeview(tf, columns=("path", "value", "type", "ts"), show="headings", style="P.Treeview")
        for col, w in [("path", 360), ("value", 130), ("type", 70), ("ts", 90)]:
            self._tree.heading(col, text=col.title());
            self._tree.column(col, width=w, stretch=(col == "path"))

        vsb = ttk.Scrollbar(tf, orient="vertical", command=self._tree.yview)
        self._tree.configure(yscrollcommand=vsb.set)
        vsb.pack(side="right", fill="y")
        self._tree.pack(side="left", fill="both", expand=True)
        self._tree.bind("<Double-1>", lambda e: self._dclick())

        # ─── BOTTOM FOOTER BAR ─────────────────────────────────────────────
        footer_bar = tk.Frame(self, bg=PANEL, pady=6)
        footer_bar.pack(fill="x", side="bottom")

        self._status = tk.Label(footer_bar, text="Status: IDLE", font=FONT_SMALL, fg=SUBTEXT, bg=PANEL)
        self._status.pack(side="left", padx=12)

        # Unused structural action hooks requested matching core setup style
        tk.Button(footer_bar, text="❓", command=self._help_click, bg=PANEL, fg=SUBTEXT, activebackground=PANEL,
                  activeforeground=TEXT, relief="flat", font=(UI_FONT, 11), cursor="hand2").pack(side="right", padx=6)
        tk.Button(footer_bar, text="⚙", command=self._settings_click, bg=PANEL, fg=SUBTEXT, activebackground=PANEL,
                  activeforeground=TEXT, relief="flat", font=(UI_FONT, 11), cursor="hand2").pack(side="right", padx=6)

    def _dclick(self):
        sel = self._tree.selection()
        if sel:
            v = self._tree.item(sel[0], "values")
            self._av.set(v[0])
            self._vv.set(str(v[1]))
            self._tc.set(v[2])

    def _start(self):
        if self._run: return
        try:
            port = int(self._lport_var.get())
        except:
            messagebox.showerror("Error", "Invalid Listen Port."); return

        self._run = True;
        self._dot.set("on")
        threading.Thread(target=self._loop, args=(port,), daemon=True).start()
        self._status.config(text=f"Status: Listening on port {port}...", fg=GREEN)

    def _stop(self):
        self._run = False
        self._dot.set("off")
        self._status.config(text="Status: Stopped", fg=SUBTEXT)

    def _clear(self):
        self._params.clear()
        self._refresh()

    def _loop(self, port):
        sock = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
        sock.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEADDR, 1)
        sock.settimeout(0.5)
        try:
            sock.bind(("0.0.0.0", port))
        except Exception as e:
            self._q.put(("err", f"Bind failed: {e}"))
            self._run = False
            return

        while self._run:
            try:
                data, _ = sock.recvfrom(4096)
                res = parse_osc(data)
                if res:
                    addr, typ, val = res
                    self._q.put(("param", addr, val, typ, time.strftime("%H:%M:%S")))
            except socket.timeout:
                continue
            except:
                break
        sock.close()

    def _refresh(self):
        filt = self._fv.get().lower()
        self._tree.delete(*self._tree.get_children())
        for path, (val, typ, ts) in sorted(self._params.items()):
            if filt and filt not in path.lower(): continue
            self._tree.insert("", "end", values=(path, val, typ, ts))

    def _settings_click(self):
        messagebox.showinfo("Settings",
                            "OSC Browser settings are fully operational via top-level IP and port entries.\nGlobal configuration adjustments coming soon!")

    def _help_click(self):
        messagebox.showinfo("Help Support",
                            "OSC Parameter Browser v0.0.1\n\nDouble-click any detected incoming packet argument row inside the visual tree view to instantly capture and clone its address data path straight down into the manual inject buffer field.")

    def _poll(self):
        try:
            while True:
                item = self._q.get_nowait()
                if item[0] == "param":
                    _, path, val, typ, ts = item
                    self._params[path] = (val, typ, ts)
                    self._refresh()
                elif item[0] == "err":
                    self._status.config(text=item[1], fg=RED)
                    self._dot.set("off")
        except queue.Empty:
            pass
        self.after(100, self._poll)

    def _send(self):
        if not PYTHON_OSC:
            messagebox.showerror("Error", "python-osc package missing.");
            return

        ip = self._ip_var.get().strip()
        try:
            port = int(self._sport_var.get().strip())
        except:
            messagebox.showerror("Error", "Invalid Send Port."); return

        c = make_client(ip, port)
        addr = self._av.get().strip()
        raw = self._vv.get().strip()
        vt = self._tc.get()

        try:
            val = float(raw) if vt == "float" else int(raw) if vt == "int" else raw.lower() in ("true", "1",
                                                                                                "yes") if vt == "bool" else raw
        except:
            self._status.config(text="Status: Typing format value error!", fg=YELLOW);
            return

        if send_osc(c, addr, val):
            self._status.config(text=f"Status: Injected {addr} -> {val}", fg=CYAN)
        else:
            self._status.config(text="Status: Packet network injection failed", fg=RED)


if __name__ == "__main__":
    OSCBrowserApp().mainloop()