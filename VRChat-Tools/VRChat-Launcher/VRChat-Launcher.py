
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import subprocess
import os

VERSION = "0.0.1"

print("VRChat Launcher")
print("Made By Boots")
print(f"Version {VERSION}")

BG = "#0f0f13"
PANEL = "#17171f"
PANEL2 = "#1f1f2b"
BORDER = "#2a2a3d"
TEXT = "#e8e4ff"
SUBTEXT = "#8878cc"
ACCENT = "#7c5cfc"
ACCENT2 = "#5c3fd4"
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

DEFAULT_LAUNCH_EXE = r"C:\Program Files (x86)\Steam\steamapps\common\VRChat\launch.exe"
PROFILE_COLORS = ["#7c5cfc", "#4cf5ff", "#4cff91", "#ffd166", "#ff4c6a", "#a87fff", "#ff9f45", "#ff6eb4"]

LIMIT_NOTE = (
    "VRChat limits 3 simultaneous instances per public IP address.\n"
    "This is enforced server-side and cannot be bypassed via launch args.\n\n"
    "Workarounds:\n"
    "  VPN per extra instance (each gets a different public IP)\n"
    "  Run extra instances on a different network/hotspot\n\n"
    "The limit is per public IP, not per machine."
)

def default_profile(idx):
    return {
        "name": f"Alt {idx}" if idx > 0 else "Main",
        "osc_ip": "127.0.0.1",
        "osc_port": 9000 + idx * 10,
        "listen_port": 9001 + idx * 10,
        "color": PROFILE_COLORS[idx % len(PROFILE_COLORS)],
        "exe_args": ""
    }

class Dot(tk.Canvas):
    def __init__(self, parent, **kw):
        super().__init__(parent, width=10, height=10, bg=parent["bg"], highlightthickness=0, **kw)
        self._s = "off"; self._draw()
    def _draw(self):
        self.delete("all")
        c = {"on": GREEN, "off": RED, "warn": YELLOW}.get(self._s, SUBTEXT)
        self.create_oval(1, 1, 9, 9, fill=c, outline="")
    def set(self, s): self._s = s; self._draw()

class VRCLauncherApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("VRChat Multi-Instance Launcher")
        self.configure(bg=BG)
        self.geometry("900x680")

        self.profiles = [default_profile(i) for i in range(3)]
        self._vprocs = {}
        self._rows = {}

        self._build_ui()
        self._poll()

    def _build_ui(self):
        # ─── TOP BAR ───────────────────────────────────────────────────────
        top_bar = tk.Frame(self, bg=PANEL, height=48)
        top_bar.pack(fill="x", side="top")
        top_bar.pack_propagate(False)

        lbl_title = tk.Label(top_bar, text="🚀 VRChat Launcher", font=FONT_BIG, fg=TEXT, bg=PANEL)
        lbl_title.pack(side="left", padx=12, pady=8)

        lbl_ver = tk.Label(top_bar, text="v0.0.1", font=(UI_FONT, 9, "italic"), fg=SUBTEXT, bg=PANEL)
        lbl_ver.pack(side="left", padx=(0, 12), pady=(14, 0))

        # ─── WARNING BANNER ────────────────────────────────────────────────
        wb = tk.Frame(self, bg="#1a1208")
        wb.pack(fill="x", padx=12, pady=(12, 4))
        tk.Label(wb, text="  ⚠  VRChat allows max 3 instances per public IP", font=FONT_SMALL, fg=YELLOW, bg="#1a1208").pack(side="left", padx=6, pady=6)
        tk.Button(wb, text="Why? / Workarounds", command=self._show_limit, bg=BTN_BG, fg=TEXT, relief="flat", font=FONT_SMALL, cursor="hand2").pack(side="right", padx=6, pady=4)

        # ─── MAIN CONTENT SPLIT ────────────────────────────────────────────
        main_split = tk.Frame(self, bg=BG)
        main_split.pack(fill="both", expand=True, padx=12, pady=6)

        # Left Column: Profiles List
        self._pf = tk.Frame(main_split, bg=BG)
        self._pf.pack(side="left", fill="both", expand=True, padx=(0, 6))

        # Right Column: Config Panel
        self._config_panel = tk.Frame(main_split, bg=PANEL, width=320, bd=1, relief="solid", highlightbackground=BORDER)
        self._config_panel.pack(side="right", fill="both", expand=False, padx=(6, 0))
        self._config_panel.pack_propagate(False)
        self._active_edit_idx = None
        self._hide_config_panel()

        self._rebuild_rows()

        # Add Profile Button
        ar = tk.Frame(self, bg=BG)
        ar.pack(fill="x", padx=12, pady=4)
        tk.Button(ar, text="+ Add Profile", command=self._add_profile, bg=ACCENT2, fg=TEXT, relief="flat", font=FONT_SMALL, cursor="hand2", padx=10).pack(side="left")

        # ─── EXE CONFIG PATH BLOCK ─────────────────────────────────────────
        tk.Frame(self, bg=BORDER, height=1).pack(fill="x", padx=12, pady=8)

        ep_lbl = tk.Label(self, text="LAUNCH EXE PATH", font=FONT_LABEL, fg=SUBTEXT, bg=BG)
        ep_lbl.pack(anchor="w", padx=12)

        ep = tk.Frame(self, bg=BG)
        ep.pack(fill="x", padx=12, pady=4)
        self._exe = tk.StringVar(value=DEFAULT_LAUNCH_EXE)
        tk.Entry(ep, textvariable=self._exe, font=FONT_MONO, bg=BG, fg=TEXT, insertbackground=ACCENT, relief="flat", highlightthickness=1, highlightbackground=BORDER, highlightcolor=ACCENT).pack(side="left", fill="x", expand=True, padx=(0, 6), ipady=3)
        tk.Button(ep, text="Browse", command=self._browse, bg=BTN_BG, fg=TEXT, relief="flat", font=FONT_SMALL, cursor="hand2", padx=10).pack(side="left")

        tk.Label(self, text="* Must use launch.exe — launching VRChat.exe directly forces offline test mode.", font=FONT_SMALL, fg=YELLOW, bg=BG).pack(anchor="w", padx=12, pady=(2, 12))

        # ─── BOTTOM FOOTER BAR ─────────────────────────────────────────────
        footer_bar = tk.Frame(self, bg=PANEL, pady=6)
        footer_bar.pack(fill="x", side="bottom")

        self.lbl_status = tk.Label(footer_bar, text="Ready", font=FONT_SMALL, fg=SUBTEXT, bg=PANEL)
        self.lbl_status.pack(side="left", padx=12)

        # Unused structural action hooks requested matching core setup style
        tk.Button(footer_bar, text="❓", command=self._help_click, bg=PANEL, fg=SUBTEXT, activebackground=PANEL, activeforeground=TEXT, relief="flat", font=(UI_FONT, 11), cursor="hand2").pack(side="right", padx=6)
        tk.Button(footer_bar, text="⚙", command=self._settings_click, bg=PANEL, fg=SUBTEXT, activebackground=PANEL, activeforeground=TEXT, relief="flat", font=(UI_FONT, 11), cursor="hand2").pack(side="right", padx=6)

    def _rebuild_rows(self):
        for w in self._pf.winfo_children(): w.destroy()
        self._rows.clear()
        for i, prof in enumerate(self.profiles): self._build_row(i, prof)

    def _build_row(self, idx, prof):
        color = prof.get("color", ACCENT)
        outer = tk.Frame(self._pf, bg=PANEL, bd=1, relief="solid", highlightbackground=BORDER)
        outer.pack(fill="x", pady=3)
        top = tk.Frame(outer, bg=PANEL)
        top.pack(fill="x", padx=8, pady=8)

        dot = Dot(top)
        dot.pack(side="left", padx=(0, 6))

        lbl = tk.Label(top, text=prof["name"], font=FONT_LABEL, fg=color, bg=PANEL, width=12, anchor="w", cursor="hand2")
        lbl.pack(side="left")
        lbl.bind("<Button-1>", lambda e, i=idx: self._show_config_panel(i))

        port_info = f"OSC: {prof['listen_port']} -> {prof['osc_port']}"
        tk.Label(top, text=port_info, font=FONT_SMALL, fg=SUBTEXT, bg=PANEL, width=20, anchor="w").pack(side="left")

        pid = tk.Label(top, text="Stopped", font=FONT_SMALL, fg=SUBTEXT, bg=PANEL, width=12, anchor="w")
        pid.pack(side="left")

        tk.Button(top, text="Launch", command=lambda i=idx: self._launch(i), bg=GREEN, fg=BG, activebackground=GREEN, relief="flat", font=FONT_SMALL, cursor="hand2", width=10).pack(side="left", padx=3)
        tk.Button(top, text="Kill", command=lambda i=idx: self._kill(i), bg=BTN_BG, fg=RED, activebackground=BTN_BG, relief="flat", font=FONT_SMALL, cursor="hand2", width=8).pack(side="left", padx=3)
        tk.Button(top, text="Remove", command=lambda i=idx: self._remove(i), bg=PANEL, fg=SUBTEXT, activebackground=PANEL, relief="flat", font=FONT_SMALL, cursor="hand2", width=8).pack(side="right", padx=3)

        self._rows[idx] = {"dot": dot, "pid": pid}

    def _show_config_panel(self, idx):
        self._active_edit_idx = idx
        prof = self.profiles[idx]

        for w in self._config_panel.winfo_children(): w.destroy()

        ch = tk.Frame(self._config_panel, bg=PANEL2, height=36)
        ch.pack(fill="x"); ch.pack_propagate(False)
        tk.Label(ch, text=f"⚙ Profile: {prof['name']}", font=FONT_LABEL, fg=prof["color"], bg=PANEL2).pack(side="left", padx=8)
        tk.Button(ch, text="✕", command=self._hide_config_panel, bg=PANEL2, fg=SUBTEXT, relief="flat", font=FONT_SMALL, cursor="hand2").pack(side="right", padx=6, pady=4)

        form = tk.Frame(self._config_panel, bg=PANEL, padx=10, pady=10)
        form.pack(fill="both", expand=True)

        tk.Label(form, text="Profile Name:", font=FONT_SMALL, fg=SUBTEXT, bg=PANEL).pack(anchor="w", pady=(4,0))
        name_var = tk.StringVar(value=prof["name"])
        tk.Entry(form, textvariable=name_var, font=FONT_MONO, bg=BG, fg=TEXT, relief="flat", highlightthickness=1, highlightbackground=BORDER, highlightcolor=prof["color"]).pack(fill="x", pady=2, ipady=2)

        tk.Label(form, text="UI Theme Color Hex:", font=FONT_SMALL, fg=SUBTEXT, bg=PANEL).pack(anchor="w", pady=(8,0))
        color_var = tk.StringVar(value=prof["color"])
        tk.Entry(form, textvariable=color_var, font=FONT_MONO, bg=BG, fg=TEXT, relief="flat", highlightthickness=1, highlightbackground=BORDER, highlightcolor=prof["color"]).pack(fill="x", pady=2, ipady=2)

        tk.Label(form, text="OSC Destination Port (VRC Input):", font=FONT_SMALL, fg=SUBTEXT, bg=PANEL).pack(anchor="w", pady=(8,0))
        os_var = tk.StringVar(value=str(prof["osc_port"]))
        tk.Entry(form, textvariable=os_var, font=FONT_MONO, bg=BG, fg=TEXT, relief="flat", highlightthickness=1, highlightbackground=BORDER, highlightcolor=prof["color"]).pack(fill="x", pady=2, ipady=2)

        tk.Label(form, text="OSC Source Bind Port (VRC Output):", font=FONT_SMALL, fg=SUBTEXT, bg=PANEL).pack(anchor="w", pady=(8,0))
        ol_var = tk.StringVar(value=str(prof["listen_port"]))
        tk.Entry(form, textvariable=ol_var, font=FONT_MONO, bg=BG, fg=TEXT, relief="flat", highlightthickness=1, highlightbackground=BORDER, highlightcolor=prof["color"]).pack(fill="x", pady=2, ipady=2)

        tk.Label(form, text="Custom Launch Args (Optional):", font=FONT_SMALL, fg=SUBTEXT, bg=PANEL).pack(anchor="w", pady=(8,0))
        args_var = tk.StringVar(value=prof["exe_args"])
        tk.Entry(form, textvariable=args_var, font=FONT_MONO, bg=BG, fg=TEXT, relief="flat", highlightthickness=1, highlightbackground=BORDER, highlightcolor=prof["color"]).pack(fill="x", pady=2, ipady=2)

        def save():
            try:
                prof["name"] = name_var.get().strip()
                prof["color"] = color_var.get().strip()
                prof["osc_port"] = int(os_var.get().strip())
                prof["listen_port"] = int(ol_var.get().strip())
                prof["exe_args"] = args_var.get().strip()
                self._rebuild_rows()
                self._show_config_panel(idx)
            except ValueError:
                messagebox.showerror("Error", "OSC ports must be numbers!")

        tk.Button(form, text="Save Changes", command=save, bg=ACCENT, fg=TEXT, relief="flat", font=FONT_SMALL, cursor="hand2", pady=4).pack(fill="x", pady=(16, 0))
        self._config_panel.pack(side="right", fill="both", expand=False, padx=(6,0))

    def _hide_config_panel(self):
        self._config_panel.pack_forget()
        self._active_edit_idx = None

    def _launch(self, idx):
        prof = self.profiles[idx]
        exe = self._exe.get().strip()
        if not os.path.exists(exe):
            messagebox.showwarning("Error", f"launch.exe not found:\n{exe}"); return
        cmd = f'"{exe}" --osc={prof["listen_port"]}:{prof["osc_port"]}'
        if prof.get("exe_args"): cmd += f' {prof["exe_args"]}'
        try:
            self._vprocs[idx] = subprocess.Popen(cmd, shell=True)
            self.lbl_status.config(text=f"Launched {prof['name']}", fg=GREEN)
        except Exception as e:
            messagebox.showerror("Error", f"Launch failed: {e}")

    def _kill(self, idx):
        p = self._vprocs.get(idx)
        if p and p.poll() is None:
            try: p.terminate()
            except: pass
        self._vprocs[idx] = None

    def _remove(self, idx):
        if len(self.profiles) <= 1:
            messagebox.showwarning("Can't Remove", "Need at least 1 profile."); return
        if messagebox.askyesno("Remove?", f"Remove '{self.profiles[idx]['name']}'?"):
            self._kill(idx)
            self.profiles.pop(idx)
            if self._active_edit_idx == idx: self._hide_config_panel()
            self._rebuild_rows()

    def _add_profile(self):
        idx = len(self.profiles)
        self.profiles.append(default_profile(idx))
        self._rebuild_rows()

    def _browse(self):
        p = filedialog.askopenfilename(title="Select launch.exe", filetypes=[("EXE", "*.exe"), ("All", "*.*")])
        if p: self._exe.set(p)

    def _show_limit(self):
        w = tk.Toplevel(self); w.title("Limit Info"); w.configure(bg=BG); w.resizable(False, False)
        tk.Label(w, text="VRChat 3-Instance Limit", font=FONT_LABEL, fg=YELLOW, bg=BG).pack(padx=20, pady=(16, 4))
        tk.Label(w, text=LIMIT_NOTE, font=FONT_SMALL, fg=TEXT, bg=BG, justify="left").pack(padx=20, pady=8)
        tk.Button(w, text="Close", command=w.destroy, bg=BTN_BG, fg=TEXT, relief="flat", font=FONT_SMALL, padx=12, pady=4).pack(pady=(4, 16))

    def _settings_click(self):
        messagebox.showinfo("Settings", "Launcher configuration profiles are fully customized inline inside the main panel list views.\nCustom system global toggles coming soon in future updates!")

    def _help_click(self):
        messagebox.showinfo("Help Support", "VRChat Launcher Tool v0.0.1\n\nClick on any profile name to modify its local dynamic ports or custom launch flags inside the editor block panels.")

    def _poll(self):
        for idx, proc in list(self._vprocs.items()):
            row = self._rows.get(idx)
            if not row: continue
            if proc and proc.poll() is None:
                row["dot"].set("on")
                row["pid"].config(text="Running", fg=GREEN)
            else:
                self._vprocs[idx] = None
                row["dot"].set("off")
                row["pid"].config(text="Stopped", fg=SUBTEXT)
        self.after(1000, self._poll)

if __name__ == "__main__":
    VRCLauncherApp().mainloop()