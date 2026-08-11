"""
ui/chatbox_tab.py
─────────────────
Chatbox tab: live preview, start/stop/restart, config fields, forced text.

For flag themes, STRIPE_COLOURS drives repeating diagonal stripes drawn on a
canvas that fills the root window. Widgets sit in a normal Frame on top with
their own PANEL backgrounds so they remain fully readable. The stripe canvas
is managed by the App root window, not by this tab — this tab just uses BG
as its own background and the stripes show in the gaps between panels.
"""

import glob
import tkinter as tk
import webbrowser

from ui.theme import BG, PANEL, BORDER, ACCENT, ACCENT2, TEXT, SUBTEXT, GREEN, RED, FONT, STRIPE_COLOURS, draw_stripes

try:
    from PIL import Image, ImageTk
    _PIL_AVAILABLE = True
except ImportError:
    _PIL_AVAILABLE = False

# Folder scanned for banner PNGs (e.g. assets/banners/voicemod.png, assets/banners/ko-fi.png ...).
# Drop any PNGs in here and they'll be picked up automatically, sorted by filename.
BANNER_DIR = "assets/banners"
BANNER_HOLD_MS = 15000   # how long each banner stays on screen
BANNER_SLIDE_MS = 400    # total duration of the slide transition
BANNER_SLIDE_STEPS = 24  # animation smoothness

# The banner box itself is a fixed 8:3 rectangle (not stretched to fill the
# bottom bar) so it keeps a consistent shape and the stripe pattern shows
# around it instead of the box growing/shrinking with the window.
BANNER_ASPECT_W = 16
BANNER_ASPECT_H = 3
BANNER_HEIGHT   = 75
BANNER_WIDTH    = round(BANNER_HEIGHT * BANNER_ASPECT_W / BANNER_ASPECT_H)

ICON_SIZE = 52  # square size (px) both the Discord and GitHub buttons are scaled to


def _load_icon(path, size=ICON_SIZE):
    """Load a button icon resized to size x size, preserving transparency."""
    if _PIL_AVAILABLE:
        img = Image.open(path).convert("RGBA")
        img = img.resize((size, size), Image.LANCZOS)
        return ImageTk.PhotoImage(img)
    # Fallback if Pillow isn't installed: crude integer downscale via Tk.
    photo = tk.PhotoImage(file=path)
    factor = max(1, round(max(photo.width(), photo.height()) / size))
    return photo.subsample(factor, factor) if factor > 1 else photo




class ChatboxTab(tk.Frame):
    def __init__(self, parent, cfg: dict, state, save_cb, start_cb, stop_cb,
                 restart_cb, settings_cb, help_cb):
        super().__init__(parent, bg=BG)
        self._cfg         = cfg
        self._state       = state
        self._save_cb     = save_cb
        self._start_cb    = start_cb
        self._stop_cb     = stop_cb
        self._restart_cb  = restart_cb
        self._settings_cb = settings_cb
        self._help_cb     = help_cb

        self.columnconfigure(0, weight=1)

        if STRIPE_COLOURS:
            # Draw stripes directly on this frame's background via a canvas
            # that sits at z-order bottom; all child widgets grid on top normally.
            self._stripe_canvas = tk.Canvas(self, bg=BG, highlightthickness=0, bd=0)
            self._stripe_canvas.place(x=0, y=0, relwidth=1, relheight=1)
            self.bind("<Configure>", self._on_resize)

        self._build()

    def _on_resize(self, event):
        w, h = event.width, event.height
        draw_stripes(self._stripe_canvas, w, h, STRIPE_COLOURS)
        self._stripe_canvas.tk.call("lower", self._stripe_canvas._w)

    def _build(self):
        row = 0

        # ── Status bar ────────────────────────────────────────────────────────
        status_frame = tk.Frame(self, bg=PANEL, pady=6)
        status_frame.grid(row=row, column=0, sticky="ew", padx=8, pady=(8, 4))
        status_frame.columnconfigure(1, weight=1)
        row += 1

        tk.Label(status_frame, text="Status:", bg=PANEL, fg=SUBTEXT,
                 font=(FONT, 9)).grid(row=0, column=0, padx=(10, 4))
        self._status_lbl = tk.Label(status_frame, text="Stopped", bg=PANEL, fg=RED,
                                    font=(FONT, 9, "bold"))
        self._status_lbl.grid(row=0, column=1, sticky="w")

        # ── Control buttons ───────────────────────────────────────────────────
        btn_frame = tk.Frame(self, bg=BG)
        btn_frame.grid(row=row, column=0, sticky="ew", padx=8, pady=4)
        row += 1

        for text, cmd, fg in (
                ("▶  Start",   self._start_cb,   ACCENT),
                ("■  Stop",    self._stop_cb,    ACCENT),
                ("↺  Restart", self._restart_cb, ACCENT),
        ):
            tk.Button(
                btn_frame, text=text, bg=PANEL, fg=fg,
                relief="flat", cursor="hand2", font=(FONT, 10, "bold"),
                activebackground=BORDER, activeforeground=TEXT,
                width=12, pady=6, command=cmd,
            ).pack(side="left", padx=4)

        tk.Button(
            btn_frame, text="⚙ Settings", bg=PANEL, fg=SUBTEXT,
            relief="flat", cursor="hand2", font=(FONT, 9),
            activebackground=BORDER, activeforeground=TEXT,
            command=self._settings_cb,
        ).pack(side="right", padx=4)

        tk.Button(
            btn_frame, text="? Help", bg=PANEL, fg=SUBTEXT,
            relief="flat", cursor="hand2", font=(FONT, 9),
            activebackground=BORDER, activeforeground=TEXT,
            command=self._help_cb,
        ).pack(side="right", padx=4)

        tk.Frame(self, bg=BORDER, height=1).grid(row=row, column=0, sticky="ew", padx=8, pady=4)
        row += 1

        # ── Live preview ──────────────────────────────────────────────────────
        tk.Label(self, text="Live Chatbox Preview", bg=BG, fg=ACCENT2,
                 font=(FONT, 9, "bold")).grid(row=row, column=0, sticky="w", padx=12)
        row += 1

        preview_frame = tk.Frame(self, bg=PANEL, highlightthickness=1, highlightbackground=BORDER)
        preview_frame.grid(row=row, column=0, sticky="ew", padx=8, pady=(0, 4))
        preview_frame.columnconfigure(0, weight=1)
        row += 1

        self._preview = tk.Text(
            preview_frame, bg=PANEL, fg=TEXT,
            font=(FONT, 10), relief="flat",
            height=8, wrap="word",
            state="disabled", padx=10, pady=8,
        )
        self._preview.pack(fill="x")

        self._page_lbl = tk.Label(preview_frame, text="", bg=PANEL, fg=SUBTEXT, font=(FONT, 8))
        self._page_lbl.pack(anchor="e", padx=8, pady=(0, 4))

        tk.Frame(self, bg=BORDER, height=1).grid(row=row, column=0, sticky="ew", padx=8, pady=4)
        row += 1

        # ── Config fields ─────────────────────────────────────────────────────
        tk.Label(self, text="Configuration", bg=BG, fg=ACCENT2,
                 font=(FONT, 9, "bold")).grid(row=row, column=0, sticky="w", padx=12)
        row += 1

        cfg_frame = tk.Frame(self, bg=BG)
        cfg_frame.grid(row=row, column=0, sticky="ew", padx=12, pady=4)
        cfg_frame.columnconfigure(1, weight=1)
        cfg_frame.columnconfigure(3, weight=1)
        row += 1

        self._entries = {}

        fields = [
            ("OSC IP",       "osc_ip",         0, 0, 1),
            ("OSC Port",     "osc_port",        0, 2, 3),
            ("Interface",    "interface",       1, 0, 1),
            ("useless block","temp_var1",       1, 2, 3),
            ("LHM URL",      "lhm_api",         2, 0, 1),
            ("Location",     "location",        2, 2, 3),
        ]

        for label, key, r, cl, ce in fields:
            tk.Label(cfg_frame, text=label, bg=BG, fg=SUBTEXT,
                     font=(FONT, 9), anchor="e").grid(row=r, column=cl, sticky="e",
                                                      padx=(8, 4), pady=3)
            e = tk.Entry(
                cfg_frame, bg=PANEL, fg=TEXT, insertbackground=ACCENT,
                relief="flat", font=(FONT, 9),
                highlightthickness=1, highlightbackground=BORDER, highlightcolor=ACCENT,
            )
            e.insert(0, str(self._cfg.get(key, "")))
            e.grid(row=r, column=ce, sticky="ew", pady=3)
            self._entries[key] = e

            def _on_change(event, k=key, entry=e):
                self._cfg[k] = entry.get()
                self._save_cb()

            e.bind("<FocusOut>", _on_change)
            e.bind("<Return>",   _on_change)

        tk.Frame(self, bg=BORDER, height=1).grid(row=row, column=0, sticky="ew", padx=8, pady=4)
        row += 1

        # ── Forced text ───────────────────────────────────────────────────────
        tk.Label(self, text="Forced Text (overrides all pages)",
                 bg=BG, fg=ACCENT2, font=(FONT, 9, "bold")).grid(
            row=row, column=0, sticky="w", padx=12)
        row += 1

        forced_frame = tk.Frame(self, bg=BG)
        forced_frame.grid(row=row, column=0, sticky="ew", padx=12, pady=(0, 8))
        forced_frame.columnconfigure(0, weight=1)

        self._forced_var = tk.StringVar()
        forced_entry = tk.Entry(
            forced_frame, textvariable=self._forced_var,
            bg=PANEL, fg=TEXT, insertbackground=ACCENT,
            relief="flat", font=(FONT, 10),
            highlightthickness=1, highlightbackground=BORDER, highlightcolor=ACCENT,
        )
        forced_entry.grid(row=0, column=0, sticky="ew")
        tk.Label(forced_frame, text="Leave blank to use pages",
                 bg=BG, fg=SUBTEXT, font=(FONT, 8)).grid(row=1, column=0, sticky="w")

        def _forced_changed(*_):
            self._state.forced_text = self._forced_var.get()

        self._forced_var.trace_add("write", _forced_changed)

        # discord + github images from "https://icons8.com

        # ── Bottom bar: pinned 20 px above window bottom via place ──────────────────
        bottom_frame = tk.Frame(self, bg=BG)
        bottom_frame.columnconfigure(1, weight=1)

        # Square Discord logo button
        _discord_img = _load_icon("assets/discord.png")
        _discord_btn = tk.Button(
            bottom_frame, image=_discord_img,
            bg="#5865F2", activebackground="#4752C4",
            relief="flat", cursor="hand2", bd=0,
            padx=6, pady=6,
            command=lambda: webbrowser.open("https://discord.gg/YDXpQPF6g9"),
        )
        _discord_btn.image = _discord_img  # keep reference
        _discord_btn.grid(row=0, column=0, sticky="w", padx=(0, 8))

        # Scrolling banner canvas — cycles through PNGs in BANNER_DIR,
        # holding each for BANNER_HOLD_MS then sliding to the next.
        self._banner_canvas = tk.Canvas(bottom_frame, width=BANNER_WIDTH, height=BANNER_HEIGHT,
                                        highlightthickness=0, bd=0, bg=BG)
        self._banner_canvas.grid(row=0, column=1, padx=4)  # no sticky → stays fixed-size & centered

        self._banner_paths      = sorted(glob.glob(f"{BANNER_DIR}/*.png"))
        self._banner_index      = 0
        self._banner_photo      = None   # keep a ref so Tk doesn't garbage-collect it
        self._banner_photo_next = None
        self._banner_after_id   = None
        self._banner_anim_id    = None

        if not self._banner_paths:
            tk.Label(self._banner_canvas, text="(no banners found in assets/banners)",
                     bg=BG, fg=SUBTEXT, font=(FONT, 8)).place(relx=0.5, rely=0.5, anchor="center")
        elif not _PIL_AVAILABLE:
            tk.Label(self._banner_canvas, text="(Pillow not installed — pip install pillow)",
                     bg=BG, fg=SUBTEXT, font=(FONT, 8)).place(relx=0.5, rely=0.5, anchor="center")

        self._banner_canvas.bind("<Configure>", self._on_banner_resize)
        self.bind("<Destroy>", self._on_banner_widget_destroyed)

        # Square GitHub logo button
        _github_img = _load_icon("assets/github.png")
        _github_btn = tk.Button(
            bottom_frame, image=_github_img,
            bg="#24292e", activebackground="#444d56",
            relief="flat", cursor="hand2", bd=0,
            padx=6, pady=6,
            command=lambda: webbrowser.open("https://github.com/CaptainBoots/VRChat-ToolBox"),
        )
        _github_btn.image = _github_img  # keep reference
        _github_btn.grid(row=0, column=2, sticky="e", padx=(8, 0))

        # Pin bottom_frame 20 px above the bottom edge of the tab
        bottom_frame.place(relx=0, rely=1.0, anchor="sw", relwidth=1.0, y=-20)

    # ── Banner rotation ──────────────────────────────────────────────────────

    def _draw_banner_bg(self, w, h):
        """Redraw the diagonal stripe pattern as the banner canvas's own
        background so the letterboxed space around a non-stretched image
        still matches the rest of the theme, instead of a flat block."""
        self._banner_canvas.delete("all")
        if STRIPE_COLOURS:
            draw_stripes(self._banner_canvas, w, h, STRIPE_COLOURS)

    def _on_banner_widget_destroyed(self, event):
        if event.widget is self and self._banner_after_id is not None:
            try:
                self.after_cancel(self._banner_after_id)
            except Exception:
                pass

    def _on_banner_resize(self, event):
        w, h = event.width, event.height
        if w <= 1 or h <= 1:
            return
        if not self._banner_paths or not _PIL_AVAILABLE:
            self._draw_banner_bg(w, h)
            return
        # (Re)draw the current image at the new size. First resize also kicks
        # off the rotation timer.
        first_draw = self._banner_after_id is None
        self._show_banner_image(self._banner_index, animate=False)
        if first_draw:
            self._schedule_next_banner()

    def _fit_banner_image(self, path, w, h):
        if w <= 1 or h <= 1:
            return None
        img = Image.open(path).convert("RGBA")
        scale = min(w / img.width, h / img.height)
        new_w = max(1, int(img.width * scale))
        new_h = max(1, int(img.height * scale))
        img = img.resize((new_w, new_h), Image.LANCZOS)
        return ImageTk.PhotoImage(img)

    def _show_banner_image(self, index, animate=True):
        if not self._banner_canvas.winfo_exists():
            return
        w = self._banner_canvas.winfo_width()
        h = self._banner_canvas.winfo_height()
        if w <= 1 or h <= 1:
            return

        photo = self._fit_banner_image(self._banner_paths[index], w, h)
        if photo is None:
            return

        # No current image yet (first draw / resize before rotation started)
        if not animate or not self._banner_canvas.find_withtag("current_banner"):
            if self._banner_anim_id is not None:
                self.after_cancel(self._banner_anim_id)
                self._banner_anim_id = None
            self._draw_banner_bg(w, h)
            self._banner_photo = photo
            self._banner_canvas.create_image(w // 2, h // 2, image=photo,
                                             anchor="center", tags="current_banner")
            return

        self._banner_photo_next = photo
        self._banner_canvas.create_image(w + w // 2, h // 2, image=photo,
                                         anchor="center", tags="incoming_banner")
        self._animate_banner_slide(w, step=0)

    def _animate_banner_slide(self, w, step):
        if not self._banner_canvas.winfo_exists():
            return
        if step >= BANNER_SLIDE_STEPS:
            h = self._banner_canvas.winfo_height()
            self._draw_banner_bg(w, h)
            self._banner_photo = self._banner_photo_next
            self._banner_canvas.create_image(w // 2, h // 2, image=self._banner_photo,
                                             anchor="center", tags="current_banner")
            self._banner_anim_id = None
            return
        dx = -(w / BANNER_SLIDE_STEPS)
        self._banner_canvas.move("current_banner", dx, 0)
        self._banner_canvas.move("incoming_banner", dx, 0)
        self._banner_anim_id = self.after(
            max(1, BANNER_SLIDE_MS // BANNER_SLIDE_STEPS),
            lambda: self._animate_banner_slide(w, step + 1),
        )

    def _schedule_next_banner(self):
        self._banner_after_id = self.after(BANNER_HOLD_MS, self._advance_banner)

    def _advance_banner(self):
        if not self.winfo_exists() or not self._banner_paths:
            return
        self._banner_index = (self._banner_index + 1) % len(self._banner_paths)
        self._show_banner_image(self._banner_index, animate=True)
        self._schedule_next_banner()

    # ── Public update methods ─────────────────────────────────────────────────

    def set_status(self, text: str):
        colour = GREEN if "running" in text.lower() else RED
        self._status_lbl.config(text=text, fg=colour)

    def set_preview(self, text: str):
        self._preview.config(state="normal")
        self._preview.delete("1.0", tk.END)
        self._preview.insert("1.0", text)
        self._preview.config(state="disabled")

    def set_page_label(self, text: str):
        self._page_lbl.config(text=text)

    def get_forced_text(self) -> str:
        return self._forced_var.get()