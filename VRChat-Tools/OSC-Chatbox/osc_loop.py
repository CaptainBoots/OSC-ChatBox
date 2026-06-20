"""
osc_loop.py
───────────
Background OSC loop that runs on a daemon thread.

  start_loop(cfg, state, status_cb, preview_cb)
  stop_loop()

All mutable state lives in AppState (state.py).
Page layout is read from cfg["pages"] each tick so live edits take effect.
"""

import asyncio
import threading
import time
from typing import Callable, Optional

import psutil
from pythonosc.udp_client import SimpleUDPClient

import hardware.lhm as lhm_mod
from hardware.cpu import detect_cpu, get_cpu_temp, get_cpu_power, get_cpu_load
from hardware.gpu import detect_gpu, detect_vram_type, get_gpu_temp, get_gpu_power, get_gpu_load
from hardware.lhm import get_lhm_data
from hardware.memory import (
    detect_dram_type, get_dram_used, get_dram_total,
    get_vram_used, get_vram_total,
)
from modules.registry import render_page
from monitors import media as media_mod
from monitors.media import clean_title, clean_value, estimate_position
from monitors.network import sample as net_sample
from monitors.weather import fetch as weather_fetch
from monitors import steamvr, vrchat
from state import AppState, CHATBOX_MAX_CHARS

# Polling intervals
_LHM_INTERVAL     = 1.0   # seconds between LHM REST calls / hardware sensor reads
_WEATHER_INTERVAL = 300   # seconds between weather refreshes
_MEDIA_INTERVAL   = 1.0   # seconds between media polls

_loop_thread: Optional[threading.Thread] = None


def start_loop(
    cfg:        dict,
    state:      AppState,
    status_cb:  Callable[[str], None],   # called on the main thread via root.after
    preview_cb: Callable[[str], None],   # called with the latest chatbox text
):
    global _loop_thread
    state.running = True
    _loop_thread = threading.Thread(
        target=_run, args=(cfg, state, status_cb, preview_cb), daemon=True
    )
    _loop_thread.start()


def stop_loop(state: AppState):
    state.running = False


def _run(cfg, state: AppState, status_cb, preview_cb):
    # ── Validate interface ────────────────────────────────────────────────────
    interface = cfg.get("interface", "Ethernet")
    all_stats = psutil.net_io_counters(pernic=True)
    if interface not in all_stats:
        status_cb(f"Error: interface '{interface}' not found")
        state.running = False
        return

    # ── Set LHM URL from config ───────────────────────────────────────────────
    lhm_mod.LHM_URL = cfg.get("lhm_api", "http://localhost:8085/data.json")

    # ── One-time hardware detection ───────────────────────────────────────────
    state.cpu_name  = detect_cpu()
    state.gpu_name  = detect_gpu()
    state.dram_type = detect_dram_type()
    state.vram_type = detect_vram_type(state.gpu_name)

    init_lhm = get_lhm_data()
    state.dram_total = get_dram_total(init_lhm)
    state.vram_total = get_vram_total(init_lhm)

    print(f"CPU: {state.cpu_name}  GPU: {state.gpu_name}")
    print(f"RAM: {state.dram_total}GB  VRAM: {state.vram_total}GB")

    # ── Start VR / VRChat background monitors ─────────────────────────────────
    steamvr.start()
    vrchat.start()

    # ── LHM background poller ─────────────────────────────────────────────────
    lhm_cache = {"data": init_lhm, "lock": threading.Lock()}

    def _poll_lhm():
        while state.running:
            data = get_lhm_data()
            if data:
                with lhm_cache["lock"]:
                    lhm_cache["data"] = data
            time.sleep(_LHM_INTERVAL)

    threading.Thread(target=_poll_lhm, daemon=True).start()

    # ── Media background poller ───────────────────────────────────────────────
    media_cache = {"info": media_mod.empty(), "lock": threading.Lock()}

    def _poll_media():
        async def _loop():
            while state.running:
                info = await media_mod.fetch()
                with media_cache["lock"]:
                    media_cache["info"] = info if isinstance(info, dict) else media_mod.empty()
                await asyncio.sleep(_MEDIA_INTERVAL)
        asyncio.run(_loop())

    threading.Thread(target=_poll_media, daemon=True).start()

    # ── OSC client ────────────────────────────────────────────────────────────
    client = SimpleUDPClient(cfg.get("osc_ip", "127.0.0.1"), int(cfg.get("osc_port", 9000)))

    # ── Network baseline ──────────────────────────────────────────────────────
    prev_net  = all_stats[interface]
    prev_time = time.time()

    # ── Weather ───────────────────────────────────────────────────────────────
    def _do_weather():
        t, h, d = weather_fetch(cfg.get("location", "0,0"))
        state.update_weather(t, h, d)

    _do_weather()
    last_weather = time.time()

    # ── Page rotation state ───────────────────────────────────────────────────
    page_index       = 0
    page_start_time  = time.time()
    media_pos_state: dict = {}

    status_cb("Running")

    while state.running:
        try:
            now = time.time()

            # ── Sleep mode ────────────────────────────────────────────────────
            if state.slow_mode:
                sleep = 5.0
            elif state.speed_mode:
                sleep = 0.1
            else:
                sleep = 1.0
            state.sleep_delay = sleep

            # ── Hardware sensors (read from LHM cache, refreshed every ~1s) ────
            with lhm_cache["lock"]:
                lhm_data = lhm_cache["data"]
            if lhm_data:
                state.update_hardware(
                    cpu_temp  = get_cpu_temp(lhm_data),
                    cpu_power = get_cpu_power(lhm_data),
                    cpu_load  = get_cpu_load(lhm_data),
                    gpu_temp  = get_gpu_temp(lhm_data),
                    gpu_power = get_gpu_power(lhm_data),
                    gpu_load  = get_gpu_load(lhm_data),
                    dram_used = get_dram_used(lhm_data),
                    vram_used = get_vram_used(lhm_data),
                )
                # Retry totals if they came up as "?"
                if state.vram_total == "?":
                    state.vram_total = get_vram_total(lhm_data)
                if state.dram_total == "?":
                    state.dram_total = get_dram_total(lhm_data)

            # ── Network ───────────────────────────────────────────────────────
            prev_net, up, down, prev_time = net_sample(prev_net, prev_time, interface)
            state.update_network(up, down)

            # ── Weather (interval) ────────────────────────────────────────────
            if now - last_weather >= _WEATHER_INTERVAL:
                _do_weather()
                last_weather = now

            # ── Media snapshot ────────────────────────────────────────────────
            with media_cache["lock"]:
                media_info = dict(media_cache["info"])

            # Smooth the playback position so the progress bar advances
            # continuously between media polls instead of stepping.
            estimate_position(media_info, media_pos_state, now)

            state.update_media(media_info)

            # ── Build snap dict for module renderers ──────────────────────────
            snap = state.snapshot()
            snap["media_info"] = media_info

            # Merge SteamVR and VRChat monitor data into snap
            snap.update(steamvr.snapshot())
            snap.update(vrchat.snapshot())

            # Pre-build media title clean so all media modules share same result
            raw_title = clean_value(media_info.get("title"))
            snap["media_title_clean"] = (
                clean_title(raw_title) if state.media_title_trim else raw_title
            )

            # ── Forced text override ──────────────────────────────────────────
            forced = snap.get("forced_text", "").strip()
            if forced:
                text = forced
            else:
                # ── Page rotation ─────────────────────────────────────────────
                pages = cfg.get("pages", [])
                enabled = [p for p in pages if p.get("enabled", True)]

                if not enabled:
                    text = "No pages enabled"
                else:
                    # Advance page if its duration has elapsed
                    current_page = enabled[page_index % len(enabled)]
                    duration = float(current_page.get("duration", cfg.get("switch_interval", 20)))
                    if now - page_start_time >= duration:
                        page_index   = (page_index + 1) % len(enabled)
                        page_start_time = now
                        current_page = enabled[page_index % len(enabled)]

                    text = render_page(current_page, snap)

            text = text[:CHATBOX_MAX_CHARS]
            preview_cb(text)
            client.send_message("/chatbox/input", [text, True])
            print(text)

        except Exception as e:
            print(f"[OSC] Error: {e}")

        time.sleep(state.sleep_delay)

    status_cb("Stopped")