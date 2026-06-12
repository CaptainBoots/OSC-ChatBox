"""
config.py
─────────
Config file I/O, defaults, and migration.

Config is a JSON file stored next to main.py.
Pages are stored as a list of {enabled, duration, slots} dicts.
"""

import json
import os
import sys

from state import (
    DEFAULT_PROGRESS_FILLED, DEFAULT_PROGRESS_BORDER, DEFAULT_PROGRESS_EMPTY,
    DEFAULT_SLEEP,
)

SCRIPT_DIR  = os.path.dirname(os.path.abspath(__file__))
CONFIG_FILE = os.path.join(SCRIPT_DIR, "chatbox_config.json")

# ── Default pages ─────────────────────────────────────────────────────────────
DEFAULT_PAGES = [
    {
        "enabled":  True,
        "duration": 20,
        "slots": [
            {"module": "custom_text", "text": "Boots's OSC Chatbox"},
            {"module": "time"},
            {"module": "net_down"},
            {"module": "net_up"},
            {"module": "media_progress"},
            {"module": "media_title"},
        ],
    },
    {
        "enabled":  True,
        "duration": 20,
        "slots": [
            {"module": "custom_text", "text": "Hardware"},
            {"module": "time"},
            {"module": "cpu_name"},
            {"module": "cpu_load"},
            {"module": "cpu_temp"},
            {"module": "cpu_power"},
            {"module": "gpu_name"},
            {"module": "gpu_load"},
            {"module": "gpu_temp"},
            {"module": "gpu_power"},
        ],
    },
    {
        "enabled":  True,
        "duration": 20,
        "slots": [
            {"module": "custom_text", "text": "Memory"},
            {"module": "time"},
            {"module": "ram_used_of_total"},
            {"module": "vram_used_of_total"},
            {"module": "media_progress"},
            {"module": "media_title"},
        ],
    },
    {
        "enabled":  True,
        "duration": 20,
        "slots": [
            {"module": "custom_text", "text": "Local Weather"},
            {"module": "time"},
            {"module": "weather_temp"},
            {"module": "weather_humidity"},
            {"module": "weather_desc"},
            {"module": "media_progress"},
            {"module": "media_title"},
        ],
    },
    {
        "enabled":  True,
        "duration": 20,
        "slots": [
            {"module": "custom_text", "text": "Now Playing"},
            {"module": "time"},
            {"module": "media_progress"},
            {"module": "media_title"},
            {"module": "media_artist"},
            {"module": "media_album"},
            {"module": "media_detail"},
        ],
    },
]


def get_defaults() -> dict:
    return {
        "osc_ip":          "127.0.0.1",
        "osc_port":        9000,
        "interface":       _default_interface(),
        "switch_interval": 20,
        "lhm_api":         "http://localhost:8085/data.json",
        "location":        "0,0",
        "slow_mode":       False,
        "speed_mode":      False,
        "media_title_trim": True,
        "cat_mode":        False,
        "progress_filled": DEFAULT_PROGRESS_FILLED,
        "progress_border": DEFAULT_PROGRESS_BORDER,
        "progress_empty":  DEFAULT_PROGRESS_EMPTY,
        "ui_scale":        1.0,
        "pages":           DEFAULT_PAGES,
    }


def _default_interface() -> str:
    try:
        import psutil
        ifaces = list(psutil.net_io_counters(pernic=True).keys())
        if sys.platform == "win32":
            for preferred in ("Ethernet", "Wi-Fi"):
                if preferred in ifaces:
                    return preferred
        for iface in ifaces:
            if not iface.lower().startswith("lo"):
                return iface
    except Exception:
        pass
    return "Ethernet" if sys.platform == "win32" else "eth0"


# ── Load / Save ───────────────────────────────────────────────────────────────

def load_config() -> dict:
    defaults = get_defaults()
    try:
        with open(CONFIG_FILE, "r", encoding="utf-8") as f:
            loaded = json.load(f)
        merged = {**defaults, **loaded}
        # Ensure pages list always exists and has duration on each page
        if not isinstance(merged.get("pages"), list) or not merged["pages"]:
            merged["pages"] = DEFAULT_PAGES
        else:
            for page in merged["pages"]:
                if "duration" not in page:
                    page["duration"] = merged.get("switch_interval", 20)
                if "slots" not in page:
                    page["slots"] = []
                if "enabled" not in page:
                    page["enabled"] = True
        return merged
    except (FileNotFoundError, json.JSONDecodeError):
        return defaults


def save_config(cfg: dict):
    with open(CONFIG_FILE, "w", encoding="utf-8") as f:
        json.dump(cfg, f, indent=2)


def normalize_char(value, fallback: str) -> str:
    text = str(value).strip() if value else ""
    return text[0] if text else fallback
