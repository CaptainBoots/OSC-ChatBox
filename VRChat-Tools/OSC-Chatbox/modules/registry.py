"""
modules/registry.py
───────────────────
Every chatbox module is defined here as a simple render function.

A module is a dict:
  {
    "id":       str,          # unique key used in config slots
    "label":    str,          # display name shown in the palette
    "category": str,          # palette grouping
    "render":   callable,     # fn(snap: dict, slot: dict) -> str
    "has_text": bool,         # True if slot has a user-editable "text" field
  }

Adding a new module:
  1. Write a render function below.
  2. Add an entry to MODULES list.
  Nothing else needs changing.

snap is the dict returned by AppState.snapshot() plus:
  snap["progress_bar_str"]  — pre-built progress bar string
  snap["media_time_str"]    — pre-built time string
  snap["media_title_clean"] — title after trim/clean
  snap["media_icon"]        — "⏸" or "🎵"
"""

from monitors.media import progress_bar, fmt_time, clean_value, detail_line
from monitors.network import fmt_net

# ── Render helpers ────────────────────────────────────────────────────────────

def _fmt(val, suffix="", fallback="N/A") -> str:
    try:
        f = float(val)
    except (TypeError, ValueError):
        return fallback
    if f <= 0:
        return fallback
    return f"{int(f)}{suffix}" if f == int(f) else f"{f:.1f}{suffix}"


# ── Render functions ──────────────────────────────────────────────────────────

def _render_time(snap, slot):
    import time
    return time.strftime("%I:%M %p")

def _render_custom(snap, slot):
    return slot.get("text", "")

def _render_divider(snap, slot):
    char = "─"
    return char * 15

# ── CPU ───────────────────────────────────────────────────────────────────────
def _render_cpu_name(snap, slot):   return snap.get("cpu_name", "CPU Unknown")
def _render_cpu_load(snap, slot):   return _fmt(snap.get("cpu_load"), "%")
def _render_cpu_temp(snap, slot):   return _fmt(snap.get("cpu_temp"), "℃")
def _render_cpu_power(snap, slot):  return _fmt(snap.get("cpu_power"), "w")

# ── GPU ───────────────────────────────────────────────────────────────────────
def _render_gpu_name(snap, slot):   return snap.get("gpu_name", "GPU Unknown")
def _render_gpu_load(snap, slot):   return _fmt(snap.get("gpu_load"), "%")
def _render_gpu_temp(snap, slot):   return _fmt(snap.get("gpu_temp"), "℃")
def _render_gpu_power(snap, slot):  return _fmt(snap.get("gpu_power"), "w")

# ── VRAM ──────────────────────────────────────────────────────────────────────
def _render_vram_used(snap, slot):
    return f"VRAM {snap.get('vram_used', 0.0):.1f}GB"

def _render_vram_total(snap, slot):
    return f"VRAM Total {snap.get('vram_total', '?')}GB"

def _render_vram_combined(snap, slot):
    return f"{snap.get('vram_type','GDDR')} {snap.get('vram_used',0.0):.1f}GB/{snap.get('vram_total','?')}GB"

# ── RAM ───────────────────────────────────────────────────────────────────────
def _render_ram_used(snap, slot):
    return f"RAM {snap.get('dram_used', 0.0):.1f}GB"

def _render_ram_total(snap, slot):
    return f"RAM Total {snap.get('dram_total', '?')}GB"

def _render_ram_combined(snap, slot):
    return f"{snap.get('dram_type','DDR')} {snap.get('dram_used',0.0):.1f}GB/{snap.get('dram_total','?')}GB"

# ── SteamVR ───────────────────────────────────────────────────────────────────
def _render_fps_vr(snap, slot):
    v = snap.get("vr_fps")
    return f"VR FPS: {v}" if v else "VR FPS: N/A"

def _render_vr_frametime(snap, slot):
    v = snap.get("vr_frametimes")
    return f"Frame: {v}ms" if v else "Frame: N/A"

def _render_vr_reprojection(snap, slot):
    v = snap.get("vr_reprojection")
    if v is None: return "Reproj: N/A"
    return f"Reproj: {int(v*100)}%"

def _render_vr_headset(snap, slot):
    return snap.get("vr_headset") or "Headset: N/A"

def _render_vr_connected(snap, slot):
    return "🟢 VR On" if snap.get("vr_connected") else "🔴 VR Off"

# ── VR Battery helpers ─────────────────────────────────────────────────────────
def _batt_str(pct, charging):
    if pct is None:
        return "N/A"
    icon = "⚡" if charging else "🔋"
    return f"{icon}{pct}%"

def _render_vr_hmd_battery(snap, slot):
    return f"HMD {_batt_str(snap.get('vr_hmd_battery'), snap.get('vr_hmd_charging'))}"

def _render_vr_lc_battery(snap, slot):
    return f"LC {_batt_str(snap.get('vr_lc_battery'), snap.get('vr_lc_charging'))}"

def _render_vr_rc_battery(snap, slot):
    return f"RC {_batt_str(snap.get('vr_rc_battery'), snap.get('vr_rc_charging'))}"

def _render_vr_all_battery(snap, slot):
    parts = [
        f"HMD:{_batt_str(snap.get('vr_hmd_battery'), snap.get('vr_hmd_charging'))}",
        f"L:{_batt_str(snap.get('vr_lc_battery'), snap.get('vr_lc_charging'))}",
        f"R:{_batt_str(snap.get('vr_rc_battery'), snap.get('vr_rc_charging'))}",
    ]
    return "  ".join(parts)

# ── Tracker battery — dynamic, one module per tracker index (0-based) ──────────
_MAX_TRACKERS = 8

def _make_tracker_render(idx):
    def _render(snap, slot):
        trackers = snap.get("vr_trackers", [])
        if idx >= len(trackers):
            return f"T{idx + 1}: N/A"
        t = trackers[idx]
        return f"T{idx + 1}: {_batt_str(t.get('battery'), t.get('charging'))}"
    return _render

# Pre-build render functions and module entries for all tracker slots
_TRACKER_MODULES = [
    {
        "id":       f"vr_tracker_{i + 1}_battery",
        "label":    f"Tracker {i + 1} Battery (Beta)",
        "category": "VR Trackers",
        "render":   _make_tracker_render(i),
        "has_text": False,
    }
    for i in range(_MAX_TRACKERS)
]

# ── VRChat ────────────────────────────────────────────────────────────────────
def _render_fps_desktop(snap, slot):
    v = snap.get("vrc_fps")
    return f"FPS: {v}" if v else "FPS: N/A"

def _render_vrc_world(snap, slot):
    return snap.get("vrc_world") or "World: N/A"

def _render_vrc_players(snap, slot):
    n = snap.get("vrc_player_count", 0)
    return f"Players: {n}"

def _render_vrc_avatar(snap, slot):
    return snap.get("vrc_avatar") or "Avatar: N/A"

def _render_vrc_ping(snap, slot):
    v = snap.get("vrc_ping")
    return f"Ping: {v}ms" if v else "Ping: N/A"


# ── Network ───────────────────────────────────────────────────────────────────
def _render_net_down(snap, slot):
    return f"Download {fmt_net(snap.get('net_down', 0))}"

def _render_net_up(snap, slot):
    return f"Upload {fmt_net(snap.get('net_up', 0))}"

# ── Weather ───────────────────────────────────────────────────────────────────
def _render_weather_temp(snap, slot):
    return f"{snap.get('weather_temp', '?')}℃"

def _render_weather_humidity(snap, slot):
    return f"{snap.get('weather_humidity', '?')}% humidity"

def _render_weather_desc(snap, slot):
    return snap.get("weather_desc", "Unavailable")

def _render_weather_full(snap, slot):
    return (
        f"{snap.get('weather_temp','?')}℃  "
        f"{snap.get('weather_humidity','?')}% humidity  "
        f"{snap.get('weather_desc','Unavailable')}"
    )

# ── Media ─────────────────────────────────────────────────────────────────────
def _render_media_title(snap, slot):
    mi = snap.get("media_info", {})
    title = snap.get("media_title_clean") or clean_value(mi.get("title"))
    if not title:
        return ""
    icon = "⏸" if mi.get("is_paused") else "🎵"
    return f"{icon} {title}"

def _render_media_artist(snap, slot):
    mi = snap.get("media_info", {})
    artist = clean_value(mi.get("artist"))
    return f"- {artist}" if artist else ""

def _render_media_album(snap, slot):
    mi = snap.get("media_info", {})
    return clean_value(mi.get("album"))

def _render_media_source(snap, slot):
    mi = snap.get("media_info", {})
    return clean_value(mi.get("source"))

def _render_media_progress(snap, slot):
    mi  = snap.get("media_info", {})
    return progress_bar(
        mi.get("position_ms", 0),
        mi.get("duration_ms", 0),
        snap.get("progress_filled", "▓"),
        snap.get("progress_border", "▒"),
        snap.get("progress_empty",  "░"),
    )

def _render_media_time(snap, slot):
    mi = snap.get("media_info", {})
    return fmt_time(mi.get("position_ms", 0), mi.get("duration_ms", 0))

def _render_media_detail(snap, slot):
    mi = snap.get("media_info", {})
    return detail_line(mi)

# ── Fun ───────────────────────────────────────────────────────────────────────
def _render_ascii_cat(snap, slot):
    return "/|_/|\n(＞.＜)\n|     \\\n      | || |ノ"

def _render_ascii_dog_1(snap, slot):
    return "  __      _\no''')}____//\n `_/      )\n (_(_/-(_/"

def _render_ascii_dog_2(snap, slot):
    return (
        f"""
        __
   (___()'`;
    /,    /`
   \\"--\\"
"""
    )

def _render_ascii_fish(snap, slot):
    return "<`)))><"



# ── Module registry ───────────────────────────────────────────────────────────

MODULES: list[dict] = [
    # ── Time / misc ─────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────
    {"id": "divider",            "label": "Visual Divider Line",           "category": "General",       "render": _render_divider,              "has_text": False},
    {"id": "time",               "label": "Time",                          "category": "General",       "render": _render_time,                 "has_text": False},
    {"id": "custom_text",        "label": "Custom Text",                   "category": "General",       "render": _render_custom,               "has_text": True},

    # ── CPU ─────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────
    {"id": "cpu_name",           "label": "CPU Name",                       "category": "CPU",          "render": _render_cpu_name,             "has_text": False},
    {"id": "cpu_load",           "label": "CPU Load %",                     "category": "CPU",          "render": _render_cpu_load,             "has_text": False},
    {"id": "cpu_temp",           "label": "CPU Temp ℃",                    "category": "CPU",          "render": _render_cpu_temp,             "has_text": False},
    {"id": "cpu_power",          "label": "CPU Power W",                    "category": "CPU",          "render": _render_cpu_power,            "has_text": False},

    # ── GPU ─────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────
    {"id": "gpu_name",           "label": "GPU Name",                       "category": "GPU",          "render": _render_gpu_name,             "has_text": False},
    {"id": "gpu_load",           "label": "GPU Load %",                     "category": "GPU",          "render": _render_gpu_load,             "has_text": False},
    {"id": "gpu_temp",           "label": "GPU Temp ℃",                    "category": "GPU",          "render": _render_gpu_temp,             "has_text": False},
    {"id": "gpu_power",          "label": "GPU Power W",                    "category": "GPU",          "render": _render_gpu_power,            "has_text": False},

    # ── VRAM ────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────
    {"id": "vram_used",          "label": "VRAM Used GB",                   "category": "GPU",          "render": _render_vram_used,            "has_text": False},
    {"id": "vram_total",         "label": "VRAM Total GB",                  "category": "GPU",          "render": _render_vram_total,           "has_text": False},
    {"id": "vram_used_of_total", "label": "VRAM Used/Total",                "category": "GPU",          "render": _render_vram_combined,        "has_text": False},

    # ── RAM ─────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────
    {"id": "ram_used",           "label": "RAM Used GB",                    "category": "Memory",       "render": _render_ram_used,             "has_text": False},
    {"id": "ram_total",          "label": "RAM Total GB",                   "category": "Memory",       "render": _render_ram_total,            "has_text": False},
    {"id": "ram_used_of_total",  "label": "RAM Used/Total",                 "category": "Memory",       "render": _render_ram_combined,         "has_text": False},

    # ── VR Data ─────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────
    {"id": "vr_fps",             "label": "SteamVR FPS",                    "category": "VR",           "render": _render_fps_vr,               "has_text": False},
    {"id": "vr_frame-time",      "label": "VR Frame Time",                  "category": "VR",           "render": _render_vr_frametime,         "has_text": False},
    {"id": "vr_reprojection",    "label": "VR Reprojection %",              "category": "VR",           "render": _render_vr_reprojection,      "has_text": False},
    {"id": "vr_headset",         "label": "VR Headset Name",                "category": "VR",           "render": _render_vr_headset,           "has_text": False},
    {"id": "vr_connected",       "label": "VR Connected Status",            "category": "VR",           "render": _render_vr_connected,         "has_text": False},


    # ── VR Battery ──────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────
    {"id": "vr_hmd_battery",     "label": "Headset Battery",                "category": "VR",           "render": _render_vr_hmd_battery,       "has_text": False},
    {"id": "vr_lc_battery",      "label": "Left Controller Batt",           "category": "VR",           "render": _render_vr_lc_battery,        "has_text": False},
    {"id": "vr_rc_battery",      "label": "Right Controller Batt",          "category": "VR",           "render": _render_vr_rc_battery,        "has_text": False},
    {"id": "vr_all_battery",     "label": "All Batteries",                  "category": "VR",           "render": _render_vr_all_battery,       "has_text": False},

    # ── VRChat ──────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────
    {"id": "desktop_fps",       "label": "Desktop FPS",                     "category": "VRChat",       "render": _render_fps_desktop,          "has_text": False},
    {"id": "vrc_world",         "label": "World Name",                      "category": "VRChat",       "render": _render_vrc_world,            "has_text": False},
    {"id": "vrc_players",       "label": "Player Count",                    "category": "VRChat",       "render": _render_vrc_players,          "has_text": False},
    {"id": "vrc_avatar",        "label": "Avatar Name (Beta)",              "category": "VRChat",       "render": _render_vrc_avatar,           "has_text": False},
    {"id": "vrc_ping",          "label": "VRChat Ping ",                    "category": "VRChat",       "render": _render_vrc_ping,             "has_text": False},

    # ── Network ─────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────
    {"id": "net_down",          "label": "Download Speed",                  "category": "Network",      "render": _render_net_down,             "has_text": False},
    {"id": "net_up",            "label": "Upload Speed",                    "category": "Network",      "render": _render_net_up,               "has_text": False},

    # ── Weather ─────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────
    {"id": "weather_temp",      "label": "Weather Temp",                    "category": "Weather",      "render": _render_weather_temp,         "has_text": False},
    {"id": "weather_humidity",  "label": "Weather Humidity",                "category": "Weather",      "render": _render_weather_humidity,     "has_text": False},
    {"id": "weather_desc",      "label": "Weather Description",             "category": "Weather",      "render": _render_weather_desc,         "has_text": False},
    {"id": "weather_full",      "label": "Weather Full Line",               "category": "Weather",      "render": _render_weather_full,         "has_text": False},

    # ── Media ───────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────
    {"id": "media_title",       "label": "Media Title",                     "category": "Media",        "render": _render_media_title,          "has_text": False},
    {"id": "media_artist",      "label": "Media Artist",                    "category": "Media",        "render": _render_media_artist,         "has_text": False},
    {"id": "media_album",       "label": "Media Album",                     "category": "Media",        "render": _render_media_album,          "has_text": False},
    {"id": "media_source",      "label": "Media Source App",                "category": "Media",        "render": _render_media_source,         "has_text": False},
    {"id": "media_progress",    "label": "Media Progress Bar",              "category": "Media",        "render": _render_media_progress,       "has_text": False},
    {"id": "media_time",        "label": "Media Time",                      "category": "Media",        "render": _render_media_time,           "has_text": False},
    {"id": "media_detail",      "label": "Media Detail Line",               "category": "Media",        "render": _render_media_detail,         "has_text": False},

    # ── Fun ─────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────
    {"id": "ascii_cat",         "label": "ASCII Cat",                       "category": "Fun",          "render": _render_ascii_cat,            "has_text": False},
    {"id": "ascii_dog_1",       "label": "ASCII Dog 1",                     "category": "Fun",          "render": _render_ascii_dog_1,          "has_text": False},
    {"id": "ascii_dog_2",       "label": "ASCII Dog 2",                     "category": "Fun",          "render": _render_ascii_dog_2,          "has_text": False},
    {"id": "ascii_fish",        "label": "ASCII Fish",                      "category": "Fun",          "render": _render_ascii_fish,           "has_text": False},
]

# Append dynamic tracker modules (T1–T8, "VR Trackers" category)
MODULES.extend(_TRACKER_MODULES)

# Fast lookup by id
MODULE_BY_ID: dict[str, dict] = {m["id"]: m for m in MODULES}

# Palette grouped by category (preserving insertion order)
CATEGORIES: dict[str, list[dict]] = {}
for _m in MODULES:
    CATEGORIES.setdefault(_m["category"], []).append(_m)


def render_slot(slot: dict, snap: dict) -> str:
    """Render a single slot dict against a state snapshot.
    Supports nested horizontal modules if 'modules' is a list."""
    # Check if this slot contains multiple side-by-side modules
    if "modules" in slot and isinstance(slot["modules"], list):
        sub_texts = []
        for sub_slot in slot["modules"]:
            mod = MODULE_BY_ID.get(sub_slot.get("module", ""))
            if mod:
                try:
                    t = mod["render"](snap, sub_slot)
                    if t:
                        sub_texts.append(t)
                except Exception:
                    sub_texts.append(f"[{sub_slot.get('module','?')} error]")
        return " ".join(sub_texts)

    # Standard fallback for single original modules
    mod = MODULE_BY_ID.get(slot.get("module", ""))
    if mod is None:
        return ""
    try:
        return mod["render"](snap, slot)
    except Exception:
        return f"[{slot.get('module','?')} error]"


def render_page(page: dict, snap: dict) -> str:
    """Render all slots in a page and join with newlines."""
    lines = []
    for slot in page.get("slots", []):
        line = render_slot(slot, snap)
        if line:
            lines.append(line)
    return "\n".join(lines)