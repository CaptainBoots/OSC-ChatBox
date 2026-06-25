"""monitors/media.py — Media session info fetch + helpers."""

import re
import subprocess
import sys
from typing import Optional


# ── Platform import ───────────────────────────────────────────────────────────
if sys.platform == "win32":
    try:
        import winrt.windows.media.control as wmc
    except ImportError:
        wmc = None
else:
    wmc = None


# ── Global State Tracking ──────────────────────────────────────────────────────
# Keeps track of the last source ID that was actively playing across function calls
_LAST_PLAYING_SOURCE: Optional[str] = None


# ── Priority Configuration ────────────────────────────────────────────────────
# Apps ordered by strict preference. If multiple items are playing simultaneously,
# items appearing earlier in this list will take precedence.
PRIORITY_ORDER = [
    # 1. Dedicated Music Apps
    "spotify", "itunes", "applemusic", "tidal", "deezer", "foobar2000", "winamp", "musicbee", "aimp",

    # 2. Web Browsers
    "firefox", "308046b0", "chrome", "googlechrome", "msedge", "edge", "brave", "opera", "vivaldi", "waterfox", "librewolf",

    # 3. Video & Media Players
    "vlc", "mpc-hc", "mpchc", "mpv", "potplayer", "plex", "netflix", "primevideo",

    # 4. Communication Utilities
    "discord", "telegram", "whatsapp",

    # 5. Native OS Players
    "zune", "microsoft.windows.music", "microsoft.zunevideo",
]


def _get_priority_score(raw_id: str) -> int:
    """Returns an integer representing priority. Lower numbers = Higher priority."""
    if not raw_id:
        return len(PRIORITY_ORDER) + 1
    raw_lower = raw_id.lower()
    for index, key in enumerate(PRIORITY_ORDER):
        if key in raw_lower:
            return index
    return len(PRIORITY_ORDER)  # Default fallback priority


def empty() -> dict:
    return {
        "title": "", "artist": "", "album": "", "album_artist": "",
        "track_number": None, "track_count": None, "source": "",
        "position_ms": 0, "duration_ms": 0, "is_paused": False,
    }


def clean_value(v) -> str:
    if v is None:
        return ""
    s = str(v).strip()
    return "" if s.lower() in ("none", "unknown", "null") else s


def clean_title(raw: str) -> str:
    if not raw:
        return ""
    t = re.sub(r"\(.*?\)|\[.*?]|\{.*?}", "", raw)
    junk = r"\b(official|video|lyrics|audio|hd|4k|remastered|live|visualizer|explicit|clean|version|mix)\b"
    t = re.sub(junk, "", t, flags=re.IGNORECASE)
    t = re.sub(r"\b(ft\.|feat\.|featuring).*", "", t, flags=re.IGNORECASE)
    parts = [p.strip() for p in re.split(r"[-–|•]", t) if len(p.strip()) > 2]
    t = parts[0] if parts else t
    return re.sub(r"\s+", " ", t).strip()


def source_name(raw: str) -> str:
    if not raw:
        return ""

    raw_lower = raw.lower()

    # Explicit human-readable conversions
    mappings = {
        "firefox": "Firefox", "308046b0": "Firefox",
        "chrome": "Chrome", "googlechrome": "Chrome",
        "msedge": "Edge", "edge": "Edge", "brave": "Brave", "opera": "Opera", "vivaldi": "Vivaldi", "waterfox": "Waterfox", "librewolf": "LibreWolf",
        "spotify": "Spotify", "itunes": "iTunes", "applemusic": "Apple Music", "tidal": "Tidal", "deezer": "Deezer", "foobar2000": "foobar2000", "winamp": "Winamp", "musicbee": "MusicBee", "aimp": "AIMP",
        "vlc": "VLC", "mpc-hc": "MPC-HC", "mpchc": "MPC-HC", "mpv": "mpv", "potplayer": "PotPlayer", "plex": "Plex", "netflix": "Netflix", "primevideo": "Prime Video",
        "discord": "Discord", "telegram": "Telegram", "whatsapp": "WhatsApp",
        "zune": "Windows Media Player", "microsoft.windows.music": "Media Player", "microsoft.zunevideo": "Movies & TV",
    }

    for key, label in mappings.items():
        if key in raw_lower:
            return label

    # Advanced Regex Cleanup Fallback
    name = raw.split("!")[-1]
    name = name.split("/")[-1].split("\\")[-1]
    name = name.replace(".exe", "")
    name = name.split(".")[0] if "." in name and "_" in name else name
    name = re.sub(r"_[a-z0-9]{13}$", "", name, flags=re.IGNORECASE)
    name = re.sub(r"[0-9a-f]{8,}", "", name, flags=re.IGNORECASE)

    cleaned = re.sub(r"[._-]+", " ", name).strip()
    return cleaned.title() if cleaned else "System Media"


def progress_bar(pos_ms: float, dur_ms: float, filled: str, border: str, empty: str, length: int = 15) -> str:
    if dur_ms <= 0:
        return empty * length
    pct = min(max(pos_ms / dur_ms, 0), 1)
    n   = int(length * pct)
    if 0 < n < length:
        return filled * n + border + empty * (length - n - 1)
    return filled * n + empty * (length - n)


def fmt_time(pos_ms, dur_ms) -> str:
    try:
        ps = max(0, int(float(pos_ms) / 1000))
        ds = max(0, int(float(dur_ms) / 1000))
    except (TypeError, ValueError):
        return ""
    if ds <= 0:
        return ""
    def clk(s):
        m, s = divmod(s, 60)
        h, m = divmod(m, 60)
        return f"{h}:{m:02d}:{s:02d}" if h else f"{m}:{s:02d}"
    return f"{clk(ps)} / {clk(ds)}"


def _ms(value, fallback: float = 0.0) -> float:
    try:
        return max(0.0, float(value))
    except (TypeError, ValueError):
        return fallback


def estimate_position(info: dict, pos_state: dict, now: float) -> float:
    raw_pos   = _ms(info.get("position_ms"))
    duration  = _ms(info.get("duration_ms"))
    is_paused = bool(info.get("is_paused", False))

    signature = (
        clean_value(info.get("title")),
        clean_value(info.get("artist")),
        clean_value(info.get("album")),
        clean_value(info.get("track_number")),
        duration,
    )

    if not signature[0]:
        pos_state.clear()
        info["position_ms"] = 0
        return 0

    prev_sig  = pos_state.get("signature")
    prev_pos  = _ms(pos_state.get("position_ms"))
    prev_raw  = pos_state.get("raw_position_ms")
    prev_seen = pos_state.get("seen_at", now)

    if signature != prev_sig:
        estimated = raw_pos
    else:
        elapsed_ms = max(0.0, (now - prev_seen) * 1000.0)
        raw_delta  = raw_pos - _ms(prev_raw) if prev_raw is not None else None
        raw_stale  = raw_delta is not None and abs(raw_delta) <= 250.0

        if is_paused:
            estimated = prev_pos if raw_stale else raw_pos
        elif raw_stale:
            estimated = prev_pos + elapsed_ms
        else:
            estimated = raw_pos

    if duration > 0:
        estimated = min(estimated, duration)

    pos_state["signature"]       = signature
    pos_state["position_ms"]     = estimated
    pos_state["raw_position_ms"] = raw_pos
    pos_state["seen_at"]         = now
    info["position_ms"] = estimated
    return estimated


def detail_line(info: dict) -> str:
    parts = []
    album = clean_value(info.get("album"))
    track = info.get("track_number")
    count = info.get("track_count")
    src   = clean_value(info.get("source"))
    t     = fmt_time(info.get("position_ms", 0), info.get("duration_ms", 0))

    if track:
        parts.append(f"Track {track}/{count}" if count else f"Track {track}")
    if t:
        parts.append(t)
    if src:
        parts.append(src)
    return " | ".join(parts)


async def fetch() -> dict:
    global _LAST_PLAYING_SOURCE
    info = empty()

    # ── Windows Platform ──────────────────────────────────────────────────────
    if sys.platform == "win32" and wmc is not None:
        try:
            mgr = await wmc.GlobalSystemMediaTransportControlsSessionManager.request_async()
            sessions = mgr.get_sessions()
            if not sessions:
                return info

            playing_sessions = []
            paused_sessions = []

            for s in sessions:
                raw_id = getattr(s, "source_app_user_model_id", "") or ""
                playback = s.get_playback_info()
                status = playback.playback_status if playback else None

                if status == wmc.GlobalSystemMediaTransportControlsSessionPlaybackStatus.PLAYING:
                    playing_sessions.append((s, raw_id))
                else:
                    paused_sessions.append((s, raw_id))

            target_session = None

            # 1. If multiple are playing: Pick the highest priority matching the dictionary list
            if playing_sessions:
                playing_sessions.sort(key=lambda item: _get_priority_score(item[1]))
                target_session = playing_sessions[0][0]
                # Keep historical record of what was just active
                _LAST_PLAYING_SOURCE = playing_sessions[0][1]

            # 2. If all are paused: Try to find the one that matches our historical record
            elif paused_sessions:
                if _LAST_PLAYING_SOURCE:
                    for s, raw_id in paused_sessions:
                        if raw_id == _LAST_PLAYING_SOURCE:
                            target_session = s
                            break

                # If history doesn't exist or app was closed, fall back to dictionary priority ranking
                if not target_session:
                    paused_sessions.sort(key=lambda item: _get_priority_score(item[1]))
                    target_session = paused_sessions[0][0]

            if not target_session:
                return info

            props    = await target_session.try_get_media_properties_async()
            timeline = target_session.get_timeline_properties()
            playback = target_session.get_playback_info()

            info["position_ms"] = timeline.position.total_seconds() * 1000
            info["duration_ms"] = timeline.end_time.total_seconds() * 1000
            info["is_paused"]   = (
                    playback.playback_status ==
                    wmc.GlobalSystemMediaTransportControlsSessionPlaybackStatus.PAUSED
            )
            info["source"] = source_name(getattr(target_session, "source_app_user_model_id", "") or "")

            if props:
                info["title"]        = clean_value(getattr(props, "title", ""))
                info["artist"]       = clean_value(getattr(props, "artist", ""))
                info["album"]        = clean_value(getattr(props, "album_title", ""))
                info["album_artist"] = clean_value(getattr(props, "album_artist", ""))
                info["track_number"] = _safe_int(getattr(props, "track_number", None))
                info["track_count"]  = _safe_int(getattr(props, "album_track_count", None))
            return info
        except Exception:
            pass

    # ── Linux Platform ────────────────────────────────────────────────────────
    try:
        players = subprocess.check_output(
            ["playerctl", "-l"], encoding="utf-8", stderr=subprocess.DEVNULL, timeout=2
        ).splitlines()
        if not players:
            return info

        playing_players = []
        paused_players = []

        for p in players:
            status = subprocess.check_output(
                ["playerctl", "-p", p, "status"],
                encoding="utf-8", stderr=subprocess.DEVNULL, timeout=2,
            ).strip().lower()

            if status == "playing":
                playing_players.append(p)
            else:
                paused_players.append(p)

        player = None

        # 1. If multiple are playing: Pick the highest priority matching the dictionary list
        if playing_players:
            playing_players.sort(key=_get_priority_score)
            player = playing_players[0]
            _LAST_PLAYING_SOURCE = player

        # 2. If all are paused: Fall back to our last playing history tracker
        elif paused_players:
            if _LAST_PLAYING_SOURCE and _LAST_PLAYING_SOURCE in paused_players:
                player = _LAST_PLAYING_SOURCE
            else:
                paused_players.sort(key=_get_priority_score)
                player = paused_players[0]

        if not player:
            return info

        out = subprocess.check_output(
            ["playerctl", "-p", player, "metadata", "--format",
             "{{title}}\n{{artist}}\n{{album}}\n{{xesam:trackNumber}}\n{{position}}\n{{mpris:length}}"],
            encoding="utf-8", stderr=subprocess.DEVNULL, timeout=2,
        ).strip().split("\n")

        if len(out) >= 6:
            info["title"]       = clean_value(out[0])
            info["artist"]      = clean_value(out[1])
            info["album"]       = clean_value(out[2])
            info["track_number"] = _safe_int(out[3])
            info["position_ms"] = int(out[4]) / 1000
            info["duration_ms"] = int(out[5]) / 1000
            status = subprocess.check_output(
                ["playerctl", "-p", player, "status"],
                encoding="utf-8", stderr=subprocess.DEVNULL, timeout=2,
            ).strip().lower()
            info["is_paused"] = (status == "paused")
            info["source"]    = source_name(player)
    except Exception:
        pass

    return info


def _safe_int(v) -> Optional[int]:
    try:
        n = int(v)
        return n if n > 0 else None
    except (TypeError, ValueError):
        return None