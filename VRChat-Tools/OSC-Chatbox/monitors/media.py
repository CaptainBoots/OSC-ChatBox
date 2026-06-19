"""monitors/media.py — Media session info fetch + helpers."""

import asyncio
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
    for key, label in (
        ("spotify", "Spotify"), ("chrome", "Chrome"), ("firefox", "Firefox"),
        ("msedge", "Edge"), ("vlc", "VLC"),
    ):
        if key in raw.lower():
            return label
    name = raw.split("!")[-1].replace(".exe", "").split(".")[-1].strip()
    return re.sub(r"[_-]+", " ", name).strip().title()


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
    """
    Smooth the reported playback position using elapsed wall-clock time,
    so the progress bar advances continuously between media polls instead
    of jumping in steps. Mutates info["position_ms"] in-place and returns
    the estimated position in ms.

    pos_state is a small dict the caller keeps between calls (per media
    source) holding: signature, position_ms, raw_position_ms, seen_at.
    """
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
        # New track — trust the reported position immediately
        estimated = raw_pos
    else:
        elapsed_ms = max(0.0, (now - prev_seen) * 1000.0)
        raw_delta  = raw_pos - _ms(prev_raw) if prev_raw is not None else None
        raw_stale  = raw_delta is not None and abs(raw_delta) <= 250.0

        if is_paused:
            # While paused, keep our smoothed value unless the source jumped
            estimated = prev_pos if raw_stale else raw_pos
        elif raw_stale:
            # Source hasn't updated position yet — extrapolate from elapsed time
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
    info = empty()

    if sys.platform == "win32" and wmc is not None:
        try:
            mgr     = await wmc.GlobalSystemMediaTransportControlsSessionManager.request_async()
            session = mgr.get_current_session()
            if session:
                props    = await session.try_get_media_properties_async()
                timeline = session.get_timeline_properties()
                playback = session.get_playback_info()

                info["position_ms"] = timeline.position.total_seconds() * 1000
                info["duration_ms"] = timeline.end_time.total_seconds() * 1000
                info["is_paused"]   = (
                    playback.playback_status ==
                    wmc.GlobalSystemMediaTransportControlsSessionPlaybackStatus.PAUSED
                )
                info["source"] = source_name(
                    getattr(session, "source_app_user_model_id", "") or ""
                )
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

    # Linux: playerctl
    try:
        players = subprocess.check_output(
            ["playerctl", "-l"], encoding="utf-8", stderr=subprocess.DEVNULL, timeout=2
        ).splitlines()
        if not players:
            return info
        player = players[0]
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