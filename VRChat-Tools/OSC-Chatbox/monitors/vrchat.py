"""monitors/vrchat.py — tails VRChat output log."""
import threading, time, os, re, glob

_data = {
    "vrc_fps":          None,
    "vrc_world":        None,
    "vrc_player_count": 0,
    "vrc_avatar":       None,
    "vrc_ping":         None,
}
_lock    = threading.Lock()
_players = set()

def _find_log() -> str | None:
    base = os.path.expandvars(r"%APPDATA%\..\LocalLow\VRChat\VRChat")
    logs = sorted(
        glob.glob(os.path.join(base, "output_log_*.txt")),
        key=os.path.getmtime,
    )
    return logs[-1] if logs else None

# ── Regexes (tested against real VRChat log lines) ────────────────────────────
# 2024.01.01 12:00:00 Log        - [Behaviour] Joining or Creating Room: My World
_RE_WORLD  = re.compile(r"\[Behaviour\] Joining or Creating Room:\s+(.+)")
# 2024.01.01 12:00:00 Log        - [Behaviour] OnPlayerJoined SomeName (usr_xxx)
_RE_JOIN   = re.compile(r"\[Behaviour\] OnPlayerJoined\s+(.+?)(?:\s+\(usr_[^)]+\))?$")
_RE_LEAVE  = re.compile(r"\[Behaviour\] OnPlayerLeft\s+(.+?)(?:\s+\(usr_[^)]+\))?$")
# Avatar switching — VRChat logs the avatar NAME here:
# [Behaviour] Switching into avatar avtr_xxx (My Avatar Name)
_RE_AVATAR = re.compile(r"\[Behaviour\] Switching into avatar\s+\S+\s+\((.+?)\)")
# FPS — VRChat prints lines like:
# [VRCFlowManager] FPS: 72  or  Fps: 90.0
_RE_FPS    = re.compile(r"[Ff][Pp][Ss][:\s]+(\d+(?:\.\d+)?)")

def _poll():
    global _players
    last_log = None

    while True:
        log = _find_log()
        if not log:
            time.sleep(5)
            continue

        try:
            with open(log, "r", encoding="utf-8", errors="ignore") as f:
                # If it's a new log file, reset state and read from the start
                # so we pick up the current world/avatar from this session.
                if log != last_log:
                    last_log = log
                    with _lock:
                        _data["vrc_world"]        = None
                        _data["vrc_avatar"]       = None
                        _data["vrc_player_count"] = 0
                        _data["vrc_fps"]          = None
                        _data["vrc_ping"]         = None
                    _players = set()
                    # Read from beginning to catch current session state
                    f.seek(0)
                else:
                    # Already processed this file before — jump to end to tail
                    f.seek(0, 2)

                while True:
                    line = f.readline()
                    if not line:
                        # Check if a newer log file appeared (VRChat restarted)
                        new_log = _find_log()
                        if new_log and new_log != last_log:
                            break  # restart outer loop with new file
                        time.sleep(0.3)
                        continue

                    line = line.strip()

                    m = _RE_WORLD.search(line)
                    if m:
                        with _lock:
                            _data["vrc_world"] = m.group(1).strip()
                            _data["vrc_player_count"] = 0
                        _players = set()
                        continue

                    m = _RE_JOIN.search(line)
                    if m:
                        name = m.group(1).strip()
                        _players.add(name)
                        with _lock:
                            _data["vrc_player_count"] = len(_players)
                        continue

                    m = _RE_LEAVE.search(line)
                    if m:
                        name = m.group(1).strip()
                        _players.discard(name)
                        with _lock:
                            _data["vrc_player_count"] = len(_players)
                        continue

                    m = _RE_AVATAR.search(line)
                    if m:
                        with _lock:
                            _data["vrc_avatar"] = m.group(1).strip()
                        continue

                    m = _RE_FPS.search(line)
                    if m:
                        with _lock:
                            _data["vrc_fps"] = int(float(m.group(1)))
                        continue

        except Exception as e:
            print(f"[vrchat monitor] error: {e}")
            time.sleep(2)

def start():
    threading.Thread(target=_poll, daemon=True).start()

def snapshot() -> dict:
    with _lock:
        return dict(_data)