"""monitors/vrchat.py — tails VRChat output log + listens for OSC feedback."""
import threading, time, os, re, glob, json

from pythonosc.dispatcher import Dispatcher
from pythonosc.osc_server import BlockingOSCUDPServer

_data = {
    "vrc_fps":          None,
    "vrc_world":        None,
    "vrc_player_count": 0,
    "vrc_avatar":       None,
    "vrc_ping":         None,
}
_lock    = threading.Lock()
_players = set()

# ── OSC feedback listener (VRChat → us on port 9001) ─────────────────────────

def _on_fps(address, *args):
    if args:
        with _lock:
            _data["vrc_fps"] = int(args[0])

def _start_osc_server():
    dispatcher = Dispatcher()
    dispatcher.map("/avatar/parameters/FPS", _on_fps)
    dispatcher.set_default_handler(lambda *a: None)
    try:
        server = BlockingOSCUDPServer(("127.0.0.1", 9001), dispatcher)
        server.serve_forever()
    except Exception as e:
        print(f"[vrchat monitor] OSC server error: {e}")

# ── Log regexes ───────────────────────────────────────────────────────────────
_RE_WORLD  = re.compile(r"\[Behaviour\] Joining or Creating Room:\s+(.+)")
_RE_JOIN   = re.compile(r"\[Behaviour\] OnPlayerJoined\s+(.+?)(?:\s+\(usr_[^)]+\))?$")
_RE_LEAVE  = re.compile(r"\[Behaviour\] OnPlayerLeft\s+(.+?)(?:\s+\(usr_[^)]+\))?$")
_RE_AVATAR = re.compile(r"\[Behaviour\] Switching Avatar To\s+(\S+)")
# VRChat periodically writes a JSON stats blob on a single line:
# {"runningTime": 97.6, "stats": [{"name": "fps", ...}, {"name": "ping", ...}, ...]}
_RE_STATS  = re.compile(r'^\{"runningTime":')

def _parse_stats(line: str):
    """Extract fps and ping tw-mean from a VRChat stats JSON line."""
    try:
        obj = json.loads(line)
        fps = ping = None
        for stat in obj.get("stats", []):
            name = stat.get("name")
            if name == "fps":
                fps = int(stat["tw-mean"])
            elif name == "ping":
                ping = int(stat["tw-mean"])
        return fps, ping
    except Exception:
        return None, None

def _find_log() -> str | None:
    base = os.path.expandvars(r"%APPDATA%\..\LocalLow\VRChat\VRChat")
    logs = sorted(
        glob.glob(os.path.join(base, "output_log_*.txt")),
        key=os.path.getmtime,
    )
    return logs[-1] if logs else None

def _poll():
    global _players
    last_log = None
    last_pos = 0

    while True:
        log = _find_log()
        if not log:
            time.sleep(5)
            continue

        try:
            with open(log, "r", encoding="utf-8", errors="ignore") as f:
                if log != last_log:
                    last_log = log
                    last_pos = 0
                    with _lock:
                        _data["vrc_world"]        = None
                        _data["vrc_avatar"]       = None
                        _data["vrc_player_count"] = 0
                        _data["vrc_fps"]          = None
                        _data["vrc_ping"]         = None
                    _players = set()

                f.seek(last_pos)

                while True:
                    line = f.readline()
                    if not line:
                        last_pos = f.tell()
                        new_log = _find_log()
                        if new_log and new_log != last_log:
                            break
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
                        _players.add(m.group(1).strip())
                        with _lock:
                            _data["vrc_player_count"] = len(_players)
                        continue

                    m = _RE_LEAVE.search(line)
                    if m:
                        _players.discard(m.group(1).strip())
                        with _lock:
                            _data["vrc_player_count"] = len(_players)
                        continue

                    m = _RE_AVATAR.search(line)
                    if m:
                        with _lock:
                            _data["vrc_avatar"] = m.group(1).strip()
                        continue

                    if _RE_STATS.match(line):
                        fps, ping = _parse_stats(line)
                        with _lock:
                            if fps is not None:
                                _data["vrc_fps"] = fps
                            if ping is not None:
                                _data["vrc_ping"] = ping
                        continue

        except Exception as e:
            print(f"[vrchat monitor] error: {e}")
            time.sleep(2)

_started = False

def start():
    global _started
    if _started:
        return
    _started = True
    threading.Thread(target=_poll, daemon=True).start()
    threading.Thread(target=_start_osc_server, daemon=True).start()

def snapshot() -> dict:
    with _lock:
        return dict(_data)