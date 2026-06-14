"""
ui/theme.py
───────────
Static color definitions loaded from configuration on startup.
"""

from config import load_config

# Load config to determine color palette immediately on import
_cfg = load_config()
colour_mode = _cfg.get("colour_mode", "new")

FONT = "Consolas"
TITLE_PREFIX = "◈"

if colour_mode == "old":
    BG = "#0f0f13"
    PANEL = "#17171f"
    BORDER = "#2a2a38"
    ACCENT = "#7c5cfc"
    ACCENT2 = "#a78bfa"
    TEXT = "#e2e0f0"
    SUBTEXT = "#7e7b9a"
    GREEN = "#4ade80"
    RED = "#f87171"
    YELLOW = "#facc15"
    CYAN = "#67e8f9"
    ORANGE = "#fb923c"

elif colour_mode == "new":
    BG = "#0f0f13"
    PANEL = "#1f102a"
    BORDER = "#2a2a38"
    ACCENT = "#9D00FF"
    ACCENT2 = "#b44bff"
    TEXT = "#e2e0f0"
    SUBTEXT = "#7e7b9a"
    GREEN = "#4ade80"
    RED = "#f87171"
    YELLOW = "#facc15"
    CYAN = "#67e8f9"
    ORANGE = "#fb923c"

else:  # light
    BG = "#f5f4fa"
    PANEL = "#eae6f3"
    BORDER = "#d3cfe2"
    ACCENT = "#7a00cc"
    ACCENT2 = "#9d00ff"
    TEXT = "#1a1829"
    SUBTEXT = "#625f7a"
    GREEN = "#16a34a"
    RED = "#dc2626"
    YELLOW = "#ca8a04"
    CYAN = "#0891b2"
    ORANGE = "#ea580c"