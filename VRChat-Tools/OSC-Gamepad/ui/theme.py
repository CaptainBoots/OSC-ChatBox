"""
ui/theme.py
───────────
Colour definitions. The active palette is selected by colour_mode,
which is set at startup based on the saved config (theme_mode key,
default "new").

To change the theme: Settings → Theme. Takes effect after restart.
"""

colour_mode = "new"

FONT = "Consolas"
TITLE_PREFIX = "◈"

THEMES: dict[str, dict[str, str]] = {
    "old": {
        "BG":      "#0f0f13",
        "PANEL":   "#17171f",
        "BORDER":  "#2a2a38",
        "ACCENT":  "#7c5cfc",
        "ACCENT2": "#a78bfa",
        "TAB":     "#4ade80",
        "TEXT":    "#e2e0f0",
        "TEXT2":   "#E0E0E0",
        "SUBTEXT": "#7e7b9a",
        "GREEN":   "#4ade80",
        "RED":     "#f87171",
        "YELLOW":  "#facc15",
        "CYAN":    "#67e8f9",
        "ORANGE":  "#fb923c",
    },
    "new": {
        "BG":      "#0f0f13",
        "PANEL":   "#1f102a",
        "BORDER":  "#2a2a38",
        "ACCENT":  "#9D00FF",
        "ACCENT2": "#b44bff",
        "TAB":     "#4ade80",
        "TEXT":    "#e2e0f0",
        "TEXT2":   "#E0E0E0",
        "SUBTEXT": "#7e7b9a",
        "GREEN":   "#4ade80",
        "RED":     "#f87171",
        "YELLOW":  "#facc15",
        "CYAN":    "#67e8f9",
        "ORANGE":  "#fb923c",
    },
    "light": {
        "BG":      "#F6E6FA",
        "PANEL":   "#ffffff",
        "BORDER":  "#DDCAE3",
        "ACCENT":  "#9D00FF",
        "ACCENT2": "#b44bff",
        "TAB":     "#000000",
        "TEXT":    "#1a1829",
        "TEXT2":   "#1a1829",
        "SUBTEXT": "#1a1829",
        "GREEN":   "#4ade80",
        "RED":     "#f87171",
        "YELLOW":  "#facc15",
        "CYAN":    "#67e8f9",
        "ORANGE":  "#fb923c",
    },
}

THEME_LABELS: dict[str, str] = {
    "new":   "Purple (Default)",
    "old":   "Violet (Classic)",
    "light": "Light",
}


def set_theme(mode: str):
    """Apply a theme by rebinding this module's colour globals. Call before any UI is built."""
    global colour_mode
    palette = THEMES.get(mode, THEMES["new"])
    colour_mode = mode if mode in THEMES else "new"
    g = globals()
    for key, value in palette.items():
        g[key] = value


# Apply the default palette on import so `from ui.theme import BG` etc. always works.
set_theme(colour_mode)