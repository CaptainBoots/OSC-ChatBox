"""
ui/help_dialog.py
─────────────────
Paginated help window for OSC-ScriptMaker.
"""

from PySide6.QtCore import Qt
from PySide6.QtWidgets import QDialog, QWidget, QVBoxLayout, QHBoxLayout, QLabel, QPushButton, QFrame

from ui import theme

HELP_PAGES = [
    {
        "title": "Getting Started",
        "content": (
            "OSC-ScriptMaker lets you build automations that react to\n"
            "OSC signals — press a keybind, send another OSC message,\n"
            "post to the VRChat chatbox, and more.\n\n"
            "1. Set Listen Host/Port to where incoming OSC arrives\n"
            "   (VRChat's OSC-out is 127.0.0.1:9001 by default).\n\n"
            "2. Set Send Host/Port for outgoing messages this tool\n"
            "   sends (VRChat's OSC-in is 127.0.0.1:9000 by default).\n\n"
            "3. Click Connect. The status dot turns green.\n\n"
            "4. Click + Add Script, configure a Trigger and one or\n"
            "   more Actions, and it'll fire automatically.\n\n"
            "Everything is saved automatically when you close the app."
        ),
    },
    {
        "title": "Triggers",
        "content": (
            "Every script starts with one trigger:\n\n"
            "OSC MESSAGE\n"
            "  Fires when a matching OSC address is received.\n"
            "  Use an exact address, or end with * to match a prefix\n"
            "  (e.g. /avatar/parameters/* matches any parameter).\n\n"
            "TIMER / INTERVAL\n"
            "  Fires repeatedly every N seconds, no OSC needed.\n\n"
            "VARIABLE CHANGED\n"
            "  Fires when a Set Variable action (from any script)\n"
            "  updates the named variable. This is how scripts chain\n"
            "  off each other.\n\n"
            "CONDITIONS (OSC/Variable only)\n"
            "  Any, Equals, Not equals, Greater/Less than, In range,\n"
            "  Rising edge, Falling edge, Changed."
        ),
    },
    {
        "title": "Actions",
        "content": (
            "Actions run in order, top to bottom, when a trigger fires.\n"
            "Add as many as you like to build a sequence.\n\n"
            "  Press Keybind    — tap or hold a key/combo\n"
            "  Send OSC Message — static, forwarded, or remapped value\n"
            "  VRChat Chatbox   — send text, use {value} for the\n"
            "                     trigger's value\n"
            "  Run Program      — launch a local program\n"
            "  Wait             — pause before the next action\n"
            "  Set Variable     — write a variable other scripts can\n"
            "                     trigger off\n"
            "  Play Sound       — plays a local sound file\n"
            "  Random Action    — runs one randomly-picked sub-action"
        ),
    },
    {
        "title": "Tips",
        "content": (
            "CHAINING SCRIPTS\n"
            "  Script A's Set Variable action + Script B's Variable\n"
            "  Changed trigger lets one script kick off another —\n"
            "  useful for multi-step reactions.\n\n"
            "REMAPPING VALUES\n"
            "  Send OSC Message's Remap mode linearly maps the\n"
            "  trigger's value from one range to another (e.g. a\n"
            "  0–1 float into a 0–100 range for another app), with\n"
            "  optional invert or bool-threshold output.\n\n"
            "SAFETY\n"
            "  A script's action chain automatically stops if it\n"
            "  runs too many actions in one firing — this guards\n"
            "  against accidental variable-trigger loops.\n\n"
            "The ACTIVITY log at the top of the tab shows every\n"
            "script firing and any errors, live."
        ),
    },
]


def open_help(parent):
    dlg = QDialog(parent)
    dlg.setWindowTitle("OSC-ScriptMaker Help")
    dlg.resize(parent.size())

    root = QVBoxLayout(dlg)
    root.setContentsMargins(0, 0, 0, 0)
    root.setSpacing(0)

    current = [0]

    hdr = QWidget()
    hdr.setStyleSheet(f"background-color: {theme.PANEL}; border: none;")
    hdr_layout = QHBoxLayout(hdr)
    hdr_layout.setContentsMargins(16, 10, 16, 10)

    title_lbl = QLabel("")
    title_lbl.setStyleSheet(f"color: {theme.ACCENT2}; background: transparent; border: none;")
    title_lbl.setFont(theme.qt_font(12, bold=True))
    hdr_layout.addWidget(title_lbl)
    hdr_layout.addStretch(1)

    page_lbl = QLabel("")
    page_lbl.setStyleSheet(f"color: {theme.SUBTEXT}; background: transparent; border: none;")
    page_lbl.setFont(theme.qt_font(8))
    hdr_layout.addWidget(page_lbl)

    root.addWidget(hdr)

    divider = QFrame()
    divider.setFixedHeight(1)
    divider.setStyleSheet(f"background-color: {theme.BORDER}; border: none;")
    root.addWidget(divider)

    content_panel = QFrame()
    content_panel.setStyleSheet(f"background-color: {theme.PANEL}; border: 1px solid {theme.BORDER};")
    content_layout = QVBoxLayout(content_panel)
    content_layout.setContentsMargins(16, 14, 16, 14)

    content_lbl = QLabel("")
    content_lbl.setStyleSheet(f"color: {theme.TEXT}; background: transparent; border: none;")
    content_lbl.setFont(theme.qt_font(10))
    content_lbl.setWordWrap(True)
    content_lbl.setAlignment(Qt.AlignLeft | Qt.AlignTop)
    content_layout.addWidget(content_lbl)

    body_wrap = QWidget()
    body_wrap_layout = QVBoxLayout(body_wrap)
    body_wrap_layout.setContentsMargins(20, 14, 20, 0)
    body_wrap_layout.addWidget(content_panel)
    root.addWidget(body_wrap, 1)

    nav = QHBoxLayout()
    nav.setContentsMargins(20, 8, 20, 14)

    prev_btn = QPushButton("← Back")
    prev_btn.setStyleSheet(theme.subtle_button_qss())
    prev_btn.setFont(theme.qt_font(9, bold=True))
    prev_btn.setFixedWidth(100)
    nav.addWidget(prev_btn)
    nav.addStretch(1)

    next_btn = QPushButton("Next →")
    next_btn.setStyleSheet(theme.accent_button_qss())
    next_btn.setFont(theme.qt_font(9, bold=True))
    next_btn.setFixedWidth(100)
    nav.addWidget(next_btn)

    root.addLayout(nav)

    def show(idx):
        p = HELP_PAGES[idx]
        title_lbl.setText(p["title"])
        content_lbl.setText(p["content"])
        page_lbl.setText(f"{idx + 1} / {len(HELP_PAGES)}")
        prev_btn.setEnabled(idx > 0)
        next_btn.setText("Close" if idx == len(HELP_PAGES) - 1 else "Next →")

    def go_back():
        if current[0] > 0:
            current[0] -= 1
            show(current[0])

    def next_or_close():
        if current[0] < len(HELP_PAGES) - 1:
            current[0] += 1
            show(current[0])
        else:
            dlg.close()

    prev_btn.clicked.connect(go_back)
    next_btn.clicked.connect(next_or_close)

    show(0)
    dlg.exec()
