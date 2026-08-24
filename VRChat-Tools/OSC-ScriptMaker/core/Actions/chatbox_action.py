"""
core/Actions/chatbox_action.py
──────────────────────────────────
Action module: sends text to the VRChat chatbox via /chatbox/input.
Its own action kind (rather than making people configure a raw
send_osc for it) since it's such a common thing to want, and it needs
the extra send_immediately / play_sfx booleans that /chatbox/input
takes alongside the text.

"{value}" in the text is replaced with the trigger's value, so e.g.
a script watching a float parameter can chatbox its live value.
"""

ID = "chatbox"
LABEL = "VRChat Chatbox Message"
CATEGORY = "OSC"
DESCRIPTION = "Sends text to the VRChat chatbox. Use {value} to include the trigger's value."


def run(action, ctx):
    host = action.host or ctx.default_host
    port = int(action.port or ctx.default_port)
    text = (action.text or "").replace("{value}", str(ctx.trigger_value))
    client = ctx.senders.get(host, port)
    client.send_message("/chatbox/input", [text, bool(action.send_immediately), bool(action.play_sfx)])
