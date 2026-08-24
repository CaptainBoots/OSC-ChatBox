"""
core/Actions/registry.py
────────────────────────────
Every action kind is defined in its own module in this folder and
listed here. Adding a new action kind:

  1. Write Actions/my_thing_action.py with ID / LABEL / CATEGORY /
     DESCRIPTION, plus a run(action, ctx) function (ctx is an
     ActionContext — see context.py).
  2. Add it to the ACTIONS list below.

Nothing else needs changing — the engine's dispatcher and the UI's
action-kind dropdown both read this registry rather than hardcoding
kinds.
"""

from core.Actions import (
    chatbox_action,
    keybind_action,
    play_sound_action,
    random_action,
    run_program_action,
    send_osc_action,
    set_variable_action,
    wait_action,
)
from core.Actions.context import ActionContext

ACTIONS = [
    {"id": keybind_action.ID, "label": keybind_action.LABEL, "category": keybind_action.CATEGORY,
     "description": keybind_action.DESCRIPTION, "module": keybind_action},
    {"id": send_osc_action.ID, "label": send_osc_action.LABEL, "category": send_osc_action.CATEGORY,
     "description": send_osc_action.DESCRIPTION, "module": send_osc_action},
    {"id": chatbox_action.ID, "label": chatbox_action.LABEL, "category": chatbox_action.CATEGORY,
     "description": chatbox_action.DESCRIPTION, "module": chatbox_action},
    {"id": run_program_action.ID, "label": run_program_action.LABEL, "category": run_program_action.CATEGORY,
     "description": run_program_action.DESCRIPTION, "module": run_program_action},
    {"id": wait_action.ID, "label": wait_action.LABEL, "category": wait_action.CATEGORY,
     "description": wait_action.DESCRIPTION, "module": wait_action},
    {"id": set_variable_action.ID, "label": set_variable_action.LABEL, "category": set_variable_action.CATEGORY,
     "description": set_variable_action.DESCRIPTION, "module": set_variable_action},
    {"id": play_sound_action.ID, "label": play_sound_action.LABEL, "category": play_sound_action.CATEGORY,
     "description": play_sound_action.DESCRIPTION, "module": play_sound_action},
    {"id": random_action.ID, "label": random_action.LABEL, "category": random_action.CATEGORY,
     "description": random_action.DESCRIPTION, "module": random_action},
]

# Fast lookup by id
ACTION_BY_ID: dict = {a["id"]: a for a in ACTIONS}

# Grouped by category (preserving insertion order), for the action-kind palette
CATEGORIES: dict = {}
for _a in ACTIONS:
    CATEGORIES.setdefault(_a["category"], []).append(_a)

# Kinds allowed inside a "random" action's sub-action list — everything
# except "random" itself, so nesting stays one level deep.
NON_NESTABLE_KINDS = [a["id"] for a in ACTIONS if a["id"] != random_action.ID]


def run_action(action, ctx: ActionContext):
    entry = ACTION_BY_ID.get(action.kind)
    if entry is None:
        raise ValueError(f"Unknown action kind: {action.kind!r}")
    entry["module"].run(action, ctx)
