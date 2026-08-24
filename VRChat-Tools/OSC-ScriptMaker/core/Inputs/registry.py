"""
core/Inputs/registry.py
─────────────────────────
Every trigger kind is defined in its own module in this folder and
listed here. Adding a new trigger kind:

  1. Write Inputs/my_thing_input.py with ID / LABEL / CATEGORY /
     DESCRIPTION, plus a matches(trigger, key, value) -> bool function
     (or next_interval() for a timer-style kind that doesn't need one).
  2. Add it to the INPUTS list below.

Nothing else needs changing — the engine and the UI's trigger-kind
dropdown both read this registry rather than hardcoding kinds.
"""

from core.Inputs import osc_input, timer_input, variable_input

INPUTS = [
    {"id": osc_input.ID, "label": osc_input.LABEL, "category": osc_input.CATEGORY,
     "description": osc_input.DESCRIPTION, "module": osc_input},
    {"id": timer_input.ID, "label": timer_input.LABEL, "category": timer_input.CATEGORY,
     "description": timer_input.DESCRIPTION, "module": timer_input},
    {"id": variable_input.ID, "label": variable_input.LABEL, "category": variable_input.CATEGORY,
     "description": variable_input.DESCRIPTION, "module": variable_input},
]

# Fast lookup by id
INPUT_BY_ID: dict = {i["id"]: i for i in INPUTS}

# Grouped by category (preserving insertion order), for a future palette UI
CATEGORIES: dict = {}
for _i in INPUTS:
    CATEGORIES.setdefault(_i["category"], []).append(_i)


def input_matches(trigger, key: str, value) -> bool:
    """key is the OSC address for kind=='osc', the variable name for
    kind=='variable'. Timer triggers don't go through this — the engine's
    timer loop drives them directly via next_interval()."""
    entry = INPUT_BY_ID.get(trigger.kind)
    if entry is None or not hasattr(entry["module"], "matches"):
        return False
    return entry["module"].matches(trigger, key, value)


def timer_interval(trigger) -> float:
    return timer_input.next_interval(trigger)
