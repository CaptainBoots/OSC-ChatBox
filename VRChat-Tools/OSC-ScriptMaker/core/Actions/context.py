"""
core/Actions/context.py
──────────────────────────
ActionContext is the small bundle of engine hooks every action module
needs. Keeping it as one plain dataclass (rather than passing the whole
ScriptEngine in) keeps each action module trivially testable on its
own, with no import back onto script_engine.py.
"""

from __future__ import annotations

from dataclasses import dataclass
from typing import Callable, Any

from core.osc_io import OSCSenderPool


@dataclass
class ActionContext:
    trigger_value: Any                       # the value that caused this chain to fire (None for timer)
    default_host: str                        # fallback OSC out host, for actions with no per-action override
    default_port: str                        # fallback OSC out port
    senders: OSCSenderPool
    set_variable: Callable[[str, Any], None]     # (name, value) -> None
    run_sub_action: Callable[[Any], None]         # (Action) -> None, used by the "random" action kind
