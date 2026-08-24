"""
core/script_engine.py
──────────────────────
The brain of OSC-ScriptMaker. No Qt anywhere in this file — it's driven
entirely by callbacks and a background listener/timer thread, so it can
be constructed and unit-tested with no QApplication in existence.

Trigger matching and action execution are both dispatched through the
pluggable registries in core/Inputs and core/Actions rather than being
hardcoded here — see those folders' registry.py for how to add a new
trigger or action kind.

Owns:
- the OSC input listener
- a pool of cached outgoing OSC clients
- the shared variable store
- the timer-trigger loop
- action-chain execution (each firing script runs its chain on its own
  daemon thread so a `wait` action never blocks the listener)
"""

from __future__ import annotations

import threading
import time

from core.Actions.context import ActionContext
from core.Actions.registry import run_action
from core.Inputs.registry import input_matches, timer_interval
from core.conditions import evaluate_condition
from core.osc_io import OSCListener, OSCSenderPool

MAX_ACTIONS_PER_EVENT = 200  # runaway-loop guard shared across a fire chain


class _BudgetExceeded(Exception):
    pass


class ScriptEngine:
    def __init__(self, get_scripts_cb, default_out_host="127.0.0.1",
                 default_out_port="9000", log_cb=None):
        self._get_scripts = get_scripts_cb
        self.default_out_host = default_out_host
        self.default_out_port = str(default_out_port)
        self._log = log_cb or (lambda msg: None)

        self._listener: OSCListener | None = None
        self._senders = OSCSenderPool()

        self._last_osc_values: dict[str, object] = {}
        self._variables: dict[str, object] = {}
        self._last_var_values: dict[str, object] = {}

        self._timer_next_fire: dict[int, float] = {}
        self._timer_thread: threading.Thread | None = None
        self._timer_stop = threading.Event()

    # ── Lifecycle ────────────────────────────────────────────────────────

    def start_listener(self, host: str, port: int):
        if self._listener and self._listener.is_running:
            return
        self._listener = OSCListener(host, int(port), self._on_osc_message)
        self._listener.start()
        self._start_timer_thread()

    def stop_listener(self):
        if self._listener:
            self._listener.stop()
            self._listener = None
        self._stop_timer_thread()

    @property
    def is_running(self) -> bool:
        return bool(self._listener and self._listener.is_running)

    def shutdown(self):
        self.stop_listener()
        self._senders.close_all()

    def _start_timer_thread(self):
        if self._timer_thread and self._timer_thread.is_alive():
            return
        self._timer_stop.clear()
        self._timer_next_fire.clear()
        self._timer_thread = threading.Thread(target=self._timer_loop, daemon=True)
        self._timer_thread.start()

    def _stop_timer_thread(self):
        self._timer_stop.set()
        if self._timer_thread:
            self._timer_thread.join(timeout=1)
        self._timer_thread = None

    # ── Variables ────────────────────────────────────────────────────────

    def get_variable(self, name: str, default=None):
        return self._variables.get(name, default)

    def set_variable(self, name: str, value, _budget=None):
        if not name:
            return
        prev = self._variables.get(name)
        self._variables[name] = value
        self._last_var_values[name] = prev
        self._check_variable_triggers(name, value, prev, _budget)

    # ── Trigger dispatch (via core/Inputs registry) ─────────────────────

    def _on_osc_message(self, address: str, value):
        prev = self._last_osc_values.get(address)
        self._last_osc_values[address] = value
        for script in self._safe_scripts():
            trig = script.trigger
            if not script.enabled or trig.kind != "osc":
                continue
            if not input_matches(trig, address, value):
                continue
            if evaluate_condition(trig.condition, value, trig.value, trig.value2, prev):
                self._fire(script, value)

    def _check_variable_triggers(self, name: str, value, prev, _budget):
        for script in self._safe_scripts():
            trig = script.trigger
            if not script.enabled or trig.kind != "variable":
                continue
            if not input_matches(trig, name, value):
                continue
            if evaluate_condition(trig.condition, value, trig.value, trig.value2, prev):
                self._fire(script, value, _budget)

    def _timer_loop(self):
        while not self._timer_stop.is_set():
            now = time.time()
            for script in self._safe_scripts():
                trig = script.trigger
                if not script.enabled or trig.kind != "timer":
                    continue
                interval = timer_interval(trig)
                next_fire = self._timer_next_fire.get(script.uid)
                if next_fire is None:
                    self._timer_next_fire[script.uid] = now + interval
                    continue
                if now >= next_fire:
                    self._timer_next_fire[script.uid] = now + interval
                    self._fire(script, None)
            self._timer_stop.wait(0.2)

    def _safe_scripts(self):
        try:
            return list(self._get_scripts())
        except Exception:
            return []

    # ── Firing / execution (via core/Actions registry) ──────────────────

    def _fire(self, script, trigger_value, _budget=None):
        self._log(f"\u25b6 {script.name} fired")
        t = threading.Thread(
            target=self._run_chain, args=(script, trigger_value, _budget), daemon=True
        )
        t.start()

    def _run_chain(self, script, trigger_value, _budget=None):
        budget = _budget if _budget is not None else {"n": 0}
        for action in script.actions:
            try:
                self._run_action(action, trigger_value, budget)
            except _BudgetExceeded:
                self._log(f"\u26a0 {script.name}: action limit reached, stopping (possible loop)")
                return
            except Exception as exc:
                self._log(f"\u26a0 {script.name}: action error \u2014 {exc}")

    def _run_action(self, action, trigger_value, budget):
        budget["n"] = budget.get("n", 0) + 1
        if budget["n"] > MAX_ACTIONS_PER_EVENT:
            raise _BudgetExceeded()

        ctx = ActionContext(
            trigger_value=trigger_value,
            default_host=self.default_out_host,
            default_port=self.default_out_port,
            senders=self._senders,
            set_variable=lambda name, val: self.set_variable(name, val, budget),
            run_sub_action=lambda sub: self._run_action(sub, trigger_value, budget),
        )
        run_action(action, ctx)
