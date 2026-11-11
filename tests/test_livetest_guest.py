# SPDX-License-Identifier: Apache-2.0
"""Harness start-up: wait (bounded) for the client Mac window's webview guest before the first navigate.

Pure logic with a fake clock and a fake control endpoint; nothing here touches Office, ssh or the network."""
import asyncio
import importlib.util
import json
import sys
import types
import urllib.error
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parent.parent
LIVE = ROOT / "scripts" / "livetest"


def _load(name, file):
    if name in sys.modules:
        return sys.modules[name]
    sys.path.insert(0, str(LIVE))
    spec = importlib.util.spec_from_file_location(name, LIVE / file)
    mod = importlib.util.module_from_spec(spec)
    sys.modules[name] = mod
    spec.loader.exec_module(mod)
    return mod


def _ui():
    return _load("office_ui", "office_ui.py")


class Clock:
    def __init__(self):
        self.t = 0.0

    def now(self):
        return self.t

    def sleep(self, s):
        self.t += s


def _pane(answers, log=None):
    """A Pane whose control endpoint answers `state` from a list (dict, or an exception to raise)."""
    ui = _ui()
    p = ui.Pane("http://127.0.0.1:1", (log if log is not None else []).append)
    seen = []

    def cmd(c, timeout=90, **kw):
        seen.append(c)
        a = answers.pop(0) if len(answers) > 1 else answers[0]
        if isinstance(a, BaseException):
            raise a
        return a
    p.cmd = cmd
    p.seen = seen
    return p


NO_GUEST = {"url": None, "title": None, "loading": None, "frames": [], "tab": None, "tabs": []}
GUEST = {"url": "https://excel.cloud.microsoft/", "title": "Excel", "loading": False, "frames": [], "tab": "main",
         "tabs": [{"name": "main", "visible": True}]}


def test_wait_guest_returns_once_the_guest_attaches():
    ui, clock, log = _ui(), Clock(), []
    p = _pane([NO_GUEST, NO_GUEST, NO_GUEST, GUEST], log)
    st = p.wait_guest(timeout=90, clock=clock.now, sleep=clock.sleep)
    assert st["tab"] == "main" and p.seen == ["state"] * 4 and clock.t < 10
    assert any(e.get("ev") == "guest-attached" for e in log)


def test_wait_guest_tolerates_an_unreachable_control_endpoint_while_the_window_starts():
    ui, clock = _ui(), Clock()
    p = _pane([urllib.error.URLError("refused"), ConnectionResetError("reset"), ui.HarnessError("state: boom"), GUEST])
    assert p.wait_guest(timeout=90, clock=clock.now, sleep=clock.sleep)["tab"] == "main"


def test_a_guest_that_never_attaches_is_reported_as_an_unavailable_gui_session():
    ui, clock, log = _ui(), Clock(), []
    p = _pane([NO_GUEST], log)
    with pytest.raises(ui.GuestUnavailable) as e:
        p.wait_guest(timeout=90, clock=clock.now, sleep=clock.sleep)
    msg = str(e.value)
    assert "client Mac GUI session unavailable (locked/asleep/lid closed)" in msg and "90 s" in msg
    assert isinstance(e.value, ui.HarnessError)
    assert 90 <= clock.t <= 92  # bounded
    assert "navigate" not in p.seen
    waiting = [x for x in log if x.get("ev") == "guest-waiting"]
    assert waiting and len(waiting) <= 8  # a few progress lines, not one per poll


def test_an_endpoint_that_never_answers_is_also_bounded():
    ui, clock = _ui(), Clock()
    p = _pane([urllib.error.URLError("refused")])
    with pytest.raises(ui.GuestUnavailable) as e:
        p.wait_guest(timeout=90, clock=clock.now, sleep=clock.sleep)
    assert "control endpoint" in str(e.value) and clock.t <= 92


def test_driver_waits_for_the_guest_before_the_first_navigate_and_stops_distinctly(tmp_path, monkeypatch):
    d = _load("hol_live_driver", "live_driver.py")
    ui = d.ui
    seen = []

    class FakePane:
        def __init__(self, base, log):
            self.base = base

        def wait_guest(self, timeout=90, **kw):
            seen.append(("wait_guest", timeout))
            raise ui.GuestUnavailable("client Mac GUI session unavailable (locked/asleep/lid closed): test")

        def cmd(self, c, timeout=90, **kw):
            seen.append(c)
            return {}
    monkeypatch.setattr(ui, "Pane", FakePane)
    a = types.SimpleNamespace(out=str(tmp_path / "out"), control="http://127.0.0.1:1", mac_user="x", hosts=["excel"], runs=1,
                              f3=False, loop=False, images=False, signin_timeout=0)
    rc = asyncio.run(d.main_async(a))
    assert rc == d.EXIT_GUI_UNAVAILABLE and rc not in (0, 2, 3)
    assert seen == [("wait_guest", d.GUEST_WAIT_S)] and d.GUEST_WAIT_S == 90
    summary = json.loads((tmp_path / "out" / "summary.json").read_text())
    assert summary["blocked"] == "client Mac GUI session unavailable (locked/asleep/lid closed)"


def test_driver_proceeds_to_navigate_after_the_guest_attached(tmp_path, monkeypatch):
    d = _load("hol_live_driver", "live_driver.py")
    ui = d.ui
    seen = []

    class Stop(Exception):
        pass

    class FakePane:
        def __init__(self, base, log):
            pass

        def wait_guest(self, timeout=90, **kw):
            seen.append("wait_guest")
            return GUEST

        def cmd(self, c, timeout=90, **kw):
            seen.append(c)
            if c == "navigate":
                raise Stop()
            return {}
    monkeypatch.setattr(ui, "Pane", FakePane)
    a = types.SimpleNamespace(out=str(tmp_path / "out"), control="http://127.0.0.1:1", mac_user="x", hosts=["excel"], runs=1,
                              f3=False, loop=False, images=False, signin_timeout=0)
    with pytest.raises(Stop):
        asyncio.run(d.main_async(a))
    assert seen[0] == "wait_guest" and seen.index("navigate") > 0


# ---- wait_pane_state: a pane whose Office.js failed to load is reloaded once (A8) ------------------------------

class _FakePane:
    """A pane on the control endpoint: ``states`` are what pane_state reads, one per poll (the last repeats)."""
    FRAME = "https://studio.example/office/live/excel/taskpane.html"

    def __init__(self, states):
        self.states, self.cmds, self.events = list(states), [], []

    def log(self, ev):
        self.events.append(ev)

    def frames(self):
        return [self.FRAME]

    def run_js(self, frame, js, timeout=90):
        return self.states.pop(0) if len(self.states) > 1 else self.states[0]

    def cmd(self, c, **kw):
        self.cmds.append((c, kw))
        return {"value": True}


_FAILED = {"state": "starting", "status": "Could not load Office.js from https://appsforoffice.microsoft.com/lib/1/hosted/office.js. Check the network, then reload this pane."}


def test_wait_pane_state_reloads_a_pane_whose_office_js_failed_once():
    ui, clock = _ui(), Clock()
    p = _FakePane([_FAILED, {"state": "starting", "status": ""}, {"state": "unpaired", "status": "Not connected yet."}])
    st = ui.wait_pane_state(p, "excel", ("ready", "unpaired"), timeout=60, clock=clock.now, sleep=clock.sleep)
    assert st["state"] == "unpaired"
    assert p.cmds == [("reloadFrame", {"frame": _FakePane.FRAME})]
    assert {"ev": "pane-reload-office-js-failed", "host": "excel"} in p.events


def test_wait_pane_state_reports_a_second_office_js_failure_instead_of_reloading_again():
    ui, clock = _ui(), Clock()
    p = _FakePane([_FAILED])
    with pytest.raises(ui.HarnessError, match="Could not load Office.js"):
        ui.wait_pane_state(p, "excel", ("ready", "unpaired"), timeout=30, clock=clock.now, sleep=clock.sleep)
    assert [c for c, _ in p.cmds] == ["reloadFrame"]


def test_wait_pane_state_never_reloads_a_pane_that_is_just_starting():
    ui, clock = _ui(), Clock()
    p = _FakePane([{"state": "starting", "status": "Loading Office.js"}])
    with pytest.raises(ui.HarnessError):
        ui.wait_pane_state(p, "excel", ("ready",), timeout=10, clock=clock.now, sleep=clock.sleep)
    assert p.cmds == []
