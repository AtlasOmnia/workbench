# SPDX-License-Identifier: Apache-2.0
"""OneNote add-in upload needs an open page (A46).

A45 (2026-11-09): the reopened test notebook landed in an empty section ("This section is empty. Click here or
press ENTER to create a new page."). With no page open, OneNote for the web greys out Insert > Office Add-ins
(pointer-events: none, no disabled/aria-disabled attribute), so the trusted click lands but no add-in dialog
opens, and every OneNote setup failed with "no visible control labelled ['MY ADD-INS', 'My add-ins']".
Reproduced live 2026-11-09 23:46 (probe in the A46 report): click in the empty section -> no WefGallery frame;
Add page -> pointer-events auto -> click -> WefGallery.htm frame.

Pure logic with a fake pane and a fake clock; nothing here touches Office, ssh or the network."""
import importlib.util
import sys
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parent.parent
LIVE = ROOT / "scripts" / "livetest"


def _ui():
    if "office_ui" in sys.modules:
        return sys.modules["office_ui"]
    sys.path.insert(0, str(LIVE))
    spec = importlib.util.spec_from_file_location("office_ui", LIVE / "office_ui.py")
    mod = importlib.util.module_from_spec(spec)
    sys.modules["office_ui"] = mod
    spec.loader.exec_module(mod)
    return mod


class Clock:
    def __init__(self):
        self.t = 0.0

    def now(self):
        return self.t

    def sleep(self, s):
        self.t += s


EMPTY = {"empty": True, "addins": "disabled"}
LOADING = {"empty": False, "addins": "disabled"}
READY = {"empty": False, "addins": "enabled"}
NO_TAB = {"empty": False, "addins": "missing"}


class FakePane:
    """run_js answers the page-state probe from a script; clicks are recorded and may advance the script."""

    def __init__(self, states, on_click=None):
        self.states = list(states)
        self.clicks = []
        self.on_click = on_click or {}

    def run_js(self, frame, js, timeout=90):
        assert frame == "F"
        return self.states.pop(0) if len(self.states) > 1 else self.states[0]

    def click_label(self, patterns, timeout=20, frames_like=None, sel=None):
        self.clicks.append(patterns[0])
        nxt = self.on_click.get(patterns[0])
        if nxt is not None:
            self.states = list(nxt)
        return patterns[0]


def test_empty_section_gets_one_page_then_waits_for_enabled_button():
    ui, clk, log = _ui(), Clock(), []
    p = FakePane([EMPTY], on_click={"Add page": [LOADING, LOADING, READY]})
    st = ui.ensure_onenote_page(p, log.append, frame="F", clock=clk.now, sleep=clk.sleep)
    assert st == READY
    assert p.clicks == ["Add page"]  # exactly one page, never a second while the first is loading
    assert any(e.get("ev") == "onenote-page-added" for e in log)


def test_open_page_with_enabled_button_clicks_nothing():
    ui, clk = _ui(), Clock()
    p = FakePane([READY])
    assert ui.ensure_onenote_page(p, lambda e: None, frame="F", clock=clk.now, sleep=clk.sleep) == READY
    assert p.clicks == []


def test_missing_button_reselects_insert_tab():
    ui, clk = _ui(), Clock()
    p = FakePane([NO_TAB], on_click={"Insert": [READY]})
    ui.ensure_onenote_page(p, lambda e: None, frame="F", clock=clk.now, sleep=clk.sleep)
    assert p.clicks == ["Insert"]


def test_greyed_button_is_a_bounded_harness_error_with_at_most_two_pages():
    ui, clk = _ui(), Clock()
    p = FakePane([EMPTY])  # Add page never takes
    with pytest.raises(ui.HarnessError, match="greyed out"):
        ui.ensure_onenote_page(p, lambda e: None, frame="F", timeout=60, clock=clk.now, sleep=clk.sleep)
    assert p.clicks.count("Add page") <= 2
    assert clk.t <= 65


def test_page_state_js_reads_pointer_events_and_the_empty_section_prompt():
    js = _ui().ONENOTE_PAGE_JS
    assert "pointer-events" in js or "pointerEvents" in js
    assert "This section is empty" in js
    assert "Office Add-ins" in js


def test_upload_addin_onenote_ensures_a_page_before_clicking_office_addins(monkeypatch):
    ui = _ui()
    order = []

    class P:
        def click_label(self, patterns, timeout=20, frames_like=None, sel=None):
            order.append(patterns[0])
            if patterns[0] == "MY ADD-INS":
                raise ui.HarnessError("stop here")
            return patterns[0]

        def cmd(self, *a, **k):
            return {}

    monkeypatch.setattr(ui, "ensure_onenote_page", lambda pane, log, **k: order.append("ensure-page") or READY)
    monkeypatch.setattr(ui.time, "sleep", lambda s: None)
    with pytest.raises(ui.HarnessError, match="stop here"):
        ui.upload_addin(P(), "onenote", "/m.xml", lambda e: None)
    assert order.index("ensure-page") < order.index("Office Add-ins")
