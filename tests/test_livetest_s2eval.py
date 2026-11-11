# SPDX-License-Identifier: Apache-2.0
"""Story S2's page-open eval (A20). Seen live on 40f1b3e: 3 of 4 story runs (gw2 run 1, gw3 runs 1 and 2) failed S2
with "eval in .../onenote/taskpane.html: frame script timed out (frame detached?)" in the harness's stand-in for
the person clicking the decision page, AFTER the plugin had made and opened the page; S3 then never sent its prompt
(KeyError on the decision id) and S4-S6 ran against deliverables that did not exist. The eval walked every section
of the shared test notebook (one more section per run) and hit the harness window's fixed 20 s frame-script limit
(main.js evalIn): S2's post-turn time grew 16-17 s (5d840f0) -> 23-24 s (77b5652, gw2 run 2) -> timeout.
The harness now walks only the run's own section, names a frame-script timeout as a harness fault, never retries
it, and reports the later beats that depend on S2 as skipped instead of running them. Pure logic, no Office."""
import asyncio

import pytest

from test_livetest_story import _driver


class FakePane:
    def __init__(self, value=None, exc=None):
        self.value, self.exc, self.calls = value, exc, []

    def cmd(self, cmd, timeout=90, **kw):
        self.calls.append((cmd, kw))
        if self.exc:
            raise self.exc
        return {"value": self.value}

    def show(self):
        pass

    def shot(self, name):
        raise RuntimeError("no screenshots in unit tests")

    def frames(self):
        return ["top", "https://studio.example/office/live/onenote/taskpane.html"]


def test_a_frame_script_timeout_is_its_own_harness_error():
    d = _driver()
    p = d.ui.Pane.__new__(d.ui.Pane)
    p.cmd = FakePane(value={"__error": "frame script timed out (frame detached?)"}).cmd
    with pytest.raises(d.ui.FrameScriptTimeout) as e:
        p.run_js("https://x/onenote/taskpane.html", "1")
    assert isinstance(e.value, d.ui.HarnessError)
    assert "20 s" in str(e.value) and "harness" in str(e.value)
    p.cmd = FakePane(value={"__error": "no frame matching x"}).cmd
    with pytest.raises(d.ui.HarnessError) as e2:
        p.run_js("x", "1")
    assert not isinstance(e2.value, d.ui.FrameScriptTimeout)


def test_the_goto_walks_only_the_runs_section_when_it_is_known():
    d = _driver()
    js = d.onenote_goto_js("Decisions: pilot sync (HOL s1 1107-153117)", section="HOL s1 1107-153117")
    assert "items/id,items/name" in js and '"HOL s1 1107-153117"' in js
    assert "sections" in js and "ms:" in js  # how many sections it walked and how long, for the record
    whole = d.onenote_goto_js("Decisions: x")  # the loop keeps the whole-notebook walk
    assert "x.name ===" not in whole and '"Decisions: x"' in whole


class Tps(dict):
    pass


def test_story_post_decision_reports_the_timeout_once_and_never_retries(monkeypatch):
    d = _driver()
    seen = []

    class TP(FakePane):
        def run_js(self, frame, js, timeout=90):
            seen.append(js)
            raise d.ui.FrameScriptTimeout("eval in onenote: frame script timed out (frame detached?)")

    monkeypatch.setattr(d.time, "sleep", lambda s: None)
    c = {"fx": {"decision": "Decisions: pilot sync (HOL s1 1107-153117)", "section": "HOL s1 1107-153117"}}
    with pytest.raises(d.ui.FrameScriptTimeout) as e:
        d._story_post_decision({"onenote": TP()}, {"tools": []}, c)
    assert len(seen) == 1 and '"HOL s1 1107-153117"' in seen[0]
    msg = str(e.value)
    assert "OneNote page-open" in msg and "not retried" in msg and "plugin" in msg


def _rec():
    class Rec:
        out = None

        def save(self):
            pass

        def log(self, obj):
            pass
    return Rec()


def test_a_harness_timeout_in_s2_is_marked_as_a_harness_fault(monkeypatch):
    d = _driver()

    class Chat:
        approvals = []

        async def turn(self, text):
            return {"text": "done", "tools": [], "ms": 5}

    def post(tps, t, c):
        raise d.ui.FrameScriptTimeout("OneNote page-open eval: not retried")

    step = {"name": "S1_email_decision_page", "show": "onenote", "read": [], "post": post,
            "check": lambda t, r, c: (True, "x"), "prompt": lambda c: "make the page"}
    result = {"steps": []}
    s = asyncio.run(d._story_step(Chat(), {"onenote": FakePane()}, step, {}, _rec(), result, 2, 1))
    assert s["pass"] is False and s["fault"] == "harness" and "not retried" in s["error"]


def test_beats_that_need_a_failed_beat_are_skipped_without_a_turn():
    d = _driver()
    steps = {st["name"]: st for st in d.story_steps()}
    assert steps["S2_deliverables"]["needs"] == "S1_email_decision_page"
    assert steps["S3_why"]["needs"] == steps["S4_reply"]["needs"] == "S2_deliverables"
    assert steps["S5_undo"]["needs"] == "S4_reply"

    class Chat:
        approvals = []

        async def turn(self, text):
            raise AssertionError("no prompt is sent for a skipped beat")

    result = {"steps": [{"step": "S1_email_decision_page", "pass": False, "fault": "harness"}]}
    for i, name in enumerate(("S2_deliverables", "S3_why", "S4_reply", "S5_undo"), start=2):
        s = asyncio.run(d._story_step(Chat(), {h: FakePane() for h in ("onenote", "excel", "powerpoint", "outlook")},
                                      steps[name], {}, _rec(), result, i, 1))
        assert s["pass"] is False and s["skipped"].startswith("not run: needs ")
    assert [s["step"] for s in result["steps"]][1:] == ["S2_deliverables", "S3_why", "S4_reply", "S5_undo"]
    assert "S1_email_decision_page" in result["steps"][1]["skipped"]
    beats = d.story_beats(result)
    assert [b["pass"] for b in beats] == [False] * 5
