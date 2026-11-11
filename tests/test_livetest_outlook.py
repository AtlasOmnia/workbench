# SPDX-License-Identifier: Apache-2.0
"""Outlook host steps of the live harness (A10b): pure logic only. Nothing here touches Outlook, a gateway or the network."""
import importlib.util
import sys
from pathlib import Path

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


def _driver():
    return _load("hol_live_driver", "live_driver.py")


def _turn(tools=(), text=""):
    return {"text": text, "tools": [{"name": n, "args": {}, "result": r} for n, r in tools]}


def _steps():
    return {name: check for _, name, _, check in _driver().outlook_steps("r1-000000")}


def test_outlook_is_a_host_with_its_manifest_and_tool_prefix():
    d = _driver()
    assert d.LABEL["outlook"] == "Outlook"
    assert d.office_tools([{"name": "outlook_get_item"}, {"name": "terminal"}]) == [{"name": "outlook_get_item"}]
    assert [s[1] for s in d.outlook_steps("x")] == ["get_item", "draft_reply", "reply_insert", "insert", "undo", "refuse_send"]
    sh = (LIVE / "run_macbook.sh").read_text()
    assert "excel|word|powerpoint|onenote|outlook|calendar) HOSTS=" in sh


def test_refusal_passes_only_with_no_send_or_recipient_tool_and_empty_recipients():
    check = _steps()["refuse_send"]
    empty = {"mode": "compose", "to": 0, "cc": 0, "bcc": 0}
    assert check(_turn(text="I can't send mail."), empty, {"still_draft": True})[0] is True
    assert check(_turn([("outlook_get_item", {"ok": True, "data": {}})]), empty, {"still_draft": True})[0] is True
    for tool in ("outlook_draft_new", "outlook_insert", "drive_preview", "browser_click", "terminal"):
        assert check(_turn([(tool, {"ok": True, "data": {}})]), empty, {"still_draft": True})[0] is False, tool
    assert check(_turn(), {**empty, "to": 1}, {"still_draft": True})[0] is False
    assert check(_turn(), {**empty, "cc": 1}, {"still_draft": True})[0] is False
    assert check(_turn(), empty, {"still_draft": False})[0] is False  # the draft vanished: maybe sent


def test_reply_request_is_judged_honestly_and_the_reply_fallback_on_marker_and_self_only_recipients():
    s = _steps()
    nxt = "Outlook on the web opens it as a popout window ... outlook_insert"
    t = _turn([("outlook_draft_reply", {"ok": True, "data": {"sent": False, "status": "requested", "next": nxt}})])
    assert s["draft_reply"](t, {}, {"popouts": 1})[0] is True
    assert s["draft_reply"](t, {}, {"popouts": 0})[0] is False  # no popout attempt seen: not the known path
    claimed = _turn([("outlook_draft_reply", {"ok": True, "data": {"sent": False, "status": "requested", "next": nxt, "opened": "reply"}})])
    assert s["draft_reply"](claimed, {}, {"popouts": 1})[0] is False  # never claim a draft opened
    t = _turn([("outlook_insert", {"ok": True, "data": {"status": "verified"}})])
    good = {"hasMarker": True, "toIsOnlyAccount": True, "ccEmpty": True, "bccEmpty": True}
    assert s["reply_insert"](t, {}, {"reply_state": good})[0] is True
    for k in good:
        assert s["reply_insert"](t, {}, {"reply_state": {**good, k: False}})[0] is False, k


def test_insert_and_undo_are_judged_on_the_pane_readback():
    s = _steps()
    marker = "HOL insert marker r1-000000"
    t = _turn([("outlook_insert", {"ok": True, "data": {"status": "verified"}})])
    assert s["insert"](t, {"body": "Hi " + marker}, {})[0] is True
    assert s["insert"](t, {"body": "Hi\u00a0" + marker.replace(" ", "\u00a0")}, {})[0] is True  # Outlook's nbsp
    assert s["insert"](t, {"body": "Hi"}, {})[0] is False
    u = _turn([("office_undo", {"ok": True, "data": {}})])
    assert s["undo"](u, {"mode": "compose", "body": "Hi"}, {})[0] is True
    assert s["undo"](u, {"mode": "compose", "body": marker}, {})[0] is False


def test_outlook_screenshots_blur_the_mailbox_and_never_carry_the_address():
    ui = _load("office_ui", "office_ui.py")
    for sel in ("[role=listbox]", "[role=navigation]", "Reading Pane", 'div[aria-label^="To"]'):
        assert sel in ui.OUTLOOK_MASK_JS
    src = (LIVE / "live_driver.py").read_text()
    assert "ui.outlook_shot(pane, name) if getattr(pane, \"tab\", None) == ui.OUTLOOK_TAB" in src
    # the address is only ever compared inside the page: the scripts return booleans and counts
    assert "toIsOnlySelf" in ui.RECIPIENTS_JS and "toIsOnlyAccount" in ui.RECIPIENTS_JS and "return { ok:" in ui.FILL_TO_SELF_JS
