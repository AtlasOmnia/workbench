# SPDX-License-Identifier: Apache-2.0
"""Demo-story harness mode (live_driver.py --story): pure logic only. Nothing here touches Office, a gateway or the
network."""
import asyncio
import datetime as dt
import importlib.util
import sqlite3
import struct
import subprocess
import sys
import zlib
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


EMAIL = "email:AAMkAGI2"
DTITLE = "Decisions: pilot sync (HOL s1 1107-030000)"
ITEMS = [{"id": 1, "kind": "decision", "text": "We ship the beta on Friday", "owner": "Unknown", "quoteVerified": True},
         {"id": 2, "kind": "decision", "text": "We keep the current pilot scope", "owner": "Unknown", "quoteVerified": True},
         {"id": 3, "kind": "action", "text": "Write the FAQ", "owner": "Sam", "quoteVerified": True}]


def _ctx(d, **kw):
    c = {"run": 1, "fx": d.story_fixture(1, "1107-030000", today=dt.date(2026, 11, 7), offset_minutes=-300),
         "email": {"pageId": EMAIL, "title": d.EMAIL_SOURCE}, "img": {"path": "/p/cache/images/demo/team.png"}}
    c.update(kw)
    return c


def _tool(name, data, ok=True):
    return {"name": name, "result": {"ok": ok, "data": data} if ok else {"ok": False, "error": data}}


def _with_decision(d):
    c = _ctx(d)
    c["decision"] = {"title": DTITLE, "pageId": "p-dec", "items": ITEMS}
    return c


# ---- the script is exactly docs/demo-script.md's prompts, in this order --------------------------------------
def test_story_prompts_come_from_docs_demo_script_md_in_order():
    d = _driver()
    p = d.story_prompts()
    assert len(p) == 5
    assert p[0] == ("Read this email and turn it into a decision page in my notes, quoting the email for each decision "
                    "and action.") and p[4] == "Undo that."
    steps = d.story_steps()
    assert [s["name"] for s in steps] == ["S1_email_decision_page", "S2_deliverables", "S3_why", "S4_reply", "S5_undo"]
    c = _with_decision(d)
    c["slides0"] = 1
    for s, want in zip(steps, p):
        assert s["prompt"](c).startswith(want)  # verbatim, run details only appended
    s1 = steps[0]["prompt"](c)
    assert "brain_from_email once with title" in s1 and "decisions and actions" in s1 and "index only" not in s1
    s3 = steps[1]["prompt"](c)
    assert "hosts excel, word, powerpoint, outlook and calendar" in s3 and "slideIndex 1, left 600, top 300" in s3
    assert "'/p/cache/images/demo/team.png'" in s3 and "2026-11-12T10:00:00-05:00" in s3
    assert [s["show"] for s in steps] == ["outlook", "powerpoint", "excel", "outlook", "outlook"]
    assert steps[0]["then_show"] == "onenote"  # the person types in Outlook, the result is the OneNote page


def test_next_thursday_is_never_today_and_carries_the_offset():
    d = _driver()
    assert d.next_thursday(dt.date(2026, 11, 7), -300) == ("2026-11-12T10:00:00-05:00", "2026-11-12T10:30:00-05:00")  # Sat
    assert d.next_thursday(dt.date(2026, 11, 12), 60)[0] == "2026-11-19T10:00:00+01:00"  # a Thursday: next week
    assert d.next_thursday(dt.date(2026, 11, 11), 0)[0] == "2026-11-12T10:00:00+00:00"


def test_story_fixture_is_run_tagged():
    d = _driver()
    a = d.story_fixture(1, "1107-030000", today=dt.date(2026, 11, 7), offset_minutes=-300)
    b = d.story_fixture(2, "1107-031500", today=dt.date(2026, 11, 7), offset_minutes=-300)
    assert a["section"] == "HOL s1 1107-030000" and a["decision"] == DTITLE
    assert a["decision"] != b["decision"] and a["followUp"]["location"] == "Studio"


# ---- the picture is a local PNG fixture, written by the harness (no image generation) -----------------------
def test_story_image_is_a_valid_png_in_the_profile_image_cache(tmp_path):
    d = _driver()
    info = d.story_image(tmp_path)
    p = Path(info["path"])
    assert p == tmp_path / "cache" / "images" / "demo" / "team.png" and info["rel"] == "cache/images/demo/team.png"
    data = p.read_bytes()
    assert data[:8] == b"\x89PNG\r\n\x1a\n" and len(data) == info["bytes"] < 4 * 1024 * 1024
    w, h = struct.unpack(">II", data[16:24])
    assert (w, h) == (320, 180)
    idat = data[data.index(b"IDAT") + 4:data.index(b"IEND") - 8]
    assert len(zlib.decompress(idat)) == h * (1 + 3 * w)
    assert d.story_image(tmp_path)["sha256"] == info["sha256"]  # deterministic, reused


def test_story_image_never_overwrites_a_real_picture_and_refuses_a_populated_legacy_image_cache(tmp_path):
    """A26: the path is the presenter's own (cache/images/demo/team.png); a different file there is someone's real
    picture. A populated legacy image_cache/ makes the plugin read that folder instead of cache/images."""
    d = _driver()
    real = tmp_path / "a" / "cache" / "images" / "demo" / "team.png"
    real.parent.mkdir(parents=True)
    real.write_bytes(b"\x89PNG\r\n\x1a\nreal team photo")
    try:
        d.story_image(tmp_path / "a")
        raise AssertionError("overwrote a real picture")
    except d.ui.HarnessError as exc:
        assert "never overwritten" in str(exc)
    assert real.read_bytes().endswith(b"real team photo")
    (tmp_path / "b" / "image_cache").mkdir(parents=True)
    (tmp_path / "b" / "image_cache" / "x.png").write_bytes(b"x")
    try:
        d.story_image(tmp_path / "b")
        raise AssertionError("ignored the legacy image_cache")
    except d.ui.HarnessError as exc:
        assert "image_cache" in str(exc)
    assert not (tmp_path / "b" / "cache").exists()


def test_prompt_2_names_the_picture_file_and_the_demo_script_names_the_harness_path():
    """A26: the bare prompt names the picture by file name (a cold Hermes cannot know which picture 'our team picture'
    is), and the set-up names the exact path the harness writes, so presenter and harness type the same words."""
    d = _driver()
    p = d.story_prompts()
    assert d.STORY_IMAGE_NAME == "team.png" and d.STORY_IMAGE_REL.endswith("/" + d.STORY_IMAGE_NAME)
    assert "put team.png from my Hermes images folder on the title slide" in p[1]
    assert "team picture" not in p[1]
    doc = " ".join((ROOT / "docs" / "demo-script.md").read_text(encoding="utf-8").split())
    assert f"`{d.STORY_IMAGE_REL}`" in doc
    harness = " ".join((ROOT / "docs" / "harness.md").read_text(encoding="utf-8").split())
    assert f"`{d.STORY_IMAGE_REL}`" in harness


# ---- the fixture email is reused, never sent ---------------------------------------------------------------
def test_find_fixture_reuses_inbox_then_sent_items_and_never_sends():
    d = _driver()
    asked = []

    def count(n):
        def f(folder):
            asked.append(folder)
            return n.get(folder, 0)
        return f
    assert d.find_fixture(None, count({"Inbox": 1}))["folder"] == "Inbox"
    assert d.find_fixture(None, count({"Sent Items": 1}))["folder"] == "Sent Items"
    try:
        d.find_fixture(None, count({}))
        raise AssertionError("must raise")
    except d.ui.HarnessError as exc:
        assert "never sends" in str(exc)
    assert set(asked) == {"Inbox", "Sent Items"}
    src = Path(d.__file__).read_text()
    story = src[src.index("# ---- --story"):src.index("def summarize(results)")]
    assert "send_fixture_once" not in story and "create_test_event_once" not in story


def test_addresses_are_masked_everywhere_in_a_value():
    d = _driver()
    v = {"a": "From: Pat <pat.x@example.com>", "b": ["to x_y@mail.example.org", 3], "c": None}
    assert d.mask_addresses(v) == {"a": "From: <name> <<address>>", "b": ["to <address>", 3], "c": None}  # A20: names too


# ---- beat checks -----------------------------------------------------------------------------------------
def test_email_index_check_needs_the_email_entry_and_read_mode():
    d = _driver()
    c = _ctx(d)
    c.pop("email")
    ok_t = {"tools": [_tool("brain_from_email", {"indexed": {"pageId": EMAIL, "title": d.EMAIL_SOURCE, "notebookId": "nb"}})]}
    rb = {"outlook": {"mode": "read", "subjectIsFixture": True}}
    ok, note = d.email_index_check(ok_t, rb, c)
    assert ok, note
    assert c["email"]["pageId"] == EMAIL
    with_page = {"tools": [_tool("brain_from_email", {"indexed": {"pageId": EMAIL, "title": d.EMAIL_SOURCE}, "decisionPage": {}})]}
    assert d.email_index_check(with_page, rb, _ctx(d))[0]  # A25: the same call makes the decision page
    assert not d.email_index_check(ok_t, {"outlook": {"mode": "compose"}}, _ctx(d))[0]
    assert not d.email_index_check({"tools": []}, rb, _ctx(d))[0]


def _decision_rbs(d, rows=None, page_extra=""):
    text = [f"Source: {d.EMAIL_SOURCE} Decisions D1. We ship the beta on Friday · owner: Unknown D2. We keep the current pilot scope "
            f"· owner: Unknown Actions A1. Write the FAQ · owner: Sam · due: 12 November {page_extra}"]
    return {"onenote_decision": {"pageTitle": DTITLE, "text": text},
            "db": {"rows": rows if rows is not None else [{"kind": "decision", "text": "x", "owner": None, "source": EMAIL,
                                                           "quoteVerified": True}] * 3}}


def test_story_decision_check_requires_citations_to_resolve_to_the_email_on_page_and_in_brain_db():
    d = _driver()
    t = {"tools": [_tool("brain_from_email", {"indexed": {"pageId": EMAIL}, "decisionPage": {
        "pageId": "p-dec", "title": DTITLE, "items": ITEMS, "unknownOwners": 2, "unverifiedQuotes": 0}})]}
    c = _ctx(d)
    ok, note = d.story_decision_check(t, _decision_rbs(d), c)
    assert ok, note
    assert c["decision"]["pageId"] == "p-dec"
    other = [{"kind": "decision", "text": "x", "owner": None, "source": "page-1", "quoteVerified": True}]
    assert "cite the email" in d.story_decision_check(t, _decision_rbs(d, rows=other), _ctx(d))[1]
    assert "unreadable" in d.story_decision_check(t, {**_decision_rbs(d), "db": {"error": "no ledger"}}, _ctx(d))[1]
    assert "unverified" in d.story_decision_check(t, _decision_rbs(d, page_extra="(not found in the source page)"), _ctx(d))[1]
    both = {"tools": [_tool("brain_from_email", {"indexed": {"pageId": EMAIL, "title": d.EMAIL_SOURCE}, "decisionPage": {
        "pageId": "p-dec", "title": DTITLE, "items": ITEMS, "unknownOwners": 2, "unverifiedQuotes": 0}})]}
    rbs = {**_decision_rbs(d), "outlook": {"mode": "read", "subjectIsFixture": True}}
    ok, note = d.story_email_decision_check(both, rbs, _ctx(d))  # S1: indexed AND the page, one call
    assert ok, note
    assert not d.story_email_decision_check({"tools": [_tool("brain_from_email", {"indexed": {"pageId": EMAIL, "title": d.EMAIL_SOURCE}})]},
                                            rbs, _ctx(d))[0]  # indexed only: no decision page
    wrong_owner = [dict(i, owner="Unknown") if i["kind"] == "action" else i for i in ITEMS]
    t2 = {"tools": [_tool("brain_decision_page", {"pageId": "p-dec", "title": DTITLE, "items": wrong_owner})]}
    assert "Sam" in d.story_decision_check(t2, _decision_rbs(d), _ctx(d))[1]


def _tracker_rb(d, source=None):
    return {"values": [["#", "Item", "Owner", "Due", "Status", "Source", "", "Open items", 1],
                       ["A1", "Write the FAQ", "Sam", "12 November", "Open", source or d.EMAIL_SOURCE, "", "Owner", "Open items"]],
            "formulas": [["#", "Item", "Owner", "Due", "Status", "Source", "", "Open items", '=COUNTIF(E2:E2,"Open")']],
            "charts": [{"name": "c", "title": "Open items by owner"}]}


def _brief_rb(d):
    return {"paragraphs": ["HOL live test document.", f"{DTITLE} — brief", "Decisions",
                           f"D1. We ship the beta on Friday (owner: Unknown). Source: {d.EMAIL_SOURCE}",
                           f"D2. We keep the current pilot scope (owner: Unknown). Source: {d.EMAIL_SOURCE}",
                           "Actions", f"A1. Write the FAQ (owner: Sam, due: 12 November). Source: {d.EMAIL_SOURCE}",
                           f"Decision record: {DTITLE}"],
            "links": [{"text": DTITLE, "hyperlink": True}]}


DECK_BODY = {"What we decided": "We ship the beta on Friday\nWe keep the current pilot scope",
             "Who does what": "Write the FAQ — Sam (12 November)"}


def _deck(base=1, made=None, body=None):
    made = made or [DTITLE, "What we decided", "Who does what", "Sources"]
    body = DECK_BODY if body is None else body
    return made, {"count": base + len(made), "slides": [["HOL"]] + [[m, body.get(m, "")] for m in made]}


def test_story_deck_check_rejects_decisions_listed_as_open_work():
    """A24: an ownerless decision stays on 'What we decided'; a 'Still open' slide or a decision under 'Action owner
    needed' fails the deck check."""
    d = _driver()
    c = _ctx(d)
    c["decision"] = {"title": DTITLE, "items": ITEMS}
    c["slides0"] = 1

    def check(made, body):
        _, rb = _deck(made=made, body=body)
        receipt = {"results": [{"host": "powerpoint", "slides": [{"title": m} for m in made]}]}
        return d.deck_check(rb, receipt, c)
    ok, note = check(None or [DTITLE, "What we decided", "Who does what", "Sources"], DECK_BODY)
    assert ok, note
    ok, note = check([DTITLE, "What we decided", "Who does what", "Still open", "Sources"],
                     {**DECK_BODY, "Still open": "No owner yet: We ship the beta on Friday"})
    assert not ok and "Still open" in note
    ok, note = check([DTITLE, "What we decided", "Action owner needed", "Sources"],
                     {"What we decided": "We keep the current pilot scope",
                      "Action owner needed": "We ship the beta on Friday"})
    assert not ok and "'We ship the beta on Friday' is not on 'What we decided'" in note
    assert "listed under 'Action owner needed'" in note


def _deliverables_turn(d, calendar="host", images=1):
    made, _ = _deck()
    results = [{"host": "excel", "status": "verified", "locator": "Decision tracker!A1:F2"},
               {"host": "word", "status": "verified"},
               {"host": "powerpoint", "status": "verified", "slides": [{"title": m} for m in made]},
               {"host": "outlook", "status": "applied", "requested": "reply", "sent": False, "replyText": "Thanks ..."}]
    if calendar == "host":
        results.append({"host": "calendar", "status": "applied", "requested": "appointment", "saved": False, "sent": False})
    tools = [_tool("brain_make_deliverables", {"status": "applied", "results": results})]
    if calendar == "draft_new":
        tools.append(_tool("calendar_draft_new", {"status": "requested", "saved": False, "sent": False}))
    tools += [_tool("office_insert_image", {"status": "verified"})] * images
    return {"tools": tools}


def test_story_deliverables_check_reads_every_document_and_accepts_an_honest_follow_up_request():
    d = _driver()
    _, deck = _deck()
    rbs = {"excel": _tracker_rb(d), "word": _brief_rb(d), "powerpoint": deck, "powerpoint_images": {"images": 1}}
    for cal in ("host", "draft_new"):
        c = _with_decision(d)
        c.update(slides0=1, img_before={"images": 0})
        ok, note = d.story_deliverables_check(_deliverables_turn(d, cal), rbs, c)
        assert ok, (cal, note)
        assert c["replyText"] == "Thanks ..."
    c = _with_decision(d)
    c.update(slides0=1, img_before={"images": 0})
    assert "follow-up" in d.story_deliverables_check(_deliverables_turn(d, None), rbs, c)[1]
    assert "pictures" in d.story_deliverables_check(_deliverables_turn(d), {**rbs, "powerpoint_images": {"images": 2}}, c)[1]
    assert "2 office_insert_image" in d.story_deliverables_check(_deliverables_turn(d, images=2), rbs, c)[1]
    assert "source" in d.story_deliverables_check(_deliverables_turn(d), {**rbs, "excel": _tracker_rb(d, "Q4 Launch")}, c)[1]
    saved = _deliverables_turn(d)
    saved["tools"][0]["result"]["data"]["results"][4]["saved"] = True
    assert "calendar host" in d.story_deliverables_check(saved, rbs, c)[1]


def test_story_brief_check_needs_the_email_source_and_the_decision_record_link():
    d = _driver()
    c = _with_decision(d)
    assert d.story_brief_check(_brief_rb(d), c)[0]
    rb = _brief_rb(d)
    rb["links"] = []
    assert "hyperlink" in d.story_brief_check(rb, c)[1]
    rb = _brief_rb(d)
    rb["paragraphs"] = [p.replace(d.EMAIL_SOURCE, "somewhere") for p in rb["paragraphs"]]
    assert "cites the email" in d.story_brief_check(rb, c)[1]


def test_story_why_check_follows_the_row_to_this_runs_decision_and_the_email():
    d = _driver()
    c = _with_decision(d)
    chain = {"item": "Write the FAQ", "quote": "Action: Sam writes the FAQ by 12 November.", "quoteVerified": True,
             "decisionPage": {"title": DTITLE}, "sourcePage": {"pageId": EMAIL, "title": d.EMAIL_SOURCE}}
    t = {"tools": [_tool("brain_why", {"found": True, "artifacts": [{"host": "excel", "because": [chain]}]})],
         "text": "Because of the email 'HOL live test: pilot sync' ..."}
    assert d.story_why_check(t, {}, c)[0]
    old = dict(chain, decisionPage={"title": "Decisions: pilot sync (HOL s1 1106-000000)"})
    t_old = {**t, "tools": [_tool("brain_why", {"artifacts": [{"host": "excel", "because": [old]}]})]}
    assert not d.story_why_check(t_old, {}, c)[0]
    assert "cite the email" in d.story_why_check({**t, "text": "Because."}, {}, c)[1]


def test_story_reply_and_undo_checks_need_the_draft_unsent_and_addressed_to_the_account_only():
    d = _driver()
    c = _with_decision(d)
    marker = d.story_reply_marker(c)
    good = {"hasMarker": True, "toIsOnlyAccount": True, "ccEmpty": True, "bccEmpty": True}
    # As seen live (A18 story r1): Hermes core's tool_describe (a deferred-tool schema lookup, not an Office action) runs
    # before outlook_insert, the pane reads back under 'outlook_compose', and Outlook turns spaces into non-breaking ones.
    t = {"tools": [{"name": "tool_describe", "result": {"tools": []}}, _tool("outlook_insert", {"status": "verified", "sent": False})]}
    rb = {"outlook_compose": {"mode": "compose", "body": f"Thanks for the notes. (Prepared with Hermes from the {marker}.)"
                                                         .replace(" ", "\u00a0") + "From: <address>", "cc": 0, "bcc": 0}}
    c.update(reply_state=good, still_draft=True)
    assert d.story_reply_check(t, rb, c)[0]
    assert "other tools" in d.story_reply_check({"tools": t["tools"] + [_tool("outlook_draft_new", {})]}, rb, c)[1]
    assert "no longer open" in d.story_reply_check(t, rb, {**c, "still_draft": False})[1]
    assert "on screen" in d.story_reply_check(t, rb, {**c, "reply_state": {**good, "toIsOnlyAccount": False}})[1]
    u = {"tools": [{"name": "tool_describe", "result": {}}, _tool("office_undo", {"undone": "insert"})]}
    after = {"outlook_compose": {"mode": "compose", "body": "\n\n\nFrom: <address>", "cc": 0, "bcc": 0}}
    c.update(undo_state={**good, "hasMarker": False}, still_draft=True)
    assert d.story_undo_check(u, after, c)[0]
    assert "still shows" in d.story_undo_check(u, rb, c)[1]
    assert "2 office_undo" in d.story_undo_check({"tools": u["tools"] * 2}, after, c)[1]
    assert "other tools" in d.story_undo_check({"tools": u["tools"] + [_tool("outlook_draft_reply", {})]}, after, c)[1]


def test_story_decisions_db_is_read_only(tmp_path):
    d = _driver()
    db = tmp_path / "brain.db"
    con = sqlite3.connect(db)
    con.execute("create table decisions(id integer primary key, notebook_id text, decision_page_id text, source_page_id text,"
                " kind text, text text, owner text, due text, rationale text, quote text, quote_verified integer, created_at real)")
    con.execute("insert into decisions(notebook_id, decision_page_id, source_page_id, kind, text, owner, quote_verified)"
                " values ('nb', 'p-dec', ?, 'action', 'Write the FAQ', 'Sam', 1)", (EMAIL,))
    con.commit()
    con.close()
    before = db.read_bytes()
    out = d.story_decisions_db(db, "p-dec")
    assert out == {"rows": [{"kind": "action", "text": "Write the FAQ", "owner": "Sam", "source": EMAIL, "quoteVerified": True}]}
    assert db.read_bytes() == before
    assert "error" in d.story_decisions_db(tmp_path / "missing.db", "p-dec")


def test_story_beats_and_timing_rows():
    d = _driver()
    r = {"steps": [{"step": "P0_prepare_excel", "pass": True, "stepMs": 9},
                   {"step": "S1_email_decision_page", "pass": True, "stepMs": 1200, "turnMs": 900},
                   {"step": "S2_deliverables", "pass": False, "stepMs": 5000, "turnMs": 4000}]}
    assert d.story_beats(r) == [{"beat": "S1_email_decision_page", "pass": True, "stepMs": 1200, "turnMs": 900},
                                {"beat": "S2_deliverables", "pass": False, "stepMs": 5000, "turnMs": 4000}]


def test_story_step_records_a_failed_check_with_masked_results_and_continues(monkeypatch):
    d = _driver()

    class Tab:
        def __init__(self):
            self.cmds = []

        def show(self):
            return None

        def cmd(self, *a, **k):
            self.cmds.append((a, k))

    class Chat:
        approvals = []

        async def turn(self, text):
            return {"text": "sent to pat@example.com", "ms": 7,
                    "tools": [_tool("brain_why", {"note": "pat@example.com"}, ok=False)]}

    class Rec:
        out = Path("/nonexistent")

        def save(self):
            return None

    tps = {h: Tab() for h in ("onenote", "excel", "outlook")}
    step = {"name": "S3_why", "show": "excel", "read": [], "check": d.story_why_check, "prompt": lambda c: "Why?"}
    result = {"steps": []}
    monkeypatch.setattr(d, "fetch_shot", lambda *a, **k: "shot.png")
    s = asyncio.run(d._story_step(Chat(), tps, step, _with_decision(d), Rec(), result, 4, 1))
    assert s["pass"] is False and s["turnMs"] == 7 and isinstance(s["stepMs"], int)
    assert "pat@example.com" not in repr(s) and "pat@example.com" not in repr(tps["onenote"].cmds)
    assert result["steps"] == [s]


def test_run_macbook_story_mode_passes_story_to_the_driver():
    sh = (LIVE / "run_macbook.sh").read_text()
    assert 'story) HOSTS=""; F3=(); STORY=(--story) ;;' in sh
    assert '${STORY[@]+"${STORY[@]}"}' in sh
    out = subprocess.run(["bash", "-n", str(LIVE / "run_macbook.sh")], capture_output=True, text=True)
    assert out.returncode == 0, out.stderr
    src = (LIVE / "live_driver.py").read_text()
    assert 'ap.add_argument("--story"' in src and "story_once(client, pane, n, rec, manifests, bare=a.bare)" in src


# ---- story --bare (A24): the presenter's exact prompts, outcomes only -------------------------------------------
def test_story_bare_prompts_are_exactly_the_demo_script_with_no_note():
    d = _driver()
    p = d.story_prompts()
    steps = d.story_steps(bare=True)
    c = _with_decision(d)
    c["slides0"] = 1
    assert [s["prompt"](c) for s in steps] == p  # nothing appended: no tool, title, id, path or time
    assert not any("(" in s["prompt"](c) or "Call " in s["prompt"](c) for s in steps)
    assert [s["name"] for s in steps] == [s["name"] for s in d.story_steps()]
    assert [s["check"].__name__ for s in steps] == ["bare_email_decision_check", "bare_deliverables_check",
                                                    "bare_why_check", "bare_reply_check", "bare_undo_check"]
    import inspect  # outcome-only: no bare check names a tool it expects the model to call
    for s in steps:
        src = inspect.getsource(s["check"])
        assert "last_result(" not in src and "STORY_REPLY_ALLOWED" not in src and "STORY_UNDO_ALLOWED" not in src


def test_run_macbook_story_bare_mode():
    sh = (LIVE / "run_macbook.sh").read_text()
    assert 'story-bare) HOSTS=""; F3=(); STORY=(--story --bare) ;;' in sh
    assert d_row(True)["host"] == "story-bare" and d_row(False)["host"] == "story"


def d_row(bare):
    return _driver().new_story_row(1, bare)


def _bare_db(tmp_path, rows, pages, now):
    db = tmp_path / "brain.db"
    con = sqlite3.connect(db)
    con.execute("create table pages(notebook_id text, page_id text, title text, captured_at real)")
    con.execute("create table decisions(id integer primary key, notebook_id text, decision_page_id text, source_page_id text,"
                " kind text, text text, owner text, due text, rationale text, quote text, quote_verified integer, created_at real)")
    for pid, title, at in pages:
        con.execute("insert into pages values ('nb', ?, ?, ?)", (pid, title, at))
    for dp, kind, text, owner, q, at in rows:
        con.execute("insert into decisions(notebook_id, decision_page_id, source_page_id, kind, text, owner, quote_verified,"
                    " created_at) values ('nb', ?, ?, ?, ?, ?, ?, ?)", (dp, EMAIL, kind, text, owner, q, at))
    con.commit()
    con.close()
    return db


def test_story_bare_index_and_decision_are_judged_from_the_index_and_the_page(tmp_path):
    d = _driver()
    now = 1_000_000.0
    db = _bare_db(tmp_path, [("p-old", "decision", "We ship the beta on Friday", None, 1, now - 900),
                             ("p-new", "decision", "We ship the beta on Friday", None, 1, now + 5),
                             ("p-new", "decision", "We keep the current pilot scope", None, 1, now + 5),
                             ("p-new", "action", "Write the FAQ", "Sam", 1, now + 5)],
                  [(EMAIL, d.EMAIL_SOURCE, now + 2), ("p-new", "Pilot sync decisions", now + 6),
                   ("p-old", "Older page", now - 900)], now)
    before = db.read_bytes()
    idx = d.story_index_db(db, now)
    assert idx["fresh"] == [{"pageId": EMAIL, "title": d.EMAIL_SOURCE}]
    assert d.story_index_db(db, now + 100)["fresh"] == []  # an entry from an earlier run does not count
    c = _ctx(d)
    ok, note = d.bare_email_check({"tools": []}, {"db": idx, "outlook": {"mode": "read", "subjectIsFixture": True}}, c)
    assert ok, note
    got = d.story_new_decision_page_db(db, now, EMAIL)
    assert got["pageId"] == "p-new" and got["title"] == "Pilot sync decisions" and len(got["rows"]) == 3
    assert db.read_bytes() == before  # read-only
    page = {"pageTitle": "Pilot sync decisions", "text": [f"Source: {d.EMAIL_SOURCE}", "We ship the beta on Friday",
                                                           "We keep the current pilot scope", "Write the FAQ owner Sam"]}
    ok, note = d.bare_decision_check({"tools": []}, {"db": got, "onenote_decision": page}, c)
    assert ok, note
    assert c["decision"]["title"] == "Pilot sync decisions" and c["decision"]["items"][0]["owner"] == "Unknown"
    assert "not owned by Sam" in d.bare_decision_check({"tools": []}, {"db": {**got, "rows": [
        dict(r, owner=None) for r in got["rows"]]}, "onenote_decision": page}, _ctx(d))[1]
    assert "no decision page" in d.bare_decision_check({"tools": []}, {"db": {"pages": 0}}, _ctx(d))[1]
    # S1 (A25): one beat, judged on both outcomes; which tool (or how many calls) made them is not judged
    rbs = {"index": idx, "outlook": {"mode": "read", "subjectIsFixture": True}, "db": got, "onenote_decision": page}
    c = _ctx(d)
    ok, note = d.bare_email_decision_check({"tools": []}, rbs, c)
    assert ok, note
    assert c["email"]["pageId"] == EMAIL and c["decision"]["pageId"] == "p-new"
    assert "no index entry" in d.bare_email_decision_check({"tools": []}, {**rbs, "index": d.story_index_db(db, now + 100)}, _ctx(d))[1]
    assert "no decision page" in d.bare_email_decision_check({"tools": []}, {**rbs, "db": {"pages": 0}}, _ctx(d))[1]


def test_story_bare_deliverables_judge_documents_and_receipts_not_tool_names():
    d = _driver()
    _, deck = _deck()
    rbs = {"excel": _tracker_rb(d), "word": _brief_rb(d), "powerpoint": deck, "powerpoint_images": {"images": 1}}
    c = _with_decision(d)
    c.update(slides0=1, img_before={"images": 0})
    t = _deliverables_turn(d, "draft_new", images=0)
    team = {"status": "verified", "source": {"type": "path", "name": "team.png"}}
    t["tools"].append(_tool("any_tool", team))  # the picture by any route: the deck's count and the receipt's source
    ok, note = d.bare_deliverables_check(t, rbs, c)
    assert ok, note
    assert "team.png" in note
    for other in ({**team, "source": {"type": "url", "url": "https://x/team.png"}},
                  {**team, "source": {"type": "path", "name": "hol-1.png"}}, {**team, "status": "failed"}):
        t2 = {"tools": t["tools"][:-1] + [_tool("any_tool", other)]}
        assert "no picture receipt names team.png" in d.bare_deliverables_check(t2, rbs, c)[1]
    assert "pictures" in d.bare_deliverables_check(t, {**rbs, "powerpoint_images": {"images": 0}}, c)[1]
    no_reply = {"tools": [x for x in t["tools"]]}
    no_reply["tools"][0] = _tool("brain_make_deliverables", {"status": "verified", "results": [
        r for r in t["tools"][0]["result"]["data"]["results"] if r["host"] != "outlook"]})
    assert "no reply" in d.bare_deliverables_check(no_reply, rbs, c)[1]
    sent = {"tools": t["tools"] + [_tool("outlook_draft_reply", {"status": "requested", "sent": True})]}
    assert "sent or saved" in d.bare_deliverables_check(sent, rbs, c)[1]


def test_story_bare_why_reply_and_undo_are_judged_on_the_answer_and_the_draft():
    d = _driver()
    c = _with_decision(d)
    good = "The FAQ row comes from the decision page '" + DTITLE + "', from the email 'HOL live test: pilot sync'."
    assert d.bare_why_check({"tools": [], "text": good}, {}, c)[0]  # no brain_why needed: the answer is judged
    assert "name the email" in d.bare_why_check({"tools": [], "text": "The FAQ is on the decision page."}, {}, c)[1]
    assert d._added_text("\n\nFrom: x", "Thanks, we ship the beta Friday; Sam writes the FAQ.\n\nFrom: x") == \
        "Thanks, we ship the beta Friday; Sam writes the FAQ."
    st = {"hasMarker": True, "toIsOnlyAccount": True, "ccEmpty": True, "bccEmpty": True}
    c.update(reply_before={"body": "\nFrom: x"}, reply_added="Thanks, we ship the beta Friday; Sam writes the FAQ.",
             reply_state=st, still_draft=True)
    rb = {"outlook_compose": {"mode": "compose", "body": "Thanks, we ship the beta Friday; Sam writes the FAQ.\nFrom: x",
                              "cc": 0, "bcc": 0}}
    ok, note = d.bare_reply_check({"tools": []}, rb, c)
    assert ok, note
    assert "no reply text" in d.bare_reply_check({"tools": []}, rb, {**c, "reply_added": ""})[1]
    assert "on screen" in d.bare_reply_check({"tools": []}, rb, {**c, "reply_state": {**st, "ccEmpty": False}})[1]
    after = {"outlook_compose": {"mode": "compose", "body": "\nFrom: x", "cc": 0, "bcc": 0}}
    c.update(undo_state={**st, "hasMarker": False})
    ok, note = d.bare_undo_check({"tools": []}, after, c)
    assert ok, note
    assert "still shows" in d.bare_undo_check({"tools": []}, rb, c)[1]


def test_the_demo_script_and_the_harness_fixture_quote_the_same_neutral_email():
    """A19: docs/demo-script.md's set-up quotes the fixture email the harness sends; the two stay in sync, and the
    second decision is neutral (no price)."""
    d = _driver()
    ui = sys.modules["office_ui"]
    doc = " ".join((ROOT / "docs" / "demo-script.md").read_text(encoding="utf-8").split())
    assert f'subject "{ui.OUTLOOK_FIXTURE_SUBJECT}"' in doc
    for line in ui.OUTLOOK_FIXTURE_BODY.split("\n")[1:]:  # the decision and action lines (not the date line)
        assert line in doc, line
    assert "$" not in ui.OUTLOOK_FIXTURE_BODY and "$49" not in doc
    for want in ui.OUTLOOK_FIXTURE_DECISIONS:
        assert want in ui.OUTLOOK_FIXTURE_BODY.lower()
    assert d.EMAIL_SOURCE == "Email: " + ui.OUTLOOK_FIXTURE_SUBJECT and ui.OUTLOOK_FIXTURE_TOPIC in ui.OUTLOOK_FIXTURE_SUBJECT
    assert d.story_fixture(1, stamp="x")["decision"].startswith(f"Decisions: {ui.OUTLOOK_FIXTURE_TOPIC} (")
    assert d.OUTLOOK_STATE.name == "outlook-fixture-v2-sent.json"  # a new state file: the v2 message is sent once


# ---- cold test profiles (A26) -------------------------------------------------------------------------------------
def test_self_learning_reads_real_keys_with_hermes_defaults_and_only_all_off_is_cold(tmp_path):
    d = _driver()
    assert set(d.SELF_LEARNING_KEYS) == {"curator.enabled", "auxiliary.background_review.enabled",
                                         "skills.creation_nudge_interval", "skills.write_approval",
                                         "memory.memory_enabled", "memory.user_profile_enabled", "memory.nudge_interval"}
    warm = d.self_learning({})
    assert warm["cold"] is False and warm["settings"]["curator.enabled"] is True
    assert warm["settings"]["skills.creation_nudge_interval"] == 10
    cold_cfg = {"curator": {"enabled": False}, "auxiliary": {"background_review": {"enabled": False}},
                "skills": {"creation_nudge_interval": 0, "write_approval": True},
                "memory": {"memory_enabled": False, "user_profile_enabled": False, "nudge_interval": 0}}
    assert d.self_learning(cold_cfg)["cold"] is True
    half = {**cold_cfg, "curator": {"enabled": True}}
    assert d.self_learning(half)["cold"] is False
    import yaml
    (tmp_path / "config.yaml").write_text(yaml.safe_dump(cold_cfg))
    assert d.self_learning(d.profile_config(tmp_path))["cold"] is True
    assert d.profile_config(tmp_path / "missing") == {}


def test_learned_state_lists_skills_memory_and_staged_writes(tmp_path):
    d = _driver()
    assert d.learned_state(tmp_path) == {"skills": [], "memoryFiles": [], "pendingWrites": []}
    for rel in ("skills/office/office-pane-workflows/SKILL.md", "skills/.archive/old/SKILL.md", "memories/MEMORY.md",
                "pending/skills/abc.json"):
        (tmp_path / rel).parent.mkdir(parents=True, exist_ok=True)
        (tmp_path / rel).write_text("x")
    (tmp_path / "memories" / "MEMORY.md.lock").write_text("")
    assert d.learned_state(tmp_path) == {"skills": ["office/office-pane-workflows"], "memoryFiles": ["MEMORY.md"],
                                         "pendingWrites": ["skills/abc.json"]}


def test_story_bare_refuses_a_profile_that_can_learn_unless_told_it_is_a_warm_run():
    d = _driver()
    warm = {"profile": "p", "selfLearning": d.self_learning({})}
    gate = d.cold_gate(warm, bare=True, allow_warm=False)
    assert gate.startswith("COLD GATE") and "curator.enabled" in gate and "cold_profile.sh" in gate
    assert d.cold_gate(warm, bare=True, allow_warm=True) is None
    assert d.cold_gate(warm, bare=False, allow_warm=False) is None  # routed story and other modes: recorded only
    cold = {"selfLearning": {"cold": True}}
    assert d.cold_gate(cold, bare=True, allow_warm=False) is None
    src = (LIVE / "live_driver.py").read_text()
    assert 'cold_gate(rec.identity, story and getattr(a, "bare", False))' in src
    assert '"selfLearning": self_learning(profile_config(home)), "learnedAtStart": learned_state(home)' in src
    assert '"learnedAtEnd": learned_state(' in src


def test_cold_profile_script_sets_exactly_the_driver_keys_with_hermes_config_set():
    sh = (LIVE / "cold_profile.sh").read_text()
    assert "from live_driver import SELF_LEARNING_KEYS" in sh and '"$HERMES" config set "$key" "$value"' in sh
    assert 'HERMES_HOME="$PH"' in sh and "sys.exit(0 if sl['cold'] else 1)" in sh
    assert "refusing the root Hermes home" in sh
    d = _driver()
    harness = " ".join((ROOT / "docs" / "harness.md").read_text(encoding="utf-8").split())
    for k, (_default, cold) in d.SELF_LEARNING_KEYS.items():
        assert f"`{k}: {str(cold).lower()}`" in harness, k


def test_plugin_commit_claim_is_checked_byte_for_byte_against_git_archive(tmp_path, monkeypatch):
    """A26: a newer harness drives a profile whose plugin dir is an older commit's export; the identity names both
    and records whether the dir really is that export."""
    import io
    import tarfile
    d = _driver()
    buf = io.BytesIO()
    with tarfile.open(fileobj=buf, mode="w") as tf:
        for name, body in (("plugin.yaml", b"name: x\n"), ("pkg/a.py", b"A = 1\n")):
            info = tarfile.TarInfo(name)
            info.size = len(body)
            tf.addfile(info, io.BytesIO(body))
    arch = buf.getvalue()
    pd = tmp_path / "plugin"
    (pd / "pkg" / "__pycache__").mkdir(parents=True)
    (pd / "plugin.yaml").write_bytes(b"name: x\n")
    (pd / "pkg" / "a.py").write_bytes(b"A = 1\n")
    (pd / "pkg" / "__pycache__" / "a.pyc").write_bytes(b"ignored")
    assert d.plugin_dir_matches("c", pd, archive=arch)["match"] is True
    (pd / "pkg" / "a.py").write_bytes(b"A = 2\n")
    (pd / "extra.py").write_bytes(b"")
    got = d.plugin_dir_matches("c", pd, archive=arch)
    assert got["match"] is False and got["differ"] == ["pkg/a.py"] and got["extra"] == ["extra.py"]
    assert d.plugin_dir_matches("c", tmp_path / "none", archive=arch)["match"] is False
    assert 'rec.identity.get("pluginDirMatches") is False' in (LIVE / "live_driver.py").read_text()
    # the real thing: this checkout's HEAD exported into a fake profile, claimed by HOL_PLUGIN_COMMIT. Needs a git
    # checkout; an exported tree (git archive, no .git) has no commit to claim, so the part above is all it can check.
    head = subprocess.run(["git", "-C", str(ROOT), "rev-parse", "--show-toplevel", "HEAD"], capture_output=True,
                          text=True).stdout.split()
    if len(head) != 2 or Path(head[0]).resolve() != ROOT.resolve():
        return
    head = head[1]
    home = tmp_path / "home"
    (home / "plugins" / "hermes-office-desktop").mkdir(parents=True)
    tar = subprocess.run(["git", "-C", str(ROOT), "archive", head], capture_output=True, check=True).stdout
    with tarfile.open(fileobj=io.BytesIO(tar)) as tf:
        tf.extractall(home / "plugins" / "hermes-office-desktop", filter="data")
    monkeypatch.setenv("HERMES_HOME", str(home))
    monkeypatch.setenv("HOL_PLUGIN_COMMIT", head[:7])
    ident = d.run_identity(tmp_path / "run", fetch=lambda url: b"x  f\n")
    assert ident["pluginCommit"] == head and ident["harnessCommit"] == head and ident["pluginDirMatches"] is True
    assert ident["selfLearning"]["cold"] is False and ident["learnedAtStart"]["skills"] == []
    monkeypatch.delenv("HOL_PLUGIN_COMMIT")
    assert "pluginDirMatches" not in d.run_identity(tmp_path / "run", fetch=lambda url: b"")
    assert 'rec.identity.get("pluginDirMatches") is False' in (LIVE / "live_driver.py").read_text()
