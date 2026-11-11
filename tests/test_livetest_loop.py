# SPDX-License-Identifier: Apache-2.0
"""Demo-loop harness mode (live_driver.py --loop): pure logic only. Nothing here touches Office, a gateway or the network."""
import asyncio
import importlib.util
import json
import sqlite3
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


def _chat():
    return _load("chat", "chat.py")


def _ctx(d, run=1, stamp="1106-020000"):
    return {"run": run, "fx": d.loop_fixture(run, stamp)}


# ---- the script is exactly docs/demo.md's prompts, in one chat ------------------------------------------
def test_demo_prompts_come_from_docs_demo_md():
    d = _driver()
    prompts = d.demo_prompts()
    assert len(prompts) == 7
    assert prompts[0].startswith("Start today's daily note") and prompts[-1] == "Undo the last OneNote change."


def test_loop_script_order_and_every_demo_prompt_is_used_verbatim():
    d = _driver()
    c = _ctx(d)
    steps = d.loop_steps()
    names = [s["name"] for s in steps]
    assert names == ["L06_seed", "L07_sync", "L08_daily_note", "L09_q4_links", "L10_decision_page", "L11_deliverables",
                     "L12_tab_switch", "L13_why", "L14_ask_cited", "L15_resurface", "L16_denied", "L17_undo"]
    texts = {s["name"]: s["prompt"](c) for s in steps if callable(s.get("prompt"))}
    p = d.demo_prompts()
    q4 = c["fx"]["titles"]["Q4 Launch"]
    # Each demo prompt opens its step's message (titles retargeted to this run's tagged pages, then context).
    assert texts["L08_daily_note"].startswith(p[0].replace("[[Q4 Launch]]", f"[[{q4}]]"))
    for i, name in ((1, "L09_q4_links"), (2, "L14_ask_cited"), (3, "L10_decision_page"), (4, "L11_deliverables"),
                    (5, "L13_why"), (6, "L17_undo")):
        assert texts[name].startswith(p[i]), name
    assert c["fx"]["daily"] in texts["L08_daily_note"] and c["fx"]["decision"] in texts["L10_decision_page"]
    hosts = {s["name"]: s.get("show") for s in steps}
    assert hosts["L11_deliverables"] == "onenote" and hosts["L13_why"] == "excel" and hosts["L12_tab_switch"] == "excel"
    assert {"excel", "word", "powerpoint"} <= set(next(s for s in steps if s["name"] == "L11_deliverables")["read"])


def test_loop_fixture_titles_are_run_tagged():
    d = _driver()
    a, b = d.loop_fixture(1, "1106-020000"), d.loop_fixture(2, "1106-020500")
    assert a["daily"] != b["daily"] and a["decision"] != b["decision"]
    assert a["tag"] in a["daily"] and a["tag"] in a["decision"] and a["daily"].startswith("Daily ")


# ---- the one approval the loop denies ------------------------------------------------------------------
def test_approval_choice_denies_only_a_matching_office_prompt():
    ch = _chat()
    office = {"description": "OneNote · HOL · add a block to the open page (undoable) · Workbench",
              "choices": ["once", "session", "always", "deny"]}
    assert ch.approval_choice(office) == ("once", None)
    assert ch.approval_choice(office, ["add a block to the open page"]) == ("deny", "add a block to the open page")
    other = {"description": "OneNote · write the Links & Backlinks block (undoable) · Workbench", "choices": ["once", "deny"]}
    assert ch.approval_choice(other, ["add a block to the open page"]) == ("once", None)
    # Excel/Word/PowerPoint prompts carry no "Office" word: the product name alone identifies them.
    excel = {"description": "Excel · write S!B2 (undoable) · Workbench", "choices": ["once", "deny"]}
    assert ch.approval_choice(excel) == ("once", None)
    shell = {"description": "run rm -rf /tmp/x", "choices": ["once", "deny"]}
    assert ch.approval_choice(shell) == ("deny", None) and ch.approval_choice(shell, ["rm"]) == ("deny", None)


def test_chat_records_a_rule_denial():
    ch = _chat()

    class Client:
        def __init__(self):
            self.sent = []

        async def rpc(self, method, **kw):
            self.sent.append((method, kw))
            return {}
    cl = Client()
    chat = ch.Chat(cl, {"stored_session_id": "s1"}, lambda obj: None)
    chat.deny_rules = ["add a block"]
    asyncio.run(chat._answer({"prompt_id": "p1", "description": "OneNote · add a block to the open page · Workbench",
                              "choices": ["once", "deny"]}))
    assert cl.sent[0][1]["choice"] == "deny" and chat.approvals[0]["deniedByRule"] == "add a block"


def _ledger(tmp_path, rows):
    db = tmp_path / "brain.db"
    con = sqlite3.connect(db)
    con.execute("create table writes(id integer primary key autoincrement, notebook_id text not null, page_id text, op_id text, "
                "tool text not null, kind text not null, status text not null, digest text, summary text, created json, "
                "extra json, created_at real, updated_at real)")
    for nb, tool in rows:
        con.execute("insert into writes(notebook_id, tool, kind, status) values(?,?,?,?)", (nb, tool, "append", "done"))
    con.commit()
    con.close()
    return db


def test_ledger_snapshot_is_read_only_and_per_notebook(tmp_path):
    d = _driver()
    db = _ledger(tmp_path, [("nb1", "brain_links"), ("nb2", "onenote_append"), ("nb1", "onenote_create_page")])
    snap = d.ledger_snapshot(db, "nb1")
    assert snap["rows"] == 2 and snap["maxId"] == 3 and snap["lastTool"] == "onenote_create_page"
    assert "error" in d.ledger_snapshot(tmp_path / "missing.db", "nb1")
    assert not (tmp_path / "missing.db").exists()  # read-only: never creates a database


def _page(title, user, links=0, block="Links & Backlinks (kept by Hermes) Links: x"):
    return {"pageTitle": title, "notebookId": "nb1", "linksBlocks": links, "text": list(user) + [block] * links}


def test_deny_check_needs_one_denied_prompt_no_write_and_an_unchanged_page_and_ledger():
    d = _driver()
    c = _ctx(d)
    before = _page("Pricing review", ["Options: A) $29"])
    c["deny_before"] = {"page": before, "ledger": {"rows": 9, "maxId": 40, "lastTool": "brain_links"}}
    denied = [{"choice": "deny", "deniedByRule": "add a block to the open page", "description": "add a block"}]
    t = {"tools": [{"name": "onenote_append", "result": "BLOCKED: denied by the user"}], "approvals": denied, "text": "It was declined."}
    rbs = {"onenote": before, "ledger": {"rows": 9, "maxId": 40, "lastTool": "brain_links"}}
    ok, note = d.deny_check(t, rbs, c)
    assert ok, note
    assert not d.deny_check({**t, "approvals": []}, rbs, c)[0]  # nothing was denied: proves nothing
    assert not d.deny_check({**t, "approvals": denied * 2}, rbs, c)[0]  # the model retried
    assert not d.deny_check({**t, "approvals": denied + [{"choice": "once", "description": "x"}]}, rbs, c)[0]
    landed = {**rbs, "onenote": _page("Pricing review", ["Options: A) $29", "HOL denied note"])}
    assert "document changed" in d.deny_check(t, landed, c)[1]
    assert "ledger" in d.deny_check(t, {**rbs, "ledger": {"rows": 10, "maxId": 41, "lastTool": "onenote_append"}}, c)[1]
    assert not d.deny_check(t, {**rbs, "ledger": {"error": "unreadable"}}, c)[0]  # missing evidence is not a pass
    ok_append = {**t, "tools": [{"name": "onenote_append", "result": {"ok": True, "data": {}}}]}
    assert not d.deny_check(ok_append, rbs, c)[0]


# ---- deliverables: every nested receipt and an independent readback of each document -------------------
ITEMS = [{"kind": "decision", "text": "Run a pilot in two cities", "owner": "Unknown (not in the notes)"},
         {"kind": "decision", "text": "Keep the $29 plan for the pilot", "owner": "Unknown (not in the notes)"},
         {"kind": "action", "text": "Draft the pilot plan", "owner": "Priya"},
         {"kind": "action", "text": "Book the two pilot cities", "owner": "Sam"},
         {"kind": "action", "text": "Check supplier lead times", "owner": "Unknown (not in the notes)"}]


def _deliv(d, c, **over):
    meeting = c["fx"]["titles"]["Launch meeting 2026-11-03"]
    dtitle = c["fx"]["decision"]
    c["decision"] = {"title": dtitle, "pageId": "p-dec", "items": ITEMS}
    titles = [dtitle, "What we decided", "Who does what", "Action owner needed", "Sources"]
    bodies = {"What we decided": "Run a pilot in two cities\nKeep the $29 plan for the pilot",
              "Who does what": "Draft the pilot plan — Priya (Nov 12)\nBook the two pilot cities — Sam",
              "Action owner needed": "Check supplier lead times"}
    receipt = {"status": "verified", "results": [
        {"host": "excel", "status": "verified", "locator": "Decision tracker (2)!A1:F4", "rows": 3},
        {"host": "word", "status": "verified", "locator": f"end of document: '{dtitle} — brief'"},
        {"host": "powerpoint", "status": "verified", "slides": [{"title": x, "slideIndex": i + 1} for i, x in enumerate(titles)]}]}
    receipt.update(over.pop("receipt", {}))
    t = {"tools": [{"name": "brain_make_deliverables", "result": {"ok": True, "data": receipt}}], "text": "Done."}
    rows = [["#", "Item", "Owner", "Due", "Status", "Source", "", "Open items", 3],
            ["A1", "Draft the pilot plan", "Priya", "Nov 12", "Open", meeting, "", "", ""],
            ["A2", "Book the two pilot cities", "Sam", "Unknown", "Open", meeting, "", "Owner", "Open items"],
            ["A3", "Check supplier lead times", "Unknown", "Unknown", "Open", meeting, "", "Priya", 1]]
    forms = [r[:] for r in rows]
    forms[0][8] = '=COUNTIF(E2:E4,"Open")'
    for r in forms[1:]:
        r[5] = '=HYPERLINK("https://x","' + meeting + '")'
    excel = {"sheet": "Decision tracker (2)", "values": rows, "formulas": forms,
             "charts": [{"name": "Chart 1", "title": "Open items by owner"}]}
    word = {"paragraphs": [f"{dtitle} — brief", "Decisions", "D1. Run a pilot in two cities (owner: Unknown). Source: " + meeting,
                           "D2. Keep the $29 plan for the pilot (owner: Unknown). Source: " + meeting, "Actions",
                           "A1. Draft the pilot plan (owner: Priya, due: Nov 12). Source: " + meeting,
                           "A2. Book the two pilot cities (owner: Sam, due: Unknown). Source: " + meeting,
                           "A3. Check supplier lead times (owner: Unknown, due: Unknown). Source: " + meeting,
                           f"Decision record: {dtitle}"],
            "links": [{"text": meeting, "hyperlink": True}] * 5 + [{"text": dtitle, "hyperlink": True}]}
    ppt = {"count": 6, "slides": [["HOL live test text"]] + [[x, bodies.get(x, "bullet")] for x in titles]}
    c["slides0"] = 1
    rbs = {"excel": excel, "word": word, "powerpoint": ppt}
    for k, v in over.items():
        rbs[k] = v
    return t, rbs


def test_deliverables_check_passes_only_on_verified_receipts_and_matching_documents():
    d = _driver()
    c = _ctx(d)
    t, rbs = _deliv(d, c)
    ok, note = d.deliverables_check(t, rbs, c)
    assert ok, note
    assert d.tracker_sheet(t["tools"][0]["result"]["data"]) == "Decision tracker (2)"
    t2, rbs2 = _deliv(d, c, receipt={"status": "partial"})
    assert "receipt" in d.deliverables_check(t2, rbs2, c)[1]
    t3, _ = _deliv(d, c)
    t3["tools"][0]["result"]["data"]["results"][1]["status"] = "applied"  # one nested receipt is not verified
    assert not d.deliverables_check(t3, rbs, c)[0]
    t4, _ = _deliv(d, c)
    del t4["tools"][0]["result"]["data"]["results"][2]  # a host missing from the receipt
    assert not d.deliverables_check(t4, rbs, c)[0]


def test_tracker_check_reads_the_sheet_itself():
    d = _driver()
    c = _ctx(d)
    t, rbs = _deliv(d, c)
    assert d.tracker_check(rbs["excel"], c)[0]
    no_chart = {**rbs["excel"], "charts": []}
    assert "chart" in d.tracker_check(no_chart, c)[1]
    wrong_owner = json.loads(json.dumps(rbs["excel"]))
    wrong_owner["values"][1][2] = "Sam"
    assert "pilot plan" in d.tracker_check(wrong_owner, c)[1]
    stale = json.loads(json.dumps(rbs["excel"]))
    stale["values"][0][8] = 2  # COUNTIF disagrees with the rows
    assert not d.tracker_check(stale, c)[0]
    typed = json.loads(json.dumps(rbs["excel"]))
    typed["formulas"][0][8] = 3  # a typed number is not the live formula
    assert not d.tracker_check(typed, c)[0]
    missing = json.loads(json.dumps(rbs["excel"]))
    missing["values"] = missing["values"][:3]
    missing["formulas"] = missing["formulas"][:3]
    assert "Check supplier lead times" in d.tracker_check(missing, c)[1]


def test_brief_check_needs_headings_items_and_real_source_links():
    d = _driver()
    c = _ctx(d)
    t, rbs = _deliv(d, c)
    assert d.brief_check(rbs["word"], c)[0]
    no_links = {**rbs["word"], "links": []}
    assert "link" in d.brief_check(no_links, c)[1]
    unreadable = {**rbs["word"], "links": None}
    assert not d.brief_check(unreadable, c)[0]
    no_heading = {**rbs["word"], "paragraphs": [p for p in rbs["word"]["paragraphs"] if p != "Actions"]}
    assert "Actions" in d.brief_check(no_heading, c)[1]
    trailing = {**rbs["word"], "paragraphs": rbs["word"]["paragraphs"] + ["something after the brief"]}
    assert not d.brief_check(trailing, c)[0]


def test_deck_check_counts_new_slides_and_titles():
    d = _driver()
    c = _ctx(d)
    t, rbs = _deliv(d, c)
    receipt = t["tools"][0]["result"]["data"]
    assert d.deck_check(rbs["powerpoint"], receipt, c)[0]
    assert not d.deck_check({**rbs["powerpoint"], "count": 5}, receipt, c)[0]
    swapped = json.loads(json.dumps(rbs["powerpoint"]))
    swapped["slides"][-1] = ["Appendix"]
    assert "Sources" in d.deck_check(swapped, receipt, c)[1]


# ---- OneNote beats --------------------------------------------------------------------------------------
def _created(title):
    return {"name": "onenote_create_page", "args": json.dumps({"title": title}), "result": {"ok": True, "data": {"title": title}}}


def test_daily_and_q4_links_checks():
    d = _driver()
    c = _ctx(d)
    fx = c["fx"]
    q4, daily = fx["titles"]["Q4 Launch"], fx["daily"]
    t = {"tools": [_created(daily), {"name": "brain_links", "result": {"ok": True, "data": {"title": daily}}}], "text": ""}
    rb = _page(daily, ["Focus"], 1, f"Links & Backlinks Links: {q4}")
    assert d.daily_check(t, {"onenote": rb}, c)[0]
    assert not d.daily_check(t, {"onenote": _page(daily, ["Focus"], 0)}, c)[0]
    assert not d.daily_check({"tools": t["tools"][1:], "text": ""}, {"onenote": rb}, c)[0]
    meeting, pricing = fx["titles"]["Launch meeting 2026-11-03"], fx["titles"]["Pricing review 2026-10-20"]
    t2 = {"tools": [{"name": "brain_links", "result": {"ok": True, "data": {"title": q4}}}], "text": ""}
    rb2 = _page(q4, ["Goal"], 1, f"Links & Backlinks Links: {pricing}, {meeting} Backlinks: {daily}, {meeting}")
    assert d.q4_links_check(t2, {"onenote": rb2}, c)[0]
    rb3 = _page(q4, ["Goal"], 1, f"Links & Backlinks Links: {pricing}, {meeting} Backlinks: {meeting}")
    assert "daily" in d.q4_links_check(t2, {"onenote": rb3}, c)[1]


def test_decision_beat_records_the_meeting_block_and_checks_the_decision_page():
    d = _driver()
    c = _ctx(d)
    fx = c["fx"]
    meeting, dtitle = fx["titles"]["Launch meeting 2026-11-03"], fx["decision"]
    items = [{"kind": "decision", "text": "Run a pilot in two cities", "owner": "Unknown (not in the notes)"},
             {"kind": "action", "text": "Draft the pilot plan", "owner": "Priya"},
             {"kind": "action", "text": "Check supplier lead times", "owner": "Unknown (not in the notes)"}]
    data = {"pageId": "p-dec", "title": dtitle, "items": items, "unknownOwners": 2, "unverifiedQuotes": 0}
    t = {"tools": [{"name": "brain_decision_page", "result": {"ok": True, "data": data}},
                   {"name": "brain_links", "result": {"ok": True, "data": {"title": meeting, "pageId": "p-meet"}}}], "text": ""}
    meeting_rb = _page(meeting, ["Attendees: Priya"], 1, f"Links & Backlinks Links: Q4 Backlinks: {dtitle}")
    dec_rb = {"pageTitle": dtitle, "text": [f"Source: {meeting} Decisions D1. Run a pilot in two cities · owner: Unknown (not in the notes) "
                                            "Actions A1. Draft the pilot plan · owner: Priya · due: Nov 12 "
                                            "A2. Check supplier lead times · owner: Unknown (not in the notes) · due: Unknown (not in the notes)"]}
    ok, note = d.decision_beat_check(t, {"onenote": meeting_rb, "onenote_decision": dec_rb}, c)
    assert ok, note
    assert c["meeting_links"]["user"] == ["Attendees: Priya"] and c["decision"]["pageId"] == "p-dec"
    c2 = _ctx(d)
    assert "decision page" in d.decision_beat_check(t, {"onenote": _page(meeting, ["Attendees: Priya"], 1, "Links & Backlinks Links: Q4"),
                                                        "onenote_decision": dec_rb}, c2)[1]


def test_final_undo_removes_exactly_the_meeting_block():
    d = _driver()
    c = _ctx(d)
    meeting = c["fx"]["titles"]["Launch meeting 2026-11-03"]
    c["meeting_links"] = {"pageTitle": meeting, "user": ["Attendees: Priya"], "links": ["Links & Backlinks x"]}
    t = {"tools": [{"name": "office_undo", "result": {"ok": True, "data": {"undone": "brain_links", "pageId": "p-meet"}}}], "text": ""}
    assert d.final_undo_check(t, {"onenote": _page(meeting, ["Attendees: Priya"], 0)}, c)[0]
    assert not d.final_undo_check(t, {"onenote": _page(meeting, ["Attendees: Priya"], 1)}, c)[0]
    assert "user text" in d.final_undo_check(t, {"onenote": _page(meeting, ["Attendees: Sam"], 0)}, c)[1]
    wrong = {"tools": [{"name": "office_undo", "result": {"ok": True, "data": {"undone": "brain_decision_page"}}}], "text": ""}
    assert not d.final_undo_check(wrong, {"onenote": _page(meeting, ["Attendees: Priya"], 0)}, c)[0]
    assert not d.final_undo_check(t, {"onenote": _page(meeting, ["Attendees: Priya"], 0)}, _ctx(d))[0]  # nothing recorded


def test_why_ask_and_resurface_checks():
    d = _driver()
    c = _ctx(d)
    fx = c["fx"]
    meeting, pricing = fx["titles"]["Launch meeting 2026-11-03"], fx["titles"]["Pricing review 2026-10-20"]
    why = {"found": True, "artifacts": [{"host": "excel", "because": [
        {"item": "Draft the pilot plan", "quote": "Priya drafts the pilot plan by Nov 12.", "sourcePage": {"title": meeting},
         "decisionPage": {"title": fx["decision"]}}]}]}
    t = {"tools": [{"name": "brain_why", "result": {"ok": True, "data": why}}], "text": f"It comes from {meeting}."}
    assert d.why_check(t, {}, c)[0]
    assert not d.why_check({**t, "text": "No idea."}, {}, c)[0]
    nowhere = {"found": True, "artifacts": [{"host": "word", "because": why["artifacts"][0]["because"]}]}
    assert not d.why_check({**t, "tools": [{"name": "brain_why", "result": {"ok": True, "data": nowhere}}]}, {}, c)[0]

    ask = {"results": [{"title": pricing, "excerpt": "Rejected option B"}, {"title": meeting, "excerpt": "keep the $29"}],
           "decisions": [], "coverage": {"label": "indexed 7 of 9 pages"}}
    t2 = {"tools": [{"name": "brain_ask", "result": {"ok": True, "data": ask}},
                    {"name": "onenote_open_page", "result": {"ok": True, "data": {"title": meeting}}}],
          "text": f"Keep $29 for the pilot ({meeting}; {pricing}). (indexed 7 of 9 pages, local index)"}
    assert d.ask_check(t2, {"onenote": _page(meeting, ["x"])}, c)[0]
    assert "coverage" in d.ask_check({**t2, "text": f"Keep $29 ({meeting})."}, {"onenote": _page(meeting, ["x"])}, c)[1]
    assert not d.ask_check(t2, {"onenote": _page(fx["titles"]["Risks"], ["x"])}, c)[0]  # opened a page it did not cite

    pages = [{"title": "Decisions: older run", "decisions": 3}, {"title": meeting, "decisions": 5}, {"title": "Risks", "decisions": 0}]
    t3 = {"tools": [{"name": "brain_resurface", "result": {"ok": True, "data": {"pages": pages}}}], "text": "Decisions: older run, ..."}
    assert d.resurface_check(t3, {}, c)[0]
    bad = [pages[2], pages[0]]
    assert not d.resurface_check({**t3, "tools": [{"name": "brain_resurface", "result": {"ok": True, "data": {"pages": bad}}}]}, {}, c)[0]


def test_status_all_check_requires_every_host_live():
    d = _driver()

    def st(app, live=True):
        return {"name": "office_status", "args": {"app": app},
                "result": {"ok": True, "data": {"app": app, "paired": True, "live": live, "paneState": "live" if live else "unloaded",
                                                "lastPollAgeMs": 900 if live else 61000}}}
    hosts = ["onenote", "excel", "word", "powerpoint"]
    ok, note, per = d.status_all_check({"tools": [st(h) for h in hosts]}, hosts)
    assert ok and per["word"]["live"] is True
    ok, note, per = d.status_all_check({"tools": [st(h, h != "word") for h in hosts]}, hosts)
    assert not ok and "word" in note and per["word"]["paneState"] == "unloaded" and per["word"]["lastPollAgeMs"] == 61000
    ok, note, _ = d.status_all_check({"tools": [st(h) for h in hosts[:3]]}, hosts)
    assert not ok and "powerpoint" in note


def test_tab_pane_routes_frame_commands_to_its_tab():
    ui = _load("office_ui", "office_ui.py")
    sent = []

    class Fake(ui.TabPane):
        def __init__(self, tab):
            super().__init__("http://x", lambda o: None, tab)

    def fake_cmd(self, cmd, timeout=90, **kw):
        sent.append((cmd, kw))
        return {"value": 1, "file": "f"}
    orig = ui.Pane.cmd
    ui.Pane.cmd = fake_cmd
    try:
        p = Fake("word")
        p.run_js("top", "1")
        p.cmd("screenshot", file="a.png")
        p.cmd("chat", text="hi")
        p.show()
        p.cmd("docs", op="open", host="word")
    finally:
        ui.Pane.cmd = orig
    assert sent[0] == ("eval", {"frame": "top", "js": "1", "tab": "word"})
    assert "tab" not in sent[1][1] and "tab" not in sent[2][1]
    assert sent[3] == ("tab", {"op": "show", "name": "word"}) and sent[4][1]["tab"] == "word"


def test_results_row_shape_matches_the_nightly_summary():
    d = _driver()
    row = d.new_loop_row(2)
    assert row["host"] == "loop" and row["run"] == 2 and row["steps"] == [] and row["pass"] is False


# ---- A9: OneNote's own page-link labels, this run's pages only, the PowerPoint readback ----------------------
def test_links_checks_accept_onenotes_shortened_and_colonless_page_link_labels():
    """A8 live: the Links block showed 'Pricing review 2026-10-20 (HOL r1 ', 'Daily 2026-11-06 (HOL r1 1106-' and
    'Decisions launch meeting (HOL r1 1106-121224)' (colon dropped); the base title and this run's number must be there."""
    d = _driver()
    c = _ctx(d, run=1, stamp="1106-121224")
    fx = c["fx"]
    q4 = fx["titles"]["Q4 Launch"]
    t = {"tools": [{"name": "brain_links", "result": {"ok": True, "data": {"title": q4}}}], "text": ""}
    live = ("Links & Backlinks (kept by Hermes; regenerated on request; your text is never changed) Links: Pricing review "
            "2026-10-20 (HOL r1 , Launch meeting 2026-11-03 (HOL r1 , Risks (HOL r1 1106-121224) Backlinks: Launch meeting "
            "2026-11-03 (HOL r1 , Pricing review 2026-10-20 (HOL r1 , Daily 2026-11-06 (HOL r1 1106- (within indexed 37 of 81 pages)")
    fx["daily"] = "Daily 2026-11-06 (HOL r1 1106-121224)"
    ok, note = d.q4_links_check(t, {"onenote": _page(q4, ["Goal"], 1, live)}, c)
    assert ok, note
    assert not d.q4_links_check(t, {"onenote": _page(q4, ["Goal"], 1, live.replace("(HOL r1 ", "(HOL r2 "))}, c)[0]
    assert "daily" in d.q4_links_check(t, {"onenote": _page(q4, ["Goal"], 1, live.replace("Daily 2026-11-06", "Daily 2026-11-05"))}, c)[1]
    assert d.link_listed(fx["decision"], "Backlinks: Decisions launch meeting (HOL r1 1106-121224) (within")
    assert not d.link_listed(fx["decision"], "Backlinks: Decisions launch meeting (HOL r12 1106-121224)")
    assert not d.link_listed(fx["decision"], "Backlinks: Q4 Launch (HOL r1 1106-121224)")


def test_ask_prompt_and_check_use_only_this_runs_pages():
    """A8 live (L14 r1): brain_ask returned older runs' pricing pages; only this run's pages count as cited or opened."""
    d = _driver()
    c = _ctx(d, run=1, stamp="1106-121224")
    fx = c["fx"]
    step = next(s for s in d.loop_steps() if s["name"] == "L14_ask_cited")
    prompt = step["prompt"](c)
    assert fx["tag"] in prompt and "brain_ask" in prompt
    mine = fx["titles"]["Pricing review 2026-10-20"]
    older = "Pricing review 2026-10-20 (HOL r1 1106-111023)"
    ask = {"results": [{"title": older}, {"title": mine}], "decisions": [], "coverage": {"label": "indexed 38 of 82 pages"}}
    def turn(text):
        return {"tools": [{"name": "brain_ask", "result": {"ok": True, "data": ask}},
                          {"name": "onenote_open_page", "result": {"ok": True, "data": {}}}], "text": text}
    assert d.ask_check(turn(f"Keep $29 ({mine}). indexed 38 of 82 pages"), {"onenote": _page(mine, ["x"])}, c)[0]
    ok, note = d.ask_check(turn(f"Keep $29 ({older}). indexed 38 of 82 pages"), {"onenote": _page(older, ["x"])}, c)
    assert not ok and "this run" in note
    # this run's meeting page opened, but brain_ask never returned it: not a cited source
    meeting = fx["titles"]["Launch meeting 2026-11-03"]
    assert not d.ask_check(turn(f"Keep $29 ({mine}; {meeting}). indexed 38 of 82 pages"), {"onenote": _page(meeting, ["x"])}, c)[0]


def test_powerpoint_readback_reads_text_only_from_text_shapes():
    """A8 live (L11): loading textFrame on a picture fails the whole batch with RichApi InvalidArgument."""
    d = _driver()
    js = d.READBACK["powerpoint"]
    assert "textFrame/textRange" not in js
    assert "'GeometricShape', 'TextBox', 'Placeholder'" in js and "catch" in js
    assert "x.type === 'Image'" in d.RESET["powerpoint"]  # the disposable deck starts without pictures


def _undo(ok=True, msg=None, page=None):
    if ok:
        return {"name": "office_undo", "result": {"ok": True, "data": {"undone": "brain_link_both_ways", "pageId": page}}}
    return {"name": "office_undo", "result": {"ok": False, "error": {"category": "Unavailable", "message": msg}}}


def test_both_ways_undo_retries_only_navigation_that_did_not_land():
    """A8 live: the 2nd of 3 undos failed with 'OneNote did not open the page' (nothing written). The harness may open the
    page itself and ask for the missing undos once; any other failure is never retried."""
    d = _driver()
    nav = "OneNote did not open the page"
    t = {"tools": [_undo(page="r"), _undo(False, nav), _undo(page="l")], "text": ""}
    assert d.undo_retry_count(t, 3) == 1
    assert d.undo_retry_count({"tools": [_undo(page="r"), _undo(False, "Undo failed: boom"), _undo(page="l")]}, 3) == 0
    assert d.undo_retry_count({"tools": [_undo(page="r"), _undo(page="l"), _undo(page="p")]}, 3) == 0
    assert d.undo_retry_count({"tools": [_undo(page="r"), _undo(False, nav)]}, 2) == 1
    t2 = {"tools": [_undo(page="p")], "text": "done", "ms": 5}
    merged = d.merge_undo_retry(t, t2)
    assert [x["result"]["data"]["pageId"] for x in merged["tools"]] == ["r", "l", "p"]
    assert merged["navRetry"]["notLanded"] == 1 and merged["navRetry"]["retried"] == 1
    c = {"bw_order": ["p", "l", "r"], "bw_before": {}, "bw_undo_after": {}, "fx": {"titles": {}}}
    ok, note = d.both_ways_undo_check(merged, {}, c)  # judged on the merged results: newest first, all ok
    assert ok, note
    assert not d.both_ways_undo_check(t, {}, c)[0]  # without the retry the same turn fails


def test_ledger_next_undo_page_is_read_only(tmp_path):
    d = _driver()
    db = tmp_path / "brain.db"
    con = sqlite3.connect(db)
    con.execute("create table writes (id integer primary key, notebook_id text, page_id text, tool text, status text)")
    con.executemany("insert into writes (notebook_id, page_id, tool, status) values (?,?,?,?)",
                    [("nb1", "p", "brain_link_both_ways", "done"), ("nb1", "l", "brain_link_both_ways", "done"),
                     ("nb1", "r", "brain_link_both_ways", "undone"), ("nb2", "z", "brain_links", "done")])
    con.commit()
    con.close()
    assert d.ledger_next_undo_page(db, "nb1") == "l"
    assert d.ledger_next_undo_page(db, "nb3") is None
    assert d.ledger_next_undo_page(tmp_path / "missing.db", "nb1") is None


def test_deck_check_reports_an_unreadable_slide_instead_of_raising():
    d = _driver()
    c = _ctx(d)
    t, rbs = _deliv(d, c)
    rbs["powerpoint"]["slides"][3] = 5  # not a list of texts
    ok, note = d.deliverables_check(t, rbs, c)
    assert not ok and "does not show" in note


def test_error_where_names_the_failing_frame_without_values():
    d = _driver()
    try:
        [x for x in 5]
    except TypeError as exc:
        where = d.error_where(exc)
    assert where and where[-1].startswith("test_livetest_loop.py:") and "5" not in where[-1].split(" ", 1)[1]


def test_deliverables_judged_per_host_across_the_receipts_retry_for_one_host():
    """Live A11: PowerPoint failed in the first call; the model retried hosts=['powerpoint'] as the receipt says."""
    d = _driver()
    c = _ctx(d)
    t, rbs = _deliv(d, c)
    full = t["tools"][0]["result"]["data"]
    first = {**full, "status": "partial", "results": full["results"][:2] + [{"host": "powerpoint", "status": "failed"}]}
    retry = {**full, "results": [full["results"][2]]}
    t["tools"] = [{"name": "brain_make_deliverables", "result": {"ok": True, "data": first}},
                  {"name": "brain_make_deliverables", "result": {"ok": True, "data": retry}}]
    ok, note = d.deliverables_check(t, rbs, c)
    assert ok, note
    assert "2 brain_make_deliverables calls" in note
    c2 = {}
    d._after_deliverables(t, c2)
    assert c2["tracker_sheet"] == "Decision tracker (2)"
    t["tools"] = t["tools"][:1]  # only the failed call
    assert not d.deliverables_check(t, rbs, c)[0]


def test_tracker_check_refuses_a_readback_that_is_not_the_tracker_sheet():
    d = _driver()
    c = _ctx(d)
    t, rbs = _deliv(d, c)
    ok, note = d.tracker_check({"values": [["#", "Item"]], "formulas": [["#", "Item"]], "charts": 1}, c)
    assert not ok and "not the tracker sheet" in note
