# SPDX-License-Identifier: Apache-2.0
"""Live-test harness: the OneNote fixture is fresh per run and a reused page is reported as stale."""
import importlib.util
import json
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent


def _driver():
    name = "hol_live_driver"
    if name in sys.modules:
        return sys.modules[name]
    spec = importlib.util.spec_from_file_location(name, ROOT / "scripts" / "livetest" / "live_driver.py")
    mod = importlib.util.module_from_spec(spec)
    sys.modules[name] = mod
    spec.loader.exec_module(mod)
    return mod


def _brain():
    return sys.modules["hermes_office_desktop.brain"]  # loaded by conftest's plugin fixture


def test_fixture_titles_and_links_are_run_tagged(plugin):
    d = _driver()
    a, b = d.onenote_fixture(1, "1105-100000"), d.onenote_fixture(2, "1105-100500")
    assert a["section"] == "HOL r1 1105-100000" and a["section"] != b["section"]
    assert set(a["titles"]) == {"Q4 Launch", "Risks", "Launch meeting 2026-11-03", "Pricing review 2026-10-20", "Daily 2026-11-04"}
    assert not set(a["titles"].values()) & set(b["titles"].values())
    for v in a["titles"].values():
        assert f"- title: {v}\n" in a["seed"]
    links = [t for t, _ in _brain().parse_links(a["seed"])]
    assert links and all(t in a["titles"].values() for t in links), links  # no link to an older, untagged page
    assert "[[Pricing review 2026-10-20 (HOL r1 1105-100000)|October pricing review]]" in a["seed"]
    assert f"Section '{a['section']}':" in a["seed"] and "Section 'Projects'" not in a["seed"]


def test_reset_creates_only_a_new_section():
    js = _driver().onenote_reset_js('HOL r1 "x"')
    assert json.dumps('HOL r1 "x"') in js and "addSection(name)" in js and "stale" in js
    assert "delete" not in js  # the harness never deletes notebook content


def test_links_precondition_flags_a_reused_page():
    d = _driver()
    c = {"fx": d.onenote_fixture(1, "1105-100000")}
    q4 = c["fx"]["titles"]["Q4 Launch"]
    assert d.links_precondition({"pageTitle": q4, "linksBlocks": 0}, c) == (True, "clean")
    ok, why = d.links_precondition({"pageTitle": q4, "linksBlocks": 1}, c)
    assert not ok and why.startswith("stale fixture:") and "another profile" in why
    ok, why = d.links_precondition({"pageTitle": "Q4 Launch", "linksBlocks": 0}, c)
    assert not ok and why.startswith("stale fixture:")


def test_seed_check_requires_every_page_and_rejects_duplicates():
    d = _driver()
    titles = d.onenote_fixture(1, "1105-100000")["titles"]
    ok_calls = [{"name": "onenote_create_page", "args": json.dumps({"title": v}), "result": {"ok": True, "data": {"title": v}}}
                for v in titles.values()]
    assert d.seed_created({"tools": ok_calls}, titles)[0]
    assert not d.seed_created({"tools": ok_calls[1:]}, titles)[0]
    dup = {"name": "onenote_create_page", "args": {"title": "x"}, "result": {"ok": False, "error": {"category": "DuplicateTitle"}}}
    ok, note = d.seed_created({"tools": ok_calls + [dup]}, titles)
    assert not ok and note.startswith("stale fixture:")


# ---- stronger assertions: no truncated, partial or self-reported evidence counts as a pass ----------
def test_ok_rejects_truncated_results_and_unverified_composites():
    d = _driver()
    assert d._ok({"ok": True, "data": {}}) and d._ok('{"ok": true, "data": {"a": 1}}')
    assert not d._ok('{"ok": true, "data": {"pages": [{"title": "Q4 La')  # truncated in the event stream
    assert not d._ok(None) and not d._ok({"ok": False})
    good = {"status": "verified", "results": [{"host": "excel", "status": "verified"}, {"host": "word", "status": "verified"}]}
    assert d._ok({"ok": True, "data": good})
    assert not d._ok({"ok": True, "data": {**good, "status": "partial"}})
    assert not d._ok({"ok": True, "data": {**good, "results": good["results"][:1] + [{"host": "word", "status": "applied"}]}})
    assert d._ok({"ok": True, "data": {"results": [{"title": "a search hit"}]}})  # not a composite receipt
    tools = [{"name": "brain_sync", "result": '{"ok": true, "da'}, {"name": "brain_ask", "result": {"ok": True}}]
    assert d.truncated_results(tools) == ["brain_sync"]


def _links_rb(d, user="Launch hub. Background...", block="Links & Backlinks (kept by Hermes) Links: Risks Backlinks: Kickoff",
              n=1, title=None, c=None):
    return {"pageTitle": title or c["fx"]["titles"]["Q4 Launch"], "linksBlocks": n, "text": [user] + [block] * n}


def test_onenote_undo_requires_the_exact_previous_block_and_untouched_user_text():
    d = _driver()
    c = {"fx": d.onenote_fixture(1, "1105-100000")}
    twice = {"tools": [{"name": "brain_links", "result": {"ok": True, "data": {}}}] * 2}
    assert d.links_twice_check(twice, _links_rb(d, c=c), c)[0]
    undo = {"tools": [{"name": "office_undo", "result": {"ok": True, "data": {"undone": "brain_links",
                                                                             "restoredPreviousLinksBlock": True}}}]}
    assert d.undo_links_check(undo, _links_rb(d, c=c), c) == (True, "exactly one block, identical to the first; user text unchanged")
    assert not d.undo_links_check(undo, _links_rb(d, n=0, c=c), c)[0]  # block removed, previous one not restored
    assert not d.undo_links_check(undo, _links_rb(d, n=2, c=c), c)[0]
    assert not d.undo_links_check(undo, _links_rb(d, user="edited by Hermes", c=c), c)[0]
    assert not d.undo_links_check(undo, _links_rb(d, block="Links & Backlinks other", c=c), c)[0]
    wrong = {"tools": [{"name": "office_undo", "result": {"ok": True, "data": {"undone": "onenote_append"}}}]}
    assert not d.undo_links_check(wrong, _links_rb(d, c=c), c)[0]
    cut = {"tools": [{"name": "office_undo", "result": '{"ok": true, "data": {"undone": "brain_li'}]}
    assert not d.undo_links_check(cut, _links_rb(d, c=c), c)[0]
    assert not d.undo_links_check(undo, _links_rb(d, c=c), {"fx": c["fx"]})[0]  # nothing recorded before the undo


def _decision(d, c, **over):
    meeting = c["fx"]["titles"]["Launch meeting 2026-11-03"]
    items = [{"kind": "decision", "text": "Run a pilot in two cities", "owner": "Unknown (not in the notes)"},
             {"kind": "action", "text": "Draft the pilot plan", "owner": "Priya"},
             {"kind": "action", "text": "Check supplier lead times", "owner": "Unknown (not in the notes)"}]
    data = {"title": "Decisions: launch meeting (HOL run 1 101500)", "items": items, "unknownOwners": 2, "unverifiedQuotes": 0}
    data.update(over)
    text = [f"Source: {meeting} Decisions D1. Run a pilot in two cities · owner: Unknown (not in the notes) "
            "Actions A1. Draft the pilot plan · owner: Priya · due: Nov 12 "
            "A2. Check supplier lead times · owner: Unknown (not in the notes) · due: Unknown (not in the notes)"]
    rb = {"pageTitle": data["title"], "text": text}
    return {"tools": [{"name": "brain_decision_page", "result": {"ok": True, "data": data}}]}, rb


def test_decision_page_step_checks_the_page_content_not_the_envelope():
    d = _driver()
    c = {"fx": d.onenote_fixture(1, "1105-100000")}
    t, rb = _decision(d, c)
    ok, note = d.decision_page_check(t, rb, c)
    assert ok, note
    t2, rb2 = _decision(d, c, unknownOwners=1)  # the tool under-reports unknown owners
    assert not d.decision_page_check(t2, rb2, c)[0]
    t3, rb3 = _decision(d, c)
    t3["tools"][0]["result"]["data"]["items"][2]["owner"] = "Sam"  # an invented owner
    assert "supplier" in d.decision_page_check(t3, rb3, c)[1]
    t4, rb4 = _decision(d, c)
    rb4["text"] = [rb4["text"][0].replace("Draft the pilot plan", "Draft a plan")]  # page does not show the item
    assert "not on the page" in d.decision_page_check(t4, rb4, c)[1]
    t5, rb5 = _decision(d, c)
    rb5["pageTitle"] = "Q4 Launch"
    assert not d.decision_page_check(t5, rb5, c)[0]
    assert not d.decision_page_check({"tools": [{"name": "brain_decision_page", "result": '{"ok": true'}]}, rb, c)[0]


def test_decision_page_source_line_accepts_onenotes_shortened_page_link_label_but_not_a_missing_one():
    """A8 live: OneNote for the web shows the link to the meeting page as 'Launch meeting 2026-11-03 (HOL r1'."""
    d = _driver()
    c = {"fx": d.onenote_fixture(1, "1105-100000")}
    meeting = c["fx"]["titles"]["Launch meeting 2026-11-03"]
    t, rb = _decision(d, c)
    rb["text"] = [rb["text"][0].replace(meeting, "Launch meeting 2026-11-03 (HOL r1\u00a0")]
    ok, note = d.decision_page_check(t, rb, c)
    assert ok, note
    t2, rb2 = _decision(d, c)
    rb2["text"] = [rb2["text"][0].replace("Source: " + meeting, "Source: Pricing review 2026-10-20")]
    assert "source line missing" in d.decision_page_check(t2, rb2, c)[1]


def test_onenote_readback_reads_list_items_nested_under_their_heading():
    """A8 live: <ul><li> items are child paragraphs in OneNote for the web; the readback reads one level down."""
    js = _driver().READBACK["onenote"]
    assert "x.paragraphs" in js and "kids[i].items" in js


def test_results_record_plugin_commit_served_sums_and_run_id(tmp_path):
    import hashlib
    import subprocess
    d = _driver()
    body = b"abc  index.html\n"
    ident = d.run_identity(tmp_path / "20261106-101500", fetch=lambda url: body)
    head = subprocess.run(["git", "-C", str(ROOT), "rev-parse", "HEAD"], capture_output=True, text=True).stdout.strip()
    assert ident["runId"] == "20261106-101500" and ident["pluginCommit"] == head and isinstance(ident["pluginDirty"], bool)
    assert ident["servedSums"]["sha256"] == hashlib.sha256(body).hexdigest()
    assert ident["servedSums"]["url"].endswith("/office/live/SHA256SUMS")

    def down(url):
        raise OSError("offline")
    assert d.run_identity(tmp_path / "r", fetch=down)["servedSums"] == {
        "url": ident["servedSums"]["url"], "sha256": None, "error": "OSError: offline"}
    rec = d.Recorder(tmp_path / "out", identity=ident)
    r = rec.add({"host": "onenote", "run": 1, "fixture": {"section": "HOL r1 x", "tag": "HOL r1 x"}, "steps": []})
    rec.save()
    saved = json.loads((tmp_path / "out" / "results.json").read_text())
    assert saved[0]["identity"]["pluginCommit"] == head and saved[0]["fixture"]["tag"] == "HOL r1 x" and r is rec.results[0]


# ---- --images steps (implemented, run live only by FRIDAY) ------------------------------------------
def test_image_fixtures_match_what_the_plugin_accepts_and_refuses(plugin, tmp_path, monkeypatch):
    d = _driver()
    images = sys.modules["hermes_office_desktop.images"]
    home = tmp_path / "hermes"
    monkeypatch.setenv("HERMES_HOME", str(home))
    fx = d.image_fixtures(home, tmp_path / "out", "excel-r1-120000")
    assert images.load_path(fx["path"]).kind == "png"
    for key, cat in (("oversize", "Capacity"), ("wrongType", "WrongType"), ("outside", "Blocked")):
        try:
            images.load_path(fx[key])
        except images.ImageError as exc:
            assert exc.category == cat, key
        else:
            raise AssertionError(f"{key} was accepted")
    try:
        images.fetch_url(fx["privateUrl"])  # an IP literal: refused before any lookup or connection
    except images.ImageError as exc:
        assert exc.category == "Blocked"
    else:
        raise AssertionError("private URL accepted")
    assert fx["url"].startswith("https://")


def test_image_steps_per_host_and_checks():
    d = _driver()
    for host in ("excel", "word", "powerpoint", "onenote"):
        names = [n for n, _, _ in d.image_steps(host)]
        assert names == ["img_path", "img_url", "img_html", "neg_oversize", "neg_wrong_type", "neg_outside_path",
                         "neg_private_url", "img_undo"], host
        assert all(n.startswith(d.IMAGE_STEP_PREFIXES) for n in names) and host in d.IMAGE_READBACK
    c = {"sheet": "Sheet1", "img": {"path": "/p.png", "url": "https://u/x.png", "credit": "C", "oversize": "/o",
                                    "wrongType": "/w", "outside": "/x", "privateUrl": "https://127.0.0.1/x.png"},
         "img_before": {"images": 2}}
    steps = {n: (p, chk) for n, p, chk in d.image_steps("excel")}
    assert "cell 'D2'" in steps["img_path"][0](c) and "/p.png" in steps["img_path"][0](c)
    ok_turn = {"tools": [{"name": "office_insert_image", "result": {"ok": True, "data": {"status": "verified"}}}]}
    assert steps["img_path"][1](ok_turn, {"images": 3}, c)[0] is True
    assert steps["img_path"][1](ok_turn, {"images": 2}, c)[0] is False  # the tool said verified, the document disagrees
    raster = {"tools": [{"name": "office_insert_image", "result": {"ok": False, "error": {"category": "NeedsRaster"}}}]}
    assert steps["img_html"][1](raster, {"images": 2}, c)[0] is True
    assert steps["neg_private_url"][1](raster, {"images": 2}, c)[0] is False  # wrong category
    blocked = {"tools": [{"name": "office_insert_image", "result": {"ok": False, "error": {"category": "Blocked"}}}]}
    assert steps["neg_private_url"][1](blocked, {"images": 2}, c)[0] is True
    assert steps["neg_private_url"][1](blocked, {"images": 3}, c)[0] is False  # something landed anyway
    undo = {"tools": [{"name": "office_undo", "result": {"ok": True, "data": {"undone": "insert_image"}}}]}
    assert steps["img_undo"][1](undo, {"images": 1}, c)[0] is True
    word_html = {n: chk for n, _, chk in d.image_steps("word")}["img_html"]
    assert word_html(ok_turn, {"images": 3}, c)[0] is True  # Word inserts HTML natively


# ---- two-way links: panel + brain_link_both_ways live steps (logic only; not run live) ---------------
def test_onenote_script_ends_with_panel_both_ways_and_undo_and_targets_match_the_seed(plugin):
    d = _driver()
    names = [s[0] for s in d.onenote_steps()]
    assert names[-3:] == ["14_panel_backlink", "15_both_ways", "16_both_ways_undo"]
    c = {"fx": d.onenote_fixture(1, "1105-100000"), "run": 1}
    steps = {s[0]: s for s in d.onenote_steps()}
    assert steps["14_panel_backlink"][1](c).startswith(f"PANEL:{c['fx']['section']}/")
    q4 = c["fx"]["titles"]["Q4 Launch"]
    assert f"brain_link_both_ways once with title {q4!r}" in steps["15_both_ways"][1](c)
    assert "exactly 3 times" in steps["16_both_ways_undo"][1](c)
    q4_seed = next(p for s in __import__("json").loads((ROOT / "addin" / "dev" / "demo-notebook.json").read_text())["sections"]
                   for p in s["pages"] if p["title"] == "Q4 Launch")
    assert sorted(t for t, _ in _brain().parse_links("".join(q4_seed["html"]))) == sorted(d.BOTH_WAYS_TARGETS)
    assert set(d.STEP_HOOKS) == {"15_both_ways", "16_both_ways_undo"}


def _panel_ctx(d, **over):
    c = {"fx": d.onenote_fixture(1, "1105-100000")}
    t = c["fx"]["titles"]
    c["panel"] = {"present": True, "hidden": False, "groups": {"Links": ["October pricing review", t["Risks"]],
                                                               "Backlinks": [t["Pricing review 2026-10-20"]]},
                  "text": "Links … indexed 4 of 6 pages. Ask Hermes to run brain_sync (with crawl) …", "markup": False}
    c["panel_ledger_before"] = c["panel_ledger_after"] = {"rows": 3, "maxId": 3, "lastTool": "brain_links"}
    for k, v in over.items():
        c[k] = v
    return c, {"pageTitle": t["Q4 Launch"]}


def test_panel_check_needs_the_backlink_coverage_and_no_write():
    d = _driver()
    c, rb = _panel_ctx(d)
    assert d.panel_check({"tools": []}, rb, c)[0]
    assert not d.panel_check({"tools": [{"name": "brain_links"}]}, rb, c)[0]
    c2, _ = _panel_ctx(d, panel_ledger_after={"rows": 4})
    assert "ledger changed" in d.panel_check({"tools": []}, rb, c2)[1]
    c3, _ = _panel_ctx(d)
    c3["panel"]["groups"]["Backlinks"] = []
    assert "not shown" in d.panel_check({"tools": []}, rb, c3)[1]
    c4, _ = _panel_ctx(d)
    c4["panel"]["hidden"] = True
    assert not d.panel_check({"tools": []}, rb, c4)[0]
    c5, _ = _panel_ctx(d)
    c5["panel"]["markup"] = True
    assert "markup" in d.panel_check({"tools": []}, rb, c5)[1]


def _bw(d):
    c = {"fx": d.onenote_fixture(1, "1105-100000")}
    t = c["fx"]["titles"]
    q4 = t["Q4 Launch"]
    c["bw_before"] = {k: {"pageTitle": t[k], "linksBlocks": 0, "text": [f"user text of {k}"]} for k in d.BOTH_WAYS_TARGETS}
    c["bw_after"] = {k: {"pageTitle": t[k], "linksBlocks": 1,
                         "text": [f"user text of {k}", f"Links & Backlinks (kept by Hermes) Links: none Backlinks: {q4} (within …)"]}
                     for k in d.BOTH_WAYS_TARGETS}
    pages = [{"pageId": f"id-{i}", "title": t[k], "status": "verified"} for i, k in enumerate(d.BOTH_WAYS_TARGETS)]
    turn = {"tools": [{"name": "brain_link_both_ways", "result": {"ok": True, "data": {
        "status": "verified", "pages": pages, "returnedTo": {"title": q4, "ok": True}, "undo": {"calls": 3}}}}],
        "approvals": [{"choice": "once", "description": f"OneNote · n · visibly open 3 pages that '{q4}' links to: …"}]}
    return c, turn, {"pageTitle": q4}


def test_both_ways_check_requires_every_page_verified_one_approval_and_return():
    d = _driver()
    c, turn, rb = _bw(d)
    assert d.both_ways_check(turn, rb, c)[0] and c["bw_order"] == ["id-0", "id-1", "id-2"] and c["bw_undo_calls"] == 3
    c, turn, rb = _bw(d)
    assert not d.both_ways_check(turn, {"pageTitle": "elsewhere"}, c)[0]
    c, turn, rb = _bw(d)
    turn["approvals"] = turn["approvals"] * 2
    assert "approval" in d.both_ways_check(turn, rb, c)[1]
    c, turn, rb = _bw(d)
    turn["tools"][0]["result"]["data"]["pages"][1]["status"] = "applied"
    assert not d.both_ways_check(turn, rb, c)[0]
    c, turn, rb = _bw(d)
    c["bw_after"]["Risks"]["text"][0] = "user text rewritten"
    assert "user text changed" in d.both_ways_check(turn, rb, c)[1]
    c, turn, rb = _bw(d)
    c["bw_after"]["Risks"]["text"][1] = "Links & Backlinks Links: none Backlinks: none"
    assert not d.both_ways_check(turn, rb, c)[0]


def test_both_ways_undo_check_requires_lifo_and_the_earlier_state():
    d = _driver()
    c, turn, rb = _bw(d)
    d.both_ways_check(turn, rb, c)
    c["bw_undo_after"] = {k: dict(v) for k, v in c["bw_before"].items()}

    def undos(ids):
        return {"tools": [{"name": "office_undo", "result": {"ok": True, "data": {"undone": "brain_link_both_ways", "pageId": i}}}
                          for i in ids]}
    assert d.both_ways_undo_check(undos(["id-2", "id-1", "id-0"]), rb, c)[0]
    assert "LIFO" in d.both_ways_undo_check(undos(["id-0", "id-1", "id-2"]), rb, c)[1]
    assert not d.both_ways_undo_check(undos(["id-2", "id-1"]), rb, c)[0]
    c["bw_undo_after"]["Risks"] = c["bw_after"]["Risks"]
    assert "differs" in d.both_ways_undo_check(undos(["id-2", "id-1", "id-0"]), rb, c)[1]
    assert not d.both_ways_undo_check(undos([]), rb, {"fx": c["fx"]})[0]
