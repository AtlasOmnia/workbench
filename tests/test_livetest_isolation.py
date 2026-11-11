# SPDX-License-Identifier: Apache-2.0
"""scripts/livetest/isolation.py verify(): the two-profile live check's judgement, from recorded phases."""
import importlib.util
import json
import sqlite3
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent


def _iso():
    sys.path.insert(0, str(ROOT / "scripts" / "livetest"))
    spec = importlib.util.spec_from_file_location("hol_isolation", ROOT / "scripts" / "livetest" / "isolation.py")
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    return mod


def _db(path, ops):
    path.parent.mkdir(parents=True, exist_ok=True)
    c = sqlite3.connect(path)
    c.execute("create table writes(op_id text, status text, kind text, page_id text, notebook_id text)")
    c.executemany("insert into writes values(?,?,?,?,?)", [(o, "done", "links_outline", "p", "nb") for o in ops])
    c.commit()
    c.close()


def _world(tmp_path, monkeypatch, *, b_ops=("op-b",), probe_blocks=None, undos=None, same=False):
    iso = _iso()
    monkeypatch.setattr(iso, "PROFILES_ROOT", tmp_path / "profiles")
    a_db = tmp_path / "profiles" / "a" / iso.DB
    b_db = a_db if same else tmp_path / "profiles" / "b" / iso.DB
    _db(a_db, ["op-a1", "op-a2"])
    if not same:
        _db(b_db, list(b_ops))
    root = tmp_path / "root.db"
    _db(root, ["old"])
    block_a = "Links & Backlinks ... A"
    phases = {
        "write": {"profile": "a", "dbFile": str(a_db), "blocks": [block_a], "tag": "t", "pageId": "p"},
        "probe": {"profile": "b", "dbFile": str(b_db), "blocks": probe_blocks if probe_blocks is not None else [block_a],
                  "undos": undos if undos is not None else [{"ok": True}, {"ok": False, "category": "NothingToUndo"}],
                  "ledger": {"rows": [[o, "undone", "links_outline", "p"] for o in b_ops]}},
        "cleanup": {"profile": "a", "dbFile": str(a_db), "blocks": []}}
    for k, v in phases.items():
        (tmp_path / f"{k}.json").write_text(json.dumps(v))
    return iso, root


def test_verify_passes_two_separate_ledgers_and_an_untouched_block(tmp_path, monkeypatch):
    iso, root = _world(tmp_path, monkeypatch)
    v = iso.verify(tmp_path, iso.table_counts(root), root)
    assert v["pass"], v["problems"]
    assert v["aRows"] == 2 and v["bRows"] == 1 and v["rootUnchanged"]


def test_verify_fails_a_shared_file_a_replaced_block_or_an_undo_of_a(tmp_path, monkeypatch):
    iso, root = _world(tmp_path, monkeypatch, same=True)
    assert "A and B share one brain.db" in iso.verify(tmp_path, iso.table_counts(root), root)["problems"]
    iso, root = _world(tmp_path / "2", monkeypatch, probe_blocks=["Links & Backlinks ... B"])
    assert any("A's block is not on the page" in p for p in iso.verify(tmp_path / "2", iso.table_counts(root), root)["problems"])
    iso, root = _world(tmp_path / "3", monkeypatch, undos=[{"ok": True}, {"ok": True}])
    assert any("B's undos" in p for p in iso.verify(tmp_path / "3", iso.table_counts(root), root)["problems"])


def test_verify_fails_when_the_shared_root_db_changed(tmp_path, monkeypatch):
    iso, root = _world(tmp_path, monkeypatch)
    before = iso.table_counts(root)
    c = sqlite3.connect(root)
    c.execute("insert into writes values('new','done','k','p','nb')")
    c.commit()
    c.close()
    assert "the shared default-profile brain.db changed" in iso.verify(tmp_path, before, root)["problems"]
