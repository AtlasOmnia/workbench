# SPDX-License-Identifier: Apache-2.0
"""Two-profile isolation, live: one gateway, two Hermes profiles, one OneNote page.

Each phase runs in its own process with HERMES_HOME set to the profile it drives (scripts/livetest/
run_isolation.sh runs them in order against the client Mac harness window):

  write   (profile A) prepare the OneNote tab and a fresh section, pair a new chat, create one page and
          write A's Links & Backlinks block there with brain_links.
  probe   (profile B) pair a new chat in B, run brain_links on THAT page, then office_undo twice. B may add
          its own block, but must never remove or replace A's; its undos may only touch B's own block.
  cleanup (profile A) pair a new chat in A again and undo A's block.
  verify  (no gateway) the two profiles' brain.db files are different files, neither holds the other's
          ledger rows, and the shared default-profile brain.db did not change.

Every phase writes <out>/<phase>.json; verify reads them. Nothing here edits a database: ledgers are
opened read-only.
"""
from __future__ import annotations

import argparse
import asyncio
import json
import os
import sqlite3
import sys
import time
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
import live_driver as d  # noqa: E402
import office_ui as ui  # noqa: E402
from chat import Chat, connect_live, pairing_code  # noqa: E402

DB = Path("plugin-data") / "hermes-office-desktop" / "brain.db"
ROOT_DB = Path(os.environ.get("HOL_HERMES_ROOT", str(Path.home() / ".hermes"))) / DB
PROFILES_ROOT = Path(os.environ.get("HOL_PROFILES_ROOT", str(Path.home() / ".hermes" / "profiles")))


def profile_home() -> Path:
    return Path(os.environ["HERMES_HOME"])


def ledger(db: Path, nb: str | None = None) -> dict:
    """Read-only: every ledger row's op id/status/kind/page (no content)."""
    if not db.exists():
        return {"exists": False, "rows": []}
    con = sqlite3.connect(f"file:{db}?mode=ro", uri=True, timeout=5)
    try:
        q = "select op_id, status, kind, page_id from writes" + (" where notebook_id=?" if nb else "")
        rows = [list(r) for r in con.execute(q, (nb,) if nb else ())]
    finally:
        con.close()
    return {"exists": True, "rows": rows}


def table_counts(db: Path) -> dict:
    con = sqlite3.connect(f"file:{db}?mode=ro", uri=True, timeout=5)
    try:
        names = [r[0] for r in con.execute("select name from sqlite_master where type='table' and name not like 'sqlite_%' "
                                           "and name not like 'pages_fts_%' order by name")]
        return {n: con.execute(f'select count(*) from "{n}"').fetchone()[0] for n in names}
    finally:
        con.close()


def links_blocks(rb: dict) -> list[str]:
    return [t for t in (rb or {}).get("text") or [] if "Links & Backlinks" in t]


async def pair(client, pane, rec, title):
    chat = await Chat.create(client, title, rec.log)
    t = await chat.turn("connect to OneNote")
    code, source = pairing_code(t["text"], t["tools"])
    st = await asyncio.to_thread(ui.type_code, pane, "onenote", code, rec.log) if code else {"state": "no-code"}
    if st.get("state") != "paired":
        raise ui.HarnessError(f"pairing failed: {st}")
    return chat


def results(t, name):
    return [d.parsed(x.get("result")) for x in t["tools"] if x.get("name") == name]


async def phase(a) -> dict:
    out = Path(a.out)
    rec = d.Recorder(out / a.phase)
    pane = ui.Pane(a.control, rec.log)
    pane.wait_guest(timeout=d.GUEST_WAIT_S)
    res = {"phase": a.phase, "profile": profile_home().name, "started": d.now(), "identity": rec.identity}
    async with connect_live() as client:
        if a.phase == "write":
            tag = f"HOL iso {time.strftime('%m%d-%H%M%S')}"
            manifests = {h: f"/Users/{a.mac_user}/{d.MAC_ROOT}/manifests/manifest-{h}.xml" for h in d.LABEL}
            pane.cmd("clearCache")
            await asyncio.to_thread(d.prepare_host, pane, "onenote", rec, manifests, 1, {"section": tag})
            chat = await pair(client, pane, rec, f"HOL isolation A write {tag}")
            title = f"Isolation page ({tag})"
            t = await chat.turn(f"Call onenote_list_sections, then call onenote_create_page with title '{title}', html "
                                f"'<p>Isolation check. Links: [[Isolation target ({tag})]]</p>' and the sectionId of the "
                                f"section '{tag}'. Then call brain_links on that new page. Do nothing else.")
            created = [r for r in results(t, "onenote_create_page") if d._ok(r)]
            if not created:
                raise ui.HarnessError("onenote_create_page did not succeed")
            data = (created[-1] or {}).get("data") or {}
            res.update(tag=tag, title=title, pageId=data.get("id") or data.get("pageId"),
                       linksOk=any(d._ok(r) for r in results(t, "brain_links")))
        else:
            prev = json.loads((out / "write.json").read_text())
            chat = await pair(client, pane, rec, f"HOL isolation {a.phase} {prev['tag']}")
            if a.phase == "probe":
                t = await chat.turn(f"Call brain_links with pageId '{prev['pageId']}'. Then call office_undo with app "
                                    "'onenote'. Then call office_undo with app 'onenote' once more. Report each tool "
                                    "result exactly. Do nothing else.")
                res["links"] = [r and {"ok": r.get("ok"), "replacedPreviousBlock": (r.get("data") or {}).get("replacedPreviousBlock")}
                                for r in results(t, "brain_links")]
                res["undos"] = [r and {"ok": r.get("ok"), "category": (r.get("error") or {}).get("category"),
                                       "what": (r.get("data") or {}).get("undone") or (r.get("data") or {}).get("kind")}
                                for r in results(t, "office_undo")]
            else:  # cleanup: A undoes its own block
                t = await chat.turn("Call office_undo with app 'onenote' once. Report the result exactly. Do nothing else.")
                res["undos"] = [r and {"ok": r.get("ok"), "category": (r.get("error") or {}).get("category")}
                                for r in results(t, "office_undo")]
        rb = d.readback(pane, "onenote")
        res.update(pageTitle=rb.get("pageTitle"), notebookId=rb.get("notebookId"), blocks=links_blocks(rb),
                   tools=[x["name"] for x in d.office_tools(t["tools"])], reply=d.ui_redact(t["text"])[:800])
    res["ledger"] = ledger(profile_home() / DB, res.get("notebookId"))
    res["dbFile"] = str((profile_home() / DB).resolve())
    res["ended"] = d.now()
    return res


def verify(out: Path, root_before: dict, root_db: Path | None = None) -> dict:
    w, p, c = (json.loads((out / f"{n}.json").read_text()) for n in ("write", "probe", "cleanup"))
    a_db, b_db = Path(w["dbFile"]), Path(p["dbFile"])
    problems = []
    if w["profile"] == p["profile"]:
        problems.append("A and B are the same profile")
    if a_db == b_db or (a_db.exists() and b_db.exists() and a_db.stat().st_ino == b_db.stat().st_ino):
        problems.append("A and B share one brain.db")
    for name, db in (("A", a_db), ("B", b_db)):
        if not (db.exists() and db.is_relative_to(PROFILES_ROOT)):
            problems.append(f"{name}'s brain.db is not inside its own profile ({db})")
    a_ops = {r[0] for r in ledger(a_db)["rows"]}
    b_ops = {r[0] for r in ledger(b_db)["rows"]}
    if a_ops & b_ops:
        problems.append(f"{len(a_ops & b_ops)} ledger row(s) in both profiles")
    if len(w["blocks"]) != 1:
        problems.append(f"after A's write the page has {len(w['blocks'])} Links & Backlinks blocks (1 expected)")
    a_block = w["blocks"][0] if w["blocks"] else None
    if a_block not in p["blocks"]:
        problems.append("A's block is not on the page, unchanged, after B's brain_links and undos")
    undos = p.get("undos") or []
    if len(undos) != 2 or not (undos[0] or {}).get("ok") or (undos[1] or {}).get("category") != "NothingToUndo":
        problems.append(f"B's undos were {undos}; expected B's own block undone, then NothingToUndo")
    if len(p["blocks"]) != 1:
        problems.append(f"after B's undos the page has {len(p['blocks'])} blocks (A's only expected)")
    a_rows_after_probe = {r[0]: r[1] for r in p.get("ledger", {}).get("rows", [])}
    if any(op in a_rows_after_probe for op in a_ops):
        problems.append("B's ledger names one of A's writes")
    if c["blocks"]:
        problems.append(f"after A's cleanup undo the page still has {len(c['blocks'])} block(s)")
    root_db = root_db or ROOT_DB
    root_after = table_counts(root_db) if root_db.exists() else {}
    if root_after != root_before:
        problems.append("the shared default-profile brain.db changed")
    return {"pass": not problems, "problems": problems, "aDb": str(a_db), "bDb": str(b_db), "aRows": len(a_ops),
            "bRows": len(b_ops), "rootUnchanged": root_after == root_before, "rootCounts": root_after}


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("phase", choices=("root-counts", "write", "probe", "cleanup", "verify"))
    ap.add_argument("--out", required=True)
    ap.add_argument("--control", default="http://127.0.0.1:19334")
    ap.add_argument("--mac-user", default=os.environ.get("HOL_MAC_USER", "demo-user"))
    a = ap.parse_args()
    out = Path(a.out)
    out.mkdir(parents=True, exist_ok=True)
    if a.phase == "root-counts":
        (out / "root-before.json").write_text(json.dumps(table_counts(ROOT_DB) if ROOT_DB.exists() else {}))
        return 0
    if a.phase == "verify":
        v = verify(out, json.loads((out / "root-before.json").read_text()))
        (out / "verify.json").write_text(json.dumps(v, indent=2))
        print(json.dumps(v, indent=2))
        return 0 if v["pass"] else 1
    try:
        res = asyncio.run(phase(a))
    except Exception as exc:  # recorded; verify then reports the missing phase
        res = {"phase": a.phase, "profile": profile_home().name, "error": f"{type(exc).__name__}: {str(exc)[:300]}",
               "where": d.error_where(exc)}
        (out / f"{a.phase}.json").write_text(json.dumps(d.mask_addresses(res), indent=2))
        print(json.dumps(res), flush=True)
        return 1
    (out / f"{a.phase}.json").write_text(json.dumps(d.mask_addresses(res), indent=2, ensure_ascii=False))  # A20 names
    print(json.dumps({k: v for k, v in res.items() if k not in ("ledger", "reply", "identity")}, ensure_ascii=False), flush=True)
    return 0


if __name__ == "__main__":
    sys.exit(main())
