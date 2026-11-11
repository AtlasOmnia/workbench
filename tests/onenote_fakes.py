# SPDX-License-Identifier: Apache-2.0
"""Python stand-ins for paired panes (tests only).

FakeOneNotePane mirrors the contract of addin/onenote/executor.js (active-page-only reads, opId
journal, notebook re-check, owned-outline deletes gated on the write-time fingerprint) closely enough to test the plugin side; the
real executor is covered by tests/js/onenote.test.mjs and the Electron harness.
"""
from __future__ import annotations

import html as _html
import re
import sys
import threading
import uuid


def strip(h: str) -> str:
    t = re.sub(r"<br\s*/?>|</(p|li|h[1-6]|tr|div)>", "\n", h, flags=re.I)
    return _html.unescape(re.sub(r"<[^>]+>", "", t)).strip()


class PaneThread:
    def __init__(self, registry, binding, handle):
        self.registry, self.binding, self.handle = registry, binding, handle
        self.stop = threading.Event()
        self.seen: list[tuple[str, dict]] = []
        self.drop_result_for: set[str] = set()
        self.t = threading.Thread(target=self._run, daemon=True)
        self.t.start()

    def _run(self):
        while not self.stop.is_set():
            try:
                cmd = self.registry.next_command(self.binding, 0.2)
            except Exception:  # noqa: BLE001 - binding closed
                return
            if cmd is None:
                continue
            self.seen.append((cmd.tool, dict(cmd.arguments)))
            result = self.handle(cmd.tool, cmd.arguments)
            if cmd.tool in self.drop_result_for:
                self.drop_result_for.discard(cmd.tool)
                continue  # executed, but the result never reaches Hermes
            self.registry.complete(self.binding, cmd.request_id, result)

    def tools(self):
        return [t for t, _ in self.seen]

    def close(self):
        self.stop.set()
        self.t.join(2)


class FakeOneNote:
    MARKER = "Links & Backlinks"

    def __init__(self, notebook_id="nb-north", spec=None):
        self.id = notebook_id
        self.pages: dict[str, dict] = {}
        self.order: list[str] = []
        self.journal: dict[str, dict] = {}
        self.locked = set()
        n = 0
        for section, pages in (spec or {}).items():
            for title, paras in pages:
                n += 1
                pid = f"{notebook_id}-p{n}"
                self.pages[pid] = {"id": pid, "title": title, "section": section,
                                   "outlines": [{"id": f"{pid}-o0", "top": 90, "paras": [{"id": f"{pid}-o0-{i}", "text": t}
                                                                                        for i, t in enumerate(paras)]}]}
                self.order.append(pid)
        self.active = self.order[0] if self.order else None
        self.navigations: list[str] = []
        # Live (A9 L10): OneNote for the web reports a new page as "Untitled Page" for a moment. lag = how many
        # list_pages/get_page reads still show that placeholder for a newly created page (0: none, the default).
        self.lag = 0
        self.title_lag: dict[str, int] = {}
        self.list_calls = 0

    def shown(self, p):
        return "Untitled Page" if self.title_lag.get(p["id"], 0) > 0 else p["title"]

    def url(self, pid):
        # Like OneNote for the web, the address names the page in target(...): the placeholder while it lags.
        lag = "?wd=target(s.one%7C/Untitled%20Page%7Cg/)" if self.title_lag.get(pid, 0) > 0 else ""
        return f"https://onenote.test/{self.id}/{pid}{lag}"

    def by_title(self, title):
        return next(p for p in self.pages.values() if p["title"] == title)

    def info(self, p):
        return {"id": p["id"], "title": self.shown(p), "webUrl": self.url(p["id"]), "clientUrl": "onenote:x", "level": 0,
                "sectionId": "sec-" + p["section"], "sectionPath": p["section"]}

    def read(self, p):
        outlines = [{"id": o["id"], "top": o["top"],
                     "paragraphs": [{"id": x["id"], "type": x.get("type", "RichText"), "text": x["text"],
                                     **({"rows": x["rows"]} if "rows" in x else {})} for x in o["paras"]]}
                    for o in p["outlines"]]
        text = "\n\n".join("\n".join(x["text"] for x in o["paras"]) for o in p["outlines"])
        return {"notebookId": self.id, "sectionId": "sec-" + p["section"], "pageId": p["id"], "title": self.shown(p),
                "webUrl": self.url(p["id"]), "outlines": outlines, "text": text, "truncated": False, "omitted": {}}

    def add_outline(self, p, html):
        oid = "o-" + uuid.uuid4().hex[:8]
        top = max([o["top"] for o in p["outlines"]] or [0]) + 240
        paras = [{"id": oid + f"-{i}", "text": strip(m)} for i, m in enumerate([x[1] for x in re.findall(r"<(p|li|h[1-6])>(.*?)</\1>", html)] or [html])]
        p["outlines"].append({"id": oid, "top": top, "paras": paras})
        return oid

    def fps(self, p) -> dict:
        """The pane's integrity fingerprints (executor.js outlineFingerprint) for this page's outlines/paragraphs."""
        return sys.modules["hermes_office_desktop.brain"].page_fingerprints(self.read(p))

    def created(self, p, oids) -> dict:
        cur = self.fps(p)
        return {"outlineIds": oids, "fingerprints": {o: cur[o] for o in oids if o in cur}}

    def fingerprint(self):
        return repr(sorted((p["title"], [[x["text"] for x in o["paras"]] for o in p["outlines"]]) for p in self.pages.values()))

    def __call__(self, tool, a):
        try:
            return {"ok": True, "data": self._do(tool, dict(a))} if tool != "op_status" else \
                {"ok": True, "data": {"opId": a.get("opId"), **self.journal.get(a.get("opId"), {"state": "unknown"})}}
        except KeyError as exc:
            return {"ok": False, "error": {"category": "NotFound", "message": str(exc)}}
        except ValueError as exc:
            cat, _, code = str(exc).partition(":")
            return {"ok": False, "error": {"category": cat, "code": code, "message": str(exc)}}

    def _do(self, tool, a):
        if a.get("expectNotebookId") and a["expectNotebookId"] != self.id:
            raise ValueError("Conflict:IdentityMismatch")
        mutating = tool in ("create_page", "append", "insert_table", "delete_owned", "replace_owned")
        op = a.get("opId")
        if mutating:
            if not op:
                raise ValueError("Validation:noop")
            if op in self.journal:
                return self.journal[op]["result"]["data"]
            self.journal[op] = {"state": "started"}
        out = self._run(tool, a)
        if mutating:
            self.journal[op] = {"state": "done", "result": {"ok": True, "data": out}}
        return out

    def _run(self, tool, a):
        act = self.pages.get(self.active)
        if tool in ("list_pages", "get_page"):
            self.list_calls += tool == "list_pages"
            out = self._read_tool(tool, a, act)
            self.title_lag = {k: v - 1 for k, v in self.title_lag.items() if v > 1}
            return out
        return self._write_tool(tool, a, act)

    def _read_tool(self, tool, a, act):
        if tool == "list_pages":
            pages = [self.info(self.pages[i]) for i in self.order if self.pages[i]["section"] not in self.locked]
            if a.get("sectionId"):
                pages = [p for p in pages if p["sectionId"] == a["sectionId"]]
            return {"pages": pages, "complete": not self.locked, "skipped": [{"path": s} for s in self.locked]}
        if a.get("pageId") and a["pageId"] != self.active:
            raise ValueError("Conflict:NotActivePage")
        return self.read(act)

    def _write_tool(self, tool, a, act):
        if tool == "get_context":
            return {"notebook": {"id": self.id, "name": "Northwind"}, "section": {"id": "sec-" + act["section"]},
                    "page": {"id": act["id"], "title": act["title"], "webUrl": self.url(act["id"])} if act else None}
        if tool == "open_page":
            if a["pageId"] not in self.pages:
                raise KeyError("no page")
            prev, self.active = self.active, a["pageId"]
            self.navigations.append(a["pageId"])
            return {**self.info(self.pages[a["pageId"]]), "previousPageId": prev}
        if tool == "create_page":
            here = {p["title"].strip().lower(): p for p in self.pages.values() if p["section"] == act["section"]}
            dup, title, adjusted = here.get(a["title"].strip().lower()), a["title"], None
            if dup and not a.get("allowDuplicateTitle"):
                if a.get("onDuplicateTitle") != "next":
                    raise ValueError("Conflict:DuplicateTitle")
                title = next((f"{a['title']} ({i})" for i in range(2, 21) if f"{a['title']} ({i})".lower() not in here), None)
                if title is None:
                    raise ValueError("Conflict:DuplicateTitle")
                adjusted = {"requested": a["title"], "used": title, "existingPageId": dup["id"], "existingTitle": dup["title"]}
            a = {**a, "title": title}
            pid = f"{self.id}-new-{len(self.order) + 1}"
            self.pages[pid] = {"id": pid, "title": a["title"], "section": act["section"], "outlines": []}
            self.order.append(pid)
            self.active = pid
            oid = self.add_outline(self.pages[pid], a["html"]) if a.get("html") else None
            if self.lag:
                self.title_lag[pid] = self.lag
            return {**self.info(self.pages[pid]), "title": a["title"], "outlineId": oid,
                    **({"titleAdjusted": adjusted} if adjusted else {}),
                    "created": {"pageId": pid, **self.created(self.pages[pid], [oid] if oid else [])}}
        if tool in ("append", "insert_table"):
            if a.get("pageId") and a["pageId"] != self.active:
                raise ValueError("Conflict:NotActivePage")
            if tool == "insert_table":  # like executor.js: a bold heading paragraph, then the table
                oid = self.add_outline(act, "<p>" + _html.escape(a.get("title") or "Table") + "</p>")
                rows = [[str(v) for v in r] for r in a["rows"]]
                act["outlines"][-1]["paras"].append({"id": oid + "-t", "type": "Table", "rows": rows,
                                                     "text": "\n".join("\t".join(r) for r in rows)})
                return {"pageId": act["id"], "outlineId": oid, "created": self.created(act, [oid])}
            oid = self.add_outline(act, a["html"])
            return {"pageId": act["id"], "outlineId": oid, "created": self.created(act, [oid])}
        if tool == "delete_owned":
            if a["pageId"] != self.active:
                raise ValueError("Conflict:NotActivePage")
            want = set(a.get("outlineIds") or [])
            have = {o["id"] for o in act["outlines"]}
            cur, expect = self.fps(act), a.get("expect") or {}
            kept = sorted(i for i in want & have if expect.get(i) is None or expect.get(i) != cur.get(i))
            gone = (want & have) - set(kept)
            act["outlines"] = [o for o in act["outlines"] if o["id"] not in gone]
            return {"pageId": act["id"], "deleted": sorted(gone), "missing": sorted(want - have), "notOwned": kept}
        if tool == "replace_owned":
            if a["pageId"] != self.active:
                raise ValueError("Conflict:NotActivePage")
            deleted, not_owned = [], []
            cur, expect = self.fps(act), a.get("expect") or {}
            for oid in a.get("deleteOutlineIds") or []:
                o = next((x for x in act["outlines"] if x["id"] == oid), None)
                if not o:
                    continue
                if not (o["paras"] and o["paras"][0]["text"].startswith(self.MARKER)) or expect.get(oid) != cur.get(oid):
                    not_owned.append(oid)
                    continue
                act["outlines"].remove(o)
                deleted.append(oid)
            oid = self.add_outline(act, a["html"])
            return {"pageId": act["id"], "outlineId": oid, "deleted": deleted, "notOwned": not_owned,
                    "created": self.created(act, [oid])}
        raise ValueError(f"Unsupported:{tool}")

    def links_block(self, pid):
        return next(("\n".join(x["text"] for x in o["paras"]) for o in self.pages[pid]["outlines"]
                     if o["paras"] and o["paras"][0]["text"].startswith(self.MARKER)), None)


class FakeHost:
    """Excel / Word / PowerPoint panes for deliverables: records commands, answers like the executors."""

    def __init__(self, app):
        self.app, self.calls, self.cells, self.slides, self.text = app, [], {}, [], ""

    def __call__(self, tool, a):
        self.calls.append((tool, a))
        if self.app == "excel":
            if tool == "get_workbook_overview":
                return {"ok": True, "data": {"activeSheet": "Sheet1", "sheets": [{"name": "Sheet1"}, {"name": "Decision tracker"}]}}
            if tool == "write_range":
                self.cells[a["address"]] = a["values"]
                return {"ok": True, "data": {"address": a["address"], "undo": {"available": True}}}
            if tool == "read_range":
                rows = next(v for k, v in self.cells.items() if k.startswith("A1:"))
                n = sum(1 for r in rows[1:] if r[4] == "Open")  # =COUNTIF(E2:E<n>,"Open")
                return {"ok": True, "data": {"values": [["Open items", n]]}}
            return {"ok": True, "data": {"undo": {"available": True}}}
        if self.app == "word":
            if tool == "insert_ooxml":
                self.text += " ".join(re.findall(r"<w:t[^>]*>([^<]*)</w:t>", a["ooxml"]))
                return {"ok": True, "data": {"inserted": True}}
            return {"ok": True, "data": {"text": _html.unescape(self.text)}}
        if tool == "create_text_slide":
            self.slides.append([a["title"], *(a.get("bullets") or []), a.get("body") or ""])
            return {"ok": True, "data": {"slideIndex": len(self.slides) - 1, "slideId": f"s{len(self.slides)}"}}
        return {"ok": True, "data": {"slideCount": len(self.slides), "truncated": False,
                                     "slides": [{"index": i, "id": f"s{i + 1}", "text": s} for i, s in enumerate(self.slides)]}}
