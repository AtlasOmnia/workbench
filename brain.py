# SPDX-License-Identifier: Apache-2.0
"""Local second-brain index for the OneNote host: pages, links, tags, write ledger, events.

Storage is one SQLite file from Hermes' documented ``plugins.plugin_storage.plugin_db`` (that is
``<Hermes home>/plugin-data/hermes-office-desktop/brain.db`` for the profile that OWNS this plugin
instance, resolved per call by profile_scope.py, also on bridge threads), never Hermes memory and never
the notebook itself. Everything is partitioned by the OneNote notebook id.

Honest coverage model (OneNote JS API): titles/ids/urls of every page are cheap (the "directory");
page CONTENT is known only for pages the user opened while paired ("visit-indexing"), pages Hermes
wrote, or an explicit, approved crawl. Every answer reports "indexed N of M pages".

Page text is the user's data. It is stored and returned as quoted data and never interpreted as an
instruction by this module.
"""
from __future__ import annotations

import hashlib
import html as _html
import json
import re
import sqlite3
import threading
import time
import unicodedata
from typing import Any, Callable, Iterable, Optional

PLUGIN_ID = "hermes-office-desktop"
DB_FILE = "brain.db"
MARKER = "Links & Backlinks"  # must match addin/onenote/executor.js MARKER
MAX_TEXT = 50000
WIKILINK = re.compile(r"\[\[([^\[\]|\n]{1,200}?)(?:\|([^\[\]\n]{1,200}?))?\]\]")
HASHTAG = re.compile(r"(?<![\w&#/])#([A-Za-z][\w-]{0,48})")
OWNED_KINDS = ("links_outline", "append", "table", "page_body")

SCHEMA = """
create table if not exists meta(key text primary key, value text);
create table if not exists pages(
  notebook_id text not null, page_id text not null, title text not null, title_norm text not null,
  section_id text, section_path text, web_url text, client_url text, level integer default 0,
  discovered_at real, captured_at real, opened_at real, content_hash text, text text,
  deleted integer default 0, hermes_created integer default 0,
  primary key(notebook_id, page_id));
create index if not exists pages_title on pages(notebook_id, title_norm);
create table if not exists links(
  notebook_id text not null, src_page_id text not null, dst_title text not null, dst_norm text not null,
  alias text, kind text not null, dst_page_id text);
create index if not exists links_src on links(notebook_id, src_page_id);
create index if not exists links_dst on links(notebook_id, dst_norm);
create table if not exists tags(notebook_id text not null, page_id text not null, tag text not null);
create table if not exists writes(
  id integer primary key autoincrement, notebook_id text not null, page_id text, op_id text unique,
  tool text not null, kind text not null, status text not null, digest text, summary text,
  created json, extra json, created_at real, updated_at real);
create index if not exists writes_nb on writes(notebook_id, status);
create table if not exists events(notebook_id text not null, page_id text, kind text not null, ts real not null);
create table if not exists decisions(
  id integer primary key autoincrement, notebook_id text not null, decision_page_id text not null,
  source_page_id text, kind text not null, text text not null, owner text, due text, rationale text,
  quote text, quote_verified integer default 0, created_at real);
create table if not exists artifacts(
  id integer primary key autoincrement, notebook_id text not null, decision_page_id text, host text not null,
  document text, locator text, decision_ids json, summary text, created_at real);
"""


def norm(title: str) -> str:
    """Title key for resolution: NFKC, case-folded, whitespace collapsed."""
    return " ".join(unicodedata.normalize("NFKC", str(title)).casefold().split())


def esc(s: Any) -> str:
    return _html.escape(str(s), quote=True)


# What OneNote for the web reports for a page whose title it has not applied yet (seen live right after a page is
# created, A9 L10). A page Hermes created keeps the title it asked for instead.
PLACEHOLDER_TITLES = frozenset({"", "untitled page"})


def placeholder_title(title: Any) -> bool:
    return not isinstance(title, str) or title.strip().casefold() in PLACEHOLDER_TITLES


def safe_url(url: Any) -> Optional[str]:
    """Only https URLs reported by OneNote itself go into an href."""
    if isinstance(url, str) and url.startswith("https://") and len(url) < 2000 and not re.search(r"[\s\"'<>]", url):
        return url
    return None


def parse_links(text: str) -> list[tuple[str, Optional[str]]]:
    out = []
    for m in WIKILINK.finditer(text or ""):
        target = m.group(1).strip()
        alias = (m.group(2) or "").strip() or None
        if target:
            out.append((target, alias))
    return out


def parse_tags(text: str) -> list[str]:
    return sorted({t.lower() for t in HASHTAG.findall(text or "")})


def digest(*parts: Any) -> str:
    return hashlib.sha256(json.dumps(parts, sort_keys=True, ensure_ascii=False, default=str).encode()).hexdigest()


# Integrity fingerprint of a Hermes-written outline (A6 #6): sha256 of its canonical text, per paragraph
# "type<US>text" (NFC, whitespace collapsed) joined by <RS>. addin/onenote/executor.js outlineFingerprint
# computes the same value; the ledger stores it at write time and a delete/replace needs a match.
_FP_WS = re.compile("[\t\n\v\f\r \u00a0\u1680\u2000-\u200a\u2028\u2029\u202f\u205f\u3000\ufeff]+")


def outline_fingerprint(paragraphs: Iterable[dict]) -> str:
    parts = []
    for p in paragraphs or []:
        p = p if isinstance(p, dict) else {}
        text = unicodedata.normalize("NFC", str(p.get("text") or ""))
        parts.append(f"{p.get('type') or 'RichText'}\x1f" + " ".join(x for x in _FP_WS.split(text) if x))
    return hashlib.sha256("\x1e".join(parts).encode()).hexdigest()


def _level(p: Any) -> int:
    return p.get("level") or 0 if isinstance(p, dict) else 0


def page_fingerprints(page: dict) -> dict:
    """id -> fingerprint for every outline of a get_page read, and for each paragraph with the paragraphs under it
    (the rows after it with a deeper `level`, F8: deleting a paragraph deletes them too)."""
    out = {}
    for o in (page or {}).get("outlines") or []:
        paras = o.get("paragraphs") or []
        if o.get("id"):
            out[o["id"]] = outline_fingerprint(paras)
        for i, p in enumerate(paras):
            if isinstance(p, dict) and p.get("id"):
                j = i + 1
                while j < len(paras) and _level(paras[j]) > _level(p):
                    j += 1
                out[p["id"]] = outline_fingerprint(paras[i:j])
    return out


def _open_default() -> sqlite3.Connection:
    from plugins.plugin_storage import plugin_db  # documented public plugin surface (catalog rule 9)
    return plugin_db(PLUGIN_ID, DB_FILE)


def db_file(conn: sqlite3.Connection) -> Optional[str]:
    """The main database file a connection has open (``pragma database_list``)."""
    for row in conn.execute("pragma database_list").fetchall():
        if row[1] == "main":
            return row[2] or None
    return None


def open_for_home(home) -> sqlite3.Connection:
    """Open ``<home>/plugin-data/hermes-office-desktop/brain.db`` for the OWNING profile's home.

    ``plugin_db`` resolves the Hermes home of the calling context, so it runs under that home bound
    explicitly (a bridge thread would otherwise resolve the gateway's launch profile). The opened file is
    checked to be inside that home before anything is read or written."""
    from pathlib import Path
    from .bridge_server import BridgeError
    from .profile_scope import CATEGORY, scoped
    want = (Path(str(home)).expanduser().resolve(strict=False) / "plugin-data" / PLUGIN_ID / DB_FILE)
    with scoped(home):
        conn = _open_default()
    got = db_file(conn)
    if not got or Path(got).resolve(strict=False) != want:
        conn.close()
        raise BridgeError(CATEGORY, "The local index opened outside this profile's Hermes home; it was not used.")
    return conn


class Brain:
    """Thread-safe wrapper over one profile's brain.db."""

    def __init__(self, connect: Callable[[], sqlite3.Connection] = _open_default, use_fts: Optional[bool] = None,
                 clock: Callable[[], float] = time.time):
        self._conn = connect()
        self._lock = threading.RLock()
        self._clock = clock
        with self._lock:
            self._conn.executescript(SCHEMA)
            # A30: a held (uncertain) write is released only by a reconciliation that a successful read of the
            # affected page (or listing, for a new page) wrote; older brain.db files gain the columns here.
            cols = {r[1] for r in self._conn.execute("pragma table_info(writes)")}
            for col, typ in (("reconciled", "text"), ("reconciled_at", "real"), ("reconciled_by", "text")):
                if col not in cols:
                    self._conn.execute(f"alter table writes add column {col} {typ}")
            self.fts = False
            if use_fts is not False:
                try:
                    self._conn.execute("create virtual table if not exists pages_fts using fts5("
                                       "notebook_id unindexed, page_id unindexed, title, text, tokenize='unicode61')")
                    self.fts = True
                except sqlite3.OperationalError:
                    self.fts = False  # this SQLite has no FTS5: LIKE search below
            self._conn.commit()

    # ---- helpers ---------------------------------------------------------------------------
    def _q(self, sql: str, args: Iterable = ()) -> list[dict]:
        cur = self._conn.execute(sql, tuple(args))
        cols = [c[0] for c in cur.description] if cur.description else []
        return [dict(zip(cols, r)) for r in cur.fetchall()]

    def _one(self, sql: str, args: Iterable = ()) -> Optional[dict]:
        rows = self._q(sql, args)
        return rows[0] if rows else None

    def close(self) -> None:
        with self._lock:
            self._conn.close()

    def set_meta(self, key: str, value: str) -> None:
        with self._lock:
            self._conn.execute("insert into meta(key,value) values(?,?) on conflict(key) do update set value=excluded.value",
                               (key, value))
            self._conn.commit()

    def get_meta(self, key: str) -> Optional[str]:
        with self._lock:
            row = self._one("select value from meta where key=?", (key,))
            return row["value"] if row else None

    def event(self, nb: str, page_id: Optional[str], kind: str) -> None:
        self._conn.execute("insert into events(notebook_id,page_id,kind,ts) values(?,?,?,?)", (nb, page_id, kind, self._clock()))

    def _known_title(self, nb: str, pid: str, title: str) -> str:
        """OneNote's placeholder title never replaces the real title of a page Hermes created."""
        if not placeholder_title(title):
            return title
        with self._lock:
            row = self._one("select title, hermes_created from pages where notebook_id=? and page_id=?", (nb, pid))
        if row and row["hermes_created"] and not placeholder_title(row["title"]):
            return row["title"]
        return title

    # ---- directory (cheap whole-notebook structure sweep) -------------------------------------
    def upsert_directory(self, nb: str, pages: list[dict], complete: bool) -> dict:
        now = self._clock()
        with self._lock:
            seen = set()
            for p in pages:
                pid, title = p.get("id"), p.get("title")
                if not isinstance(pid, str) or not isinstance(title, str):
                    continue
                seen.add(pid)
                title = self._known_title(nb, pid, title)
                self._conn.execute(
                    "insert into pages(notebook_id,page_id,title,title_norm,section_id,section_path,web_url,client_url,level,"
                    "discovered_at,deleted) values(?,?,?,?,?,?,?,?,?,?,0) on conflict(notebook_id,page_id) do update set "
                    "title=excluded.title,title_norm=excluded.title_norm,section_id=excluded.section_id,"
                    "section_path=excluded.section_path,web_url=excluded.web_url,client_url=excluded.client_url,"
                    "level=excluded.level,deleted=0",
                    (nb, pid, title, norm(title), p.get("sectionId"), p.get("sectionPath"), safe_url(p.get("webUrl")),
                     p.get("clientUrl"), int(p.get("level") or 0), now))
            removed = 0
            if complete:  # only a COMPLETE sweep may mark pages gone (never after an error)
                # Emails and calendar events indexed by brain_from_email (page ids "email:<key>", "event:<key>") are not
                # OneNote pages: a notebook sweep never sees them, so it must not mark them gone.
                for row in self._q("select page_id from pages where notebook_id=? and deleted=0 and page_id not like 'email:%' "
                                   "and page_id not like 'event:%'", (nb,)):
                    if row["page_id"] not in seen:
                        self._conn.execute("update pages set deleted=1 where notebook_id=? and page_id=?", (nb, row["page_id"]))
                        removed += 1
            self.event(nb, None, "sweep")
            self._conn.commit()
            return {"pages": len(seen), "markedDeleted": removed, "complete": bool(complete)}

    # ---- capture (visit-indexing, crawl, Hermes writes) --------------------------------------
    def owned_outline_ids(self, nb: str, page_id: str) -> set[str]:
        with self._lock:
            out: set[str] = set()
            for w in self._q("select created from writes where notebook_id=? and page_id=? and status='done' and kind in "
                             f"({','.join('?' * len(OWNED_KINDS))})", (nb, page_id, *OWNED_KINDS)):
                created = json.loads(w["created"] or "{}")
                out.update(created.get("outlineIds") or [])
            return out

    def capture(self, nb: str, page: dict, source: str = "visit") -> dict:
        """Index one page snapshot read from the ACTIVE page. Replaces that page's outgoing edges."""
        pid, title = page.get("pageId"), page.get("title")
        if not isinstance(pid, str) or not isinstance(title, str):
            raise ValueError("capture needs pageId and title")
        if page.get("notebookId") and page["notebookId"] != nb:
            raise ValueError("IdentityMismatch")
        title = self._known_title(nb, pid, title)
        owned = self.owned_outline_ids(nb, pid)
        parts = []
        for o in page.get("outlines") or []:
            if o.get("id") in owned:
                continue  # Hermes' own outlines are not the user's notes (no self-links/backlink loops)
            paras = o.get("paragraphs") or []  # a marker-looking outline we did not write stays user content
            parts.append("\n".join(str(p.get("text") or "") for p in paras))
        text = "\n\n".join(parts)[:MAX_TEXT]
        h = hashlib.sha256(text.encode()).hexdigest()
        now = self._clock()
        with self._lock:
            self._conn.execute(
                "insert into pages(notebook_id,page_id,title,title_norm,section_id,web_url,discovered_at,captured_at,"
                "content_hash,text,deleted) values(?,?,?,?,?,?,?,?,?,?,0) on conflict(notebook_id,page_id) do update set "
                "title=excluded.title,title_norm=excluded.title_norm,web_url=coalesce(excluded.web_url,pages.web_url),"
                "section_id=coalesce(excluded.section_id,pages.section_id),captured_at=excluded.captured_at,"
                "content_hash=excluded.content_hash,text=excluded.text,deleted=0",
                (nb, pid, title, norm(title), page.get("sectionId"), safe_url(page.get("webUrl")), now, now, h, text))
            if source == "visit":
                self._conn.execute("update pages set opened_at=? where notebook_id=? and page_id=?", (now, nb, pid))
            # Recapture REPLACES this page's user-authored edges (hermes edges survive).
            self._conn.execute("delete from links where notebook_id=? and src_page_id=? and kind='wikilink'", (nb, pid))
            links = parse_links(text)
            for target, alias in links:
                self._conn.execute("insert into links(notebook_id,src_page_id,dst_title,dst_norm,alias,kind) values(?,?,?,?,?,?)",
                                   (nb, pid, target, norm(target), alias, "wikilink"))
            self._conn.execute("delete from tags where notebook_id=? and page_id=?", (nb, pid))
            tags = parse_tags(text)
            for t in tags:
                self._conn.execute("insert into tags(notebook_id,page_id,tag) values(?,?,?)", (nb, pid, t))
            if self.fts:
                self._conn.execute("delete from pages_fts where notebook_id=? and page_id=?", (nb, pid))
                self._conn.execute("insert into pages_fts(notebook_id,page_id,title,text) values(?,?,?,?)", (nb, pid, title, text))
            self.event(nb, pid, "capture:" + source)
            self._conn.commit()
        return {"pageId": pid, "links": len(links), "tags": tags, "chars": len(text)}

    def add_hermes_edge(self, nb: str, src: str, dst_page_id: str, dst_title: str) -> None:
        with self._lock:
            self._conn.execute("insert into links(notebook_id,src_page_id,dst_title,dst_norm,kind,dst_page_id) values(?,?,?,?,?,?)",
                               (nb, src, dst_title, norm(dst_title), "hermes", dst_page_id))
            self._conn.commit()

    def page(self, nb: str, page_id: str) -> Optional[dict]:
        with self._lock:
            return self._one("select * from pages where notebook_id=? and page_id=?", (nb, page_id))

    def coverage(self, nb: str) -> dict:
        with self._lock:
            row = self._one("select count(*) m, sum(case when captured_at is not null then 1 else 0 end) n, "
                            "min(captured_at) oldest from pages where notebook_id=? and deleted=0", (nb,))
            m, n = int(row["m"] or 0), int(row["n"] or 0)
            oldest = row["oldest"]
            out = {"indexed": n, "discovered": m, "label": f"indexed {n} of {m} pages"}
            if oldest:
                age_h = max(0.0, (self._clock() - oldest) / 3600)
                out["oldestCaptureHours"] = round(age_h, 1)
            out["fts"] = self.fts
            return out

    # ---- resolution / links -------------------------------------------------------------------
    def resolve(self, nb: str, target: str) -> dict:
        """[[Title]] -> exact title, then normalized title; duplicates are ambiguous; Section/Title qualifies."""
        with self._lock:
            live = "notebook_id=? and deleted=0"
            exact = self._q(f"select * from pages where {live} and title=?", (nb, target))
            if len(exact) == 1:
                return {"status": "resolved", "page": exact[0]}
            cands = exact or self._q(f"select * from pages where {live} and title_norm=?", (nb, norm(target)))
            if len(cands) == 1:
                return {"status": "resolved", "page": cands[0]}
            if len(cands) > 1:
                return {"status": "ambiguous", "candidates": cands}
            if "/" in target:  # "Section/Title" or "Group / Section/Title"
                sec, _, title = target.rpartition("/")
                rows = [r for r in self._q(f"select * from pages where {live} and title_norm=?", (nb, norm(title)))
                        if norm((r.get("section_path") or "")).endswith(norm(sec))]
                if len(rows) == 1:
                    return {"status": "resolved", "page": rows[0]}
                if len(rows) > 1:
                    return {"status": "ambiguous", "candidates": rows}
            return {"status": "unresolved"}

    def links_for(self, nb: str, page_id: str) -> dict:
        with self._lock:
            me = self.page(nb, page_id)
            if not me:
                raise KeyError(page_id)
            outbound, unresolved, ambiguous, seen = [], [], [], set()
            for e in self._q("select * from links where notebook_id=? and src_page_id=? order by rowid", (nb, page_id)):
                if e["kind"] == "hermes" and e["dst_page_id"]:
                    tgt = self.page(nb, e["dst_page_id"])
                    r = {"status": "resolved", "page": tgt} if tgt and not tgt["deleted"] else {"status": "unresolved"}
                else:
                    r = self.resolve(nb, e["dst_title"])
                key = (e["dst_norm"], e["alias"])
                if key in seen:
                    continue
                seen.add(key)
                if r["status"] == "resolved":
                    if r["page"]["page_id"] != page_id:
                        outbound.append({"pageId": r["page"]["page_id"], "title": r["page"]["title"],
                                         "label": e["alias"] or r["page"]["title"], "webUrl": r["page"]["web_url"],
                                         "kind": e["kind"]})
                elif r["status"] == "ambiguous":
                    ambiguous.append({"target": e["dst_title"], "candidates": [
                        {"pageId": c["page_id"], "title": c["title"], "section": c["section_path"]} for c in r["candidates"]]})
                else:
                    unresolved.append({"target": e["dst_title"], "label": e["alias"]})
            backlinks, bseen = [], set()
            for e in self._q("select l.*, p.title src_title, p.web_url src_url, p.deleted src_deleted from links l "
                             "join pages p on p.notebook_id=l.notebook_id and p.page_id=l.src_page_id "
                             "where l.notebook_id=? and (l.dst_norm=? or l.dst_page_id=?) and l.src_page_id<>?",
                             (nb, me["title_norm"], page_id, page_id)):
                if e["src_deleted"] or e["src_page_id"] in bseen:
                    continue
                if e["kind"] == "wikilink":
                    r = self.resolve(nb, e["dst_title"])
                    if r["status"] != "resolved" or r["page"]["page_id"] != page_id:
                        continue  # points at a same-titled other page, or is ambiguous
                bseen.add(e["src_page_id"])
                backlinks.append({"pageId": e["src_page_id"], "title": e["src_title"], "webUrl": e["src_url"]})
            tags = [r["tag"] for r in self._q("select tag from tags where notebook_id=? and page_id=? order by tag", (nb, page_id))]
            return {"page": {"pageId": page_id, "title": me["title"]}, "links": outbound, "unresolved": unresolved,
                    "ambiguous": ambiguous, "backlinks": backlinks, "tags": tags}

    @staticmethod
    def render_links_outline(info: dict, coverage: dict) -> str:
        def a(item):
            url = safe_url(item.get("webUrl"))
            label = esc(item.get("label") or item.get("title"))
            return f'<a href="{esc(url)}">{label}</a>' if url else label
        parts = [f"<p><b>{esc(MARKER)}</b> <i>(kept by Hermes; regenerated on request; your text is never changed)</i></p>"]
        parts.append("<p>Links: " + (", ".join(a(x) for x in info["links"]) or "none") + "</p>")
        if info["unresolved"]:
            parts.append("<p>Not created yet: " + ", ".join(esc(x["target"]) for x in info["unresolved"]) + "</p>")
        if info["ambiguous"]:
            parts.append("<p>Ambiguous (several pages share this title): "
                         + ", ".join(esc(x["target"]) for x in info["ambiguous"]) + "</p>")
        parts.append("<p>Backlinks: " + (", ".join(a(x) for x in info["backlinks"]) or "none")
                     + f" <i>(within {esc(coverage['label'])})</i></p>")
        if info["tags"]:
            parts.append("<p>Tags: " + " ".join("#" + esc(t) for t in info["tags"]) + "</p>")
        return "".join(parts)

    # ---- search (notebook Q&A) ---------------------------------------------------------------
    def search(self, nb: str, question: str, limit: int = 5) -> list[dict]:
        words = [w for w in re.findall(r"[\w]+", unicodedata.normalize("NFKC", question).casefold()) if len(w) > 2][:12]
        words = [w for w in words if w not in _STOP] or words
        if not words:
            return []
        with self._lock:
            rows: list[dict] = []
            if self.fts:
                q = " OR ".join('"' + w.replace('"', "") + '"' for w in words)
                try:
                    rows = self._q("select f.page_id, snippet(pages_fts, 3, '[', ']', ' … ', 18) excerpt, bm25(pages_fts) score "
                                   "from pages_fts f join pages p on p.notebook_id=f.notebook_id and p.page_id=f.page_id "
                                   "where pages_fts match ? and f.notebook_id=? and p.deleted=0 order by score limit ?",
                                   (q, nb, limit * 3))
                except sqlite3.OperationalError:
                    rows = []
                # The user's own #tags are explicit topic labels: a tag that matches the question ranks first.
                tagged = {r["page_id"] for r in self._q(
                    f"select distinct page_id from tags where notebook_id=? and tag in ({','.join('?' * len(words))})", (nb, *words))}
                rows = sorted(rows, key=lambda r: r["score"] - (5.0 if r["page_id"] in tagged else 0.0))[:limit]
            if not rows:
                tagged = {r["page_id"] for r in self._q(
                    f"select distinct page_id from tags where notebook_id=? and tag in ({','.join('?' * len(words))})", (nb, *words))}
                scored = []
                for p in self._q("select page_id, title, text from pages where notebook_id=? and deleted=0 and text is not null", (nb,)):
                    hay = (p["title"] + "\n" + (p["text"] or "")).casefold()
                    s = sum(hay.count(w) for w in words)
                    if s and p["page_id"] in tagged:
                        s += 5
                    if s:
                        i = min((hay.find(w) for w in words if w in hay), default=0)
                        txt = p["title"] + "\n" + (p["text"] or "")
                        scored.append({"page_id": p["page_id"], "score": -s,
                                       "excerpt": ("… " if i > 60 else "") + txt[max(0, i - 60):i + 160].replace("\n", " ") + " …"})
                rows = sorted(scored, key=lambda r: r["score"])[:limit]
            out = []
            for r in rows:
                p = self.page(nb, r["page_id"])
                out.append({"pageId": p["page_id"], "title": p["title"], "section": p["section_path"], "webUrl": p["web_url"],
                            "excerpt": r["excerpt"], "capturedAtHoursAgo": round((self._clock() - (p["captured_at"] or 0)) / 3600, 1)
                            if p["captured_at"] else None,
                            "open": {"tool": "onenote_open_page", "args": {"pageId": p["page_id"]}}})
            return out

    # ---- ledger ---------------------------------------------------------------------------------
    def begin_write(self, nb: str, page_id: Optional[str], op_id: str, tool: str, kind: str, dig: str, summary: str,
                    extra: Optional[dict] = None) -> int:
        now = self._clock()
        with self._lock:
            cur = self._conn.execute("insert into writes(notebook_id,page_id,op_id,tool,kind,status,digest,summary,created,extra,"
                                     "created_at,updated_at) values(?,?,?,?,?,?,?,?,?,?,?,?)",
                                     (nb, page_id, op_id, tool, kind, "pending", dig, summary[:300], "{}",
                                      json.dumps(extra or {}), now, now))
            self._conn.commit()
            return int(cur.lastrowid)

    def finish_write(self, op_id: str, status: str, page_id: Optional[str] = None, created: Optional[dict] = None,
                     extra: Optional[dict] = None) -> None:
        with self._lock:
            row = self._one("select * from writes where op_id=?", (op_id,))
            if not row:
                return
            ex = json.loads(row["extra"] or "{}")
            ex.update(extra or {})
            self._conn.execute("update writes set status=?, page_id=coalesce(?,page_id), created=?, extra=?, updated_at=? where op_id=?",
                               (status, page_id, json.dumps(created if created is not None else json.loads(row["created"] or "{}")),
                                json.dumps(ex), self._clock(), op_id))
            if status == "done":
                self.event(row["notebook_id"], page_id or row["page_id"], "write:" + row["kind"])
            self._conn.commit()

    def write(self, op_id: str) -> Optional[dict]:
        with self._lock:
            return self._decode(self._one("select * from writes where op_id=?", (op_id,)))

    @staticmethod
    def _decode(row: Optional[dict]) -> Optional[dict]:
        if row:
            row = {**row, "created": json.loads(row["created"] or "{}"), "extra": json.loads(row["extra"] or "{}")}
        return row

    def open_uncertain(self, nb: str, tool: str, dig: str) -> Optional[dict]:
        """The newest held identical write. 'checked' is what F6 left after a blocked retry: still held (A30)."""
        with self._lock:
            return self._decode(self._one("select * from writes where notebook_id=? and tool=? and digest=? and status in "
                                          "('uncertain','checked') order by id desc limit 1", (nb, tool, dig)))

    def held_writes(self, nb: str) -> list[dict]:
        with self._lock:
            return [self._decode(r) for r in self._q("select * from writes where notebook_id=? and status in "
                                                     "('uncertain','checked') order by id", (nb,))]

    def reconcile_write(self, op_id: str, outcome: str, by: str, page_id: Optional[str] = None) -> None:
        """Record what a successful read showed for a held write: 'landed' or 'absent' (A30). Only read paths call
        this; a later read overwrites an earlier one (the newest evidence wins). Never changes the status."""
        with self._lock:
            self._conn.execute("update writes set reconciled=?, reconciled_at=?, reconciled_by=?, page_id=coalesce(page_id,?), "
                               "updated_at=? where op_id=? and status in ('uncertain','checked')",
                               (outcome, self._clock(), by[:200], page_id, self._clock(), op_id))
            self._conn.commit()

    def annotate_write(self, op_id: str, extra: dict) -> None:
        """Add to a ledger row's extra without changing its status."""
        with self._lock:
            row = self._one("select extra from writes where op_id=?", (op_id,))
            if not row:
                return
            ex = json.loads(row["extra"] or "{}")
            ex.update(extra)
            self._conn.execute("update writes set extra=?, updated_at=? where op_id=?", (json.dumps(ex), self._clock(), op_id))
            self._conn.commit()

    def page_ids_titled(self, nb: str, norms: set) -> list[str]:
        """Ids of indexed pages whose title is one of these norms or still a placeholder."""
        with self._lock:
            return [r["page_id"] for r in self._q("select page_id, title, title_norm from pages where notebook_id=?", (nb,))
                    if r["title_norm"] in norms or placeholder_title(r["title"])]

    def last_undoable(self, nb: str) -> Optional[dict]:
        with self._lock:
            return self._decode(self._one("select * from writes where notebook_id=? and status='done' order by id desc limit 1", (nb,)))

    def last_change(self, nb: str) -> Optional[dict]:
        """The newest change office_undo acts on (F8 D2): a done write, or a held one (uncertain / checked) that undo
        must settle with a read first. A held write whose landed block undo could not identify is passed over."""
        with self._lock:
            for row in self._q("select * from writes where notebook_id=? and status in ('done','uncertain','checked') "
                               "order by id desc", (nb,)):
                row = self._decode(row) or {}
                if row.get("status") == "done" or not (row.get("extra") or {}).get("undoUnidentified"):
                    return row
            return None

    def superseded_for(self, nb: str, page_id: str) -> Optional[dict]:
        with self._lock:
            return self._decode(self._one("select * from writes where notebook_id=? and page_id=? and kind='links_outline' "
                                          "and status='superseded' order by id desc limit 1", (nb, page_id)))

    def writes_for_page(self, nb: str, page_id: str) -> list[dict]:
        with self._lock:
            return [self._decode(r) for r in self._q("select * from writes where notebook_id=? and page_id=? order by id", (nb, page_id))]

    def pending_manual_deletes(self, nb: str) -> list[dict]:
        with self._lock:
            return [self._decode(r) for r in self._q("select * from writes where notebook_id=? and status='manual_delete' order by id", (nb,))]

    def mark_page_deleted(self, nb: str, page_id: str) -> None:
        with self._lock:
            self._conn.execute("update pages set deleted=1 where notebook_id=? and page_id=?", (nb, page_id))
            if self.fts:
                self._conn.execute("delete from pages_fts where notebook_id=? and page_id=?", (nb, page_id))
            self._conn.commit()

    # ---- decisions / artifacts / provenance ---------------------------------------------------
    def add_decisions(self, nb: str, decision_page_id: str, source_page_id: str, items: list[dict]) -> list[int]:
        now = self._clock()
        ids = []
        with self._lock:
            for it in items:
                cur = self._conn.execute(
                    "insert into decisions(notebook_id,decision_page_id,source_page_id,kind,text,owner,due,rationale,quote,"
                    "quote_verified,created_at) values(?,?,?,?,?,?,?,?,?,?,?)",
                    (nb, decision_page_id, source_page_id, it["kind"], it["text"], it.get("owner"), it.get("due"),
                     it.get("rationale"), it.get("quote"), 1 if it.get("quoteVerified") else 0, now))
                ids.append(int(cur.lastrowid))
            self._conn.commit()
        return ids

    def decisions(self, nb: str, decision_page_id: str) -> list[dict]:
        with self._lock:
            return self._q("select * from decisions where notebook_id=? and decision_page_id=? order by id", (nb, decision_page_id))

    def search_decisions(self, nb: str, question: str, limit: int = 5) -> list[dict]:
        words = [w for w in re.findall(r"[\w]+", question.casefold()) if len(w) > 2 and w not in _STOP]
        if not words:
            return []
        with self._lock:
            rows = self._q("select d.*, p.title page_title, p.web_url page_url from decisions d join pages p on "
                           "p.notebook_id=d.notebook_id and p.page_id=d.decision_page_id where d.notebook_id=? and p.deleted=0", (nb,))
        scored = [(sum((r["text"] + " " + (r["rationale"] or "")).casefold().count(w) for w in words), r) for r in rows]
        return [r for s, r in sorted(scored, key=lambda x: -x[0]) if s][:limit]

    def add_artifact(self, nb: str, decision_page_id: str, host: str, document: str, locator: str,
                     decision_ids: list[int], summary: str) -> int:
        with self._lock:
            cur = self._conn.execute("insert into artifacts(notebook_id,decision_page_id,host,document,locator,decision_ids,summary,"
                                     "created_at) values(?,?,?,?,?,?,?,?)",
                                     (nb, decision_page_id, host, document, locator, json.dumps(decision_ids), summary[:300],
                                      self._clock()))
            self._conn.commit()
            return int(cur.lastrowid)

    def provenance(self, nb: str, host: Optional[str] = None, text: Optional[str] = None, page_id: Optional[str] = None) -> dict:
        with self._lock:
            out: dict = {"artifacts": [], "writes": []}
            arts = self._q("select * from artifacts where notebook_id=?" + (" and host=?" if host else "") + " order by id desc",
                           (nb, host) if host else (nb,))
            words = [w for w in re.findall(r"[\w]+", (text or "").casefold()) if len(w) > 2]
            for a in arts:
                ids = json.loads(a["decision_ids"] or "[]")
                ds = [d for d in self._q(f"select * from decisions where id in ({','.join('?' * len(ids))})", ids)] if ids else []
                if words:
                    phrase = " ".join((text or "").casefold().split())

                    def score(d):
                        hay = (d["text"] + " " + (d["owner"] or "")).casefold()
                        return sum(w in hay for w in words) + (10 if phrase and phrase in hay else 0)
                    ds = sorted((d for d in ds if score(d)), key=score, reverse=True)
                    if not ds:
                        continue
                dp = self.page(nb, a["decision_page_id"]) if a["decision_page_id"] else None
                chain = []
                for d in ds[:10]:
                    src = self.page(nb, d["source_page_id"]) if d["source_page_id"] else None
                    chain.append({"item": d["text"], "kind": d["kind"], "owner": d["owner"] or "Unknown",
                                  "quote": d["quote"], "quoteVerified": bool(d["quote_verified"]),
                                  "decisionPage": dp and {"pageId": dp["page_id"], "title": dp["title"], "webUrl": dp["web_url"]},
                                  "sourcePage": src and {"pageId": src["page_id"], "title": src["title"], "webUrl": src["web_url"]}})
                out["artifacts"].append({"host": a["host"], "document": a["document"], "locator": a["locator"],
                                         "summary": a["summary"], "createdAt": a["created_at"], "because": chain})
                if len(out["artifacts"]) >= 5:
                    break
            if page_id:
                for w in self.writes_for_page(nb, page_id):
                    out["writes"].append({"tool": w["tool"], "kind": w["kind"], "status": w["status"], "summary": w["summary"],
                                          "at": w["created_at"]})
            return out

    def resurface(self, nb: str, count: int = 3) -> list[dict]:
        """Least-recently-opened pages that carry decisions first, then any captured page. Index only."""
        with self._lock:
            rows = self._q("select p.*, (select count(*) from decisions d where d.notebook_id=p.notebook_id and "
                           "(d.decision_page_id=p.page_id or d.source_page_id=p.page_id)) dc from pages p "
                           "where p.notebook_id=? and p.deleted=0 and p.captured_at is not null "
                           "order by (dc>0) desc, coalesce(p.opened_at, p.captured_at) asc limit ?", (nb, count))
            return [{"pageId": r["page_id"], "title": r["title"], "webUrl": r["web_url"], "decisions": r["dc"],
                     "lastOpenedHoursAgo": round((self._clock() - (r["opened_at"] or r["captured_at"])) / 3600, 1)} for r in rows]

    def notebooks(self) -> list[str]:
        with self._lock:
            return [r["notebook_id"] for r in self._q("select distinct notebook_id from pages")]


_STOP = frozenset("""the and for are but not you all any can had her was one our out has have what when where which who why
how did does this that with from they them then than into about would could should there their decide decided
decision we us our""".split())
