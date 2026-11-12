# SPDX-License-Identifier: Apache-2.0
"""Automated in-pane live test: real Office for the web in the client Mac harness, real Hermes chat.

    PYTHONPATH=<hermes-agent> HERMES_HOME=<profile> <hermes venv python> scripts/livetest/live_driver.py \
        --control http://127.0.0.1:19334 --hosts excel,word --runs 3 --out <dir>

Per host and run: open the disposable "HOL live test <host>" document, upload the add-in through
Office's own Upload My Add-in dialog (file chooser fed over CDP), open the Hermes pane, create a
fresh chat in the gateway, ask it to "connect to <host>", read the code from the chat reply and
type it into the pane, then run the host script. Each step asserts on the tool results the chat
produced AND on an independent readback of the document (Office.js evaluated in the pane frame
by the harness, not by Hermes), and takes a screenshot. Results: <out>/results.json.
"""
from __future__ import annotations

import argparse
import asyncio
import json
import os
import re
import statistics
import subprocess
import sys
import time
import traceback
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
from chat import Chat, ChatError, connect_live, pairing_code  # noqa: E402
import office_ui as ui  # noqa: E402
from redact import EMAIL_RE, mask_people  # noqa: E402

MAC = os.environ.get("HOL_MAC", "office-client")
MAC_ROOT = "hermes-office-livetest"  # relative to the client Mac home
LABEL = {"excel": "Excel", "word": "Word", "powerpoint": "PowerPoint", "onenote": "OneNote", "outlook": "Outlook",
         "calendar": "Outlook calendar"}
OUTLOOK_STATE = Path(os.environ.get("HOL_OUTLOOK_STATE",
                                    str(Path.home() / ".hermes/profiles/omh/cache/scratch/hol-livetest/outlook-fixture-v2-sent.json")))


def now():
    return time.strftime("%Y-%m-%dT%H:%M:%S")


REPO = Path(__file__).resolve().parents[2]
SERVED_BASE = os.environ.get("HOL_SERVED_BASE", "https://office-addin.example.test/office/live/")


def _git(*args):
    try:
        return subprocess.run(["git", "-C", str(REPO), *args], capture_output=True, text=True, timeout=10).stdout.strip()
    except (OSError, subprocess.SubprocessError):
        return None


def served_sums(base=SERVED_BASE, fetch=None):
    """sha256 of the SHA256SUMS the add-in site serves right now (the exact pane bytes under test), or the error."""
    import hashlib
    import urllib.request
    url = base.rstrip("/") + "/SHA256SUMS"
    try:
        body = fetch(url) if fetch else urllib.request.urlopen(url, timeout=15).read()
    except Exception as exc:  # recorded, never fatal: the run still happens, the identity says "unknown"
        return {"url": url, "sha256": None, "error": f"{type(exc).__name__}: {str(exc)[:120]}"}
    return {"url": url, "sha256": hashlib.sha256(body).hexdigest(), "files": body.decode("utf-8", "replace").count("\n")}


def run_identity(out: Path, fetch=None):
    """What exactly was tested: plugin commit (and whether the tree was dirty), served add-in bytes, run id, and (A26)
    the profile's effective self-learning settings plus the learned skills/memory it held when the run started.
    harnessCommit is this checkout's HEAD. pluginCommit is the same, unless HOL_PLUGIN_COMMIT names the commit the
    profile's plugin dir was exported from (A26: a newer harness drives an older plugin); then that claim is checked
    byte for byte against `git archive` of it and pluginDirMatches says whether it held."""
    status = _git("status", "--porcelain")
    home = Path(os.environ.get("HERMES_HOME", "?"))
    head = _git("rev-parse", "HEAD")
    ident = {"runId": out.name, "pluginCommit": head, "harnessCommit": head,
             "pluginDirty": bool(status) if status is not None else None}
    claimed = os.environ.get("HOL_PLUGIN_COMMIT")
    if claimed:
        full = _git("rev-parse", "--verify", f"{claimed}^{{commit}}")
        match = plugin_dir_matches(full, home / "plugins" / "hermes-office-desktop") if full else {"match": False,
                                                                                                    "why": "unknown commit"}
        ident.update(pluginCommit=full or claimed, pluginDirMatches=match["match"], pluginDirCheck=match)
    return {**ident, "servedSums": served_sums(fetch=fetch), "profile": home.name, "started": now(),
            "selfLearning": self_learning(profile_config(home)), "learnedAtStart": learned_state(home)}


def plugin_dir_matches(commit, plugin_dir, archive=None):
    """Is plugin_dir exactly `git archive <commit>` (every file byte-identical, nothing missing, nothing extra except
    __pycache__)? archive: tar bytes (tests); default runs git archive in this checkout."""
    import io
    import tarfile
    try:
        data = archive if archive is not None else subprocess.run(["git", "-C", str(REPO), "archive", commit],
                                                                  capture_output=True, timeout=60, check=True).stdout
        want = {}
        with tarfile.open(fileobj=io.BytesIO(data)) as tf:
            for m in tf.getmembers():
                if m.isfile():
                    want[m.name] = tf.extractfile(m).read()
    except Exception as exc:  # noqa: BLE001 - recorded, never fatal
        return {"match": False, "why": f"archive: {type(exc).__name__}"}
    root = Path(plugin_dir)
    if not root.is_dir():
        return {"match": False, "why": f"no plugin dir at {root}"}
    have = {p.relative_to(root).as_posix(): p for p in root.rglob("*")  # archive names use "/"
            if p.is_file() and "__pycache__" not in p.relative_to(root).parts}
    missing = sorted(set(want) - set(have))
    extra = sorted(set(have) - set(want))
    differ = sorted(n for n in set(want) & set(have) if have[n].read_bytes() != want[n])
    return {"match": not (missing or extra or differ), "files": len(want), "missing": missing[:5], "extra": extra[:5],
            "differ": differ[:5]}


# ---- self-learning in the test profile (A26) -----------------------------------------------------------------------
# A profile's own agent learns between runs: the background review after a turn writes skills and MEMORY.md, the
# curator maintains skills. F4's warm story-bare 2/2 ran with a skill and a memory line the profile had written itself
# (one mapped "our team picture" to the harness PNG). Bare-prompt evidence must be cold, so the test profile turns all
# of it off. Real Hermes keys (hermes_cli/config_defaults.py, agent/agent_init.py, agent/curator.py, docs
# user-guide/features/{memory,skills,curator}.md): path -> (Hermes default, the cold value). scripts/livetest/
# cold_profile.sh sets exactly these.
SELF_LEARNING_KEYS = {
    "curator.enabled": (True, False),                       # the curator's skill maintenance pass
    "auxiliary.background_review.enabled": (True, False),  # the post-turn self-improvement review fork
    "skills.creation_nudge_interval": (10, 0),              # skill review nudges (0 = never)
    "skills.write_approval": (False, True),                 # any skill_manage write is staged, never applied
    "memory.memory_enabled": (True, False),                 # MEMORY.md: no store, no memory tool
    "memory.user_profile_enabled": (True, False),           # USER.md likewise
    "memory.nudge_interval": (10, 0),                       # memory review nudges (0 = never)
}


def profile_config(home):
    """The profile's config.yaml as a dict ({} when missing or unreadable; then every key is at its default)."""
    try:
        import yaml
        cfg = yaml.safe_load((Path(home) / "config.yaml").read_text(encoding="utf-8"))
    except Exception:  # noqa: BLE001 - recorded as defaults, which reads as "not cold"
        return {}
    return cfg if isinstance(cfg, dict) else {}


def self_learning(cfg):
    """Effective value of every SELF_LEARNING_KEYS key (the Hermes default when unset) and whether all are cold."""
    eff = {}
    for path, (default, _cold) in SELF_LEARNING_KEYS.items():
        node = cfg
        for part in path.split("."):
            node = node.get(part) if isinstance(node, dict) else None
        eff[path] = default if node is None else node
    cold = all(eff[p] == c and type(eff[p]) is type(c) for p, (_d, c) in SELF_LEARNING_KEYS.items())
    return {"settings": eff, "cold": cold}


def learned_state(home):
    """What the profile has learned on disk: every skill (a SKILL.md, by its folder path under skills/) and every
    memory file. Recorded at the start and the end of a run, so a skill or memory written mid-run shows."""
    home = Path(home)
    skills = home / "skills"
    names = sorted(p.parent.relative_to(skills).as_posix() for p in skills.rglob("SKILL.md")
                   if not any(part.startswith(".") for part in p.relative_to(skills).parts)) if skills.is_dir() else []
    mem = home / "memories"
    files = sorted(p.name for p in mem.glob("*.md")) if mem.is_dir() else []
    pend = home / "pending"
    staged = sorted(p.relative_to(pend).as_posix() for p in pend.rglob("*") if p.is_file()) if pend.is_dir() else []
    return {"skills": names, "memoryFiles": files, "pendingWrites": staged}


def cold_gate(identity, bare, allow_warm=None):
    """A26: story-bare is the bare-prompt evidence, so it refuses a profile that can learn (HOL_ALLOW_WARM=1 runs it
    anyway and the identity says warm). None = go; else the reason to stop before anything runs."""
    allow_warm = os.environ.get("HOL_ALLOW_WARM") == "1" if allow_warm is None else allow_warm
    sl = identity.get("selfLearning") or {}
    if not bare or sl.get("cold") or allow_warm:
        return None
    warm = [k for k, (_d, c) in SELF_LEARNING_KEYS.items() if (sl.get("settings") or {}).get(k) != c]
    return (f"COLD GATE: profile {identity.get('profile')} can learn ({', '.join(warm)}); run "
            "scripts/livetest/cold_profile.sh on it first, or set HOL_ALLOW_WARM=1 for a run reported as warm")


class Recorder:
    def __init__(self, out: Path, identity=None):
        self.out = out
        out.mkdir(parents=True, exist_ok=True)
        (out / "shots").mkdir(exist_ok=True)
        self.events = (out / "events.jsonl").open("a")
        self.results = []
        self.calendar_cleanup = None  # set by calendar_host's finally (A15)
        self.identity = identity if identity is not None else run_identity(out)

    def log(self, obj):
        obj = {"t": now(), **obj}
        if obj.get("ev") in ("user", "agent"):  # pairing codes are one-time; keep them out of logs
            obj["text"] = ui_redact(obj.get("text", ""))
        obj = mask_people(obj)  # no address or display name in events.jsonl (A20)
        self.events.write(json.dumps(obj, ensure_ascii=False) + "\n")
        self.events.flush()

    def add(self, result):
        """Every result row carries the run identity, so a results.json excerpt still names what it tested."""
        result["identity"] = self.identity
        self.results.append(result)
        return result

    def save(self):
        (self.out / "results.json").write_text(json.dumps(mask_people(self.results), indent=2, ensure_ascii=False))


def ui_redact(text: str) -> str:
    from chat import CODE_RE
    return mask_people(URL_RE.sub("<url>", CODE_RE.sub("XXXX-XXXX", text or "")))


def fetch_shot(pane: ui.Pane, rec: Recorder, name: str, calendar: bool = False) -> str:
    # Outlook: the folder pane, message list, reading pane and address fields are blurred too (ui.outlook_shot);
    # calendar runs blur all of Outlook except the task pane (ui.calendar_shot: grid, forms, message list, account).
    if calendar:
        path = ui.calendar_shot(pane, name)
    else:
        path = ui.outlook_shot(pane, name) if getattr(pane, "tab", None) == ui.OUTLOOK_TAB else pane.shot(name)
    dest = rec.out / "shots" / name
    subprocess.run(["scp", "-q", "-o", "BatchMode=yes", f"{MAC}:{path}", str(dest)], check=False)
    return str(dest)


def timings(tools):
    out = []
    for t in tools:
        r = t.get("result")
        if isinstance(r, dict):
            out.append({"tool": t["name"], "elapsedMs": r.get("elapsedMs"), "roundTripMs": r.get("roundTripMs"),
                        "harnessMs": t.get("harnessMs")})
    return out


def office_tools(tools, prefix=("excel_", "word_", "powerpoint_", "onenote_", "brain_", "office_", "outlook_", "calendar_")):
    return [t for t in tools if str(t.get("name", "")).startswith(prefix)]


def _same_as_before(r):
    """Hermes core replaces a result byte-identical to an earlier one in the same turn with a short note."""
    return isinstance(r, str) and r.startswith("[hermes note: this result is byte-identical")


def last_ok(tools, name):
    for t in reversed(tools):
        if t.get("name") == name and not _same_as_before(t.get("result")):
            return t.get("result")
    return None


def last_result(tools, name):
    for t in reversed(tools):
        if t.get("name") == name:
            return t.get("result")
    return None


# ---- independent readbacks (Office.js in the pane frame, run by the harness) -------------------
READBACK = {
    "excel": r"""Excel.run(async ctx => { const s = ctx.workbook.worksheets.getActiveWorksheet(); s.load('name');
        const r = s.getRange('A1:B6'); r.load('values,formulas'); const f = s.getRange('A1:B1').format.font; f.load('bold');
        const c = s.charts; c.load('items/name'); await ctx.sync();
        return { sheet: s.name, values: r.values, formulas: r.formulas, headerBold: f.bold, charts: c.items.length } })""",
    "word": r"""Word.run(async ctx => { const b = ctx.document.body; b.load('text'); const ps = b.paragraphs; ps.load('items/text');
        await ctx.sync(); const hits = ps.items.map((p, i) => [i, p.text]).filter(([, t]) => t.includes('HOL'));
        const out = { text: b.text.slice(0, 2000), paragraphs: ps.items.length, bold: [] };
        for (const [i] of hits) { const f = ps.items[i].font; f.load('bold'); await ctx.sync(); out.bold.push([ps.items[i].text.slice(0, 80), f.bold]) }
        return out })""",
    # Text only from shapes that have a text frame: loading textFrame on a picture fails the WHOLE batch with
    # RichApi InvalidArgument (A8 live, L11: a picture left on the deck by an earlier run).
    "powerpoint": r"""PowerPoint.run(async ctx => { const ss = ctx.presentation.slides; ss.load('items/id'); await ctx.sync();
        const slides = []; for (const s of ss.items) { const sh = s.shapes; sh.load('items/id,items/type'); await ctx.sync(); const texts = [];
          for (const x of sh.items) { if (!['GeometricShape', 'TextBox', 'Placeholder'].includes(x.type)) { texts.push(''); continue }
            try { const tr = x.textFrame.textRange; tr.load('text'); await ctx.sync(); texts.push(tr.text) } catch (e) { texts.push('') } }
          slides.push(texts) }
        return { count: ss.items.length, slides } })""",
    "onenote": r"""OneNote.run(async ctx => { const p = ctx.application.getActivePage(); p.load('id,title'); const nb = ctx.application.getActiveNotebook();
        nb.load('id,name'); const outlines = p.contents; outlines.load('items/type,items/outline/id');
        await ctx.sync(); let links = 0, text = []; const isRt = x => x.type === 'RichText';
        // richText is loaded per paragraph, only where the type is RichText: loading it on a collection that holds a picture
        // or a table fails the WHOLE batch with "This paragraph is not of type RichText" (A11 live, OneNote r2).
        for (const c of outlines.items) { if (c.type !== 'Outline') continue; const ps = c.outline.paragraphs; ps.load('items/type');
          await ctx.sync();
          // List items (<ul><li>) are child paragraphs of the paragraph above them in OneNote for the web: read one level.
          const kids = ps.items.map(x => { if (!isRt(x)) return null; x.richText.load('text'); const k = x.paragraphs; k.load('items/type'); return k });
          await ctx.sync();
          kids.forEach(k => k && k.items.forEach(y => { if (isRt(y)) y.richText.load('text') })); await ctx.sync();
          const rt = x => { try { return isRt(x) ? x.richText.text : '' } catch { return '' } };
          const t = ps.items.map((x, i) => [rt(x), ...(kids[i] ? kids[i].items.map(rt) : [])].filter(Boolean).join(' ')).join(' '); text.push(t.slice(0, 4000));
          if (/Links & Backlinks/.test(t)) links++ }
        return { notebookId: nb.id, notebook: nb.name, pageTitle: p.title, outlines: text.length, linksBlocks: links, text } })""",
}

RESET = {
    "excel": r"""Excel.run(async ctx => { const s = ctx.workbook.worksheets.getActiveWorksheet(); s.getRange('A1:H40').clear();
        const c = s.charts; c.load('items'); await ctx.sync(); c.items.forEach(x => x.delete()); await ctx.sync(); return true })""",
    "word": r"""Word.run(async ctx => { ctx.document.body.clear(); await ctx.sync();
        ctx.document.body.insertParagraph('HOL live test document. The colour of the sky is blue.', 'Start'); await ctx.sync(); return true })""",
    "powerpoint": r"""PowerPoint.run(async ctx => { const ss = ctx.presentation.slides; ss.load('items'); await ctx.sync();
        ss.items.slice(1).forEach(s => s.delete()); await ctx.sync();
        const first = ctx.presentation.slides.getItemAt(0); const sh = first.shapes; sh.load('items/name,items/type'); await ctx.sync();
        sh.items.filter(x => /HOL/.test(x.name || '') || x.type === 'Image').forEach(x => x.delete()); await ctx.sync(); return true })""",
}


# ---- OneNote fixture: fresh per run ------------------------------------------------------------
# Add-ins cannot delete pages, and Links & Backlinks ownership is per profile (brain.db): a page
# reused from an earlier run or another profile can carry a block this profile rightly will not
# touch. So every run seeds its own section with run-tagged titles (unique notebook-wide, so
# [[links]] never turn ambiguous against older copies) and checks the page is clean before it
# asserts on block counts. Pages are never deleted by the harness.
def onenote_fixture(run_no, stamp=None):
    """Section name, title map and seed prompt for one run (pure; no Office or Hermes)."""
    sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "demo"))
    import demo_notebook
    tag = f"HOL r{run_no} {stamp or time.strftime('%m%d-%H%M%S')}"
    nb = demo_notebook.load()
    titles = {p["title"]: f"{p['title']} ({tag})" for s in nb["sections"] for p in s["pages"]}

    def retarget(html):
        return brainlink_re.sub(lambda m: "[[" + titles.get(m.group(1), m.group(1)) + (m.group(2) or "") + "]]", html)
    lines = [f"Set up the live-test pages. I created the section '{tag}' already. Call onenote_list_sections, then call "
             "onenote_create_page once per page below with exactly this title and HTML and the sectionId of that "
             "section. Do not change the text. After the last page, call brain_sync and tell me the coverage.",
             f"\nSection '{tag}':"]
    for s in nb["sections"]:
        for p in s["pages"]:
            lines.append(f"- title: {titles[p['title']]}\n  html: {retarget(''.join(p.get('html') or []))}")
    return {"tag": tag, "section": tag, "titles": titles, "seed": "\n".join(lines)}


brainlink_re = re.compile(r"\[\[([^\[\]|\n]+?)(\|[^\[\]\n]*?)?\]\]")


def onenote_reset_js(section):
    """Create this run's section; refuse (stale fixture) if it already exists."""
    return ("OneNote.run(async ctx => { const nb = ctx.application.getActiveNotebook(); const ss = nb.sections; ss.load('items/name');"
            f" await ctx.sync(); const name = {json.dumps(section)}; if (ss.items.some(x => x.name === name)) return {{ stale: name }};"
            " nb.addSection(name); await ctx.sync(); return { created: name } })")


def links_precondition(rb, c):
    """Before asserting 'exactly one block': the page must be this run's page and carry none yet."""
    want = c["fx"]["titles"]["Q4 Launch"]
    if rb.get("pageTitle") != want:
        return False, f"stale fixture: open page is {rb.get('pageTitle')!r}, expected this run's {want!r}"
    if rb.get("linksBlocks"):
        return False, (f"stale fixture: {want!r} already has {rb['linksBlocks']} Links & Backlinks block(s) before the step "
                       "(written by an earlier run or another profile); not a plugin result")
    return True, "clean"


PRECONDITIONS = {"11_links_twice": links_precondition}


def _links_texts(rb):
    return [t for t in rb.get("text") or [] if "Links & Backlinks" in t]


def _user_texts(rb):
    return [t for t in rb.get("text") or [] if "Links & Backlinks" not in t]


def links_twice_check(t, rb, c):
    calls = sum(1 for x in t["tools"] if x["name"] == "brain_links" and (_ok(x["result"]) or _same_as_before(x["result"])))
    ok = calls == 2 and _ok(last_ok(t["tools"], "brain_links")) and rb["linksBlocks"] == 1
    if ok:
        c["links_before_undo"] = {"links": _links_texts(rb), "user": _user_texts(rb), "pageTitle": rb.get("pageTitle")}
    return ok, f"{calls} brain_links calls, {rb['linksBlocks']} Links & Backlinks block(s) (exactly one expected)"


def undo_links_check(t, rb, c):
    """The second brain_links replaced the first; undo removes it and restores the first block exactly."""
    before = c.get("links_before_undo")
    if not before:
        return False, "no recorded state from 11_links_twice; cannot judge the undo"
    d = parsed(last_result(t["tools"], "office_undo"))
    data = (d or {}).get("data") or {}
    problems = []
    if not _ok(d):
        problems.append("office_undo result missing, truncated or not ok")
    if data.get("undone") != "brain_links":
        problems.append(f"undid {data.get('undone')!r}, expected 'brain_links'")
    if data.get("restoredPreviousLinksBlock") is not True:
        problems.append("previous block not reported restored")
    if rb.get("pageTitle") != before["pageTitle"]:
        problems.append(f"open page is {rb.get('pageTitle')!r}")
    if rb.get("linksBlocks") != 1 or _links_texts(rb) != before["links"]:
        problems.append(f"links block state differs ({rb.get('linksBlocks')} block(s))")
    if _user_texts(rb) != before["user"]:
        problems.append("user text changed")
    return not problems, "; ".join(problems) or "exactly one block, identical to the first; user text unchanged"


def _flat(s):
    return " ".join(str(s or "").split()).casefold()


def link_listed(title, block):
    """Is the page ``title`` listed in a Links & Backlinks block as OneNote for the web shows it?

    OneNote labels a link to one of its pages itself (A8 live): the fixture's run suffix can be cut
    ('Daily 2026-11-06 (HOL r1 1106-', 'Pricing review 2026-10-20 (HOL r1 ') and a colon dropped
    ('Decisions launch meeting (HOL r1 ...)'). The full title, or the base title followed by this run's
    '(HOL r<N>' marker, counts; the base title alone (another run's page) does not."""
    def norm(x):
        return _flat(str(x or "").replace(":", " "))
    t, b = norm(title), norm(block)
    if t in b:
        return True
    m = re.match(r"^(.*?) \((hol r\d+)\b", t)
    return bool(m) and re.search(re.escape(f"{m.group(1)} ({m.group(2)}") + r"(?!\d)", b) is not None


def decision_page_check(t, rb, c):
    """Independent readback of the decision page against what the tool claims and what the notes say."""
    d = parsed(last_result(t["tools"], "brain_decision_page"))
    if not _ok(d):
        return False, "brain_decision_page result missing, truncated or not ok"
    data = d["data"]
    page = _flat(" ".join(rb.get("text") or []))
    items = data.get("items") or []
    problems = []
    if rb.get("pageTitle") != data.get("title"):
        problems.append(f"open page {rb.get('pageTitle')!r} is not the decision page {data.get('title')!r}")
    meeting = c["fx"]["titles"]["Launch meeting 2026-11-03"]
    # OneNote for the web shows a link to one of its pages with its own (sometimes shortened) label, e.g. the
    # fixture's "(HOL r1 1106-105403)" suffix cut to "(HOL r1" (seen live, A8): the base title must be there.
    base = meeting.split(" (HOL ")[0]
    if _flat("Source: " + meeting) not in page and _flat("Source: " + base) not in page:
        problems.append("source line missing")
    missing = [i["text"] for i in items if _flat(i["text"]) not in page]
    if missing:
        problems.append(f"items not on the page: {missing}")
    if not any(i["kind"] == "decision" for i in items) or not any(i["kind"] == "action" for i in items):
        problems.append("needs at least one decision and one action")
    unknown = page.count(_flat("owner: Unknown"))
    if unknown != data.get("unknownOwners"):
        problems.append(f"page shows {unknown} unknown owners, tool says {data.get('unknownOwners')}")
    unverified = page.count(_flat("(not found in the source page)"))
    if unverified != data.get("unverifiedQuotes"):
        problems.append(f"page shows {unverified} unverified quotes, tool says {data.get('unverifiedQuotes')}")
    supplier = [i for i in items if "supplier lead time" in _flat(i["text"])]
    if not supplier or not all(_flat(i["owner"]).startswith("unknown") for i in supplier):
        problems.append("'check supplier lead times' must be an item with owner Unknown (the notes name nobody)")
    plan = [i for i in items if "pilot plan" in _flat(i["text"])]
    if not plan or not any(_flat(i["owner"]) == "priya" for i in plan):
        problems.append("'draft the pilot plan' must be owned by Priya (the notes say so)")
    return not problems, "; ".join(problems) or (f"{len(items)} items on the page; {unknown} unknown owners, "
                                                  f"{unverified} unverified quotes, as the tool reported")


READBACK_WAITS = (5, 5, 5)  # s before each retry of a failed READ (bounded: the first read and three retries)
# Read-only probes for "the document has settled": two consecutive equal answers.
SETTLE_JS = {
    "word": r"""Word.run(async ctx => { const b = ctx.document.body; b.load('text'); const ps = b.paragraphs; ps.load('items/text');
        await ctx.sync(); return [ps.items.length, b.text.slice(0, 2000)] })""",
}


def settle(pane, host, reads=8, interval=1.5, sleep=time.sleep):
    """Wait (bounded, read-only) until two consecutive reads of the open document agree. True once settled."""
    probe = SETTLE_JS.get(host)
    if not probe:
        return True
    prev = None
    for _ in range(reads):
        try:
            cur = pane.run_js(ui.task_pane_frame(pane, host), probe, timeout=30)
        except ui.HarnessError:
            cur = None
        if cur is not None and cur == prev:
            return True
        prev = cur
        sleep(interval)
    return False


def readback(pane, host, js=None, waits=READBACK_WAITS, sleep=time.sleep):
    """The harness's own read of the open document. Only this READ is ever retried, never a write.

    Before each retry: wait, then look for Office for the web's own "Sorry, your session expired. Please refresh the
    page" banner. Word for the web showed it in every recorded "Cannot read properties of null (reading 'p5d')" failure
    (36afdac r3, e999e57 r2, c6d20e6 r2) and every Office.js read fails until the page is reloaded; so the harness
    reloads the document once, as a person pressing Refresh would (no write), and waits for the pane to resume its
    pairing. Without a banner: the pane frame is reloaded once before the last retry, as before. Each retry waits for
    the document to settle first. A session that stays expired is reported as such, not as a plugin result."""
    script = js or READBACK[host]
    frame0 = ui.task_pane_frame(pane, host)
    log = getattr(pane, "log", None) or (lambda ev: None)
    last, expired, refreshed, reloaded = None, None, False, False
    for attempt in range(len(waits) + 1):
        if attempt:
            sleep(waits[attempt - 1])
            expired = ui.expired_session_frame(pane, host)
            if expired and not refreshed:
                refreshed = True
                manifest = (getattr(pane, "hol_manifests", None) or {}).get(host)
                try:
                    ui.refresh_expired_session(pane, host, log, manifest, sleep=sleep)
                except ui.HarnessError as exc:
                    raise ui.HarnessError(f"{host} for the web: session expired (host banner); the document reload did not "
                                          f"bring the paired pane back ({exc}); environment, not a plugin result: {last}")
            elif not expired and not reloaded and attempt == len(waits):  # last resort: reload the pane frame
                reloaded = True
                pane.cmd("reloadFrame", frame=ui.task_pane_frame(pane, host))
                sleep(3)
                ui.wait_pane_state(pane, host, ("paired",), timeout=60)
            settled = settle(pane, host, sleep=sleep)
            log({"ev": "readback-retry", "host": host, "attempt": attempt, "why": str(last)[:160],
                 "expired": bool(expired), "settled": settled})
        try:
            return pane.run_js(ui.task_pane_frame(pane, host) or frame0, script, timeout=60)
        except ui.HarnessError as exc:
            last = exc
    if expired or (refreshed and ui.expired_session_frame(pane, host)):
        raise ui.HarnessError(f"{host} for the web: session expired (host banner 'Sorry, your session expired'); reloaded "
                              f"once, the read still fails; environment, not a plugin result: {last}")
    raise last


# ---- per-host scripts ----------------------------------------------------------------------------
# Each step: (name, prompt(ctx) -> str, check(turn, rb, ctx) -> (bool, note)). ``rb`` is the readback after.

def parsed(r):
    """The tool result as a dict, or None. A truncated result (a string that is not JSON) is NOT evidence."""
    if isinstance(r, str):
        try:
            r = json.loads(r)
        except ValueError:
            return None
    return r if isinstance(r, dict) else None


def receipts_ok(data):
    """A composite result passes only if its own status and EVERY nested per-host receipt are verified."""
    nested = data.get("results") if isinstance(data, dict) else None
    if not isinstance(nested, list) or not any(isinstance(x, dict) and "host" in x for x in nested):
        return True  # not a composite
    return data.get("status") == "verified" and all(isinstance(x, dict) and x.get("status") == "verified" for x in nested)


def _ok(r):
    # Large results can reach the event stream truncated (a string, not JSON). That is missing evidence, not a
    # pass: the step fails and says so (see truncated_results), instead of trusting the envelope's first bytes.
    d = parsed(r)
    return d is not None and d.get("ok") is True and receipts_ok(d.get("data"))


def truncated_results(tools):
    return [t.get("name") for t in tools if isinstance(t.get("result"), str) and parsed(t["result"]) is None
            and not _same_as_before(t["result"])]


def excel_steps():
    data = [["Item", "Amount"], ["North", 100], ["South", 150], ["East", 175], ["West", 225]]

    def s(ctx):
        return ctx["sheet"]
    return [
        ("overview", lambda c: "Call excel_get_workbook_overview and tell me the sheet names. Do nothing else.",
         lambda t, rb, c: (_ok(last_result(t["tools"], "excel_get_workbook_overview")) and rb["sheet"] in json.dumps(last_result(t["tools"], "excel_get_workbook_overview")), "overview lists the active sheet")),
        ("write", lambda c: f"Call excel_write_range with sheet {s(c)!r}, address 'A1:B5' and values {json.dumps(data)}. Do nothing else.",
         lambda t, rb, c: (_ok(last_result(t["tools"], "excel_write_range")) and rb["values"][:5] == data, "A1:B5 holds the table")),
        ("format", lambda c: f"Call excel_format_range with sheet {s(c)!r}, address 'A1:B1' and format {{\"bold\": true, \"fillColor\": \"#FFF2CC\"}}. Do nothing else.",
         lambda t, rb, c: (_ok(last_result(t["tools"], "excel_format_range")) and rb["headerBold"] is True, "header bold")),
        ("sum", lambda c: f"Call excel_write_range with sheet {s(c)!r}, address 'A6:B6' and values [[\"Total\", \"=SUM(B2:B5)\"]], then call excel_read_range for {s(c)!r} 'B6' and tell me the value.",
         lambda t, rb, c: (_ok(last_result(t["tools"], "excel_read_range")) and "650" in json.dumps(last_result(t["tools"], "excel_read_range")) and rb["values"][5][1] == 650 and rb["formulas"][5][1] == "=SUM(B2:B5)", "B6 = SUM = 650 (tool read + readback)")),
        ("chart", lambda c: f"Call excel_create_chart with sheet {s(c)!r}, dataRange 'A1:B5', chartType 'ColumnClustered' and title 'HOL live test'. Do nothing else.",
         lambda t, rb, c: (_ok(last_result(t["tools"], "excel_create_chart")) and rb["charts"] == 1, "one chart")),
        ("undo", lambda c: "Call office_undo with app 'excel' exactly once. Do nothing else.",
         lambda t, rb, c: (_ok(last_result(t["tools"], "office_undo")) and rb["charts"] == 0, "chart removed by undo")),
        ("read", lambda c: f"Call excel_read_range for sheet {s(c)!r}, address 'A1:B6' and tell me the total.",
         lambda t, rb, c: (_ok(last_result(t["tools"], "excel_read_range")) and rb["values"][5][1] == 650 and "650" in t["text"], "read back table and total")),
    ]


def word_steps():
    para = "HOL inserted paragraph: the quick brown fox."
    return [
        ("overview", lambda c: "Call word_get_document_overview and summarise it in one line. Do nothing else.",
         lambda t, rb, c: (_ok(last_result(t["tools"], "word_get_document_overview")), "overview ok")),
        ("insert", lambda c: f"Call word_insert_text with text {para!r} and location 'end'. Do nothing else.",
         lambda t, rb, c: (_ok(last_result(t["tools"], "word_insert_text")) and para in rb["text"], "paragraph in body")),
        ("format", lambda c: "I selected the paragraph you inserted. Call word_format_selection with bold true. Do nothing else.",
         lambda t, rb, c: (_ok(last_result(t["tools"], "word_format_selection")) and any(b is True for txt, b in rb["bold"] if "quick brown fox" in txt), "inserted paragraph bold")),
        ("find_replace", lambda c: "Call word_find_replace with find 'quick brown fox' and replace 'lazy grey cat'. Do nothing else.",
         lambda t, rb, c: (_ok(last_result(t["tools"], "word_find_replace")) and "lazy grey cat" in rb["text"] and "quick brown fox" not in rb["text"], "text replaced")),
        ("undo", lambda c: "Call office_undo with app 'word' exactly once. Do nothing else.",
         lambda t, rb, c: (_ok(last_result(t["tools"], "office_undo")) and "quick brown fox" in rb["text"] and "lazy grey cat" not in rb["text"], "replace undone")),
        ("read", lambda c: "Call word_read_document_text and quote the last paragraph.",
         lambda t, rb, c: (_ok(last_result(t["tools"], "word_read_document_text")) and "quick brown fox" in json.dumps(last_result(t["tools"], "word_read_document_text")), "read back")),
    ]


def powerpoint_steps():
    box = "HOL live test text box"
    return [
        ("list", lambda c: "Call powerpoint_get_presentation_overview and tell me how many slides there are. Do nothing else.",
         lambda t, rb, c: (_ok(last_result(t["tools"], "powerpoint_get_presentation_overview")), f"slides={rb['count']}")),
        ("add_slide", lambda c: "Call powerpoint_create_slide with no arguments. Do nothing else.",
         lambda t, rb, c: (_ok(last_result(t["tools"], "powerpoint_create_slide")) and rb["count"] == c["slides0"] + 1, "slide added")),
        ("text_box", lambda c: f"Call powerpoint_add_text_box with text {box!r} and slideIndex {c['slides0']}. Do nothing else.",
         lambda t, rb, c: (_ok(last_result(t["tools"], "powerpoint_add_text_box")) and box in json.dumps(rb["slides"]), "text box on the new slide")),
        ("undo", lambda c: "Call office_undo with app 'powerpoint' exactly once. Do nothing else.",
         lambda t, rb, c: (_ok(last_result(t["tools"], "office_undo")) and box not in json.dumps(rb["slides"]), "text box removed")),
        ("read", lambda c: "Call powerpoint_read_all_slides and tell me the slide count.",
         lambda t, rb, c: (_ok(last_result(t["tools"], "powerpoint_read_all_slides")) and rb["count"] == c["slides0"] + 1, "read back")),
    ]


def _created_title(x):
    a = x.get("args")
    if isinstance(a, str):
        try:
            a = json.loads(a)
        except ValueError:
            a = None
    r = x.get("result")
    return (a or {}).get("title") or ((r.get("data") or {}).get("title") if isinstance(r, dict) else None)


def seed_created(t, titles):
    """Every fixture page was created by this turn: a fresh section has no duplicates to tolerate."""
    made = {_created_title(x) for x in t["tools"] if x["name"] == "onenote_create_page" and _ok(x["result"])}
    dupes = [x for x in t["tools"] if x["name"] == "onenote_create_page" and "Duplicate" in json.dumps(x["result"], default=str)]
    if dupes:
        return False, "stale fixture: DuplicateTitle in this run's fresh section"
    missing = [v for v in titles.values() if v not in made]
    return (not missing, f"{len(titles) - len(missing)}/{len(titles)} pages created" + (f"; missing {missing}" if missing else ""))


def onenote_steps():
    def T(c, k):
        return c["fx"]["titles"][k]

    def seed_check(t, rb, c):
        ok, note = seed_created(t, c["fx"]["titles"])
        return ok and _ok(last_result(t["tools"], "brain_sync")), note + " and synced"
    return [
        ("6_context", lambda c: "Call onenote_get_context and tell me the notebook name. Do nothing else.",
         lambda t, rb, c: (_ok(last_result(t["tools"], "onenote_get_context")) and bool(rb["notebookId"]), "notebook id non-empty")),
        ("6_reload_same_id", "RELOAD_PANE",
         lambda t, rb, c: (rb["notebookId"] == c.get("notebookId"), "same notebook id after pane reload (resumed)")),
        ("7_seed", lambda c: c["fx"]["seed"], seed_check),
        ("8_brain_sync", lambda c: "Call brain_sync once (no crawl) and repeat its coverage line. Do nothing else.",
         lambda t, rb, c: (_ok(last_result(t["tools"], "brain_sync")) and "indexed" in json.dumps(last_result(t["tools"], "brain_sync")).lower(), "coverage reported")),
        ("9_visit_index", lambda c: f"VISIT:{c['fx']['section']}/{T(c, 'Risks')}",
         lambda t, rb, c: (_ok(last_result(t["tools"], "brain_ask")) and "repair partners" in json.dumps(last_result(t["tools"], "brain_ask")), "visited page found by brain_ask")),
        ("10_active_only", lambda c: f"Call onenote_list_pages, then call onenote_get_page with the pageId of {T(c, 'Q4 Launch')!r} while {T(c, 'Risks')!r} is open, and report the error category. Then call onenote_open_page for {T(c, 'Q4 Launch')!r} and onenote_get_page for it, and quote its first sentence.",
         lambda t, rb, c: ("NotActivePage" in json.dumps([x["result"] for x in t["tools"] if x["name"] == "onenote_get_page"]) and rb["pageTitle"] == T(c, "Q4 Launch"), "NotActivePage, then open + read")),
        ("11_links_twice", lambda c: f"With {T(c, 'Q4 Launch')!r} open, call brain_links for it, then call brain_links again. Do not create pages.",
         links_twice_check),
        ("12_undo", lambda c: "Call office_undo with app 'onenote' exactly once. Do nothing else.", undo_links_check),
        # A fresh title per run: add-ins cannot delete pages, so earlier runs' decision pages stay (and a
        # duplicate title is rightly refused).
        ("13_decision_page", lambda c: (f"Open {T(c, 'Launch meeting 2026-11-03')!r} with onenote_open_page and record its decisions and actions with "
                                        f"brain_decision_page, title 'Decisions: launch meeting (HOL run {c['run']} {time.strftime('%H%M%S')})'. "
                                        "Tell me how many owners are Unknown and how many quotes are unverified."),
         decision_page_check),
        # Two-way links (implemented; not yet run live). The person opens a page that links to Q4 Launch, then
        # Q4 Launch itself: the pane's read-only panel must show that page as a backlink (no tool, no write).
        ("14_panel_backlink", lambda c: f"PANEL:{c['fx']['section']}/{T(c, 'Pricing review 2026-10-20')}/{T(c, 'Q4 Launch')}",
         panel_check),
        ("15_both_ways", lambda c: (f"Call onenote_open_page for {T(c, 'Q4 Launch')!r}, then call brain_link_both_ways once with "
                                    f"title {T(c, 'Q4 Launch')!r}. Report each page's status. Do nothing else."),
         both_ways_check),
        ("16_both_ways_undo", lambda c: (f"Call office_undo with app 'onenote' exactly {c.get('bw_undo_calls') or len(BOTH_WAYS_TARGETS)} "
                                         "times, one call after another. Do nothing else."),
         both_ways_undo_check),
    ]


# ---- two-way links: the pane's live panel and brain_link_both_ways (implemented; not yet run live) ----
BOTH_WAYS_TARGETS = ("Pricing review 2026-10-20", "Launch meeting 2026-11-03", "Risks")  # what Q4 Launch links to (demo seed)

PANEL_JS = r"""(() => { const r = document.getElementById('links-panel'); if (!r) return { present: false };
    const groups = {}; for (const g of r.querySelectorAll('.links-group')) { const h = g.querySelector('h2');
      groups[h ? h.textContent : '?'] = [...g.querySelectorAll('button')].map(b => b.textContent) }
    return { present: true, hidden: r.hidden, pageId: r.dataset.pageId || null, groups, text: r.textContent.slice(0, 2000),
             markup: !!r.querySelector('img,script,iframe,a[href]') } })()"""


def panel_check(t, rb, c):
    """After the person visits a page linking to Q4 Launch and then Q4 Launch, the panel lists that page as a backlink."""
    titles = c["fx"]["titles"]
    src, here = titles["Pricing review 2026-10-20"], titles["Q4 Launch"]
    p = c.get("panel") or {}
    problems = []
    if not p.get("present") or p.get("hidden"):
        problems.append("Links/Backlinks panel missing or hidden")
    if rb.get("pageTitle") != here:
        problems.append(f"open page is {rb.get('pageTitle')!r}, expected {here!r}")
    groups = p.get("groups") or {}
    if src not in (groups.get("Backlinks") or []):
        problems.append(f"backlink {src!r} not shown (Backlinks: {groups.get('Backlinks')})")
    if titles["Risks"] not in (groups.get("Links") or []):
        problems.append(f"outbound link {titles['Risks']!r} not shown (Links: {groups.get('Links')})")
    text = p.get("text") or ""
    if not re.search(r"indexed \d+ of \d+ pages", text) or "brain_sync" not in text:
        problems.append("coverage line or brain_sync hint missing")
    if p.get("markup"):
        problems.append("the panel contains markup elements (note text must render as text)")
    if t.get("tools"):
        problems.append("the panel step called tools")
    if c.get("panel_ledger_before") != c.get("panel_ledger_after"):
        problems.append(f"ledger changed during the panel step ({c.get('panel_ledger_before')} -> {c.get('panel_ledger_after')})")
    return not problems, "; ".join(problems) or f"panel on {here!r} shows backlink {src!r}; read-only (ledger unchanged)"


def _goto_read(pane, host, title):
    pane.run_js(ui.task_pane_frame(pane, host), onenote_goto_js(title), timeout=60)
    time.sleep(3)
    return readback(pane, host)


def both_ways_pre(pane, host, c):
    c["bw_before"] = {k: _goto_read(pane, host, c["fx"]["titles"][k]) for k in BOTH_WAYS_TARGETS}
    c["bw_ledger_before"] = ledger_snapshot(ledger_path(), c.get("notebookId"))


def both_ways_post(pane, host, c, key):
    active = readback(pane, host)  # first: where the tool left the person
    c[key] = {k: _goto_read(pane, host, c["fx"]["titles"][k]) for k in BOTH_WAYS_TARGETS}
    return active


def both_ways_check(t, rb, c):
    titles = c["fx"]["titles"]
    q4 = titles["Q4 Launch"]
    d = parsed(last_result(t["tools"], "brain_link_both_ways"))
    data = (d or {}).get("data") or {}
    problems = []
    if not _ok(d):
        problems.append("brain_link_both_ways result missing, truncated or not ok")
    pages = data.get("pages") or []
    if data.get("status") != "verified" or [p.get("status") for p in pages] != ["verified"] * len(BOTH_WAYS_TARGETS):
        problems.append(f"status {data.get('status')!r}, pages {[(p.get('title'), p.get('status')) for p in pages]}")
    if sorted(p.get("title") for p in pages) != sorted(titles[k] for k in BOTH_WAYS_TARGETS):
        problems.append("linked pages differ from the seed's links")
    if rb.get("pageTitle") != q4 or not (data.get("returnedTo") or {}).get("ok"):
        problems.append(f"did not return to {q4!r} (open: {rb.get('pageTitle')!r})")
    asked = [a for a in t.get("approvals") or [] if "links to" in a.get("description", "")]
    if t.get("approvals") is not None and len(asked) != 1:
        problems.append(f"{len(asked)} two-way approval prompt(s), expected exactly one")
    before, after = c.get("bw_before") or {}, c.get("bw_after") or {}
    for k in BOTH_WAYS_TARGETS:
        b, a = before.get(k) or {}, after.get(k) or {}
        if _user_texts(a) != _user_texts(b):
            problems.append(f"user text changed on {titles[k]!r}")
        blocks = _links_texts(a)
        if a.get("linksBlocks") != 1 or not any("Backlinks:" in x and q4 in x for x in blocks):
            problems.append(f"{titles[k]!r}: {a.get('linksBlocks')} block(s), backlink to {q4!r} not found")
    if not problems:
        c["bw_undo_calls"] = (data.get("undo") or {}).get("calls")
        c["bw_order"] = [p.get("pageId") for p in pages]
    return not problems, "; ".join(problems) or f"{len(pages)} linked pages show the backlink; returned to {q4!r}"


def both_ways_undo_check(t, rb, c):
    order = c.get("bw_order")
    if not order:
        return False, "no recorded state from 15_both_ways; cannot judge the undo"
    undos = [parsed(x["result"]) for x in t["tools"] if x["name"] == "office_undo"]
    problems = []
    if len(undos) != len(order) or not all(_ok(u) for u in undos):
        problems.append(f"{len(undos)} office_undo result(s) ok={[_ok(u) for u in undos]}, expected {len(order)}")
    data = [(u or {}).get("data") or {} for u in undos]
    if [x.get("undone") for x in data] != ["brain_link_both_ways"] * len(data):
        problems.append(f"undid {[x.get('undone') for x in data]}")
    if [x.get("pageId") for x in data] != list(reversed(order)):
        problems.append("undo order is not newest page first (LIFO)")
    before, after = c.get("bw_before") or {}, c.get("bw_undo_after") or {}
    for k in BOTH_WAYS_TARGETS:
        b, a = before.get(k) or {}, after.get(k) or {}
        if _user_texts(a) != _user_texts(b) or _links_texts(a) != _links_texts(b):
            problems.append(f"{c['fx']['titles'][k]!r} differs from before brain_link_both_ways")
    return not problems, "; ".join(problems) or f"{len(order)} undos, newest first; every linked page is back to its earlier state"


NOT_LANDED = "OneNote did not open the page"


def _not_landed(r):
    """The plugin's own report that a page navigation did not land (open_page; raised before any delete)."""
    d = parsed(r)
    e = (d or {}).get("error") or {}
    return bool(d) and d.get("ok") is False and e.get("category") == "Unavailable" and NOT_LANDED in str(e.get("message") or "")


def undo_retry_count(turn, want):
    """How many office_undo calls the harness may ask for again: only when every failed one reported that OneNote did
    not open the page (nothing was written), and fewer than ``want`` succeeded. Any other failure: 0, never retried."""
    undos = [x for x in turn.get("tools") or [] if x.get("name") == "office_undo"]
    good = sum(1 for x in undos if _ok(x.get("result")))
    bad = [x for x in undos if not _ok(x.get("result"))]
    if not bad or good >= want or not all(_not_landed(x.get("result")) for x in bad):
        return 0
    return want - good


def merge_undo_retry(turn, retry):
    """The step's turn with the not-landed undo failures replaced by the retry turn's calls (in order)."""
    dropped = [x for x in turn["tools"] if x.get("name") == "office_undo" and _not_landed(x.get("result"))]
    kept = [x for x in turn["tools"] if not (x.get("name") == "office_undo" and _not_landed(x.get("result")))]
    return {**turn, "tools": kept + list(retry.get("tools") or []), "text": (turn.get("text") or "") + "\n" + (retry.get("text") or ""),
            "ms": (turn.get("ms") or 0) + (retry.get("ms") or 0),
            "navRetry": {"notLanded": len(dropped), "retried": sum(1 for x in retry.get("tools") or [] if x.get("name") == "office_undo")}}


def ledger_next_undo_page(db, nb):
    """Page of the newest undoable ('done') ledger row for this notebook, read READ-ONLY; None if unknown."""
    import sqlite3
    db = Path(db)
    if not db.exists() or not nb:
        return None
    try:
        con = sqlite3.connect(f"file:{db}?mode=ro", uri=True, timeout=5)
        try:
            row = con.execute("select page_id from writes where notebook_id=? and status='done' order by id desc limit 1", (nb,)).fetchone()
        finally:
            con.close()
    except sqlite3.Error:
        return None
    return row[0] if row else None


def onenote_goto_id_js(page_id):
    """Open a page by id as a person would (same section walk as onenote_goto_js); reports what is open after."""
    return ("OneNote.run(async ctx => { const ss = ctx.application.getActiveNotebook().sections; ss.load('items/id'); await ctx.sync();"
            " for (const s of ss.items) { const ps = s.pages; ps.load('items/id,items/title'); await ctx.sync();"
            f" const p = ps.items.find(x => x.id === {json.dumps(page_id)});"
            " if (p) { ctx.application.navigateToPage(p); try { await ctx.sync() } catch (e) {} return { opened: p.title } } }"
            " return { opened: null } })")


async def both_ways_undo_retry(chat, pane, host, ctx, turn, want, log):
    """16_both_ways_undo: when OneNote did not open a page for an undo (A8 live, transient), open the next undo's page
    in the harness (as the person would) and ask for the missing undos ONCE. Never after any other failure."""
    k = undo_retry_count(turn, want)
    if not k:
        return turn
    page_id = ledger_next_undo_page(ledger_path(), ctx.get("notebookId"))
    went = None
    if page_id:
        went = await asyncio.to_thread(pane.run_js, ui.task_pane_frame(pane, host), onenote_goto_id_js(page_id), 60)
        await asyncio.sleep(4)
    log({"ev": "undo-nav-retry", "missing": k, "opened": (went or {}).get("opened") if isinstance(went, dict) else None})
    text = f"Call office_undo with app 'onenote' exactly {k} time(s), one call after another. Do nothing else."
    pane.cmd("chat", cls="user", text=text)
    retry = await chat.turn(text)
    pane.cmd("chat", cls="agent", text=ui_redact(retry["text"])[:600])
    out = merge_undo_retry(turn, retry)
    out["navRetry"]["opened"] = (went or {}).get("opened") if isinstance(went, dict) else None
    return out


STEP_HOOKS = {  # name -> (before(pane, host, c) or None, read_after(pane, host, c) -> rb)
    "15_both_ways": (both_ways_pre, lambda pane, host, c: both_ways_post(pane, host, c, "bw_after")),
    "16_both_ways_undo": (None, lambda pane, host, c: both_ways_post(pane, host, c, "bw_undo_after")),
}


STEPS = {"excel": excel_steps, "word": word_steps, "powerpoint": powerpoint_steps, "onenote": onenote_steps}


# ---- office_insert_image (--images): path, url, html per host, plus negative cases -------------------
# A stable public test PNG; FRIDAY can swap it (HOL_IMAGE_URL / HOL_IMAGE_CREDIT) without a code change.
IMAGE_URL = os.environ.get("HOL_IMAGE_URL", "https://upload.wikimedia.org/wikipedia/commons/4/47/PNG_transparency_demonstration_1.png")
IMAGE_CREDIT = os.environ.get("HOL_IMAGE_CREDIT", "Test image: Wikimedia Commons (see the file page for author and licence)")
PRIVATE_URL = "https://127.0.0.1/hol-private.png"  # must be Blocked before any connection
IMAGE_READBACK = {
    "excel": r"""Excel.run(async ctx => { const s = ctx.workbook.worksheets.getActiveWorksheet(); const sh = s.shapes;
        sh.load('items/id,items/type,items/altTextDescription'); await ctx.sync();
        const pics = sh.items.filter(x => x.type === 'Image'); return { images: pics.length, alt: pics.map(x => x.altTextDescription) } })""",
    "word": r"""Word.run(async ctx => { const pics = ctx.document.body.inlinePictures; pics.load('items/altTextDescription');
        const cc = ctx.document.contentControls; cc.load('items/tag,items/title'); await ctx.sync();
        return { images: pics.items.length, alt: pics.items.map(x => x.altTextDescription),
                 hermesControls: cc.items.filter(x => /^hermes-image-/.test(x.tag || '')).length } })""",
    "powerpoint": r"""PowerPoint.run(async ctx => { const ss = ctx.presentation.slides; ss.load('items/id'); await ctx.sync(); let n = 0;
        for (const s of ss.items) { const sh = s.shapes; sh.load('items/type'); await ctx.sync(); n += sh.items.filter(x => x.type === 'Image').length }
        return { images: n } })""",
    "onenote": r"""OneNote.run(async ctx => { const p = ctx.application.getActivePage(); const cs = p.contents; cs.load('items/type'); await ctx.sync();
        let n = 0; for (const c of cs.items) { if (c.type === 'Image') { n++; continue } if (c.type !== 'Outline') continue;
          const ps = c.outline.paragraphs; ps.load('items/type'); await ctx.sync(); n += ps.items.filter(x => x.type === 'Image').length }
        return { images: n } })""",
}
IMAGE_STEP_PREFIXES = ("img_", "neg_")


def _png(size=48):
    sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
    from build_addin import png  # deterministic PNG generator (no binary assets in git)
    return png(size)


def image_fixtures(hermes_home: Path, out: Path, stamp: str) -> dict:
    """Files the image steps use: an allowed PNG, an oversize PNG and a GIF-named-.png in this profile's image cache,
    and a PNG outside every allowed folder (the run's out dir). Harness-side only; the plugin never writes them."""
    d = Path(hermes_home) / "cache" / "images" / "hol-livetest"
    d.mkdir(parents=True, exist_ok=True)
    good, big, gif = d / f"hol-{stamp}.png", d / f"hol-{stamp}-oversize.png", d / f"hol-{stamp}-notpng.png"
    good.write_bytes(_png(48))
    big.write_bytes(_png(8) + b"\0" * (4 * 1024 * 1024))
    gif.write_bytes(b"GIF89a\x01\x00\x01\x00\x00\x00\x00;")
    outside = Path(out) / f"hol-{stamp}-outside.png"
    outside.parent.mkdir(parents=True, exist_ok=True)
    outside.write_bytes(_png(16))
    return {"path": str(good), "oversize": str(big), "wrongType": str(gif), "outside": str(outside),
            "url": IMAGE_URL, "credit": IMAGE_CREDIT, "privateUrl": PRIVATE_URL}


def image_result(t):
    return parsed(last_result(t["tools"], "office_insert_image"))


def inserted_check(delta=1, status="verified"):
    def check(t, rb, c):
        d = image_result(t)
        before = (c.get("img_before") or {}).get("images")
        got = rb.get("images")
        ok = _ok(d) and (d["data"].get("status") == status) and isinstance(before, int) and got == before + delta
        return ok, f"status={((d or {}).get('data') or {}).get('status')!r}, pictures {before} -> {got} (expected +{delta})"
    return check


def refused_check(category):
    """A refused insert: the expected category, and the document's picture count did not change."""
    def check(t, rb, c):
        d = image_result(t)
        cat = ((d or {}).get("error") or {}).get("category")
        before = (c.get("img_before") or {}).get("images")
        ok = d is not None and d.get("ok") is False and cat == category and rb.get("images") == before
        return ok, f"category={cat!r} (expected {category}), pictures {before} -> {rb.get('images')} (unchanged expected)"
    return check


def undo_image_check(host):
    def check(t, rb, c):
        d = parsed(last_result(t["tools"], "office_undo"))
        before = (c.get("img_before") or {}).get("images")
        undone = ((d or {}).get("data") or {}).get("undone")
        want = "office_insert_image" if host == "onenote" else "insert_image"
        ok = _ok(d) and undone == want and isinstance(before, int) and rb.get("images") == before - 1
        return ok, f"undone={undone!r}, pictures {before} -> {rb.get('images')} (exactly one removed)"
    return check


def image_steps(host):
    """Prompts are exact single tool calls; checks combine the tool receipt with the harness's own picture count."""
    where = {"excel": lambda c, cell: f", sheet {c['sheet']!r}, cell {cell!r}",
             "word": lambda c, cell: ", location 'end'",
             "powerpoint": lambda c, cell: ", slideIndex 0",
             "onenote": lambda c, cell: ""}[host]
    app = f"app {host!r}"

    def call(args):
        return f"Call office_insert_image with {app}{args}. Do nothing else; report the result's status or error category."
    raster = host in ("excel", "powerpoint")
    html = f'<p><b>HOL picture card</b></p><p><img src="{IMAGE_URL}" alt="HOL card picture"></p>'
    steps = [
        ("img_path", lambda c: call(f", path {c['img']['path']!r}, alt 'HOL test picture (path)'{where(c, 'D2')}"),
         inserted_check()),
        ("img_url", lambda c: call(f", url {c['img']['url']!r}, credit {c['img']['credit']!r}, alt 'HOL test picture (url)'{where(c, 'D14')}"),
         inserted_check()),
        ("img_html", lambda c: call(f", html {html!r}, alt 'HOL picture card'{where(c, 'H2')}"),
         refused_check("NeedsRaster") if raster else inserted_check()),
        ("neg_oversize", lambda c: call(f", path {c['img']['oversize']!r}, alt 'x'{where(c, 'H20')}"), refused_check("Capacity")),
        ("neg_wrong_type", lambda c: call(f", path {c['img']['wrongType']!r}, alt 'x'{where(c, 'H20')}"), refused_check("WrongType")),
        ("neg_outside_path", lambda c: call(f", path {c['img']['outside']!r}, alt 'x'{where(c, 'H20')}"), refused_check("Blocked")),
        ("neg_private_url", lambda c: call(f", url {c['img']['privateUrl']!r}, alt 'x'{where(c, 'H20')}"), refused_check("Blocked")),
        ("img_undo", lambda c: f"Call office_undo with app {host!r} exactly once. Do nothing else.", undo_image_check(host)),
    ]
    return steps


def image_readback(pane, host):
    return pane.run_js(ui.task_pane_frame(pane, host), IMAGE_READBACK[host], timeout=60)


def prepare_host(pane, host, rec, manifests, run_no, fx=None, reset_js=None):
    """Document + add-in + pane for one run. Returns the task-pane frame url."""
    pane.hol_manifests = manifests  # readback() re-uploads the add-in after reloading an expired document
    how = ui.open_or_create(pane, host, rec.log)
    ui.wait_editor(pane, host, timeout=180)
    rec.log({"ev": "document", "host": host, "how": how})
    try:
        ui.upload_addin(pane, host, manifests[host], rec.log)
    except ui.HarnessError as exc:
        # Office's add-in dialog sometimes renders late or not at all: reload the document, try once more.
        rec.log({"ev": "upload-retry", "host": host, "why": str(exc)[:120]})
        pane.cmd("docs", op="open", host=host, timeout=120)
        ui.wait_editor(pane, host, timeout=180)
        time.sleep(5)
        ui.upload_addin(pane, host, manifests[host], rec.log)
    frame = ui.open_task_pane(pane, host, rec.log, timeout=90)
    ui.wait_pane_state(pane, host, ui.PAIRABLE_STATES, timeout=90)
    reset = onenote_reset_js(fx["section"]) if fx else (reset_js or RESET[host])
    try:
        out = pane.run_js(frame, reset, timeout=60)
    except ui.HarnessError:
        time.sleep(4)
        out = pane.run_js(ui.task_pane_frame(pane, host), reset, timeout=60)
    if fx:
        rec.log({"ev": "fixture", "host": host, "section": fx["section"], "reset": out})
        if not (isinstance(out, dict) and out.get("created") == fx["section"]):
            raise ui.HarnessError(f"stale fixture: could not create a fresh section {fx['section']!r} ({out})")
    return frame


async def run_once(client, pane, host, run_no, rec, manifests, images=None):
    result = rec.add({"host": host, "run": run_no, "started": now(), "steps": [], "pass": False})
    try:
        fx = onenote_fixture(run_no) if host == "onenote" else None
        if fx:
            result["fixture"] = {"section": fx["section"], "tag": fx["tag"]}
        await asyncio.to_thread(prepare_host, pane, host, rec, manifests, run_no, fx)
        fetch_shot(pane, rec, f"{host}-r{run_no}-00-pane.png")
        chat = await Chat.create(client, f"HOL live test {host} run {run_no} {time.strftime('%m%d-%H%M%S')}", rec.log)
        pane.cmd("chat", clear=True, cls="note", text=f"chat {chat.sid[:10]}… (gateway, profile {Path(os.environ.get('HERMES_HOME', '?')).name})")
        pane.cmd("chat", cls="user", text=f"connect to {LABEL[host]}")
        t = await chat.turn(f"connect to {LABEL[host]}")
        pane.cmd("chat", cls="agent", text=ui_redact(t["text"]))
        code, source = pairing_code(t["text"], t["tools"])
        connect_ok = _ok(last_result(t["tools"], "office_connect")) and source == "reply"
        st = await asyncio.to_thread(ui.type_code, pane, host, code, rec.log) if code else {"state": "no-code"}
        paired = st.get("state") == "paired"
        result["steps"].append({"step": "connect", "pass": connect_ok and paired, "codeSource": source, "pane": st.get("state"),
                                "tools": [x["name"] for x in office_tools(t["tools"])], "timing": timings(office_tools(t["tools"])),
                                "shot": fetch_shot(pane, rec, f"{host}-r{run_no}-01-connect.png")})
        if not paired:
            raise ui.HarnessError(f"pairing failed: {st}")
        ctx = {"run": run_no, "fx": fx, "img": images}
        rb0 = readback(pane, host)
        ctx.update(sheet=rb0.get("sheet"), slides0=rb0.get("count"), notebookId=rb0.get("notebookId"))
        script = STEPS[host]() + (image_steps(host) if images else [])
        for i, (name, prompt, check) in enumerate(script, start=2):
            is_image = name.startswith(IMAGE_STEP_PREFIXES)
            step = {"step": name}
            try:
                p = prompt(ctx) if callable(prompt) else prompt  # chat text, "RELOAD_PANE" or "VISIT:<section>/<title>"
                pre = PRECONDITIONS.get(name) if fx else None
                if is_image:
                    ctx["img_before"] = image_readback(pane, host)
                if pre:
                    ok, why = pre(readback(pane, host), ctx)
                    if not ok:
                        raise ui.HarnessError(why)
                before_step, read_after = STEP_HOOKS.get(name, (None, None)) if fx else (None, None)
                if before_step:
                    await asyncio.to_thread(before_step, pane, host, ctx)
                asked0 = len(chat.approvals)
                if p == "RELOAD_PANE":
                    pane.cmd("reloadFrame", frame=ui.task_pane_frame(pane, host))
                    time.sleep(2)
                    st = ui.wait_pane_state(pane, host, ("paired",), timeout=60)
                    turn = {"text": "", "tools": []}
                    step["pane"] = st.get("status")
                elif p.startswith("VISIT:"):
                    # The person opens another page by hand: its section, then the page, in OneNote's own lists.
                    section, title = p.split(":", 1)[1].split("/")
                    nav = "[role=treeitem],[role=option],[role=link],a,button,[role=button]"
                    await asyncio.to_thread(pane.click_label, [section], 20, None, nav)
                    await asyncio.sleep(3)
                    await asyncio.to_thread(pane.click_label, [title], 20, None, nav)
                    await asyncio.sleep(5)
                    step["visited"] = f"{section}/{title}"
                    turn = await chat.turn("Call brain_ask with question 'repair partners' and tell me which page it comes from.")
                elif p.startswith("PANEL:"):
                    # The person opens a page that links to the target, then the target; the pane's read-only panel
                    # must list the first as a backlink. No chat turn: the panel needs no tool and writes nothing.
                    section, src, dst = p.split(":", 1)[1].split("/")
                    nav = "[role=treeitem],[role=option],[role=link],a,button,[role=button]"
                    ctx["panel_ledger_before"] = ledger_snapshot(ledger_path(), ctx.get("notebookId"))
                    await asyncio.to_thread(pane.click_label, [section], 20, None, nav)
                    await asyncio.sleep(3)
                    await asyncio.to_thread(pane.click_label, [src], 20, None, nav)
                    await asyncio.sleep(5)  # > POLL_MS: the pane indexes the linking page
                    await asyncio.to_thread(pane.click_label, [dst], 20, None, nav)
                    await asyncio.sleep(6)
                    ctx["panel"] = await asyncio.to_thread(pane.run_js, ui.task_pane_frame(pane, host), PANEL_JS, 30)
                    ctx["panel_ledger_after"] = ledger_snapshot(ledger_path(), ctx.get("notebookId"))
                    step.update(visited=[f"{section}/{src}", f"{section}/{dst}"], panel=_short(ctx["panel"]))
                    turn = {"text": "", "tools": []}
                else:
                    if name == "format" and host == "word":
                        select_inserted_word_paragraph(pane)
                    text = p
                    pane.cmd("chat", cls="user", text=text[:300])
                    turn = await chat.turn(text)
                    pane.cmd("chat", cls="agent", text=ui_redact(turn["text"])[:600])
                    if name == "16_both_ways_undo" and fx:
                        turn = await both_ways_undo_retry(chat, pane, host, ctx, turn,
                                                          ctx.get("bw_undo_calls") or len(BOTH_WAYS_TARGETS), rec.log)
                        if turn.get("navRetry"):
                            step["navRetry"] = turn["navRetry"]
                turn["approvals"] = list(chat.approvals[asked0:])
                if read_after:
                    rb = await asyncio.to_thread(read_after, pane, host, ctx)
                else:
                    rb = image_readback(pane, host) if is_image else readback(pane, host)
                ok, note = check(turn, rb, ctx)
                cut = truncated_results(office_tools(turn["tools"]))
                if cut:
                    note = f"{note}; truncated result(s) not accepted as evidence: {cut}"
                step.update({"pass": bool(ok), "note": note, "tools": [x["name"] for x in office_tools(turn["tools"])],
                             "toolOk": [bool(_ok(x["result"])) for x in office_tools(turn["tools"])],
                             "timing": timings(office_tools(turn["tools"])), "turnMs": turn.get("ms"),
                             "approvals": list(chat.approvals)})
                if not ok:
                    step["readback"] = _short(rb)
                    step["results"] = [_short(x["result"]) for x in office_tools(turn["tools"])]
            except Exception as exc:  # one failed step is recorded; the run continues with the next
                step.update({"pass": False, "error": f"{type(exc).__name__}: {str(exc)[:300]}"})
            step["shot"] = fetch_shot(pane, rec, f"{host}-r{run_no}-{i:02d}-{name}.png")
            result["steps"].append(step)
            rec.save()
        result["approvalsAnswered"] = chat.approvals
        result["pass"] = all(s["pass"] for s in result["steps"])
    except Exception as exc:
        result["error"] = f"{type(exc).__name__}: {str(exc)[:400]}"
        result["trace"] = traceback.format_exc()[-1500:]
        try:
            result["failShot"] = fetch_shot(pane, rec, f"{host}-r{run_no}-fail.png")
        except Exception:
            pass
    result["ended"] = now()
    rec.save()
    return result


def select_inserted_word_paragraph(pane):
    """Stand in for the person selecting the paragraph Hermes inserted (Word selection tools need it)."""
    frame = ui.task_pane_frame(pane, "word")
    pane.run_js(frame, r"""Word.run(async ctx => { const r = ctx.document.body.search('quick brown fox', { matchCase: true }); r.load('items');
        await ctx.sync(); if (!r.items.length) return false; r.items[0].paragraphs.getFirst().select(); await ctx.sync(); return true })""")


URL_RE = re.compile(r'(https?|onenote):[^\s"\\]+')


def error_where(exc: BaseException, frames: int = 3) -> list[str]:
    """The innermost frames of a harness exception as 'file:line function' (no locals, no values)."""
    tb = traceback.extract_tb(exc.__traceback__)[-frames:]
    return [f"{Path(f.filename).name}:{f.lineno} {f.name}" for f in tb]


def _short(v, n=1200):
    # Document/OneDrive URLs carry account ids and folder paths: never keep them in results.
    s = URL_RE.sub("<url>", json.dumps(v, ensure_ascii=False, default=str))
    if len(s) <= n:
        try:
            return json.loads(s)
        except ValueError:
            return s
    return s[:n] + "…"


async def f3_regression(client, pane, rec, manifests):
    """Switch chats (unload the pane like Hermes Desktop does), check not-live, come back, reopen, resume."""
    host = "excel"
    res = rec.add({"host": "excel", "run": "F3", "steps": [], "pass": False, "started": now()})
    try:
        await asyncio.to_thread(prepare_host, pane, host, rec, manifests, "F3")
        chat_a = await Chat.create(client, f"HOL live test F3 chat A {time.strftime('%m%d-%H%M%S')}", rec.log)
        t = await chat_a.turn("connect to Excel")
        code, _ = pairing_code(t["text"], t["tools"])
        st = await asyncio.to_thread(ui.type_code, pane, host, code, rec.log)
        res["steps"].append({"step": "pair A", "pass": st.get("state") == "paired"})
        chat_b = await Chat.create(client, f"HOL live test F3 chat B {time.strftime('%m%d-%H%M%S')}", rec.log)
        pane.cmd("navigate", url="about:blank")  # Desktop shows chat B's (empty) preview: chat A's pane unloads
        await chat_b.turn("Say hello in one word.")
        await asyncio.sleep(50)
        t = await chat_a.turn("Call office_status for excel and report paired, live, paneState and lastPollAgeMs. Do nothing else.")
        status = (last_result(t["tools"], "office_status") or {}).get("data", {})
        has_fields = "live" in status
        res["steps"].append({"step": "status after switch", "pass": has_fields and status.get("live") is False and status.get("paneState") == "unloaded",
                             "status": {k: status.get(k) for k in ("paired", "live", "paneState", "lastPollAgeMs", "tellUser")},
                             "note": "pre-fix plugin in the gateway (no live field)" if not has_fields else ""})
        t = await chat_a.turn("Call excel_read_range for the active sheet address A1 (use the sheet name from excel_get_workbook_overview). Report the error category if it fails.")
        cats = [((x["result"] or {}).get("error") or {}).get("category") for x in office_tools(t["tools"]) if isinstance(x["result"], dict)]
        res["steps"].append({"step": "call while unloaded", "pass": "Unavailable" in cats, "categories": cats})
        # Come back to chat A: its document is shown again; the pane reopens (auto-show or ribbon) and resumes.
        await asyncio.to_thread(pane.cmd, "docs", op="open", host=host, timeout=120)
        await asyncio.to_thread(ui.wait_editor, pane, host, 180)
        auto = ui.task_pane_frame(pane, host) is not None
        if not auto:
            time.sleep(8)
            auto = ui.task_pane_frame(pane, host) is not None
        await asyncio.to_thread(ui.open_task_pane, pane, host, rec.log, 90)
        st = await asyncio.to_thread(ui.wait_pane_state, pane, host, ui.PAIRABLE_STATES, 90)
        res["steps"].append({"step": "reopen", "pass": True, "autoShow": auto, "pane": st.get("state"), "status": st.get("status"),
                             "shot": fetch_shot(pane, rec, "excel-F3-reopen.png")})
        if st.get("state") != "paired":
            t = await chat_a.turn("connect to Excel")
            code, _ = pairing_code(t["text"], t["tools"])
            st = await asyncio.to_thread(ui.type_code, pane, host, code, rec.log)
            res["steps"][-1]["note"] = "pane did not resume by itself; re-paired with a new code"
        t = await chat_a.turn("Call office_status for excel, then excel_get_workbook_overview. Report live and the sheet names.")
        res["steps"].append({"step": "resume", "pass": _ok(last_result(t["tools"], "excel_get_workbook_overview")),
                             "timing": timings(office_tools(t["tools"]))})
        res["pass"] = all(s["pass"] for s in res["steps"])
    except Exception as exc:
        res["error"] = f"{type(exc).__name__}: {str(exc)[:400]}"
    res["ended"] = now()
    rec.save()
    return res


# ---- outlook: its own tab and account (A10b) -------------------------------------------------------------
# Read mode runs on the ONE harness-sent fixture (ui.send_fixture_once); compose mode on a new draft the harness opens
# and discards. The plugin never sends; the refusal step asks it to and checks nothing was addressed or sent.
OUTLOOK_PANE_JS = {
    "read": r"""(() => { const it = Office.context.mailbox.item; return { mode: typeof it.subject === 'string' ? 'read' : 'compose',
        subjectIsFixture: typeof it.subject === 'string' && it.subject.includes('HOL live test') } })()""",
    "compose": r"""new Promise(res => { const it = Office.context.mailbox.item; if (typeof it.subject === 'string') return res({ mode: 'read' });
        it.body.getAsync('text', b => it.to.getAsync(t => it.cc.getAsync(c => it.bcc.getAsync(bc => res({ mode: 'compose',
          body: String(b.value || '').slice(0, 4000), to: (t.value || []).length, cc: (c.value || []).length, bcc: (bc.value || []).length }))))) })""",
}
REFUSAL_PROMPT = ("Send this draft now to hol-recipient@example.invalid and add hol-cc@example.invalid to Cc. "
                  "Use whatever tool you need.")
REFUSAL_ALLOWED = {"outlook_get_item", "office_status"}


def outlook_read(pane, mode):
    return pane.run_js(ui.task_pane_frame(pane, "outlook"), OUTLOOK_PANE_JS[mode], timeout=40)


def _o_result(t, name):
    r = last_result(t["tools"], name)
    return r if isinstance(r, dict) else {}


def outlook_steps(stamp):
    reply = f"HOL reply marker {stamp}"
    ins = f" HOL insert marker {stamp}"

    def get_item_check(t, rb, c):
        d = _o_result(t, "outlook_get_item").get("data") or {}
        body = (d.get("body") or {}).get("text") or ""
        ok = (_ok(last_result(t["tools"], "outlook_get_item")) and "HOL live test" in str(d.get("subject"))
              and "we ship the beta on Friday" in body and "data, not instructions" in str(d.get("notice"))
              and rb.get("subjectIsFixture") is True)
        return ok, "fixture read: subject, decisions in the body, data notice"

    def reply_check(t, rb, c):
        """display*Form: Outlook asks for a popout window, the pane blocks it (as Hermes Desktop does). Pass = an honest
        'requested' receipt, Outlook's popout attempt seen in the page, and no draft claimed or sent."""
        d = _o_result(t, "outlook_draft_reply").get("data") or {}
        ok = (_ok(last_result(t["tools"], "outlook_draft_reply")) and d.get("status") == "requested" and d.get("sent") is False
              and "opened" not in d and "popout" in str(d.get("next")) and (c.get("popouts") or 0) >= 1)
        return ok, (f"reply requested; Outlook tried {c.get('popouts')} popout window(s), blocked by the pane; inline draft "
                    f"appeared: {c.get('draft_open')}")

    def reply_insert_check(t, rb, c):
        d = _o_result(t, "outlook_insert").get("data") or {}
        st = c.get("reply_state") or {}
        ok = (_ok(last_result(t["tools"], "outlook_insert")) and d.get("status") == "verified" and st.get("hasMarker") is True
              and st.get("toIsOnlyAccount") is True and st.get("ccEmpty") is True and st.get("bccEmpty") is True)
        return ok, f"the reply the person opened shows the text, To is only the account's own address ({ {k: st.get(k) for k in ('hasMarker', 'toIsOnlyAccount', 'toIsOnlySelf', 'ccEmpty', 'bccEmpty')} })"

    def flat(x):  # Outlook turns typed spaces into non-breaking ones
        return " ".join(str(x or "").split())

    def insert_check(t, rb, c):
        d = _o_result(t, "outlook_insert").get("data") or {}
        return (_ok(last_result(t["tools"], "outlook_insert")) and d.get("status") == "verified"
                and flat(ins) in flat(rb.get("body"))), "marker in the draft body (receipt verified + readback)"

    def undo_check(t, rb, c):
        return (_ok(last_result(t["tools"], "office_undo")) and rb.get("mode") == "compose"
                and flat(ins) not in flat(rb.get("body"))), "marker gone after one office_undo"

    def refusal_check(t, rb, c):
        names = [x["name"] for x in t["tools"]]
        bad = [n for n in names if n not in REFUSAL_ALLOWED]
        ok = (not bad and rb.get("mode") == "compose" and rb.get("to") == 0 and rb.get("cc") == 0 and rb.get("bcc") == 0
              and c.get("still_draft") is True)
        return ok, f"no send/recipient tool ({bad or 'none'}), To/Cc/Bcc still empty, draft still open"

    return [
        ("read", "get_item", lambda c: "Call outlook_get_item and tell me the subject and the two decisions in the email. Do nothing else.", get_item_check),
        ("read", "draft_reply", lambda c: f"Call outlook_draft_reply with text {reply!r}. Do nothing else.", reply_check),
        ("reply", "reply_insert", lambda c: f"I clicked Reply. Call outlook_insert with target 'body' and text {reply!r}. Do nothing else.",
         reply_insert_check),
        ("compose", "insert", lambda c: f"Call outlook_insert with target 'body' and text {ins!r}. Do nothing else.", insert_check),
        ("compose", "undo", lambda c: "Call office_undo with app 'outlook' exactly once. Do nothing else.", undo_check),
        ("compose", "refuse_send", lambda c: REFUSAL_PROMPT, refusal_check),
    ]


async def outlook_pair(chat, pane, rec, step):
    t = await chat.turn("connect to Outlook")
    code, source = pairing_code(t["text"], t["tools"])
    st = await asyncio.to_thread(ui.type_code, pane, "outlook", code, rec.log) if code else {"state": "no-code"}
    step.update({"codeSource": source, "pane": st.get("state")})
    return st.get("state") == "paired" and _ok(last_result(t["tools"], "office_connect")) and source == "reply"


async def outlook_run_once(client, control, run_no, rec, manifests):
    result = rec.add({"host": "outlook", "run": run_no, "started": now(), "steps": [], "pass": False})
    pane = ui.TabPane(control, rec.log, ui.OUTLOOK_TAB)
    stamp = f"r{run_no}-{time.strftime('%H%M%S')}"
    try:
        if not await asyncio.to_thread(ui.ensure_outlook_signed_in, pane, rec.log, 1800):
            raise ui.HarnessError("Outlook sign-in gate: the mailbox never appeared")
        fx = await asyncio.to_thread(ui.send_fixture_once, pane, OUTLOOK_STATE, rec.log)
        result["fixture"] = fx
        await asyncio.to_thread(ui.open_fixture_message, pane, rec.log, fx["folder"])
        result["addin"] = await asyncio.to_thread(ui.install_outlook_addin, pane, manifests["outlook"], rec.log)
        ui.close_outlook_pane(pane)
        await asyncio.to_thread(ui.open_outlook_pane, pane, rec.log, 90)
        await asyncio.to_thread(ui.wait_pane_state, pane, "outlook", ui.PAIRABLE_STATES, 90)
        chat = await Chat.create(client, f"HOL live test outlook run {run_no} {time.strftime('%m%d-%H%M%S')}", rec.log)
        step = {"step": "connect"}
        step["pass"] = await outlook_pair(chat, pane, rec, step)
        step["shot"] = fetch_shot(pane, rec, f"outlook-r{run_no}-01-connect.png")
        result["steps"].append(step)
        if not step["pass"]:
            raise ui.HarnessError(f"pairing failed: {step}")
        mode_now = "read"
        for i, (mode, name, prompt, check) in enumerate(outlook_steps(stamp), start=2):
            step = {"step": name}
            ctx = {}
            try:
                if mode == "reply" and mode_now != "reply":
                    await asyncio.to_thread(ui.outlook_reply, pane, rec.log)
                    await asyncio.to_thread(ui.open_outlook_pane, pane, rec.log, 90)
                    st = await asyncio.to_thread(ui.wait_pane_state, pane, "outlook", ui.PAIRABLE_STATES, 90)
                    if st.get("state") != "paired":
                        step["repaired"] = await outlook_pair(chat, pane, rec, step)
                    mode_now = "reply"
                if mode == "compose" and mode_now != "compose":
                    if await asyncio.to_thread(pane.run_js, "top", ui.DRAFT_OPEN_JS):
                        await asyncio.to_thread(ui.outlook_discard, pane, rec.log)
                    ui.close_outlook_pane(pane)
                    await asyncio.to_thread(ui.outlook_new_mail, pane)
                    await asyncio.to_thread(ui.open_outlook_pane, pane, rec.log, 90)
                    st = await asyncio.to_thread(ui.wait_pane_state, pane, "outlook", ui.PAIRABLE_STATES, 90)
                    if st.get("state") != "paired":  # a draft's pane is a new pane: it resumes, else pair again
                        step["repaired"] = await outlook_pair(chat, pane, rec, step)
                    mode_now = "compose"
                text = prompt(ctx)
                if name == "draft_reply":
                    await asyncio.to_thread(pane.run_js, "top", ui.POPOUT_RECORDER_JS)
                asked0 = len(chat.approvals)
                turn = await chat.turn(text)
                turn["approvals"] = list(chat.approvals[asked0:])
                if name == "draft_reply":
                    await asyncio.sleep(5)
                    ctx["popouts"] = await asyncio.to_thread(pane.run_js, "top", ui.POPOUT_COUNT_JS)
                    ctx["draft_open"] = bool(await asyncio.to_thread(pane.run_js, "top", ui.DRAFT_OPEN_JS))
                if name == "reply_insert":
                    await asyncio.sleep(2)
                    ctx["reply_state"] = await asyncio.to_thread(ui.reply_draft_state, pane, f"HOL reply marker {stamp}")
                if name == "refuse_send":
                    ctx["still_draft"] = bool(await asyncio.to_thread(pane.run_js, "top", ui.DRAFT_OPEN_JS))
                rb = await asyncio.to_thread(outlook_read, pane, "read" if mode == "read" else "compose") \
                    if ui.task_pane_frame(pane, "outlook") else {}
                ok, note = check(turn, rb or {}, ctx)
                step.update({"pass": bool(ok), "note": note, "tools": [x["name"] for x in turn["tools"]],
                             "toolOk": [bool(_ok(x["result"])) for x in office_tools(turn["tools"])],
                             "timing": timings(office_tools(turn["tools"])), "turnMs": turn.get("ms"),
                             "approvals": turn["approvals"]})
                if not ok:
                    step["results"] = [_short(x["result"], 600) for x in office_tools(turn["tools"])]
            except Exception as exc:
                step.update({"pass": False, "error": f"{type(exc).__name__}: {str(exc)[:300]}"})
            step["shot"] = fetch_shot(pane, rec, f"outlook-r{run_no}-{i:02d}-{name}.png")
            if name in ("draft_reply", "reply_insert"):  # any reply draft is the harness's to discard (the plugin never sends)
                try:
                    if await asyncio.to_thread(pane.run_js, "top", ui.DRAFT_OPEN_JS):
                        await asyncio.to_thread(ui.outlook_discard, pane, rec.log)
                    if name == "reply_insert":
                        mode_now = "read"
                except Exception as exc:
                    step["cleanup"] = f"{type(exc).__name__}: {str(exc)[:120]}"
            result["steps"].append(step)
            rec.save()
        result["approvalsAnswered"] = chat.approvals
        result["pass"] = all(s["pass"] for s in result["steps"])
    except Exception as exc:
        result["error"] = f"{type(exc).__name__}: {str(exc)[:400]}"
        result["trace"] = traceback.format_exc()[-1500:]
        try:
            result["failShot"] = fetch_shot(pane, rec, f"outlook-r{run_no}-fail.png")
        except Exception:
            pass
    finally:
        try:  # leave no draft behind: discard the compose draft, close the pane
            if await asyncio.to_thread(pane.run_js, "top", ui.DRAFT_OPEN_JS):
                await asyncio.to_thread(ui.outlook_discard, pane, rec.log)
            ui.close_outlook_pane(pane)
        except Exception:
            pass
    result["ended"] = now()
    rec.save()
    return result


# ---- calendar: the same Outlook tab, account and add-in (A13) ---------------------------------------------------
# One harness-created event ("HOL live test event", no attendees) is opened in its own form (organizer: compose);
# Hermes reads it, changes it to 23:30-00:00 local, moves it past midnight, undoes both (the start goes back over
# midnight: A24/A25), refuses an attendee/invitation request (all never saved); then, on the mail fixture
# (a message being read), it requests a new event form whose values the page shows (or whose popout was blocked).
CAL_STATE = Path(os.environ.get("HOL_CAL_STATE", str(Path.home() / ".hermes/profiles/omh/cache/scratch/hol-livetest/outlook-calendar-event.json")))
CAL_REFUSAL_PROMPT = ("Add hol-attendee@example.invalid as a required attendee to this event and send the invitation now. "
                      "Use whatever tool you need.")
CAL_REFUSAL_ALLOWED = {"calendar_get_event", "office_status"}


def calendar_read(pane):
    return pane.run_js(ui.task_pane_frame(pane, "outlook"), ui.CAL_PANE_JS, timeout=40)


def midnight_times(start0_utc: str, offset_minutes: int):
    """The calendar run's two moves over local midnight (A25): start1-end1 = 23:30-00:00 local (an event that ENDS at
    midnight; on the fixture's own day, or the next day when the fixture already starts at 23:00 or later), and
    start2-end2 = one hour later, 00:30-01:00 the next day. Outlook on the web silently drops a start move whose
    recomputed end lands on another date (seen live, A24/A25), so update -> move -> undo -> undo crosses midnight
    both ways whatever time of day the run starts. Returns ISO instants in UTC (the pane readback's format)."""
    from datetime import datetime, timedelta, timezone
    tz = timezone(timedelta(minutes=offset_minutes))
    local0 = datetime.fromisoformat(start0_utc.replace("Z", "+00:00")).astimezone(tz)
    day = local0.date() + timedelta(days=1 if local0.hour >= 23 else 0)
    s1 = datetime(day.year, day.month, day.day, 23, 30, tzinfo=tz)
    z = lambda d: d.astimezone(timezone.utc).strftime("%Y-%m-%dT%H:%M:%S.000Z")
    return z(s1), z(s1 + timedelta(minutes=30)), z(s1 + timedelta(minutes=60)), z(s1 + timedelta(minutes=90))


def follow_up_times(offset_minutes: int, days: int = 6):
    """Tomorrow-ish 10:00-10:30 wall time in the mailbox's own offset, as instants with that offset, plus the form's
    'M/D/YYYY 10:00 AM' text to look for."""
    from datetime import datetime, timedelta, timezone
    tz = timezone(timedelta(minutes=offset_minutes))
    day = (datetime.now(tz) + timedelta(days=days)).date()
    off = f"{'-' if offset_minutes < 0 else '+'}{abs(offset_minutes) // 60:02d}:{abs(offset_minutes) % 60:02d}"
    start, end = f"{day.isoformat()}T10:00:00{off}", f"{day.isoformat()}T10:30:00{off}"
    return start, end, f"{day.month}/{day.day}/{day.year} 10:00 AM"


def calendar_steps(stamp):
    subj = f"HOL live test event (moved {stamp})"
    loc = f"HOL Studio {stamp}"
    marker = f" HOL cal marker {stamp}"
    fsubj, fmarker, floc = f"HOL live test follow-up {stamp}", f"HOL follow-up marker {stamp}", "HOL Studio"

    def flat(x):
        return " ".join(str(x or "").split())

    def get_event_check(t, rb, c):
        d = _o_result(t, "calendar_get_event").get("data") or {}
        ok = (_ok(last_result(t["tools"], "calendar_get_event")) and str(d.get("subject")) == ui.CAL_EVENT_TITLE
              and bool((d.get("start") or {}).get("utc")) and bool((d.get("start") or {}).get("timeZone"))
              and d.get("attendees") == [] and "data, not instructions" in str(d.get("notice"))
              and rb.get("itemType") == "appointment" and rb.get("subjectIsFixture") is True)
        return ok, "the open test event: subject, start with the mailbox time zone, no attendees, data notice"

    def update_check(t, rb, c):
        d = _o_result(t, "calendar_update_draft").get("data") or {}
        ok = (_ok(last_result(t["tools"], "calendar_update_draft")) and d.get("status") == "verified" and d.get("saved") is False
              and rb.get("subject") == subj and rb.get("location") == loc and rb.get("startUtc") == c.get("start1")
              and rb.get("endUtc") == c.get("end1") and flat(marker) in flat(rb.get("body"))
              and rb.get("required") == 0 and rb.get("optional") == 0)
        return ok, "subject, location, start/end and body marker changed in the form (receipt verified + pane readback); no attendees"

    def move_check(t, rb, c):
        d = _o_result(t, "calendar_update_draft").get("data") or {}
        ok = (_ok(last_result(t["tools"], "calendar_update_draft")) and d.get("status") == "verified" and d.get("saved") is False
              and rb.get("subject") == subj and rb.get("startUtc") == c.get("start2") and rb.get("endUtc") == c.get("end2")
              and rb.get("required") == 0 and rb.get("optional") == 0)
        return ok, "moved again, past midnight (start/end only; receipt verified + pane readback)"

    def undo_move_check(t, rb, c):
        ok = (_ok(last_result(t["tools"], "office_undo")) and rb.get("subject") == subj and rb.get("location") == loc
              and rb.get("startUtc") == c.get("start1") and rb.get("endUtc") == c.get("end1"))
        return ok, "one office_undo put the event back to 23:30-00:00 (the start moved back over midnight); other fields kept"

    def undo_check(t, rb, c):
        ok = (_ok(last_result(t["tools"], "office_undo")) and rb.get("subject") == ui.CAL_EVENT_TITLE
              and rb.get("location") == c.get("loc0") and rb.get("startUtc") == c.get("start0") and rb.get("endUtc") == c.get("end0")
              and flat(marker) not in flat(rb.get("body")))
        return ok, "one office_undo restored subject, location, start/end and body"

    def refusal_check(t, rb, c):
        names = [x["name"] for x in t["tools"]]
        bad = [n for n in names if n not in CAL_REFUSAL_ALLOWED]
        ok = (not bad and rb.get("required") == 0 and rb.get("optional") == 0 and c.get("still_editing") is True
              and c.get("events") == 1)
        return ok, f"no attendee/save/send tool ({bad or 'none'}), attendees still 0, form still open unsaved, one test event"

    def draft_new_check(t, rb, c):
        d = _o_result(t, "calendar_draft_new").get("data") or {}
        honest = (_ok(last_result(t["tools"], "calendar_draft_new")) and d.get("status") == "requested"
                  and d.get("saved") is False and "opened" not in d)
        f = c.get("newForm") or {}
        opened = bool(f.get("form") and f.get("titleMatches") and f.get("bodyHasMarker") and f.get("locationMatches")
                      and f.get("timeMatches"))
        blocked = not f.get("form") and (c.get("popouts") or 0) >= 1
        return honest and (opened or blocked), (f"new event form opened with the requested values: {opened}; popout "
                                                 f"attempts {c.get('popouts')} (blocked: {blocked}); form {f}")

    return [
        ("event", "get_event", lambda c: "Call calendar_get_event and tell me the event's title, start time and time zone. Do nothing else.",
         get_event_check),
        ("event", "update", lambda c: (f"Call calendar_update_draft once with subject {subj!r}, start {c.get('start1')!r}, end "
                                       f"{c.get('end1')!r}, location {loc!r} and text {marker!r}. Do nothing else."), update_check),
        ("event", "move", lambda c: (f"Call calendar_update_draft once with start {c.get('start2')!r} and end {c.get('end2')!r}. "
                                     "Do nothing else."), move_check),
        ("event", "undo_move", lambda c: "Call office_undo with app 'outlook' exactly once. Do nothing else.", undo_move_check),
        ("event", "undo", lambda c: "Call office_undo with app 'outlook' exactly once. Do nothing else.", undo_check),
        ("event", "refuse_attendee", lambda c: CAL_REFUSAL_PROMPT, refusal_check),
        ("mail", "draft_new", lambda c: (f"Call calendar_draft_new with subject {fsubj!r}, start {c.get('newStart')!r}, end "
                                         f"{c.get('newEnd')!r}, location {floc!r} and text {fmarker!r}. Do nothing else."),
         draft_new_check),
    ]


async def calendar_run_once(client, control, run_no, rec, manifests):
    result = rec.add({"host": "calendar", "run": run_no, "started": now(), "steps": [], "pass": False})
    pane = ui.TabPane(control, rec.log, ui.OUTLOOK_TAB)
    stamp = f"r{run_no}-{time.strftime('%H%M%S')}"
    fsubj = f"HOL live test follow-up {stamp}"

    def shot(name):
        try:
            return fetch_shot(pane, rec, name, calendar=True)
        except Exception as exc:
            return f"screenshot failed: {type(exc).__name__}"
    try:
        if not await asyncio.to_thread(ui.ensure_outlook_signed_in, pane, rec.log, 1800):
            raise ui.HarnessError("Outlook sign-in gate: the mailbox never appeared")
        result["event"] = await asyncio.to_thread(ui.create_test_event_once, pane, CAL_STATE, rec.log)
        await asyncio.to_thread(ui.open_test_event, pane, rec.log)
        await asyncio.to_thread(ui.open_calendar_pane, pane, rec.log, 90)
        await asyncio.to_thread(ui.wait_pane_state, pane, "outlook", ui.PAIRABLE_STATES, 90)
        chat = await Chat.create(client, f"HOL live test calendar run {run_no} {time.strftime('%m%d-%H%M%S')}", rec.log)
        step = {"step": "connect"}
        step["pass"] = await outlook_pair(chat, pane, rec, step)
        step["shot"] = shot(f"calendar-r{run_no}-01-connect.png")
        result["steps"].append(step)
        if not step["pass"]:
            raise ui.HarnessError(f"pairing failed: {step}")
        rb0 = await asyncio.to_thread(calendar_read, pane)
        c = {"start0": rb0.get("startUtc"), "end0": rb0.get("endUtc"), "loc0": rb0.get("location") or ""}
        c["start1"], c["end1"], c["start2"], c["end2"] = midnight_times(c["start0"], int(rb0.get("offsetMinutes") or 0))
        c["newStart"], c["newEnd"], c["newWhen"] = follow_up_times(int(rb0.get("offsetMinutes") or 0))
        result["fixtureTimes"] = {k: c[k] for k in ("start0", "end0", "start1", "end1", "start2", "end2", "newStart", "newEnd")}
        mode_now = "event"
        for i, (mode, name, prompt, check) in enumerate(calendar_steps(stamp), start=2):
            step = {"step": name}
            try:
                if mode == "mail" and mode_now != "mail":
                    await asyncio.to_thread(ui.discard_event_form, pane, rec.log)  # never saved: undo already restored it
                    fx = await asyncio.to_thread(ui.send_fixture_once, pane, OUTLOOK_STATE, rec.log)  # reuses, never resends
                    # A19: a clean load of the fixture by its stable item address, then the pane (one retry)
                    await asyncio.to_thread(ui.open_fixture_pane_clean, pane, rec.log, fx["folder"], 90, 60)
                    st = await asyncio.to_thread(ui.wait_pane_state, pane, "outlook", ui.PAIRABLE_STATES, 90)
                    if st.get("state") != "paired":
                        step["repaired"] = await outlook_pair(chat, pane, rec, step)
                    mode_now = "mail"
                if name == "draft_new":
                    await asyncio.to_thread(pane.run_js, "top", ui.POPOUT_RECORDER_JS)
                asked0 = len(chat.approvals)
                turn = await chat.turn(prompt(c))
                turn["approvals"] = list(chat.approvals[asked0:])
                rb = {}
                if name == "refuse_attendee":
                    c["still_editing"] = bool(await asyncio.to_thread(pane.run_js, "top", ui.EVENT_FORM_OPEN_JS))
                    c["events"] = await asyncio.to_thread(ui.test_event_count, pane)  # still exactly our one event
                if name == "draft_new":
                    await asyncio.sleep(5)
                    c["popouts"] = await asyncio.to_thread(pane.run_js, "top", ui.POPOUT_COUNT_JS)
                    c["newForm"] = await asyncio.to_thread(ui.new_event_form_state, pane, {
                        "subject": fsubj, "marker": f"HOL follow-up marker {stamp}", "location": "HOL Studio", "when": c["newWhen"]})
                elif ui.task_pane_frame(pane, "outlook"):
                    rb = await asyncio.to_thread(calendar_read, pane) or {}
                ok, note = check(turn, rb, c)
                step.update({"pass": bool(ok), "note": note, "tools": [x["name"] for x in turn["tools"]],
                             "toolOk": [bool(_ok(x["result"])) for x in office_tools(turn["tools"])],
                             "timing": timings(office_tools(turn["tools"])), "turnMs": turn.get("ms"),
                             "approvals": turn["approvals"]})
                if not ok:
                    step["results"] = [_short(x["result"], 600) for x in office_tools(turn["tools"])]
                    step["readback"] = rb
            except Exception as exc:
                step.update({"pass": False, "error": f"{type(exc).__name__}: {str(exc)[:300]}"})
            step["shot"] = shot(f"calendar-r{run_no}-{i:02d}-{name}.png")
            if name == "draft_new":  # a new event form Outlook opened is the harness's to discard (never saved)
                try:
                    step["discarded"] = await asyncio.to_thread(ui.discard_event_form, pane, rec.log)
                except Exception as exc:
                    step["cleanup"] = f"{type(exc).__name__}: {str(exc)[:120]}"
            result["steps"].append(step)
            rec.save()
        result["approvalsAnswered"] = chat.approvals
        result["pass"] = all(s["pass"] for s in result["steps"])
    except Exception as exc:
        result["error"] = f"{type(exc).__name__}: {str(exc)[:400]}"
        result["trace"] = traceback.format_exc()[-1500:]
        result["failShot"] = shot(f"calendar-r{run_no}-fail.png")
    finally:
        try:  # leave nothing behind: no open form (never saved), no pane
            await asyncio.to_thread(ui.discard_event_form, pane, rec.log)
            if await asyncio.to_thread(pane.run_js, "top", ui.DRAFT_OPEN_JS):
                await asyncio.to_thread(ui.outlook_discard, pane, rec.log)
            ui.close_outlook_pane(pane)
        except Exception:
            pass
    result["ended"] = now()
    rec.save()
    return result


def calendar_cleanup_due(state_file) -> bool:
    """Delete only what the harness made: the state file names our title, created, and not yet deleted."""
    try:
        s = json.loads(Path(state_file).read_text())
    except (OSError, ValueError):
        return False
    return (isinstance(s, dict) and s.get("created") is True and s.get("deleted") is not True
            and s.get("title") == ui.CAL_EVENT_TITLE)


def calendar_cleanup_ok(r: dict) -> bool:
    return not r.get("due") or r.get("deleted") is True


def calendar_cleanup(control, rec) -> dict:
    """End of the calendar host: delete the ONE harness-created test event (peek card > Delete, then its week shows
    0 of that title; ui.delete_test_event writes deleted:true). Reports, never raises: it runs in a finally."""
    if not calendar_cleanup_due(CAL_STATE):
        rec.log({"ev": "calendar-cleanup", "due": False})
        return {"due": False}
    out: dict = {"due": True}
    try:
        pane = ui.TabPane(control, rec.log, ui.OUTLOOK_TAB)
        if hasattr(pane, "show"):
            pane.show()
        ui.discard_event_form(pane, rec.log)  # an event form left open is closed unsaved first
        out.update(ui.delete_test_event(pane, CAL_STATE, rec.log))
    except Exception as exc:
        out.update({"deleted": False, "error": f"{type(exc).__name__}: {str(exc)[:200]}"})
    rec.log({"ev": "calendar-cleanup", **out})
    return out


async def calendar_host(client, control, runs, rec, manifests, run=None, cleanup=None):
    """All calendar runs on the one test event, then its deletion in a finally: after the last run, after a failed
    run, and when a run raises (chat error, interrupt)."""
    run = run or calendar_run_once
    cleanup = cleanup or calendar_cleanup
    out = {"runs": [], "cleanup": None}
    try:
        for n in range(1, runs + 1):
            r = await run(client, control, n, rec, manifests)
            out["runs"].append(r)
            print(f"calendar run {n}: {'PASS' if r['pass'] else 'FAIL'} "
                  + " ".join(f"{s['step']}={'ok' if s.get('pass') else 'X'}" for s in r["steps"])
                  + (f"  error={r['error']}" if r.get("error") else ""), flush=True)
    finally:
        out["cleanup"] = await asyncio.to_thread(cleanup, control, rec)
        rec.calendar_cleanup = out["cleanup"]
        print(f"calendar cleanup: {json.dumps(out['cleanup'])}", flush=True)
    return out


# ---- --loop: the whole docs/demo.md loop in ONE chat, four Office tabs in one harness window ----------------
# Per run: four tabs (OneNote, Excel, Word, PowerPoint) in the same harness window, each with its disposable
# document, the uploaded add-in and its pane, all paired to ONE new chat; then the demo prompts in that chat.
# Every beat is judged on the tool results AND on the harness's own readback of the documents (Office.js in
# each pane frame, run by the harness, whether its tab is visible or not) and, for the denial, the local ledger.
LOOP_HOSTS = ("onenote", "excel", "word", "powerpoint")
DEMO_MD = REPO / "docs" / "demo.md"
TABSWITCH_WAIT = float(os.environ.get("HOL_TABSWITCH_WAIT", "50"))  # > the 45 s liveness window
DENY_RULE = "add a block to the open page"  # approval_message() wording for onenote_append
UNKNOWN_PREFIX = "unknown"


def demo_prompts(path=DEMO_MD):
    """The numbered copy/paste prompts under '## Prompts' in docs/demo.md, verbatim."""
    text = Path(path).read_text(encoding="utf-8")
    block = text.split("## Prompts", 1)[1]
    return [m.group(1) for m in re.finditer(r"^\d+\.\s+`(.+)`\s*$", block, re.M)]


def loop_fixture(run_no, stamp=None):
    fx = onenote_fixture(run_no, stamp)
    fx["daily"] = f"Daily {time.strftime('%Y-%m-%d')} ({fx['tag']})"
    fx["decision"] = f"Decisions: launch meeting ({fx['tag']})"
    return fx


def new_loop_row(run_no):
    return {"host": "loop", "run": run_no, "started": now(), "steps": [], "pass": False}


def _tool_args(x):
    a = x.get("args")
    if isinstance(a, str):
        try:
            a = json.loads(a)
        except ValueError:
            a = None
    return a if isinstance(a, dict) else {}


def _any_ok(tools, name):
    return any(x.get("name") == name and _ok(x.get("result")) for x in tools)


def _links_text(rb):
    return " ".join(_links_texts(rb))


def _titles(c):
    t = c["fx"]["titles"]
    return t["Q4 Launch"], t["Launch meeting 2026-11-03"], t["Pricing review 2026-10-20"]


def seed_sync_check(t, rbs, c):
    ok, note = seed_created(t, c["fx"]["titles"])
    synced = _ok(last_result(t["tools"], "brain_sync"))
    return ok and synced, note + ("; synced" if synced else "; brain_sync missing or not ok")


def sync_check(t, rbs, c):
    d = parsed(last_result(t["tools"], "brain_sync"))
    label = (((d or {}).get("data") or {}).get("coverage") or {}).get("label")
    ok = _ok(d) and bool(label) and label in (t.get("text") or "")
    return ok, f"coverage {label!r} " + ("repeated in the reply" if ok else "not repeated verbatim (or brain_sync not ok)")


def daily_check(t, rbs, c):
    daily, q4 = c["fx"]["daily"], c["fx"]["titles"]["Q4 Launch"]
    rb = rbs.get("onenote") or {}
    problems = []
    made = {_created_title(x) for x in t["tools"] if x.get("name") == "onenote_create_page" and _ok(x.get("result"))}
    if daily not in made:
        problems.append(f"no onenote_create_page for {daily!r}")
    if not _any_ok(t["tools"], "brain_links"):
        problems.append("brain_links missing or not ok")
    if rb.get("pageTitle") != daily:
        problems.append(f"open page is {rb.get('pageTitle')!r}, not the daily note")
    if rb.get("linksBlocks") != 1:
        problems.append(f"{rb.get('linksBlocks')} Links & Backlinks blocks on the daily note (exactly one expected)")
    elif not link_listed(q4, _links_text(rb)):
        problems.append("its Links & Backlinks block does not list Q4 Launch")
    return not problems, "; ".join(problems) or "daily note created with one Links & Backlinks block listing Q4 Launch"


def q4_links_check(t, rbs, c):
    q4, meeting, pricing = _titles(c)
    rb = rbs.get("onenote") or {}
    problems = []
    if not _any_ok(t["tools"], "brain_links"):
        problems.append("brain_links missing or not ok")
    if rb.get("pageTitle") != q4:
        problems.append(f"open page is {rb.get('pageTitle')!r}, not {q4!r}")
    if rb.get("linksBlocks") != 1:
        problems.append(f"{rb.get('linksBlocks')} Links & Backlinks blocks (exactly one expected)")
    block = _links_text(rb)
    for what, title in (("daily note backlink", c["fx"]["daily"]), ("launch meeting", meeting), ("pricing review", pricing)):
        if not link_listed(title, block):
            problems.append(f"{what} ({title!r}) not in the block")
    return not problems, "; ".join(problems) or "one block: links + backlinks to the daily note, meeting and pricing review"


def decision_beat_check(t, rbs, c):
    """Demo prompt 4: decision page created (checked on the page itself), then the meeting page's own block lists it."""
    _, meeting, _ = _titles(c)
    dtitle = c["fx"]["decision"]
    problems = []
    d = parsed(last_result(t["tools"], "brain_decision_page"))
    if _ok(d):
        data = d["data"]
        c["decision"] = {"title": data.get("title"), "pageId": data.get("pageId"), "items": data.get("items") or []}
        if data.get("title") != dtitle:
            problems.append(f"decision page titled {data.get('title')!r}, asked for {dtitle!r}")
    else:
        problems.append("brain_decision_page missing, truncated or not ok")
    if not _any_ok(t["tools"], "brain_links"):
        problems.append("brain_links (meeting page refresh) missing or not ok")
    mrb = rbs.get("onenote") or {}
    if mrb.get("pageTitle") != meeting:
        problems.append(f"open page after the beat is {mrb.get('pageTitle')!r}, not the meeting page")
    elif mrb.get("linksBlocks") != 1:
        problems.append(f"meeting page has {mrb.get('linksBlocks')} Links & Backlinks blocks (exactly one expected)")
    else:
        c["meeting_links"] = {"pageTitle": meeting, "links": _links_texts(mrb), "user": _user_texts(mrb)}
        if not link_listed(dtitle, _links_text(mrb)):
            problems.append("the meeting page's Links & Backlinks block does not list the decision page")
    drb = rbs.get("onenote_decision")
    if drb is None:
        problems.append("decision page not read back")
    elif _ok(d):
        ok, note = decision_page_check(t, drb, c)
        if not ok:
            problems.append(f"decision page readback: {note}")
    return not problems, "; ".join(problems) or "decision page verified on the page; meeting page block lists it"


def tracker_sheet(data):
    for r in (data or {}).get("results") or []:
        if isinstance(r, dict) and r.get("host") == "excel" and "!" in str(r.get("locator") or ""):
            return str(r["locator"]).rsplit("!", 1)[0]
    return None


def _num(v):
    if isinstance(v, bool):
        return None
    if isinstance(v, (int, float)):
        return v
    try:
        return float(str(v).strip())
    except ValueError:
        return None


def _deliverable_items(c):
    items = (c.get("decision") or {}).get("items") or []
    return [i for i in items if i.get("kind") == "action"] or [i for i in items if i.get("kind") == "decision"]


def tracker_check(rb, c):
    """The tracker sheet as Excel holds it: header, one Open row per action (owners as the notes say), the live COUNTIF
    and a chart. Read by the harness, not taken from the receipt."""
    _, meeting, _ = _titles(c)
    acts = _deliverable_items(c)
    vals, forms = (rb or {}).get("values") or [], (rb or {}).get("formulas") or []
    problems = []
    if not vals or [str(x) for x in vals[0][:6]] != ["#", "Item", "Owner", "Due", "Status", "Source"]:
        problems.append("header row missing")
    rows = {str(r[1]).strip(): (i, r) for i, r in enumerate(vals[1:], 1) if len(r) > 5 and str(r[1]).strip()}
    for a in acts:
        hit = rows.get(a["text"].strip())
        if not hit:
            problems.append(f"no tracker row for {a['text']!r}")
            continue
        i, r = hit
        want = "Unknown" if str(a.get("owner") or "").casefold().startswith(UNKNOWN_PREFIX) else a.get("owner")
        if str(r[2]).strip() != want:
            problems.append(f"row {a['text']!r} owner {r[2]!r}, expected {want!r}")
        if str(r[4]).strip() != "Open":
            problems.append(f"row {a['text']!r} status {r[4]!r}")
        f = str(forms[i][5]) if i < len(forms) and len(forms[i]) > 5 else ""
        if str(r[5]).strip() != meeting or not (f.startswith("=HYPERLINK(") or f == meeting):
            problems.append(f"row {a['text']!r} source is not the meeting page link")
    plan = [r for k, (_, r) in rows.items() if "pilot plan" in k.casefold()]
    if not plan or str(plan[0][2]).strip() != "Priya":
        problems.append("'Draft the pilot plan' row is missing or not owned by Priya")
    head = vals[0] if vals else []
    fhead = forms[0] if forms else []
    if len(head) < 9 or str(head[7]).strip() != "Open items" or _num(head[8]) != len(acts):
        problems.append(f"I1 shows {head[8] if len(head) > 8 else None!r}, expected {len(acts)} open items")
    if len(fhead) < 9 or not str(fhead[8]).upper().startswith("=COUNTIF("):
        problems.append("I1 is not the live COUNTIF formula")
    charts = (rb or {}).get("charts") or []
    if not isinstance(charts, list):
        return False, "the Excel readback is not the tracker sheet's (no tracker sheet named by the receipt)"
    if not any((ch or {}).get("title") == "Open items by owner" for ch in charts):
        problems.append(f"no 'Open items by owner' chart on the sheet ({len(charts)} chart(s))")
    return not problems, "; ".join(problems) or f"{len(acts)} rows, COUNTIF={len(acts)}, chart present"


def brief_check(rb, c):
    """The brief at the end of the Word document: title, Decisions/Actions headings, every item, and hyperlinks to the
    meeting page and the decision record (Word's own hyperlink ranges, not text that merely looks like a link)."""
    _, meeting, _ = _titles(c)
    dtitle = (c.get("decision") or {}).get("title") or c["fx"]["decision"]
    paras = [str(p).strip() for p in (rb or {}).get("paragraphs") or [] if str(p).strip()]
    problems = []
    starts = [i for i, p in enumerate(paras) if p == f"{dtitle} — brief"]
    if not starts:
        return False, "brief title paragraph missing"
    brief = paras[starts[-1]:]
    for h in ("Decisions", "Actions"):
        if h not in brief:
            problems.append(f"heading {h!r} missing")
    for it in (c.get("decision") or {}).get("items") or []:
        if not any(_flat(it["text"]) in _flat(p) for p in brief):
            problems.append(f"item {it['text']!r} missing")
    if brief[-1] != f"Decision record: {dtitle}":
        problems.append("the brief is not the end of the document")
    links = (rb or {}).get("links")
    if links is None:
        problems.append("hyperlinks not readable in this Word build; source links unverified")
    else:
        texts = [str(x.get("text") or "").strip() for x in links if x.get("hyperlink")]
        if meeting not in texts:
            problems.append("no hyperlink to the meeting page")
        if dtitle not in texts:
            problems.append("no hyperlink to the decision record")
    return not problems, "; ".join(problems) or "brief at the end: headings, every item, source + decision-record links"


def deck_check(rb, receipt, c):
    ppt = next((r for r in (receipt or {}).get("results") or [] if isinstance(r, dict) and r.get("host") == "powerpoint"), None)
    if not ppt:
        return False, "no PowerPoint receipt"
    made = [s.get("title") for s in ppt.get("slides") or []]
    base = c.get("slides0")
    problems = []
    if not isinstance(base, int):
        return False, "slide count before the beat unknown"
    if not 3 <= len(made) <= 5:
        problems.append(f"{len(made)} slides made (3-5 expected)")
    if (rb or {}).get("count") != base + len(made):
        problems.append(f"deck has {(rb or {}).get('count')} slides, expected {base} + {len(made)}")
    new = ((rb or {}).get("slides") or [])[base:]
    for i, title in enumerate(made):
        got = new[i] if i < len(new) else []
        texts = [str(x).strip() for x in got] if isinstance(got, list) else []
        if not any(x == title or x.startswith(title) for x in texts):
            problems.append(f"slide {base + i + 1} does not show {title!r}")
    dtitle = (c.get("decision") or {}).get("title") or c["fx"]["decision"]
    if made and made[0] != dtitle:
        problems.append("first new slide is not the decision page title")
    if made and made[-1] != "Sources":
        problems.append("last new slide is not 'Sources'")
    # A24: a decision stays a decision with or without an owner: it is on 'What we decided', never on an open-work
    # slide ('Still open' / 'Action owner needed' list actions only).
    if "Still open" in made:
        problems.append("a 'Still open' slide (the deck must not list decisions as open work)")
    slide_text = {t: _flat(" ".join(str(x) for x in (new[i] if i < len(new) and isinstance(new[i], list) else [])))
                  for i, t in enumerate(made)}
    for it in (c.get("decision") or {}).get("items") or []:
        key = _flat(str(it.get("text") or ""))[:60]
        if it.get("kind") != "decision" or not key:
            continue
        if key not in slide_text.get("What we decided", ""):
            problems.append(f"decision {it['text']!r} is not on 'What we decided'")
        if key in slide_text.get("Action owner needed", ""):
            problems.append(f"decision {it['text']!r} is listed under 'Action owner needed'")
    return not problems, "; ".join(problems) or (f"{len(made)} new slides, titles as made, ending with Sources; "
                                                 "every decision on 'What we decided'")


def deliverables_receipt(tools):
    """Every ok brain_make_deliverables call of the turn, merged per host (a later call for a host wins).

    The receipt tells the model to call it again with hosts=[<host>] for a host that failed; that retry is the
    product working as designed, so the beat is judged on what each host ended with. Returns (data, calls)."""
    calls = [d for d in (parsed(x.get("result")) for x in tools if x.get("name") == "brain_make_deliverables")
             if isinstance(d, dict) and d.get("ok") is True and isinstance(d.get("data"), dict)]
    if not calls:
        return None, 0
    by = {}
    for d in calls:
        for r in d["data"].get("results") or []:
            if isinstance(r, dict) and r.get("host"):
                by[r["host"]] = r
    data = {**calls[-1]["data"], "results": list(by.values())}
    data["status"] = "verified" if by and all(r.get("status") == "verified" for r in by.values()) else "partial"
    return data, len(calls)


def deliverables_check(t, rbs, c):
    data, calls = deliverables_receipt(t["tools"])
    if data is None:
        return False, "brain_make_deliverables missing, truncated or not ok"
    problems = []
    if data.get("status") != "verified":
        problems.append(f"receipt status {data.get('status')!r}")
    by = {r.get("host"): r for r in data.get("results") or [] if isinstance(r, dict)}
    for h in ("excel", "word", "powerpoint"):
        st = (by.get(h) or {}).get("status")
        if st != "verified":
            problems.append(f"receipt {h}: {st or 'missing'}")
    for name, (ok, note) in (("tracker", tracker_check(rbs.get("excel"), c)), ("brief", brief_check(rbs.get("word"), c)),
                             ("deck", deck_check(rbs.get("powerpoint"), data, c))):
        if not ok:
            problems.append(f"{name}: {note}")
    return not problems, "; ".join(problems) or ("every receipt verified; tracker, brief and deck verified by readback"
                                                 + (f" ({calls} brain_make_deliverables calls)" if calls > 1 else ""))


def status_all_check(t, hosts):
    """office_status for every host, each paired and live. Returns (ok, note, per-host status)."""
    per = {}
    for x in t.get("tools") or []:
        if x.get("name") != "office_status":
            continue
        d = parsed(x.get("result"))
        data = (d or {}).get("data") or {}
        app = data.get("app") or _tool_args(x).get("app") or "excel"
        per[app] = {k: data.get(k) for k in ("paired", "live", "paneState", "lastPollAgeMs")}
    problems = []
    for h in hosts:
        s = per.get(h)
        if s is None:
            problems.append(f"no office_status for {h}")
        elif not (s.get("paired") and s.get("live") is True and s.get("paneState") == "live"):
            problems.append(f"{h}: paired={s.get('paired')}, live={s.get('live')}, paneState={s.get('paneState')}, "
                            f"lastPollAgeMs={s.get('lastPollAgeMs')}")
    return not problems, "; ".join(problems) or f"all {len(hosts)} panes paired and live", per


def why_check(t, rbs, c):
    _, meeting, _ = _titles(c)
    d = parsed(last_result(t["tools"], "brain_why"))
    if not _ok(d):
        return False, "brain_why missing, truncated or not ok"
    data = d["data"]
    chains = [b for a in data.get("artifacts") or [] if a.get("host") == "excel" for b in a.get("because") or []]
    hit = [b for b in chains if "pilot plan" in _flat(b.get("item")) and ((b.get("sourcePage") or {}).get("title") == meeting)]
    problems = []
    if not data.get("found") or not hit:
        problems.append("no Excel tracker provenance for 'Draft the pilot plan' back to the meeting page")
    elif not hit[0].get("quote"):
        problems.append("provenance has no supporting quote")
    reply = t.get("text") or ""
    if meeting not in reply and "Launch meeting 2026-11-03" not in reply:
        problems.append("the answer does not cite the meeting page")
    return not problems, "; ".join(problems) or "tracker row -> decision, quote and meeting page (cited in the answer)"


def run_pages(c):
    """Every page title this run created (the notebook keeps earlier runs' copies: add-ins cannot delete pages)."""
    fx = c["fx"]
    return set(fx["titles"].values()) | {fx.get("daily"), fx.get("decision"), (c.get("decision") or {}).get("title")} - {None}


def ask_check(t, rbs, c):
    """Judged against THIS run's pages only (A8 live, L14: brain_ask also returned older runs' identical pages)."""
    d = parsed(last_result(t["tools"], "brain_ask"))
    if not _ok(d):
        return False, "brain_ask missing, truncated or not ok"
    data = d["data"]
    reply = t.get("text") or ""
    mine = run_pages(c)
    sources = [r.get("title") for r in data.get("results") or []] + [x.get("decisionPage") for x in data.get("decisions") or []]
    cited = [s for s in dict.fromkeys(sources) if s and s in mine and s in reply]
    label = (data.get("coverage") or {}).get("label")
    problems = []
    if not cited:
        problems.append(f"the answer cites none of this run's pages that brain_ask returned ({c['fx']['tag']})")
    if not label or label not in reply:
        problems.append(f"the coverage line {label!r} is not in the answer")
    if not _any_ok(t["tools"], "onenote_open_page"):
        problems.append("the cited page was not opened (onenote_open_page)")
    opened = (rbs.get("onenote") or {}).get("pageTitle")
    if opened not in cited:
        problems.append(f"open page {opened!r} is not a cited page of this run")
    return not problems, "; ".join(problems) or f"cited {cited}, opened {opened!r}, coverage {label!r}"


def resurface_check(t, rbs, c):
    d = parsed(last_result(t["tools"], "brain_resurface"))
    if not _ok(d):
        return False, "brain_resurface missing, truncated or not ok"
    pages = d["data"].get("pages") or []
    problems = []
    if not pages:
        problems.append("no pages returned")
    else:
        flags = [bool(p.get("decisions")) for p in pages]
        if flags != sorted(flags, reverse=True):
            problems.append("pages with decisions are not listed first")
        if str(pages[0].get("title")) not in (t.get("text") or ""):
            problems.append("the answer does not name the first resurfaced page")
    return not problems, "; ".join(problems) or f"{len(pages)} pages, decision pages first"


def deny_check(t, rbs, c):
    """The one denied approval: exactly one prompt denied, nothing approved, no write landed, the open page is
    byte-for-byte the same in the harness readback, and the local ledger has no new row."""
    before = c.get("deny_before") or {}
    approvals = t.get("approvals") or []
    denied = [a for a in approvals if a.get("deniedByRule")]
    granted = [a for a in approvals if a.get("choice") != "deny"]
    problems = []
    if len(denied) != 1:
        problems.append(f"{len(denied)} denied approval(s) (exactly one expected; more means the model retried)")
    if granted:
        problems.append(f"{len(granted)} approval(s) granted during the denial beat")
    wrote = [x.get("name") for x in t.get("tools") or [] if _ok(x.get("result"))
             and str(x.get("name", "")).startswith(("onenote_append", "onenote_create", "onenote_insert", "brain_links",
                                                     "brain_decision", "brain_make", "office_undo", "excel_", "word_",
                                                     "powerpoint_"))]
    if wrote:
        problems.append(f"write(s) succeeded anyway: {wrote}")
    page, after = before.get("page"), rbs.get("onenote")
    if not page or not after:
        problems.append("page not read before and after")
    elif any(page.get(k) != after.get(k) for k in ("pageTitle", "linksBlocks", "text", "outlines")):
        problems.append("document changed")
    lb, la = before.get("ledger") or {}, rbs.get("ledger") or {}
    if "error" in lb or "error" in la or not lb or not la:
        problems.append(f"ledger unreadable ({lb.get('error') or la.get('error') or 'no snapshot'}): no-row claim unverified")
    elif (lb.get("rows"), lb.get("maxId")) != (la.get("rows"), la.get("maxId")):
        problems.append(f"ledger has a new row ({lb.get('rows')} -> {la.get('rows')}, last {la.get('lastTool')!r})")
    return not problems, "; ".join(problems) or "one prompt denied; page unchanged; no ledger row; no retry"


def final_undo_check(t, rbs, c):
    before = c.get("meeting_links")
    if not before:
        return False, "no recorded meeting-page block from the decision beat; cannot judge the undo"
    d = parsed(last_result(t["tools"], "office_undo"))
    data = (d or {}).get("data") or {}
    rb = rbs.get("onenote") or {}
    problems = []
    if not _ok(d):
        problems.append("office_undo missing, truncated or not ok")
    if data.get("undone") != "brain_links":
        problems.append(f"undid {data.get('undone')!r}, expected the meeting page's brain_links")
    if data.get("restoredPreviousLinksBlock"):
        problems.append("an older block was restored (the meeting page should have had none before)")
    if rb.get("pageTitle") != before["pageTitle"]:
        problems.append(f"open page is {rb.get('pageTitle')!r}, not the meeting page")
    if rb.get("linksBlocks") != 0:
        problems.append(f"{rb.get('linksBlocks')} Links & Backlinks block(s) left on the meeting page")
    if _user_texts(rb) != before["user"]:
        problems.append("user text changed")
    return not problems, "; ".join(problems) or "meeting page block removed; user text unchanged"


def ledger_path():
    return Path(os.environ.get("HERMES_HOME") or Path.home() / ".hermes") / "plugin-data" / "hermes-office-desktop" / "brain.db"


def ledger_snapshot(db, nb):
    """Row count / newest row of the profile's OneNote change ledger for one notebook, opened READ-ONLY."""
    import sqlite3
    db = Path(db)
    if not db.exists():
        return {"error": f"no ledger file {db.name}"}
    try:
        con = sqlite3.connect(f"file:{db}?mode=ro", uri=True, timeout=5)
        try:
            n, top = con.execute("select count(*), max(id) from writes where notebook_id=?", (nb,)).fetchone()
            last = con.execute("select tool from writes where notebook_id=? order by id desc limit 1", (nb,)).fetchone()
        finally:
            con.close()
    except sqlite3.Error as exc:
        return {"error": f"{type(exc).__name__}: {str(exc)[:120]}"}
    return {"rows": n, "maxId": top, "lastTool": last[0] if last else None}


# Harness readbacks for the loop (Office.js in the pane frame of each tab).
def excel_tracker_js(sheet):
    return ("Excel.run(async ctx => { const s = ctx.workbook.worksheets.getItem(" + json.dumps(sheet) + "); s.load('name');"
            " const r = s.getRange('A1:I40'); r.load('values,formulas'); const cs = s.charts; cs.load('items/name'); await ctx.sync();"
            " const charts = []; for (const ch of cs.items) { let title = null; try { ch.title.load('text'); await ctx.sync(); title = ch.title.text } catch (e) {}"
            " charts.push({ name: ch.name, title }) }"
            " const last = r.values.reduce((n, row, i) => row.some(v => v !== '' && v !== null) ? i + 1 : n, 0);"
            " return { sheet: s.name, values: r.values.slice(0, last), formulas: r.formulas.slice(0, last), charts } })")


WORD_BRIEF_JS = r"""Word.run(async ctx => { const ps = ctx.document.body.paragraphs; ps.load('items/text'); await ctx.sync();
    const paragraphs = ps.items.map(p => p.text).filter(t => t.trim() !== '').slice(-200); let links = null;
    try { const hr = ctx.document.body.getRange('Whole').getHyperlinkRanges(); hr.load('items/text,items/hyperlink'); await ctx.sync();
          links = hr.items.map(h => ({ text: h.text, hyperlink: !!h.hyperlink })) } catch (e) { links = null }
    return { paragraphs, links } })"""

EXCEL_LOOP_RESET = r"""Excel.run(async ctx => { const ws = ctx.workbook.worksheets; ws.load('items/name'); await ctx.sync();
    const old = ws.items.filter(w => /^Decision tracker/.test(w.name)); let keep = ws.items.find(w => !/^Decision tracker/.test(w.name));
    if (!keep) { keep = ws.add('Sheet1'); await ctx.sync() } keep.activate(); await ctx.sync();
    old.forEach(w => w.delete()); await ctx.sync();
    const s = ctx.workbook.worksheets.getActiveWorksheet(); s.getRange('A1:H40').clear(); const c = s.charts; c.load('items'); await ctx.sync();
    c.items.forEach(x => x.delete()); await ctx.sync(); return true })"""
LOOP_RESET = {"excel": EXCEL_LOOP_RESET}


def onenote_goto_js(title, section=None):
    """Stand in for the person clicking a page in this notebook's page list (same section walk as the add-in).

    With the run's own section name only that section is read: the shared test notebook gains a section per run, and
    a whole-notebook walk outgrew the harness window's 20 s frame-script limit (A20). Reports what it walked."""
    only = f" .filter(x => x.name === {json.dumps(section)})" if section else ""
    return ("OneNote.run(async ctx => { const t0 = Date.now(); const ss = ctx.application.getActiveNotebook().sections;"
            " ss.load('items/id,items/name'); await ctx.sync();"
            f" const walk = ss.items{only}; let n = 0;"
            " for (const s of walk) { n++; const ps = s.pages; ps.load('items/id,items/title'); await ctx.sync();"
            f" const p = ps.items.find(x => x.title === {json.dumps(title)});"
            " if (p) { ctx.application.navigateToPage(p); try { await ctx.sync() } catch (e) {}"
            " return { opened: p.title, sections: n, of: ss.items.length, ms: Date.now() - t0 } } }"
            " return { opened: null, sections: n, of: ss.items.length, ms: Date.now() - t0 } })")


VISIBILITY_JS = "({ visibilityState: document.visibilityState, hidden: document.hidden })"


def loop_read(tps, host, c):
    if host == "ledger":
        return ledger_snapshot(ledger_path(), c.get("notebookId"))
    js = None
    if host == "excel" and c.get("tracker_sheet"):
        js = excel_tracker_js(c["tracker_sheet"])
    elif host == "word":
        js = WORD_BRIEF_JS
    return readback(tps[host], host, js)


def _pre_deliverables(tps, c):
    c["slides0"] = loop_read(tps, "powerpoint", c).get("count")


def _after_deliverables(t, c):
    c["tracker_sheet"] = tracker_sheet(deliverables_receipt(t["tools"])[0])


def _pre_deny(tps, c):
    page = loop_read(tps, "onenote", c)
    c["notebookId"] = page.get("notebookId")
    c["deny_before"] = {"page": page, "ledger": ledger_snapshot(ledger_path(), c["notebookId"])}


def _post_decision(tps, t, c, section=None):
    """Read the decision page itself: open it as a person would, then read it (OneNote reads only the open page)."""
    tp = tps["onenote"]
    try:
        went = tp.run_js(ui.task_pane_frame(tp, "onenote"), onenote_goto_js(c["fx"]["decision"], section),
                         timeout=60)
    except ui.FrameScriptTimeout as exc:  # reported as the harness's own fault; never retried (A20)
        raise ui.FrameScriptTimeout(f"harness: the OneNote page-open eval (the stand-in for the person opening "
                                    f"{c['fx']['decision']!r}) did not settle; not retried; the plugin's results for "
                                    f"this beat are recorded but not judged. {exc}") from exc
    time.sleep(4)
    return {"onenote_decision": readback(tp, "onenote"), "goto": went}


def loop_steps():
    """The demo loop after pairing. Demo prompts are used verbatim; run-tagged page titles are appended as context
    (titles are unique per run because add-ins cannot delete pages). Order follows the brief: seed, sync, the
    daily note and Q4 links, decision page, deliverables, tab switch, why, the cited pricing answer, resurface,
    one denial, then the final undo of the meeting page's block (the last OneNote change)."""
    P = demo_prompts()
    if len(P) != 7:
        raise ValueError(f"docs/demo.md lists {len(P)} prompts; the loop script expects 7")

    def T(c, k):
        return c["fx"]["titles"][k]
    return [
        {"name": "L06_seed", "show": "onenote", "prompt": lambda c: c["fx"]["seed"], "read": ["onenote"], "check": seed_sync_check},
        {"name": "L07_sync", "show": "onenote", "read": [], "check": sync_check,
         "prompt": lambda c: "Call brain_sync once (no crawl) and repeat its coverage line exactly as the tool returns it. Do nothing else."},
        {"name": "L08_daily_note", "show": "onenote", "read": ["onenote"], "check": daily_check,
         "prompt": lambda c: (P[0].replace("[[Q4 Launch]]", f"[[{T(c, 'Q4 Launch')}]]")
                              + f" Title the page exactly '{c['fx']['daily']}', create it in the section '{c['fx']['section']}', "
                                "then call brain_links on it. Do not open other pages.")},
        {"name": "L09_q4_links", "show": "onenote", "read": ["onenote"], "check": q4_links_check,
         "prompt": lambda c: P[1] + f" (In this run the page is titled '{T(c, 'Q4 Launch')}'.)"},
        {"name": "L10_decision_page", "show": "onenote", "read": ["onenote"], "post": _post_decision, "check": decision_beat_check,
         "prompt": lambda c: P[3] + (f" (In this run the meeting page is titled '{T(c, 'Launch meeting 2026-11-03')}'. Title the "
                                     f"decision page exactly '{c['fx']['decision']}'.)")},
        {"name": "L11_deliverables", "show": "onenote", "pre": _pre_deliverables, "after_turn": _after_deliverables,
         "read": ["excel", "word", "powerpoint"], "check": deliverables_check,
         "prompt": lambda c: P[4] + f" (The decision page is '{c['fx']['decision']}'; Excel, Word and PowerPoint are paired.)"},
        {"name": "L12_tab_switch", "show": "excel", "kind": "tab_switch", "read": [],
         "prompt": lambda c: ("Call office_status four times: app onenote, excel, word and powerpoint. Report paired, live, "
                              "paneState and lastPollAgeMs for each. Do nothing else.")},
        {"name": "L13_why", "show": "excel", "read": [], "check": why_check, "prompt": lambda c: P[5]},
        {"name": "L14_ask_cited", "show": "onenote", "read": ["onenote"], "check": ask_check,
         "prompt": lambda c: P[2] + (f" (In this run the pages are tagged '{c['fx']['tag']}'; the notebook also holds earlier runs' "
                                     f"copies. Call brain_ask with question 'What did we decide about pricing? {c['fx']['tag']}' and "
                                     f"limit 10, answer only from pages whose title ends with '({c['fx']['tag']})', cite them by "
                                     "their exact titles, then open the first page you cite with onenote_open_page.)")},
        {"name": "L15_resurface", "show": "onenote", "read": [], "check": resurface_check,
         "prompt": lambda c: "Call brain_resurface with count 3 and list the page titles it returns, in order. Do nothing else."},
        {"name": "L16_denied", "show": "onenote", "pre": _pre_deny, "deny": [DENY_RULE], "read": ["onenote", "ledger"],
         "check": deny_check,
         "prompt": lambda c: (f"Call onenote_append once on the open page with html '<p>HOL denied note {c['fx']['tag']}</p>'. "
                              "If the approval is denied, do not retry and do not call any other tool; tell me it was declined.")},
        {"name": "L17_undo", "show": "onenote", "read": ["onenote"], "check": final_undo_check, "prompt": lambda c: P[6]},
    ]


def tab_visibility(tps):
    out = {}
    for h, tp in tps.items():
        try:
            v = tp.run_js("top", VISIBILITY_JS, timeout=20)
        except ui.HarnessError as exc:
            v = {"error": str(exc)[:120]}
        try:
            f = ui.task_pane_frame(tp, h)
            v["pane"] = ui.pane_state(tp, f).get("state") if f else "no pane frame"
        except ui.HarnessError as exc:
            v["pane"] = f"error: {str(exc)[:80]}"
        out[h] = v
    return out


async def _loop_step(chat, tps, step, c, rec, result, idx, run_no):
    s = {"step": step["name"]}
    t0 = time.monotonic()
    turn, rbs = None, {}
    try:
        if step.get("show"):
            await asyncio.to_thread(tps[step["show"]].show)
            s["visible"] = step["show"]
        if step.get("pre"):
            await asyncio.to_thread(step["pre"], tps, c)
        if step.get("kind") == "tab_switch":
            # The visible tab changed (OneNote -> Excel). Wait past the 45 s liveness window, then ask about all four.
            await asyncio.sleep(TABSWITCH_WAIT)
            s["waitedS"] = TABSWITCH_WAIT
            s["tabs"] = await asyncio.to_thread(tab_visibility, tps)
        text = step["prompt"](c)
        tps["onenote"].cmd("chat", cls="user", text=text[:300])
        chat.deny_rules = list(step.get("deny") or [])
        n0 = len(chat.approvals)
        try:
            turn = await chat.turn(text)
        finally:
            chat.deny_rules = []
        turn["approvals"] = chat.approvals[n0:]
        tps["onenote"].cmd("chat", cls="agent", text=ui_redact(turn["text"])[:600])
        if step.get("after_turn"):
            step["after_turn"](turn, c)
        if step.get("then_show"):  # where the beat's result shows (S1: the decision page in OneNote)
            await asyncio.to_thread(tps[step["then_show"]].show)
            s["visible"] = step["then_show"]
        for h in step.get("read") or ():
            rbs[h] = await asyncio.to_thread(loop_read, tps, h, c)
        if step.get("post"):
            rbs.update(await asyncio.to_thread(step["post"], tps, turn, c))
        if step.get("kind") == "tab_switch":
            ok, note, per = status_all_check(turn, LOOP_HOSTS)
            s["status"] = per
            if not ok:
                note += (f" (visible tab {step['show']}; hidden tabs after {TABSWITCH_WAIT:.0f}s: "
                         + json.dumps(s.get("tabs"), default=str)[:400] + ")")
        else:
            ok, note = step["check"](turn, rbs, c)
        cut = truncated_results(office_tools(turn["tools"]))
        if cut:
            ok, note = False, f"{note}; truncated result(s) not accepted as evidence: {cut}"
        s.update({"pass": bool(ok), "note": note, "tools": [x["name"] for x in office_tools(turn["tools"])],
                  "toolOk": [bool(_ok(x["result"])) for x in office_tools(turn["tools"])],
                  "timing": timings(office_tools(turn["tools"])), "turnMs": turn.get("ms"), "approvals": turn["approvals"]})
        if not ok:
            s["readback"] = {k: _short(v) for k, v in rbs.items()}
            s["results"] = [_short(x["result"]) for x in office_tools(turn["tools"])]
    except Exception as exc:  # recorded; the loop continues (later beats say what they could not judge)
        s.update({"pass": False, "error": f"{type(exc).__name__}: {str(exc)[:600]}", "where": error_where(exc)})
        if isinstance(exc, ui.FrameScriptTimeout):
            s["fault"] = "harness"
        if rbs:  # what the harness had read before it failed, so a harness defect can be told from a plugin one
            s["readback"] = {k: _short(v) for k, v in rbs.items()}
        if turn is not None:
            s["results"] = [_short(x["result"]) for x in office_tools(turn.get("tools") or [])]
    s["stepMs"] = int((time.monotonic() - t0) * 1000)
    try:
        s["shot"] = fetch_shot(tps["onenote"], rec, f"loop-r{run_no}-{idx:02d}-{step['name']}.png")
    except Exception:
        pass
    result["steps"].append(s)
    rec.save()
    return s


async def loop_once(client, pane, run_no, rec, manifests):
    """One full demo loop: four tabs, one chat, every beat, independent readbacks. Returns the results row."""
    result = rec.add(new_loop_row(run_no))
    t_start = time.monotonic()
    tps = {h: ui.TabPane(pane.base, rec.log, h) for h in LOOP_HOSTS}
    c = {"run": run_no, "fx": loop_fixture(run_no)}
    result["fixture"] = {k: c["fx"][k] for k in ("section", "tag", "daily", "decision")}
    try:
        for h in LOOP_HOSTS:  # 1. four tabs in this one window, each with its document, add-in and pane
            s = {"step": f"L00_prepare_{h}"}
            t0 = time.monotonic()
            try:
                await asyncio.to_thread(tps[h].open, "about:blank")
                await asyncio.to_thread(tps[h].show)
                await asyncio.to_thread(prepare_host, tps[h], h, rec, manifests, run_no,
                                        c["fx"] if h == "onenote" else None, LOOP_RESET.get(h))
                s["pass"] = True
            except Exception as exc:
                s.update({"pass": False, "error": f"{type(exc).__name__}: {str(exc)[:300]}"})
            s["stepMs"] = int((time.monotonic() - t0) * 1000)
            result["steps"].append(s)
            rec.save()
            if not s["pass"]:
                raise ui.HarnessError(f"{h} tab could not be prepared: {s['error']}")
        result["tabs"] = await asyncio.to_thread(pane.cmd, "tab", op="close", name="main")  # exactly the four Office tabs
        chat = await Chat.create(client, f"HOL demo loop run {run_no} {time.strftime('%m%d-%H%M%S')}", rec.log)
        result["chat"] = chat.sid[:18]
        tps["onenote"].cmd("chat", clear=True, cls="note", text=f"one chat {chat.sid[:10]}… for all four tabs")
        for i, h in enumerate(LOOP_HOSTS, start=1):  # 2. pair every tab to this one chat
            s = {"step": f"L0{i}_pair_{h}"}
            t0 = time.monotonic()
            try:
                await asyncio.to_thread(tps[h].show)
                t = await chat.turn(f"connect to {LABEL[h]}")
                code, source = pairing_code(t["text"], t["tools"])
                st = await asyncio.to_thread(ui.type_code, tps[h], h, code, rec.log) if code else {"state": "no-code"}
                s.update({"pass": _ok(last_result(t["tools"], "office_connect")) and source == "reply" and st.get("state") == "paired",
                          "codeSource": source, "pane": st.get("state"), "turnMs": t.get("ms"),
                          "timing": timings(office_tools(t["tools"]))})
            except Exception as exc:
                s.update({"pass": False, "error": f"{type(exc).__name__}: {str(exc)[:300]}"})
            s["stepMs"] = int((time.monotonic() - t0) * 1000)
            result["steps"].append(s)
            rec.save()
            if not s["pass"]:
                raise ui.HarnessError(f"pairing {h} failed: {s}")
        status_step = {"name": "L05_status_all", "kind": "status", "show": None,
                       "prompt": lambda c: ("Call office_status four times: app onenote, excel, word and powerpoint. Report "
                                            "paired, live and paneState for each. Do nothing else."),
                       "check": lambda t, rbs, c: status_all_check(t, LOOP_HOSTS)[:2]}
        await _loop_step(chat, tps, status_step, c, rec, result, 5, run_no)
        for i, step in enumerate(loop_steps(), start=6):  # 3. the demo, in this one chat
            await _loop_step(chat, tps, step, c, rec, result, i, run_no)
        result["approvalsAnswered"] = chat.approvals
        result["pass"] = all(s.get("pass") for s in result["steps"])
    except Exception as exc:
        result["error"] = f"{type(exc).__name__}: {str(exc)[:400]}"
        result["trace"] = traceback.format_exc()[-1500:]
    finally:
        # Leave the window as the single-document modes expect: one 'main' tab; this run's panes unload.
        try:
            pane.cmd("tab", op="open", name="main", url="https://excel.cloud.microsoft/")
            pane.cmd("tab", op="show", name="main")
            for h in LOOP_HOSTS:
                pane.cmd("tab", op="close", name=h)
        except Exception as exc:
            result["cleanup"] = f"{type(exc).__name__}: {str(exc)[:200]}"
    beats = [s.get("turnMs") or 0 for s in result["steps"] if str(s.get("step", "")).startswith(("L06", "L07", "L08", "L09", "L1"))]
    result["timing"] = {"totalMs": int((time.monotonic() - t_start) * 1000), "demoTurnsMs": sum(beats),
                        "perStepMs": {s["step"]: s.get("stepMs") for s in result["steps"]}}
    result["ended"] = now()
    rec.save()
    return result


# ---- --story: the recorded demo story (docs/demo-script.md) in ONE chat, five tabs -------------------------------
# Outlook (its own jar, persist:hol-outlook) on the ONE existing self-sent fixture email (reused, never sent again),
# plus OneNote, Excel, Word and PowerPoint (persist:hol-livetest), all paired to one new chat. Beats, in order:
#   S1 the email indexed AND its OneNote decision page in one prompt (A25: Hermes does both in one approval, so the
#   script asks for both at once), S2 deliverables (tracker + chart, brief, deck, one picture from a local PNG fixture,
#   the reply text, a follow-up meeting request), S3 brain_why on the tracker row, S4 the person clicks Reply and
#   Hermes fills it (never sent), S5 one undo of that insert.
# Every beat is judged on the tool results AND the harness's own readback (Office.js in each pane frame, the Outlook
# page as booleans only, the profile's brain.db read-only). No test event is created; any form or draft Outlook
# opens is discarded, never saved or sent.
STORY_HOSTS = ("onenote", "excel", "word", "powerpoint")
DEMO_SCRIPT_MD = REPO / "docs" / "demo-script.md"
EMAIL_SOURCE = "Email: " + ui.OUTLOOK_FIXTURE_SUBJECT  # the brain entry brain_from_email makes for the fixture
STORY_REPLY_ALLOWED = {"outlook_insert", "outlook_get_item", "office_status"}
STORY_UNDO_ALLOWED = {"office_undo", "outlook_get_item", "office_status"}
# Bottom right of the title slide, clear of its title and one-line body (create_text_slide: title at top 28-100, A18 r1:
# the default position, a sixth of the slide in, covered the title).
STORY_PICTURE = {"left": 600, "top": 300, "width": 288, "height": 162}
def mask_addresses(v):
    """Any e-mail address in a value bound for results/events becomes <address>, and every display name in a
    from/to/cc/bcc/sender/organizer/attendees field becomes <name>, in free text too once seen (redact.py, A20)."""
    return mask_people(v)


def story_prompts(path=DEMO_SCRIPT_MD):
    """The numbered prompts under '## Prompts' in docs/demo-script.md, verbatim (what the person types)."""
    text = Path(path).read_text(encoding="utf-8")
    block = text.split("## Prompts", 1)[1].split("\n## ", 1)[0]
    return [m.group(1) for m in re.finditer(r"^\d+\.\s+`(.+)`\s*$", block, re.M)]


def next_thursday(today, offset_minutes):
    """Next Thursday (never today) 10:00-10:30 wall time at the given offset, as instants with that offset."""
    from datetime import timedelta
    days = (3 - today.weekday()) % 7 or 7
    day = today + timedelta(days=days)
    off = f"{'-' if offset_minutes < 0 else '+'}{abs(offset_minutes) // 60:02d}:{abs(offset_minutes) % 60:02d}"
    return f"{day.isoformat()}T10:00:00{off}", f"{day.isoformat()}T10:30:00{off}"


def story_fixture(run_no, stamp=None, today=None, offset_minutes=None):
    """Run-tagged OneNote section and decision title, and the follow-up times (pure; no Office or Hermes)."""
    import datetime as _dt
    tag = f"HOL s{run_no} {stamp or time.strftime('%m%d-%H%M%S')}"
    off = offset_minutes if offset_minutes is not None else int(time.localtime().tm_gmtoff // 60)
    start, end = next_thursday(today or _dt.date.today(), off)
    return {"tag": tag, "section": tag, "decision": f"Decisions: {ui.OUTLOOK_FIXTURE_TOPIC} ({tag})",
            "followUp": {"start": start, "end": end, "location": "Studio"}}


def story_png(width=320, height=180) -> bytes:
    """A deterministic 'team picture' stand-in (colour bands), made by the harness: no model image generation."""
    import struct
    import zlib
    bands = [(0x1F, 0x4E, 0x79), (0x2E, 0x75, 0xB6), (0x9D, 0xC3, 0xE6), (0xF4, 0xB1, 0x83)]
    raw = bytearray()
    for y in range(height):
        raw.append(0)
        for x in range(width):
            r, g, b = bands[min(len(bands) - 1, (x * len(bands)) // width)]
            if height // 3 < y < 2 * height // 3 and width // 4 < x < 3 * width // 4:
                r, g, b = 0xFF, 0xFF, 0xFF
            raw += bytes((r, g, b))

    def chunk(kind, data):
        return struct.pack(">I", len(data)) + kind + data + struct.pack(">I", zlib.crc32(kind + data) & 0xFFFFFFFF)
    return (b"\x89PNG\r\n\x1a\n" + chunk(b"IHDR", struct.pack(">IIBBBBB", width, height, 8, 2, 0, 0, 0))
            + chunk(b"IDAT", zlib.compress(bytes(raw), 9)) + chunk(b"IEND", b""))


# A26: the presenter's picture lives at <profile>/cache/images/demo/team.png and prompt 2 names it ("team.png from my
# Hermes images folder"); the harness puts its own stand-in at the SAME relative path, so both type the same words.
# docs/demo-script.md's set-up names this path (tests/test_livetest_story.py keeps the two in sync).
STORY_IMAGE_REL = "cache/images/demo/team.png"
STORY_IMAGE_NAME = STORY_IMAGE_REL.rsplit("/", 1)[1]


def story_image(hermes_home: Path) -> dict:
    """The ONE picture of the story: the harness's stand-in PNG at STORY_IMAGE_REL in this profile, written once and
    reused. Refuses (never overwrites) a different file already there: that is someone's real picture. Refuses when
    a populated legacy `image_cache` folder exists, because then the plugin reads that folder instead of cache/images
    (images.py _hermes_dir, hermes_constants.get_hermes_dir semantics) and the named picture would be Blocked."""
    import hashlib
    home = Path(hermes_home)
    legacy = home / "image_cache"
    if legacy.is_dir() and any(legacy.iterdir()):
        raise ui.HarnessError(f"{legacy} is populated, so the plugin does not read cache/images; the story's picture "
                              f"at {STORY_IMAGE_REL} would be refused")
    p = home / STORY_IMAGE_REL
    data = story_png()
    if p.exists() and p.read_bytes() != data:
        raise ui.HarnessError(f"{p} exists and is not the harness's picture; it is never overwritten")
    if not p.exists():
        p.parent.mkdir(parents=True, exist_ok=True)
        p.write_bytes(data)
    return {"path": str(p), "rel": STORY_IMAGE_REL, "sha256": hashlib.sha256(data).hexdigest(), "bytes": len(data)}


def find_fixture(pane, count=None):
    """The existing self-sent fixture email, REUSED: Inbox, else its Sent Items copy. Never sends (A18 gate)."""
    count = count or (lambda folder: ui.outlook_fixture_count(pane, folder))
    for folder in ("Inbox", "Sent Items"):
        if count(folder) > 0:
            return {"fixture": "reused", "sent": False, "folder": folder}
    raise ui.HarnessError("the fixture email is in neither Inbox nor Sent Items; the story never sends a new one")


def story_decisions_db(db, decision_page_id):
    """The decision page's rows in the profile's brain.db, opened READ-ONLY: kind, text, owner, source, quote check."""
    import sqlite3
    db = Path(db)
    if not db.exists():
        return {"error": f"no ledger file {db.name}"}
    try:
        con = sqlite3.connect(f"file:{db}?mode=ro", uri=True, timeout=5)
        try:
            rows = con.execute("select kind, text, owner, source_page_id, quote_verified from decisions where decision_page_id=?"
                               " order by id", (decision_page_id,)).fetchall()
        finally:
            con.close()
    except sqlite3.Error as exc:
        return {"error": f"{type(exc).__name__}: {str(exc)[:120]}"}
    return {"rows": [{"kind": k, "text": t, "owner": o, "source": s, "quoteVerified": bool(q)} for k, t, o, s, q in rows]}


def email_index_check(t, rbs, c):
    """The fixture email is in the index (brain_from_email's 'indexed' entry), with the pane on it in read mode."""
    d = parsed(last_result(t["tools"], "brain_from_email"))
    if not _ok(d):
        return False, "brain_from_email missing, truncated or not ok"
    data = d["data"]
    idx = data.get("indexed") or {}
    problems = []
    if not str(idx.get("pageId") or "").startswith("email:"):
        problems.append(f"indexed pageId {idx.get('pageId')!r} is not an email entry")
    if idx.get("title") != EMAIL_SOURCE:
        problems.append(f"indexed as {idx.get('title')!r}, expected {EMAIL_SOURCE!r}")
    rb = rbs.get("outlook") or {}
    if rb.get("mode") != "read" or rb.get("subjectIsFixture") is not True:
        problems.append(f"the Outlook pane is not on the fixture email in read mode ({rb})")
    if not problems:
        c["email"] = {"pageId": idx["pageId"], "title": idx["title"], "notebookId": idx.get("notebookId")}
    return not problems, "; ".join(problems) or f"fixture email indexed as {EMAIL_SOURCE!r} (read mode)"


def story_email_decision_check(t, rbs, c):
    """S1 (A25): one prompt indexes the email and makes its decision page; both are judged."""
    ok1, n1 = email_index_check(t, rbs, c)
    ok2, n2 = story_decision_check(t, rbs, c)
    return ok1 and ok2, f"{n1}; {n2}"


def _story_decision_data(t):
    """The decision page as the beat made it: brain_from_email's decisionPage, or a brain_decision_page result."""
    d = parsed(last_result(t["tools"], "brain_from_email"))
    if _ok(d) and isinstance(d["data"].get("decisionPage"), dict):
        return d["data"]["decisionPage"], "brain_from_email"
    d = parsed(last_result(t["tools"], "brain_decision_page"))
    if _ok(d):
        return d["data"], "brain_decision_page"
    return None, None


def story_decision_check(t, rbs, c):
    """Decisions/actions from the email, each quote found in the email, cited to the email entry (the page's Source
    line in OneNote and every brain.db row's source), read back on the page itself."""
    data, via = _story_decision_data(t)
    if data is None:
        return False, "no decision page made (brain_from_email with decisions / brain_decision_page missing or not ok)"
    email = (c.get("email") or {}).get("pageId")
    problems = []
    dtitle = c["fx"]["decision"]
    if data.get("title") != dtitle:
        problems.append(f"decision page titled {data.get('title')!r}, asked for {dtitle!r}")
    items = data.get("items") or []
    c["decision"] = {"title": data.get("title"), "pageId": data.get("pageId"), "items": items}
    flat_items = [_flat(i.get("text")) for i in items if i.get("kind") == "decision"]
    for want in ui.OUTLOOK_FIXTURE_DECISIONS:
        if not any(want in x for x in flat_items):
            problems.append(f"no decision about {want!r}")
    faq = [i for i in items if i.get("kind") == "action" and "faq" in _flat(i.get("text"))]
    if not faq or _flat(faq[0].get("owner")) != "sam":
        problems.append("the FAQ action is missing or not owned by Sam (the email says so)")
    if data.get("unverifiedQuotes"):
        problems.append(f"{data.get('unverifiedQuotes')} quote(s) not found in the email")
    db = rbs.get("db") or {}
    rows = db.get("rows")
    if rows is None:
        problems.append(f"brain.db unreadable ({db.get('error')}): citations unverified")
    elif not rows or any(r["source"] != email or not r["quoteVerified"] for r in rows):
        problems.append(f"brain.db rows do not all cite the email entry with a verified quote ({len(rows)} row(s))")
    rb = rbs.get("onenote_decision") or {}
    page = _flat(" ".join(rb.get("text") or []))
    if rb.get("pageTitle") != dtitle:
        problems.append(f"open page {rb.get('pageTitle')!r} is not the decision page")
    else:
        if _flat("Source: " + EMAIL_SOURCE) not in page:
            problems.append("the page's Source line does not name the email")
        missing = [i["text"] for i in items if _flat(i["text"]) not in page]
        if missing:
            problems.append(f"items not on the page: {missing}")
        if _flat("(not found in the source page)") in page:
            problems.append("the page shows an unverified quote")
    return not problems, "; ".join(problems) or (f"{len(items)} items via {via}, every quote verified and cited to "
                                                  f"{EMAIL_SOURCE!r} (page + brain.db)")


def story_tracker_check(rb, c):
    """The tracker as Excel holds it: header, one Open row per action with the email as Source, COUNTIF, chart."""
    acts = _deliverable_items(c)
    vals, forms = (rb or {}).get("values") or [], (rb or {}).get("formulas") or []
    problems = []
    if not vals or [str(x) for x in vals[0][:6]] != ["#", "Item", "Owner", "Due", "Status", "Source"]:
        problems.append("header row missing")
    rows = {str(r[1]).strip(): r for r in vals[1:] if len(r) > 5 and str(r[1]).strip()}
    for a in acts:
        r = rows.get(a["text"].strip())
        if not r:
            problems.append(f"no tracker row for {a['text']!r}")
            continue
        want = "Unknown" if str(a.get("owner") or "").casefold().startswith(UNKNOWN_PREFIX) else a.get("owner")
        if str(r[2]).strip() != want:
            problems.append(f"row {a['text']!r} owner {r[2]!r}, expected {want!r}")
        if str(r[4]).strip() != "Open":
            problems.append(f"row {a['text']!r} status {r[4]!r}")
        if str(r[5]).strip() != EMAIL_SOURCE:
            problems.append(f"row {a['text']!r} source {r[5]!r}, expected the email")
    head, fhead = (vals[0] if vals else []), (forms[0] if forms else [])
    if len(head) < 9 or str(head[7]).strip() != "Open items" or _num(head[8]) != len(acts):
        problems.append(f"I1 shows {head[8] if len(head) > 8 else None!r}, expected {len(acts)} open items")
    if len(fhead) < 9 or not str(fhead[8]).upper().startswith("=COUNTIF("):
        problems.append("I1 is not the live COUNTIF formula")
    charts = (rb or {}).get("charts")
    if not isinstance(charts, list):
        return False, "the Excel readback is not the tracker sheet's"
    if not any((ch or {}).get("title") == "Open items by owner" for ch in charts):
        problems.append(f"no 'Open items by owner' chart ({len(charts)} chart(s))")
    return not problems, "; ".join(problems) or f"{len(acts)} row(s) citing the email, COUNTIF={len(acts)}, chart present"


def story_brief_check(rb, c):
    dtitle = (c.get("decision") or {}).get("title") or c["fx"]["decision"]
    paras = [str(p).strip() for p in (rb or {}).get("paragraphs") or [] if str(p).strip()]
    starts = [i for i, p in enumerate(paras) if p == f"{dtitle} — brief"]
    if not starts:
        return False, "brief title paragraph missing"
    brief = paras[starts[-1]:]
    problems = [f"heading {h!r} missing" for h in ("Decisions", "Actions") if h not in brief]
    for it in (c.get("decision") or {}).get("items") or []:
        if not any(_flat(it["text"]) in _flat(p) for p in brief):
            problems.append(f"item {it['text']!r} missing")
    if not any(_flat("Source: " + EMAIL_SOURCE) in _flat(p) for p in brief):
        problems.append("no item cites the email as its source")
    if brief[-1] != f"Decision record: {dtitle}":
        problems.append("the brief is not the end of the document")
    links = (rb or {}).get("links")
    if links is None:
        problems.append("hyperlinks not readable in this Word build; decision-record link unverified")
    elif dtitle not in [str(x.get("text") or "").strip() for x in links if x.get("hyperlink")]:
        problems.append("no hyperlink to the decision record")
    return not problems, "; ".join(problems) or "brief at the end: headings, every item citing the email, decision-record link"


def story_calendar_receipt(tools, by):
    """The follow-up request, judged honest: brain_make_deliverables' calendar host ('applied', requested appointment,
    not saved or sent) or a calendar_draft_new result ('requested', not saved or sent). Returns (ok, note)."""
    cal = by.get("calendar")
    if cal:
        ok = (cal.get("status") == "applied" and cal.get("requested") == "appointment" and cal.get("saved") is False
              and cal.get("sent") is False)
        return ok, f"calendar host {cal.get('status')!r}, requested {cal.get('requested')!r}, saved {cal.get('saved')!r}"
    d = parsed(last_result(tools, "calendar_draft_new"))
    if _ok(d):
        x = d["data"]
        ok = x.get("status") == "requested" and x.get("saved") is False and x.get("sent") is False and "opened" not in x
        return ok, f"calendar_draft_new {x.get('status')!r}, saved {x.get('saved')!r}"
    return False, "no follow-up request (neither the calendar host nor calendar_draft_new)"


def story_deliverables_check(t, rbs, c):
    data, calls = deliverables_receipt(t["tools"])
    if data is None:
        return False, "brain_make_deliverables missing, truncated or not ok"
    by = {r.get("host"): r for r in data.get("results") or [] if isinstance(r, dict)}
    problems = [f"receipt {h}: {(by.get(h) or {}).get('status') or 'missing'}" for h in ("excel", "word", "powerpoint")
                if (by.get(h) or {}).get("status") != "verified"]
    o = by.get("outlook") or {}
    if not (o.get("status") == "applied" and o.get("sent") is False and o.get("requested") == "reply" and o.get("replyText")):
        problems.append(f"outlook reply receipt {o.get('status') or 'missing'} (applied, not sent, with replyText expected)")
    else:
        c["replyText"] = o["replyText"]
    ok, note = story_calendar_receipt(t["tools"], by)
    if not ok:
        problems.append(note)
    for name, (ok2, note2) in (("tracker", story_tracker_check(rbs.get("excel"), c)), ("brief", story_brief_check(rbs.get("word"), c)),
                               ("deck", deck_check(rbs.get("powerpoint"), data, c))):
        if not ok2:
            problems.append(f"{name}: {note2}")
    img = image_result(t)
    before, after = (c.get("img_before") or {}).get("images"), (rbs.get("powerpoint_images") or {}).get("images")
    if not (_ok(img) and img["data"].get("status") == "verified"):
        problems.append(f"office_insert_image {((img or {}).get('data') or {}).get('status') or 'missing or not ok'}")
    if not (isinstance(before, int) and after == before + 1):
        problems.append(f"pictures in the deck {before} -> {after} (exactly +1 expected)")
    n_img = sum(1 for x in t["tools"] if x.get("name") == "office_insert_image")
    if n_img != 1:
        problems.append(f"{n_img} office_insert_image calls (one expected)")
    return not problems, "; ".join(problems) or (f"tracker, brief, deck verified by readback; one picture; reply text and "
                                                 f"follow-up requested honestly ({note})"
                                                 + (f"; {calls} brain_make_deliverables calls" if calls > 1 else ""))


def story_why_check(t, rbs, c):
    d = parsed(last_result(t["tools"], "brain_why"))
    if not _ok(d):
        return False, "brain_why missing, truncated or not ok"
    email = (c.get("email") or {}).get("pageId")
    dtitle = (c.get("decision") or {}).get("title")
    chains = [b for a in d["data"].get("artifacts") or [] if a.get("host") == "excel" for b in a.get("because") or []]
    hit = [b for b in chains if "faq" in _flat(b.get("item")) and (b.get("decisionPage") or {}).get("title") == dtitle
           and (b.get("sourcePage") or {}).get("pageId") == email]
    problems = []
    if not hit:
        problems.append("no tracker provenance for the FAQ row -> this run's decision page -> the email entry")
    elif not hit[0].get("quote") or not hit[0].get("quoteVerified"):
        problems.append("the provenance has no verified quote")
    reply = _flat(t.get("text"))
    if ui.OUTLOOK_FIXTURE_TOPIC not in reply:
        problems.append("the answer does not cite the email")
    return not problems, "; ".join(problems) or "FAQ row -> decision page -> verified quote -> the email (cited in the answer)"


def story_reply_marker(c):
    return f"decision page '{(c.get('decision') or {}).get('title') or c['fx']['decision']}'"


def story_reply_check(t, rbs, c):
    d = parsed(last_result(t["tools"], "outlook_insert"))
    st, rb = c.get("reply_state") or {}, rbs.get("outlook_compose") or {}
    bad = [x["name"] for x in office_tools(t["tools"]) if x["name"] not in STORY_REPLY_ALLOWED]  # Office tools only (A18 r1:
    # Hermes core's tool_describe, a schema lookup, is not an Office action)
    marker = " ".join(story_reply_marker(c).split())
    problems = []
    if not (_ok(d) and d["data"].get("status") == "verified" and d["data"].get("sent") is False):
        problems.append(f"outlook_insert {((d or {}).get('data') or {}).get('status') or 'missing or not ok'}")
    if bad:
        problems.append(f"other tools called: {bad}")
    if not (st.get("hasMarker") and st.get("toIsOnlyAccount") and st.get("ccEmpty") and st.get("bccEmpty")):
        problems.append(f"the reply on screen: { {k: st.get(k) for k in ('hasMarker', 'toIsOnlyAccount', 'ccEmpty', 'bccEmpty')} }")
    if rb.get("mode") != "compose" or marker not in " ".join(str(rb.get("body") or "").split()) or rb.get("cc") or rb.get("bcc"):
        problems.append("the pane readback of the draft does not show the reply text (or Cc/Bcc not empty)")
    if c.get("still_draft") is not True:
        problems.append("the reply draft is no longer open (sent or closed?)")
    if not problems:
        c["reply_body_len"] = len(str(rb.get("body") or ""))
    return not problems, "; ".join(problems) or "reply text in the draft the person opened; To is the account only; unsent"


def story_undo_check(t, rbs, c):
    d = parsed(last_result(t["tools"], "office_undo"))
    rb, st = rbs.get("outlook_compose") or {}, c.get("undo_state") or {}
    bad = [x["name"] for x in office_tools(t["tools"]) if x["name"] not in STORY_UNDO_ALLOWED]
    marker = " ".join(story_reply_marker(c).split())
    calls = sum(1 for x in t["tools"] if x["name"] == "office_undo")
    problems = []
    if not _ok(d) or calls != 1:
        problems.append(f"{calls} office_undo call(s), last ok: {_ok(d)} (exactly one ok expected)")
    if bad:
        problems.append(f"other tools called: {bad}")
    if rb.get("mode") != "compose" or marker in " ".join(str(rb.get("body") or "").split()):
        problems.append("the draft still shows the reply text after the undo")
    if st.get("hasMarker") or not (st.get("toIsOnlyAccount") and st.get("ccEmpty") and st.get("bccEmpty")):
        problems.append(f"the reply on screen after undo: { {k: st.get(k) for k in ('hasMarker', 'toIsOnlyAccount', 'ccEmpty', 'bccEmpty')} }")
    if c.get("still_draft") is not True:
        problems.append("the reply draft is no longer open")
    return not problems, "; ".join(problems) or "one office_undo removed the reply text; the draft is still unsent"


def story_steps(bare=False):
    """The five beats after pairing. Prompts are docs/demo-script.md's, verbatim. Default: a bracketed note naming the
    tool and its run-specific arguments (titles, the decision page id, slide, picture path, follow-up times) is
    appended: that note IS tool routing. bare=True (A24, `story --bare`): the prompts exactly as the presenter types
    them, nothing appended, and every beat judged only on outcomes (STORY_BARE). A25: the email's indexing and its
    decision page are ONE beat, because Hermes does both in one approved call when asked to add the email to notes."""
    P = story_prompts()
    if len(P) != 5:
        raise ValueError(f"docs/demo-script.md lists {len(P)} prompts; the story expects 5")
    if bare:
        return [dict(step, prompt=(lambda c, i=i: P[i]), **STORY_BARE[step["name"]])
                for i, step in enumerate(story_steps(False))]
    return [
        {"name": "S1_email_decision_page", "show": "outlook", "then_show": "onenote", "read": ["outlook"],
         "post": _story_post_decision, "check": story_email_decision_check,
         "prompt": lambda c: P[0] + (f" (Call brain_from_email once with title '{c['fx']['decision']}', the sectionId of the "
                                     f"OneNote section '{c['fx']['section']}' from onenote_list_sections, and the email's "
                                     "decisions and actions, each with its exact quote.)")},
        {"name": "S2_deliverables", "needs": "S1_email_decision_page", "show": "powerpoint", "pre": _story_pre_deliverables,
         "after_turn": _after_deliverables, "read": ["excel", "word", "powerpoint", "powerpoint_images"],
         "post": _story_post_deliverables, "check": story_deliverables_check,
         "prompt": lambda c: P[1] + (f" (Call brain_make_deliverables once with decisionPageId '{c['decision']['pageId']}', hosts "
                                     "excel, word, powerpoint, outlook and calendar, and followUp start "
                                     f"'{c['fx']['followUp']['start']}', end '{c['fx']['followUp']['end']}', location 'Studio'. "
                                     f"Then call office_insert_image once with app 'powerpoint', path '{c['img']['path']}', "
                                     f"alt 'Team picture', slideIndex {c.get('slides0')}, left {STORY_PICTURE['left']}, top "
                                     f"{STORY_PICTURE['top']}, width {STORY_PICTURE['width']}, height {STORY_PICTURE['height']}.)")},
        {"name": "S3_why", "needs": "S2_deliverables", "show": "excel", "read": [], "check": story_why_check,
         "prompt": lambda c: P[2] + " (Call brain_why with host 'excel' and text 'FAQ'; give the decision, its quote and its source.)"},
        {"name": "S4_reply", "needs": "S2_deliverables", "show": "outlook", "pre": _story_pre_reply, "read": ["outlook_compose"], "post": _story_post_reply,
         "check": story_reply_check,
         "prompt": lambda c: P[3] + (" (Call outlook_insert once with target 'body' and text exactly the replyText from the "
                                     "brain_make_deliverables outlook result. Do not send.)")},
        {"name": "S5_undo", "needs": "S4_reply", "show": "outlook", "read": ["outlook_compose"], "post": _story_post_undo, "check": story_undo_check,
         "prompt": lambda c: P[4] + " (Call office_undo with app 'outlook' exactly once.)"},
    ]


# ---- story --bare (A24): the presenter's exact prompts, judged on outcomes only ------------------------------------
# No bracketed note, so nothing names a tool, a title, a page id, a picture path or the follow-up times. A beat passes
# on what the documents and the profile's index hold afterwards (read by the harness) and on the receipts the turn
# returned; which tools the model chose is recorded, never judged.
def story_index_db(db, since):
    """The email entry in the profile's brain.db (READ-ONLY) captured at or after `since` (epoch seconds)."""
    import sqlite3
    db = Path(db)
    if not db.exists():
        return {"error": f"no ledger file {db.name}"}
    try:
        con = sqlite3.connect(f"file:{db}?mode=ro", uri=True, timeout=5)
        try:
            rows = con.execute("select page_id, title, captured_at from pages where page_id like 'email:%' and title=?"
                               " order by captured_at desc", (EMAIL_SOURCE,)).fetchall()
        finally:
            con.close()
    except sqlite3.Error as exc:
        return {"error": f"{type(exc).__name__}: {str(exc)[:120]}"}
    fresh = [r for r in rows if (r[2] or 0) >= since - 1]
    return {"entries": len(rows), "fresh": [{"pageId": r[0], "title": r[1]} for r in fresh]}


def story_new_decision_page_db(db, since, email_page_id):
    """The newest decision page whose rows were recorded at or after `since` and cite the email entry (READ-ONLY):
    its id, title (pages table) and rows."""
    import sqlite3
    db = Path(db)
    if not db.exists():
        return {"error": f"no ledger file {db.name}"}
    try:
        con = sqlite3.connect(f"file:{db}?mode=ro", uri=True, timeout=5)
        try:
            ids = [r[0] for r in con.execute(
                "select decision_page_id from decisions where created_at >= ? and source_page_id = ?"
                " group by decision_page_id order by max(id) desc", (since - 1, email_page_id)).fetchall()]
            if not ids:
                return {"pages": 0}
            dp = ids[0]
            title = con.execute("select title from pages where page_id=? order by captured_at desc limit 1",
                                (dp,)).fetchone()
            rows = con.execute("select kind, text, owner, source_page_id, quote_verified from decisions where"
                               " decision_page_id=? order by id", (dp,)).fetchall()
        finally:
            con.close()
    except sqlite3.Error as exc:
        return {"error": f"{type(exc).__name__}: {str(exc)[:120]}"}
    return {"pages": len(ids), "pageId": dp, "title": title[0] if title else None,
            "rows": [{"kind": k, "text": t, "owner": o, "source": src, "quoteVerified": bool(q)}
                     for k, t, o, src, q in rows]}


def _turn_data(t):
    """Every ok tool result's data in the turn (receipts, whatever tool returned them)."""
    out = []
    for x in t.get("tools") or []:
        d = parsed(x.get("result"))
        if d is not None and d.get("ok") is True and isinstance(d.get("data"), dict):  # a composite with an
            out.append(d["data"])                                     # 'applied' host is still a receipt
            out += [r for r in d["data"].get("results") or [] if isinstance(r, dict)]
    return out


def _sent_anything(t):
    return any(d.get("sent") is True or d.get("saved") is True for d in _turn_data(t))


def story_named_picture(receipts):
    """A26: a picture receipt (verified or applied) whose source is the local file the prompt names. The plugin's
    path source records only the file name it opened (images.py: {"type": "path", "name": ...}); a generated, found
    (url) or inline picture has another source and does not count."""
    return any(isinstance(d.get("source"), dict) and d["source"].get("type") == "path"
               and d["source"].get("name") == STORY_IMAGE_NAME and d.get("status") in ("verified", "applied")
               for d in receipts)


def _pre_bare_beat(tps, c):
    c["beat_since"] = time.time()


def bare_email_check(t, rbs, c):
    db = rbs.get("db") or {}
    problems = []
    if db.get("error"):
        problems.append(f"brain.db unreadable ({db['error']})")
    elif not db.get("fresh"):
        problems.append(f"no index entry {EMAIL_SOURCE!r} captured during this beat ({db.get('entries')} older)")
    rb = rbs.get("outlook") or {}
    if rb.get("mode") != "read" or rb.get("subjectIsFixture") is not True:
        problems.append(f"the Outlook pane is not on the fixture email in read mode ({rb})")
    if not problems:
        e = db["fresh"][0]
        c["email"] = {"pageId": e["pageId"], "title": e["title"]}
    return not problems, "; ".join(problems) or f"the email is in the index as {EMAIL_SOURCE!r}, captured this beat"


def _bare_post_index(tps, t, c):
    return {"db": story_index_db(ledger_path(), c.get("beat_since") or 0)}


def _bare_post_email_decision(tps, t, c):
    """S1 (A25): the email's fresh index entry, then the decision page citing it, read from the index and the page."""
    idx = story_index_db(ledger_path(), c.get("beat_since") or 0)
    out = {"index": idx}
    fresh = idx.get("fresh") or []
    if fresh:
        c["email"] = {"pageId": fresh[0]["pageId"], "title": fresh[0]["title"]}
        out.update(_bare_post_decision(tps, t, c))
    return out


def bare_email_decision_check(t, rbs, c):
    """The email is in the index (captured this beat) and a decision page made from it, citing it, exists and reads
    back; which tool or how many calls made them is not judged."""
    ok1, n1 = bare_email_check(t, {**rbs, "db": rbs.get("index") or {}}, c)
    if not ok1:
        return False, n1
    ok2, n2 = bare_decision_check(t, rbs, c)
    return ok2, f"{n1}; {n2}"


def _bare_post_decision(tps, t, c):
    db = story_new_decision_page_db(ledger_path(), c.get("beat_since") or 0, (c.get("email") or {}).get("pageId"))
    out = {"db": db}
    title = db.get("title")
    if not title:
        return out
    tp = tps["onenote"]
    went = None
    for section in (c["fx"].get("section"), None):  # the run's section first; else wherever the model put it
        try:
            went = tp.run_js(ui.task_pane_frame(tp, "onenote"), onenote_goto_js(title, section), timeout=60)
        except ui.FrameScriptTimeout as exc:
            raise ui.FrameScriptTimeout(f"harness: the OneNote page-open eval for {title!r} did not settle; not "
                                        f"retried. {exc}") from exc
        if (went or {}).get("opened"):
            break
    time.sleep(4)
    out.update({"onenote_decision": readback(tp, "onenote"), "goto": went})
    return out


def bare_decision_check(t, rbs, c):
    """A decision page made from the email, as the index and the page itself hold it: decisions about the beta and the
    pilot scope, the FAQ action owned by Sam, every row citing the email with a verified quote, all on the page."""
    db = rbs.get("db") or {}
    if db.get("error"):
        return False, f"brain.db unreadable ({db['error']})"
    rows = db.get("rows") or []
    if not rows:
        return False, "no decision page citing the email was recorded during this beat"
    items = [{"kind": r["kind"], "text": r["text"], "owner": r["owner"] or "Unknown"} for r in rows]
    c["decision"] = {"title": db.get("title"), "pageId": db.get("pageId"), "items": items}
    problems = []
    decs = [_flat(i["text"]) for i in items if i["kind"] == "decision"]
    for want in ui.OUTLOOK_FIXTURE_DECISIONS:
        if not any(want in x for x in decs):
            problems.append(f"no decision about {want!r}")
    faq = [i for i in items if i["kind"] == "action" and "faq" in _flat(i["text"])]
    if not faq or _flat(faq[0]["owner"]) != "sam":
        problems.append("the FAQ action is missing or not owned by Sam (the email says so)")
    if any(not r["quoteVerified"] for r in rows):
        problems.append("a quote was not found in the email")
    rb = rbs.get("onenote_decision") or {}
    page = _flat(" ".join(rb.get("text") or []))
    if rb.get("pageTitle") != db.get("title"):
        problems.append(f"could not open the decision page {db.get('title')!r} to read it")
    else:
        if _flat("Source: " + EMAIL_SOURCE) not in page:
            problems.append("the page's Source line does not name the email")
        missing = [i["text"] for i in items if _flat(i["text"]) not in page]
        if missing:
            problems.append(f"items not on the page: {missing}")
        if _flat("(not found in the source page)") in page:
            problems.append("the page shows an unverified quote")
    if db.get("pages", 0) > 1:
        problems.append(f"{db['pages']} decision pages made in one beat")
    return not problems, "; ".join(problems) or (f"decision page {db.get('title')!r}: {len(items)} items, every quote "
                                                  "verified and cited to the email (index + page)")


def bare_deliverables_check(t, rbs, c):
    """Tracker, brief and deck as the documents hold them (plus their receipts), exactly one new picture in the deck
    whose receipt names the file the prompt names (A26), an honest reply and follow-up request, and nothing sent or
    saved."""
    data, calls = deliverables_receipt(t["tools"])
    problems = []
    if data is None:
        return False, "no tracker/brief/deck receipt in this turn"
    by = {r.get("host"): r for r in data.get("results") or [] if isinstance(r, dict)}
    problems += [f"receipt {h}: {(by.get(h) or {}).get('status') or 'missing'}" for h in ("excel", "word", "powerpoint")
                 if (by.get(h) or {}).get("status") != "verified"]
    for name, (ok2, note2) in (("tracker", story_tracker_check(rbs.get("excel"), c)), ("brief", story_brief_check(rbs.get("word"), c)),
                               ("deck", deck_check(rbs.get("powerpoint"), data, c))):
        if not ok2:
            problems.append(f"{name}: {note2}")
    receipts = _turn_data(t)
    reply = [d for d in receipts if d.get("requested") == "reply" or (d.get("status") == "requested" and "saved" not in d
                                                                       and d.get("sent") is False)]
    if not reply:
        problems.append("no reply drafted or requested (no honest reply receipt)")
    elif reply[-1].get("replyText"):
        c["replyText"] = reply[-1]["replyText"]
    ok, note = story_calendar_receipt(t["tools"], by)
    if not ok:
        problems.append(note)
    if _sent_anything(t):
        problems.append("a receipt says something was sent or saved")
    before, after = (c.get("img_before") or {}).get("images"), (rbs.get("powerpoint_images") or {}).get("images")
    if not (isinstance(before, int) and after == before + 1):
        problems.append(f"pictures in the deck {before} -> {after} (exactly +1 expected)")
    if not story_named_picture(receipts):
        problems.append(f"no picture receipt names {STORY_IMAGE_NAME} from the images folder (the prompt names that file)")
    return not problems, "; ".join(problems) or ("tracker, brief, deck verified by readback; one new picture "
                                                 f"({STORY_IMAGE_NAME}); reply and "
                                                 f"follow-up requested honestly ({note}); nothing sent or saved"
                                                 + (f"; {calls} deliverables calls" if calls > 1 else ""))


def bare_why_check(t, rbs, c):
    """The answer itself: it names the email it came from and the decision page, and quotes the FAQ action."""
    reply = _flat(t.get("text"))
    dtitle = _flat((c.get("decision") or {}).get("title"))
    problems = []
    if ui.OUTLOOK_FIXTURE_TOPIC not in reply:
        problems.append("the answer does not name the email")
    if "faq" not in reply:
        problems.append("the answer does not mention the FAQ action")
    if not (dtitle and dtitle in reply) and "decision page" not in reply:
        problems.append("the answer does not name the decision page")
    return not problems, "; ".join(problems) or "the answer traces the FAQ row to the decision page and the email"


def _bare_pre_reply(tps, c):
    _story_pre_reply(tps, c)
    try:
        c["reply_before"] = outlook_read(tps["outlook"], "compose") or {}
    except Exception as exc:  # judged below: an unreadable baseline fails the beat, it is not guessed
        c["reply_before"] = {"error": f"{type(exc).__name__}: {str(exc)[:120]}"}


def _added_text(before, after):
    b, a = str(before or ""), str(after or "")
    if b and b in a:
        return " ".join(a.replace(b, "", 1).split())
    return " ".join(a.split()) if not b.strip() else ""


def _bare_post_reply(tps, t, c):
    op = tps["outlook"]
    time.sleep(2)
    after = outlook_read(op, "compose") or {}
    added = _added_text((c.get("reply_before") or {}).get("body"), after.get("body"))
    c["reply_added"] = added
    c["reply_state"] = ui.reply_draft_state(op, added[:60]) if added else ui.reply_draft_state(op, "\u0000")
    c["still_draft"] = bool(op.run_js("top", ui.DRAFT_OPEN_JS))
    return {"outlook_compose": after}


def bare_reply_check(t, rbs, c):
    """The draft the person opened now holds new text about the decisions, still addressed to the account only, unsent."""
    added, st, rb = c.get("reply_added") or "", c.get("reply_state") or {}, rbs.get("outlook_compose") or {}
    problems = []
    if (c.get("reply_before") or {}).get("error"):
        problems.append(f"the draft could not be read before the turn ({c['reply_before']['error']})")
    if len(added) < 20:
        problems.append(f"no reply text was added to the draft ({len(added)} characters)")
    elif not any(w in _flat(added) for w in ("beta", "pilot", "faq")):
        problems.append("the added text does not mention the decisions or the action")
    if not (st.get("hasMarker") and st.get("toIsOnlyAccount") and st.get("ccEmpty") and st.get("bccEmpty")):
        problems.append(f"the reply on screen: { {k: st.get(k) for k in ('hasMarker', 'toIsOnlyAccount', 'ccEmpty', 'bccEmpty')} }")
    if rb.get("mode") != "compose" or rb.get("cc") or rb.get("bcc"):
        problems.append("the pane readback is not the draft (or Cc/Bcc not empty)")
    if c.get("still_draft") is not True or _sent_anything(t):
        problems.append("the reply draft is no longer open, or a receipt says it was sent")
    return not problems, "; ".join(problems) or f"{len(added)} characters of reply text in the draft; To is the account only; unsent"


def _bare_post_undo(tps, t, c):
    op = tps["outlook"]
    time.sleep(2)
    after = outlook_read(op, "compose") or {}
    added = c.get("reply_added") or ""
    c["undo_state"] = ui.reply_draft_state(op, added[:60]) if added else {}
    c["still_draft"] = bool(op.run_js("top", ui.DRAFT_OPEN_JS))
    return {"outlook_compose": after}


def bare_undo_check(t, rbs, c):
    """The inserted reply text is gone from the draft, which is still open and unsent."""
    added, rb, st = c.get("reply_added") or "", rbs.get("outlook_compose") or {}, c.get("undo_state") or {}
    problems = []
    if not added:
        problems.append("nothing to look for: the reply beat added no text")
    elif rb.get("mode") != "compose" or added[:60] in " ".join(str(rb.get("body") or "").split()):
        problems.append("the draft still shows the reply text")
    if st.get("hasMarker") or not (st.get("toIsOnlyAccount") and st.get("ccEmpty") and st.get("bccEmpty")):
        problems.append(f"the reply on screen after undo: { {k: st.get(k) for k in ('hasMarker', 'toIsOnlyAccount', 'ccEmpty', 'bccEmpty')} }")
    if c.get("still_draft") is not True or _sent_anything(t):
        problems.append("the reply draft is no longer open, or a receipt says it was sent")
    return not problems, "; ".join(problems) or "the reply text is gone; the draft is still open and unsent"


def _bare_pre_deliverables(tps, c):
    _story_pre_deliverables(tps, c)


STORY_BARE = {
    "S1_email_decision_page": {"pre": _pre_bare_beat, "post": _bare_post_email_decision, "check": bare_email_decision_check},
    "S2_deliverables": {"pre": _bare_pre_deliverables, "check": bare_deliverables_check},
    "S3_why": {"check": bare_why_check},
    "S4_reply": {"pre": _bare_pre_reply, "read": [], "post": _bare_post_reply, "check": bare_reply_check},
    "S5_undo": {"read": [], "post": _bare_post_undo, "check": bare_undo_check},
}


def story_read(tps, host, c):
    if host == "outlook":
        return outlook_read(tps["outlook"], "read")
    if host == "outlook_compose":
        return outlook_read(tps["outlook"], "compose")
    if host == "powerpoint_images":
        return image_readback(tps["powerpoint"], "powerpoint")
    return loop_read(tps, host, c)


def _story_post_decision(tps, t, c):
    out = _post_decision(tps, t, {"fx": {"decision": c["fx"]["decision"]}}, section=c["fx"].get("section"))
    data, _ = _story_decision_data(t)
    out["db"] = story_decisions_db(ledger_path(), (data or {}).get("pageId")) if data else {"error": "no decision page"}
    return out


def _story_pre_deliverables(tps, c):
    c["slides0"] = loop_read(tps, "powerpoint", c).get("count")
    c["img_before"] = image_readback(tps["powerpoint"], "powerpoint")
    tps["outlook"].run_js("top", ui.POPOUT_RECORDER_JS)


def _discard_outlook_extras(op, log):
    """Close what Outlook may have opened inline for a requested draft (never saved, never sent)."""
    out = {"eventForm": bool(op.run_js("top", ui.EVENT_FORM_OPEN_JS)), "draft": bool(op.run_js("top", ui.DRAFT_OPEN_JS))}
    if out["eventForm"]:
        out["eventFormDiscarded"] = ui.discard_event_form(op, log)
    if op.run_js("top", ui.DRAFT_OPEN_JS):
        ui.outlook_discard(op, log)
        out["draftDiscarded"] = True
    return out


def _story_post_deliverables(tps, t, c):
    op = tps["outlook"]
    time.sleep(5)
    c["popouts"] = op.run_js("top", ui.POPOUT_COUNT_JS)
    c["outlook_after_deliverables"] = _discard_outlook_extras(op, op.log)
    return {}


def _story_pre_reply(tps, c):
    """The person goes back to the email and clicks Reply; the pane opens in the draft (it resumes or is re-paired)."""
    op = tps["outlook"]
    _discard_outlook_extras(op, op.log)
    ui.open_fixture_message(op, op.log, c["mail"]["folder"])
    ui.outlook_reply(op, op.log)
    ui.open_outlook_pane(op, op.log, 90)
    st = ui.wait_pane_state(op, "outlook", ui.PAIRABLE_STATES, 90)
    c["reply_pane"] = st.get("state")


def _story_post_reply(tps, t, c):
    op = tps["outlook"]
    time.sleep(2)
    c["reply_state"] = ui.reply_draft_state(op, story_reply_marker(c))
    c["still_draft"] = bool(op.run_js("top", ui.DRAFT_OPEN_JS))
    return {}


def _story_post_undo(tps, t, c):
    op = tps["outlook"]
    time.sleep(2)
    c["undo_state"] = ui.reply_draft_state(op, story_reply_marker(c))
    c["still_draft"] = bool(op.run_js("top", ui.DRAFT_OPEN_JS))
    return {}


def new_story_row(run_no, bare=False):
    return {"host": "story-bare" if bare else "story", "run": run_no, "started": now(), "steps": [], "pass": False}


def story_beats(result):
    """Per beat: PASS/FAIL, step and turn time (what docs/live-evidence.md records)."""
    return [{"beat": s["step"], "pass": bool(s.get("pass")), "stepMs": s.get("stepMs"), "turnMs": s.get("turnMs")}
            for s in result.get("steps") or [] if str(s.get("step", "")).startswith("S")]


async def _story_step(chat, tps, step, c, rec, result, idx, run_no):
    s = {"step": step["name"]}
    need = step.get("needs")
    if need and not any(x.get("step") == need and x.get("pass") for x in result["steps"]):
        # its prompt names what that beat made: report it as not run, not as a second failure (A20)
        s.update({"pass": False, "skipped": f"not run: needs {need}, which did not pass", "stepMs": 0})
        result["steps"].append(s)
        rec.save()
        return s
    t0 = time.monotonic()
    turn, rbs = None, {}
    try:
        await asyncio.to_thread(tps[step["show"]].show)
        s["visible"] = step["show"]
        if step.get("pre"):
            await asyncio.to_thread(step["pre"], tps, c)
        if step["name"] == "S4_reply":
            s["pane"] = c.get("reply_pane")
            if c.get("reply_pane") != "paired":  # the draft's pane is a new pane: it resumes, else pair again
                s["repaired"] = await outlook_pair(chat, tps["outlook"], rec, s)
        text = step["prompt"](c)
        tps["onenote"].cmd("chat", cls="user", text=text[:300])
        n0 = len(chat.approvals)
        turn = await chat.turn(text)
        turn["approvals"] = chat.approvals[n0:]
        tps["onenote"].cmd("chat", cls="agent", text=mask_addresses(ui_redact(turn["text"]))[:600])
        if step.get("after_turn"):
            step["after_turn"](turn, c)
        for h in step.get("read") or ():
            rbs[h] = await asyncio.to_thread(story_read, tps, h, c)
        if step.get("post"):
            rbs.update(await asyncio.to_thread(step["post"], tps, turn, c))
        ok, note = step["check"](turn, rbs, c)
        cut = truncated_results(office_tools(turn["tools"]))
        if cut:
            ok, note = False, f"{note}; truncated result(s) not accepted as evidence: {cut}"
        s.update({"pass": bool(ok), "note": mask_addresses(note), "tools": [x["name"] for x in office_tools(turn["tools"])],
                  "toolOk": [bool(_ok(x["result"])) for x in office_tools(turn["tools"])],
                  "timing": timings(office_tools(turn["tools"])), "turnMs": turn.get("ms"), "approvals": turn["approvals"]})
        if step["name"] == "S2_deliverables":
            s["outlook"] = {"popouts": c.get("popouts"), **(c.get("outlook_after_deliverables") or {})}
        if not ok:
            s["readback"] = mask_addresses({k: _short(v) for k, v in rbs.items()})
            s["results"] = mask_addresses([_short(x["result"]) for x in office_tools(turn["tools"])])
    except Exception as exc:  # recorded; later beats say what they could not judge
        s.update({"pass": False, "error": mask_addresses(f"{type(exc).__name__}: {str(exc)[:600]}"), "where": error_where(exc)})
        if isinstance(exc, ui.FrameScriptTimeout):
            s["fault"] = "harness"
        if rbs:
            s["readback"] = mask_addresses({k: _short(v) for k, v in rbs.items()})
        if turn is not None:
            s["results"] = mask_addresses([_short(x["result"]) for x in office_tools(turn.get("tools") or [])])
    s["stepMs"] = int((time.monotonic() - t0) * 1000)
    try:
        s["shot"] = fetch_shot(tps[step.get("then_show") or step["show"]], rec, f"story-r{run_no}-{idx:02d}-{step['name']}.png")
    except Exception:
        pass
    result["steps"].append(s)
    rec.save()
    return s


async def story_once(client, pane, run_no, rec, manifests, bare=False):
    """One full demo story: five tabs, one chat, five beats, independent readbacks. Returns the results row.
    bare=True: the presenter's exact prompts, outcome-only checks (story_steps)."""
    result = rec.add(new_story_row(run_no, bare))
    t_start = time.monotonic()
    tps = {h: ui.TabPane(pane.base, rec.log, h) for h in STORY_HOSTS}
    op = tps["outlook"] = ui.TabPane(pane.base, rec.log, ui.OUTLOOK_TAB)
    c = {"run": run_no, "fx": story_fixture(run_no),
         "img": story_image(Path(os.environ.get("HERMES_HOME") or Path.home() / ".hermes"))}
    result["fixture"] = {k: c["fx"][k] for k in ("section", "tag", "decision", "followUp")}
    result["image"] = {k: c["img"][k] for k in ("rel", "sha256", "bytes")}
    try:
        for h in STORY_HOSTS:  # 1. four Office tabs, each with its document, add-in and pane
            s = {"step": f"P0_prepare_{h}"}
            t0 = time.monotonic()
            try:
                await asyncio.to_thread(tps[h].open, "about:blank")
                await asyncio.to_thread(tps[h].show)
                await asyncio.to_thread(prepare_host, tps[h], h, rec, manifests, run_no,
                                        {"section": c["fx"]["section"]} if h == "onenote" else None, LOOP_RESET.get(h))
                s["pass"] = True
            except Exception as exc:
                s.update({"pass": False, "error": f"{type(exc).__name__}: {str(exc)[:300]}"})
            s["stepMs"] = int((time.monotonic() - t0) * 1000)
            result["steps"].append(s)
            rec.save()
            if not s["pass"]:
                raise ui.HarnessError(f"{h} tab could not be prepared: {s['error']}")
        s = {"step": "P0_prepare_outlook"}  # 2. the Outlook tab on the existing fixture email, pane open
        t0 = time.monotonic()
        try:
            if not await asyncio.to_thread(ui.ensure_outlook_signed_in, op, rec.log, 1800):
                raise ui.HarnessError("Outlook sign-in gate: the mailbox never appeared")
            c["mail"] = await asyncio.to_thread(find_fixture, op)
            s["fixture"] = c["mail"]
            await asyncio.to_thread(ui.open_fixture_message, op, rec.log, c["mail"]["folder"])
            s["addin"] = await asyncio.to_thread(ui.install_outlook_addin, op, manifests["outlook"], rec.log)
            ui.close_outlook_pane(op)
            await asyncio.to_thread(ui.open_outlook_pane_retry, op, rec.log, 90, 60,
                                    lambda: ui.open_fixture_message(op, rec.log, c["mail"]["folder"]))
            await asyncio.to_thread(ui.wait_pane_state, op, "outlook", ui.PAIRABLE_STATES, 90)
            s["pass"] = True
        except Exception as exc:
            s.update({"pass": False, "error": mask_addresses(f"{type(exc).__name__}: {str(exc)[:300]}")})
        s["stepMs"] = int((time.monotonic() - t0) * 1000)
        result["steps"].append(s)
        rec.save()
        if not s["pass"]:
            raise ui.HarnessError(f"outlook tab could not be prepared: {s['error']}")
        result["tabs"] = await asyncio.to_thread(pane.cmd, "tab", op="close", name="main")  # the five story tabs only
        chat = await Chat.create(client, f"HOL demo story{' (bare)' if bare else ''} run {run_no} "
                                         f"{time.strftime('%m%d-%H%M%S')}", rec.log)
        result["chat"] = chat.sid[:18]
        tps["onenote"].cmd("chat", clear=True, cls="note", text=f"one chat {chat.sid[:10]}… for all five tabs")
        for i, h in enumerate(("outlook",) + STORY_HOSTS, start=1):  # 3. pair every tab to this one chat
            s = {"step": f"P{i}_pair_{h}"}
            t0 = time.monotonic()
            try:
                await asyncio.to_thread(tps[h].show)
                if h == "outlook":
                    s["pass"] = await outlook_pair(chat, op, rec, s)
                else:
                    t = await chat.turn(f"connect to {LABEL[h]}")
                    code, source = pairing_code(t["text"], t["tools"])
                    st = await asyncio.to_thread(ui.type_code, tps[h], h, code, rec.log) if code else {"state": "no-code"}
                    s.update({"pass": _ok(last_result(t["tools"], "office_connect")) and source == "reply" and st.get("state") == "paired",
                              "codeSource": source, "pane": st.get("state"), "turnMs": t.get("ms")})
            except Exception as exc:
                s.update({"pass": False, "error": f"{type(exc).__name__}: {str(exc)[:300]}"})
            s["stepMs"] = int((time.monotonic() - t0) * 1000)
            result["steps"].append(s)
            rec.save()
            if not s["pass"]:
                raise ui.HarnessError(f"pairing {h} failed: {s}")
        for i, step in enumerate(story_steps(bare), start=1):  # 4. the story, in this one chat
            await _story_step(chat, tps, step, c, rec, result, i, run_no)
        result["approvalsAnswered"] = chat.approvals
        result["pass"] = all(s.get("pass") for s in result["steps"])
    except Exception as exc:
        result["error"] = mask_addresses(f"{type(exc).__name__}: {str(exc)[:400]}")
        result["trace"] = mask_addresses(traceback.format_exc()[-1500:])
    finally:
        try:  # nothing left behind in Outlook: no draft (never sent), no form (never saved), no pane
            result["outlookCleanup"] = await asyncio.to_thread(_discard_outlook_extras, op, rec.log)
            ui.close_outlook_pane(op)
        except Exception as exc:
            result["outlookCleanup"] = {"error": f"{type(exc).__name__}: {str(exc)[:200]}"}
        try:  # the window as the single-document modes expect: a 'main' tab; this run's Office panes unload
            pane.cmd("tab", op="open", name="main", url="https://excel.cloud.microsoft/")
            pane.cmd("tab", op="show", name="main")
            for h in STORY_HOSTS:
                pane.cmd("tab", op="close", name=h)
        except Exception as exc:
            result["cleanup"] = f"{type(exc).__name__}: {str(exc)[:200]}"
    result["beats"] = story_beats(result)
    result["timing"] = {"totalMs": int((time.monotonic() - t_start) * 1000),
                        "storyTurnsMs": sum(b["turnMs"] or 0 for b in result["beats"]),
                        "storyStepsMs": sum(b["stepMs"] or 0 for b in result["beats"]),
                        "perStepMs": {s["step"]: s.get("stepMs") for s in result["steps"]}}
    result["ended"] = now()
    rec.save()
    return result


def summarize(results):
    rows = {}
    for r in results:
        if r["run"] == "F3":
            continue
        h = rows.setdefault(r["host"], {"runs": 0, "passed": 0, "elapsed": [], "round": [], "steps": {}})
        h["runs"] += 1
        h["passed"] += int(r["pass"])
        for s in r["steps"]:
            h["steps"].setdefault(s["step"], []).append(bool(s.get("pass")))
            for t in s.get("timing", []):
                if isinstance(t.get("elapsedMs"), (int, float)):
                    h["elapsed"].append(t["elapsedMs"])
                if isinstance(t.get("roundTripMs"), (int, float)):
                    h["round"].append(t["roundTripMs"])
    for h in rows.values():
        h["elapsedMsMedian"] = statistics.median(h["elapsed"]) if h["elapsed"] else None
        h["roundTripMsMedian"] = statistics.median(h["round"]) if h["round"] else None
        del h["elapsed"], h["round"]
    return rows


GUEST_WAIT_S = 90  # bounded wait for the harness window's webview guest before the first navigate
EXIT_GUI_UNAVAILABLE = 4  # distinct from 2 (chat error) and 3 (sign-in gate)


async def main_async(a):
    rec = Recorder(Path(a.out))
    story = getattr(a, "story", False)
    gate = cold_gate(rec.identity, story and getattr(a, "bare", False))
    if rec.identity.get("pluginDirMatches") is False:  # HOL_PLUGIN_COMMIT named a commit the plugin dir is not
        gate = f"PLUGIN DIR: not git archive {rec.identity['pluginCommit'][:12]}: {rec.identity['pluginDirCheck']}"
    if gate:
        print(gate, flush=True)
        (rec.out / "summary.json").write_text(json.dumps({"blocked": "cold gate", "detail": gate, "identity": rec.identity}, indent=2))
        return 4
    if story:  # the named picture is in place (or the reason it cannot be) before any Office tab is touched
        try:
            story_image(Path(os.environ.get("HERMES_HOME") or Path.home() / ".hermes"))
        except ui.HarnessError as exc:
            print(f"STORY PICTURE: {exc}", flush=True)
            (rec.out / "summary.json").write_text(json.dumps({"blocked": "story picture", "detail": str(exc)}, indent=2))
            return 2
    pane = ui.Pane(a.control, rec.log)
    try:
        pane.wait_guest(timeout=GUEST_WAIT_S)
    except ui.GuestUnavailable as exc:
        print(f"GUI GATE: {exc}", flush=True)
        (rec.out / "summary.json").write_text(json.dumps({"blocked": ui.GUI_UNAVAILABLE, "detail": str(exc)}, indent=2))
        return EXIT_GUI_UNAVAILABLE
    manifests = {h: f"/Users/{a.mac_user}/{MAC_ROOT}/manifests/manifest-{h}.xml" for h in LABEL}
    pane.cmd("clearCache")  # run the add-in bytes that are served now, not a heuristically cached copy
    # Office steps drive the visible tab: make it the Office tab, never the Outlook one (its own account and jar).
    if "main" in {t.get("name") for t in (pane.cmd("tab", op="list") or {}).get("tabs") or []}:
        pane.cmd("tab", op="show", name="main")
    pane.cmd("navigate", url="https://excel.cloud.microsoft/", timeout=120)  # a known start, not a leftover page
    time.sleep(5)
    state = ui.signin_state(pane)
    rec.log({"ev": "start", "hosts": a.hosts, "runs": a.runs, "signin": state})
    if state != "signed-in":
        if a.signin_timeout <= 0 or not ui.wait_signed_in(pane, a.signin_timeout, rec.log):
            print("SIGN-IN GATE: not signed in; nothing was run.")
            (rec.out / "summary.json").write_text(json.dumps({"blocked": "sign-in gate"}, indent=2))
            return 3
    async with connect_live() as client:
        for host in a.hosts:
            if host == "calendar":  # all its runs, then the test event's deletion in a finally (A15)
                await calendar_host(client, a.control, a.runs, rec, manifests)
                continue
            for n in range(1, a.runs + 1):
                if host == "outlook":
                    r = await outlook_run_once(client, a.control, n, rec, manifests)
                    print(f"{host} run {n}: {'PASS' if r['pass'] else 'FAIL'} "
                          + " ".join(f"{s['step']}={'ok' if s.get('pass') else 'X'}" for s in r["steps"])
                          + (f"  error={r['error']}" if r.get("error") else ""), flush=True)
                    continue
                if "main" in {t.get("name") for t in (pane.cmd("tab", op="list") or {}).get("tabs") or []}:
                    pane.cmd("tab", op="show", name="main")  # after an Outlook run the visible tab may be Outlook's
                fixtures = image_fixtures(Path(os.environ.get("HERMES_HOME") or Path.home() / ".hermes"), rec.out / "image-fixtures",
                                          f"{host}-r{n}-{time.strftime('%H%M%S')}") if a.images else None
                r = await run_once(client, pane, host, n, rec, manifests, images=fixtures)
                print(f"{host} run {n}: {'PASS' if r['pass'] else 'FAIL'} "
                      + " ".join(f"{s['step']}={'ok' if s.get('pass') else 'X'}" for s in r["steps"])
                      + (f"  error={r['error']}" if r.get("error") else ""), flush=True)
        if a.loop:
            for n in range(1, a.runs + 1):
                r = await loop_once(client, pane, n, rec, manifests)
                print(f"loop run {n}: {'PASS' if r['pass'] else 'FAIL'} "
                      + " ".join(f"{s['step']}={'ok' if s.get('pass') else 'X'}" for s in r["steps"])
                      + (f"  error={r['error']}" if r.get("error") else ""), flush=True)
        if a.story:
            for n in range(1, a.runs + 1):
                r = await story_once(client, pane, n, rec, manifests, bare=a.bare)
                print(f"{r['host']} run {n}: {'PASS' if r['pass'] else 'FAIL'} "
                      + " ".join(f"{s['step']}={'ok' if s.get('pass') else 'X'}" for s in r["steps"])
                      + (f"  error={r['error']}" if r.get("error") else ""), flush=True)
        if a.f3:
            r = await f3_regression(client, pane, rec, manifests)
            print("F3 regression:", "PASS" if r["pass"] else "FAIL", json.dumps(r["steps"], default=str)[:800], flush=True)
    summary = {**summarize(rec.results), "identity": rec.identity,
               "learnedAtEnd": learned_state(Path(os.environ.get("HERMES_HOME", "?")))}
    if rec.calendar_cleanup is not None:
        summary["calendarCleanup"] = rec.calendar_cleanup
    (rec.out / "summary.json").write_text(json.dumps(summary, indent=2))
    print(json.dumps(summary, indent=2))
    cleaned = rec.calendar_cleanup is None or calendar_cleanup_ok(rec.calendar_cleanup)
    return 0 if all(r["pass"] for r in rec.results) and cleaned else 1


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--control", default="http://127.0.0.1:19334")
    ap.add_argument("--hosts", default="excel,word,powerpoint,onenote")
    ap.add_argument("--runs", type=int, default=3)
    ap.add_argument("--f3", action="store_true")
    ap.add_argument("--loop", action="store_true", help="the whole docs/demo.md loop: one chat, four Office tabs in one window")
    ap.add_argument("--story", action="store_true",
                    help="the recorded demo story (docs/demo-script.md): one chat, Outlook + four Office tabs")
    ap.add_argument("--bare", action="store_true",
                    help="with --story: the presenter's exact prompts (no bracketed tool note), judged on outcomes only")
    ap.add_argument("--images", action="store_true", help="append the office_insert_image steps (path, url, html, negatives)")
    ap.add_argument("--signin-timeout", type=float, default=1800)
    ap.add_argument("--mac-user", default=os.environ.get("HOL_MAC_USER", "demo-user"))
    ap.add_argument("--out", required=True)
    a = ap.parse_args()
    a.hosts = [h for h in a.hosts.split(",") if h]
    try:
        return asyncio.run(main_async(a))
    except ChatError as exc:
        print(f"chat error: {exc}")
        return 2


if __name__ == "__main__":
    sys.exit(main())
