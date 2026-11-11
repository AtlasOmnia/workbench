# SPDX-License-Identifier: Apache-2.0
"""Live-test harness readbacks (A11 deferred, harness only; nothing here touches Office, a gateway or the network).

1. OneNote: loading richText on a paragraph collection that holds any other paragraph type (a picture, a table) fails
   the whole batch with "This paragraph is not of type RichText" (A11 live, images all 3, OneNote r2). The readback
   loads the type first and richText only on RichText paragraphs. Run in node against a fake that fails the same way.
2. Word for the web: "Cannot read properties of null (reading 'p5d')" right after the pane reset. In every recorded
   case (36afdac r3, e999e57 r2, c6d20e6 r2) Word showed its own banner "Sorry, your session expired. Please refresh
   the page to continue." The harness retries the READ only: it waits for the document to settle, refreshes an expired
   document once (a reload, no write) and never runs a write.
"""
import importlib.util
import json
import shutil
import subprocess
import sys
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parent.parent
LIVE = ROOT / "scripts" / "livetest"


def _driver():
    name = "hol_live_driver"
    if name in sys.modules:
        return sys.modules[name]
    sys.path.insert(0, str(LIVE))
    spec = importlib.util.spec_from_file_location(name, LIVE / "live_driver.py")
    mod = importlib.util.module_from_spec(spec)
    sys.modules[name] = mod
    spec.loader.exec_module(mod)
    return mod


# ---- 1. OneNote: richText only on RichText paragraphs ------------------------------------------------------
ONENOTE_FAKE = r"""
const NOT_RT = 'This paragraph is not of type RichText. Check the OfficeExtension.Error.debugInfo for more information.'
const pending = new Set()
class RichText {
  constructor(par, text) { this.par = par; this._text = text; this.loaded = false }
  load() { pending.add(this); return this }
  get text() { if (!this.loaded) throw new Error('PropertyNotLoaded: richText.text'); return this._text }
}
class Coll {
  constructor(items) { this._items = items; this.loaded = false; this.spec = '' }
  load(spec) { this.spec = spec || ''; pending.add(this); return this }
  get items() { if (!this.loaded) throw new Error('PropertyNotLoaded: items'); return this._items }
}
function par(type, text, kids = []) {
  const p = { type: null, _type: type, paragraphs: new Coll(kids) }
  p._rt = new RichText(p, text)
  Object.defineProperty(p, 'richText', { get() { return p._rt } })
  return p
}
const ctx = {
  async sync() {
    const batch = [...pending]; pending.clear()
    for (const x of batch) {  // like OneNote for the web: one bad load fails the WHOLE batch
      if (x instanceof Coll && x.spec.includes('richText') && x._items.some(i => i._type && i._type !== 'RichText')) throw new Error(NOT_RT)
      if (x instanceof RichText && x.par._type !== 'RichText') throw new Error(NOT_RT)
    }
    for (const x of batch) {
      x.loaded = true
      if (x instanceof Coll) for (const i of x._items) { if (i._type) i.type = i._type; if (x.spec.includes('richText')) i._rt.loaded = true }
    }
  },
}
const outline = new Coll([
  par('RichText', 'Q4 Launch notes', [par('RichText', 'item a'), par('Image', '')]),
  par('Image', ''),
  par('Table', ''),
  par('RichText', 'Links & Backlinks Links: Risks'),
])
const contents = new Coll([{ _type: 'Outline', outline: { id: 'o1', paragraphs: outline } }, { _type: 'Image' }])
const page = { id: 'p1', title: 'Q4 Launch (HOL r2 x)', contents, load() {} }
const nb = { id: 'nb1', name: 'HOL', load() {} }
ctx.application = { getActivePage: () => page, getActiveNotebook: () => nb }
const OneNote = { run: f => f(ctx) }
"""


def _node_eval(fake, script):
    node = shutil.which("node")
    if not node:
        pytest.skip("node not on PATH")
    prog = fake + "\n;(async () => { const out = await (" + script + "); console.log(JSON.stringify(out)) })()" \
        ".catch(e => console.log(JSON.stringify({ __error: String(e && e.message || e) })))"
    out = subprocess.run([node, "-e", prog], capture_output=True, text=True, timeout=30)
    assert out.returncode == 0, out.stderr
    return json.loads(out.stdout.strip().splitlines()[-1])


def test_onenote_fake_reproduces_the_live_failure_of_a_richtext_load_on_a_picture_paragraph():
    old = ("OneNote.run(async ctx => { const ps = ctx.application.getActivePage().contents; ps.load('items/type'); await ctx.sync();"
           " const p = ps.items[0].outline.paragraphs; p.load('items/type,items/richText/text'); await ctx.sync(); return true })")
    assert "not of type RichText" in _node_eval(ONENOTE_FAKE, old)["__error"]


def test_onenote_readback_loads_rich_text_only_on_rich_text_paragraphs():
    rb = _node_eval(ONENOTE_FAKE, _driver().READBACK["onenote"])
    assert "__error" not in rb, rb
    assert rb["pageTitle"] == "Q4 Launch (HOL r2 x)" and rb["notebookId"] == "nb1" and rb["outlines"] == 1
    assert [" ".join(t.split()) for t in rb["text"]] == ["Q4 Launch notes item a Links & Backlinks Links: Risks"]  # list item read
    assert rb["linksBlocks"] == 1


# ---- 2. Word for the web: bounded retry of the READ, refresh of an expired session -----------------------------
P5D = "RichApi.Error: Cannot read properties of null (reading 'p5d')"
TASKPANE = "https://studio.example/office/live/word/taskpane.html"
EDITOR = "https://word-edit.officeapps.live.com/we/wordeditorframe.aspx"


class FakeWord:
    """A harness pane over Word for the web. ``expired``: Word shows its session-expired banner and every Office.js read
    fails with the p5d error until the document is reloaded (``fixes_on_reload``). ``transient``: that many reads fail
    without a banner."""

    def __init__(self, expired=False, fixes_on_reload=True, transient=0):
        self.expired, self.fixes, self.transient = expired, fixes_on_reload, transient
        self.cmds, self.js, self.events, self.reads = [], [], [], 0
        self.log = self.events.append

    def frames(self):
        return ["top", EDITOR, TASKPANE]

    def cmd(self, c, timeout=90, **kw):
        self.cmds.append((c, kw))
        if c == "docs" and kw.get("op") == "open" and self.fixes:
            self.expired = False
        return {"value": True}

    def run_js(self, frame, js, timeout=90):
        d, ui = _driver(), _driver().ui
        self.js.append(js)
        if js == ui.SESSION_EXPIRED_JS:
            return self.expired and frame == EDITOR
        if "__hermesPaneState" in js:
            return {"state": "paired", "status": "Reconnected to your Hermes chat after the pane reloaded."}
        if "__HOL.find" in js:
            return frame == EDITOR
        if js in (d.READBACK["word"], d.SETTLE_JS["word"]):
            if js == d.READBACK["word"]:
                self.reads += 1
            if self.expired:
                raise ui.HarnessError(f"eval in {frame}: {P5D}")
            if js == d.READBACK["word"] and self.transient:
                self.transient -= 1
                raise ui.HarnessError(f"eval in {frame}: {P5D}")
            return [2, "HOL live test document."] if js == d.SETTLE_JS["word"] else {"text": "HOL live test document.", "paragraphs": 2, "bold": []}
        raise AssertionError(f"unexpected script run by the readback: {js[:80]}")


def _no_writes(p):
    d = _driver()
    writes = [j for j in p.js if j == d.RESET["word"] or "insertParagraph" in j or ".clear()" in j or "insert" in j.lower()]
    assert writes == [], writes


def _read(p):
    return _driver().readback(p, "word", waits=(0, 0, 0), sleep=lambda s: None)


def test_word_readback_refreshes_an_expired_session_once_and_reads_again():
    p = FakeWord(expired=True)
    rb = _read(p)
    assert rb["paragraphs"] == 2
    assert [c for c in p.cmds if c[0] == "docs"] == [("docs", {"op": "open", "host": "word"})]
    assert any(e.get("ev") == "session-expired" and e.get("host") == "word" for e in p.events)
    assert any(e.get("ev") == "session-refreshed" and e.get("pane") == "paired" for e in p.events)
    _no_writes(p)


def test_word_readback_reports_a_session_that_stays_expired_and_stops_bounded():
    p = FakeWord(expired=True, fixes_on_reload=False)
    with pytest.raises(_driver().ui.HarnessError, match="session expired") as exc:
        _read(p)
    assert "not a plugin result" in str(exc.value) and "p5d" in str(exc.value)
    assert [c[0] for c in p.cmds].count("docs") == 1  # refreshed once, not in a loop
    assert p.reads == 4  # the first read and three bounded retries
    _no_writes(p)


def test_word_readback_retries_a_transient_error_without_refreshing():
    p = FakeWord(transient=1)
    assert _read(p)["paragraphs"] == 2
    assert p.reads == 2 and not [c for c in p.cmds if c[0] == "docs"]
    assert any(e.get("ev") == "readback-retry" and e.get("settled") is True for e in p.events)
    _no_writes(p)
