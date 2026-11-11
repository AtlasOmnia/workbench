# SPDX-License-Identifier: Apache-2.0
"""Office-for-the-web UI helpers for the client Mac live harness (driver side).

Everything here goes through the harness control endpoint (scripts/livetest/macbook/main.js):
JS evaluated in a frame, trusted typing/clicks, the CDP file chooser and screenshots. Office's
UI is located by visible labels (aria-label / title / text), never by private ids, and a step
that cannot find its control fails with the visible control labels of that frame (document
pages only; nothing is dumped from the Office start page, which lists the user's files).
"""
from __future__ import annotations

import json
import re
import time
import urllib.request

# excel.new / word.new / powerpoint.new time out from this network, so blank documents come from each
# app's start page ("Create blank ..."); only that one control is looked up there, nothing is logged.
HOME_URLS = {"excel": "https://excel.cloud.microsoft/", "word": "https://word.cloud.microsoft/",
             "powerpoint": "https://powerpoint.cloud.microsoft/"}
BLANK_LABELS = {"excel": ["Create blank workbook", "^Blank workbook"], "word": ["Create blank document", "^Blank document"],
                "powerpoint": ["Create blank presentation", "^Blank presentation"]}
ONENOTE_HOME = "https://www.onenote.com/notebooks"
DEFAULT_TITLES = {"excel": "Book", "word": "Document", "powerpoint": "Presentation"}

# Injected before every UI query: visibility, labels, and a label search across clickable elements.
HELPERS = r"""
var __HOL = ((() => {
  const vis = el => { const r = el.getBoundingClientRect(); if (r.width < 2 || r.height < 2) return false;
    const s = getComputedStyle(el); return s.visibility !== 'hidden' && s.display !== 'none' && s.opacity !== '0' }
  const label = el => (el.getAttribute('aria-label') || el.getAttribute('title') || el.innerText || el.value || '').trim().replace(/\s+/g, ' ')
  const names = el => [el.getAttribute('aria-label'), el.getAttribute('title'), el.innerText, el.value].filter(Boolean).map(x => String(x).trim().replace(/\s+/g, ' '))
  const SEL = 'button,[role=button],[role=menuitem],[role=menuitemcheckbox],[role=menuitemradio],[role=tab],[role=link],a,[role=option],[role=treeitem],input[type=button],input[type=submit],label'
  const match = (txt, pat) => { const t = txt.toLowerCase(); if (pat.startsWith('re:')) return new RegExp(pat.slice(3), 'i').test(txt);
    if (pat.startsWith('^')) return t.startsWith(pat.slice(1).toLowerCase()); return t === pat.toLowerCase() }
  const find = (pats, sel) => { const els = [...document.querySelectorAll(sel || SEL)].filter(vis);
    for (const p of pats) { const hit = els.find(e => names(e).some(n => match(n, p))); if (hit) return hit } return null }
  const labels = (sel) => [...new Set([...document.querySelectorAll(sel || SEL)].filter(vis).map(label).filter(Boolean).map(s => s.slice(0, 60)))].slice(0, 120)
  return { vis, label, find, labels }
})());
"""


class HarnessError(RuntimeError):
    pass


GUI_UNAVAILABLE = "client Mac GUI session unavailable (locked/asleep/lid closed)"


class GuestUnavailable(HarnessError):
    """The harness window never got a webview guest: the client Mac's GUI session cannot draw (nothing was run)."""


FRAME_SCRIPT_LIMIT_S = 20  # main.js evalIn: a frame script that has not settled by then is abandoned


class FrameScriptTimeout(HarnessError):
    """A harness eval did not settle within the window's fixed frame-script limit (main.js evalIn). A harness
    fault, not a plugin result: whether the frame detached or the script was just slow is not known (A20)."""


class Pane:
    """Client of the client Mac harness control endpoint (reached through an ssh -L forward)."""

    def __init__(self, base: str, log):
        self.base = base.rstrip("/")
        self.log = log

    def cmd(self, cmd: str, timeout: float = 90, **kw):
        data = json.dumps({"cmd": cmd, **kw}).encode()
        req = urllib.request.Request(self.base + "/", data=data, headers={"Content-Type": "application/json"})
        with urllib.request.urlopen(req, timeout=timeout) as r:
            out = json.loads(r.read() or b"null")
        if isinstance(out, dict) and "__error" in out:
            raise HarnessError(f"{cmd}: {out['__error']}")
        return out

    def state(self):
        return self.cmd("state")

    def wait_guest(self, timeout: float = 90, poll: float = 1.0, clock=time.monotonic, sleep=time.sleep) -> dict:
        """Wait (bounded) until the harness window's start tab has a webview guest, before the first navigate.

        The window creates its <webview> right after launch, but a guest attaches only while the client Mac's GUI
        session can draw; a locked, asleep or lid-closed client Mac never attaches one (A5: 'navigate: no guest
        yet'). The control endpoint may also be unreachable for a moment while the window starts."""
        start = clock()
        last_note, why = start, "the window reported no guest"
        self.log({"ev": "guest-wait", "timeoutS": timeout})
        while True:
            try:
                st = self.cmd("state", timeout=min(10.0, max(1.0, timeout)))
                if isinstance(st, dict) and st.get("tab"):
                    self.log({"ev": "guest-attached", "tab": st.get("tab"), "waitedS": round(clock() - start, 1)})
                    return st
                why = "the window reported no guest"
            except (HarnessError, OSError) as exc:  # URLError, refused/reset: the window is still starting
                why = f"the harness control endpoint did not answer ({type(exc).__name__})"
            waited = clock() - start
            if waited >= timeout:
                self.log({"ev": "guest-unavailable", "waitedS": round(waited, 1), "why": why})
                raise GuestUnavailable(f"{GUI_UNAVAILABLE}: the harness window's webview guest did not attach within "
                                       f"{timeout:.0f} s ({why}); nothing was run. Wake and unlock the client Mac with its "
                                       "lid open and the user logged in at the screen, then rerun. If it already is, look for "
                                       "tab-open-failed in ~/hermes-office-livetest/logs/electron.log on the client Mac.")
            if clock() - last_note >= 15:
                last_note = clock()
                self.log({"ev": "guest-waiting", "waitedS": round(waited, 1), "why": why})
                print(f"waiting for the harness window's webview guest ({waited:.0f}/{timeout:.0f} s): {why}", flush=True)
            sleep(min(poll, max(0.0, timeout - waited)) or poll)

    def run_js(self, frame: str, js: str, timeout: float = 90):
        v = self.cmd("eval", frame=frame, js=js, timeout=timeout)["value"]
        if isinstance(v, dict) and "__error" in v:
            if str(v["__error"]).startswith("frame script timed out"):
                raise FrameScriptTimeout(f"eval in {frame}: {v['__error']}: the harness window's {FRAME_SCRIPT_LIMIT_S} s "
                                         "frame-script limit passed (a harness eval, not a plugin call)")
            raise HarnessError(f"eval in {frame}: {v['__error']}")
        return v

    def frames(self):
        return [f["url"] for f in self.state()["frames"]]

    def wait_frame(self, substr: str, timeout: float = 60) -> str:
        end = time.time() + timeout
        while time.time() < end:
            for u in self.frames():
                if substr in u:
                    return u
            time.sleep(1)
        raise HarnessError(f"no frame containing {substr!r} after {timeout:.0f}s")

    def click_label(self, patterns, timeout: float = 20, frames_like=None, sel=None):
        """Trusted (CDP) click on the first visible control whose label matches, in any frame (top first)."""
        js_pats = json.dumps(list(patterns))
        js = HELPERS + (f"(() => {{ const el = __HOL.find({js_pats}, {json.dumps(sel)}); if (!el) return null; el.scrollIntoView({{block: 'center'}});"
                        " const r = el.getBoundingClientRect(); return { x: r.x + r.width / 2, y: r.y + r.height / 2, label: __HOL.label(el) } })()")
        end = time.time() + timeout
        while time.time() < end:
            for u in ["top"] + [f for f in self.frames()[1:]]:
                if frames_like and u != "top" and not any(s in u for s in frames_like):
                    continue
                try:
                    hit = self.cmd("clickIn", frame=u, js=js, timeout=30)
                except HarnessError:
                    continue
                if hit.get("ok"):
                    label = hit.get("label") or ""
                    self.log({"ev": "click", "label": label[:60], "frame": u if u == "top" else u.split("/")[2]})
                    return label
            time.sleep(1)
        raise HarnessError(f"no visible control labelled {list(patterns)}")

    def labels(self, frame="top"):
        """Visible control labels for debugging; account-manager / e-mail labels are never returned."""
        out = self.run_js(frame, HELPERS + "__HOL.labels()") or []
        return [x for x in out if not re.search(r"account|@|profile|signed in as", x, re.I)]

    def shot(self, name: str, mask: bool = True) -> str:
        return self.cmd("screenshot", file=name, mask=mask)["file"]


class TabPane(Pane):
    """A Pane bound to one harness tab (demo loop: one tab per Office document in the same window).

    Frame, eval, navigation and input commands carry ``tab`` so they reach that tab's guest whether it is
    visible or not; window-wide commands (chat column, gate, screenshot, file chooser, tabs) do not."""

    WINDOW_WIDE = frozenset({"tab", "version", "quit", "clearCache", "host", "gate", "chat", "focus", "screenshot",
                             "armFiles", "awaitFiles"})

    def __init__(self, base: str, log, tab: str):
        super().__init__(base, log)
        self.tab = tab

    def cmd(self, cmd: str, timeout: float = 90, **kw):
        if cmd not in self.WINDOW_WIDE:
            kw.setdefault("tab", self.tab)
        return super().cmd(cmd, timeout=timeout, **kw)

    def open(self, url: str = "about:blank"):
        return super().cmd("tab", op="open", name=self.tab, url=url, timeout=90)

    def show(self):
        """Make this tab the visible one (a person clicking its tab); UI typing/clicks need it visible."""
        return super().cmd("tab", op="show", name=self.tab)

    def close(self):
        return super().cmd("tab", op="close", name=self.tab)


# ---- sign-in gate ---------------------------------------------------------------------------
SIGNED_IN_JS = HELPERS + r"""(() => {
  const host = location.hostname
  if (/login\.(microsoftonline|live)\.com$/.test(host)) return 'signing-in'
  const signIn = __HOL.find(['Sign in', 're:^sign in$'], 'a,button,[role=button]')
  if (signIn || location.pathname.startsWith('/en-us')) return 'signed-out'
  return 'signed-in'
})()"""


def signin_state(pane: Pane) -> str:
    try:
        return pane.run_js("top", SIGNED_IN_JS)
    except HarnessError:
        return "unknown"


def wait_signed_in(pane: Pane, timeout: float, log) -> bool:
    pane.cmd("gate", text=("ONE-TIME SIGN-IN (live-test harness, its own profile): click Sign in on this Excel page and sign "
                           "in to your Microsoft account. Do not open any files. The harness notices and continues by itself."))
    pane.cmd("focus")
    end = time.time() + timeout
    last = None
    while time.time() < end:
        s = signin_state(pane)
        if s != last:
            log({"ev": "signin", "state": s})
            last = s
        if s == "signed-in":
            pane.cmd("gate", text="")
            return True
        time.sleep(5)
    pane.cmd("gate", text="Sign-in timed out; the harness stopped. Rerun scripts/livetest/run_macbook.sh to try again.")
    return False


# ---- documents ------------------------------------------------------------------------------
def doc_name(host: str) -> str:
    return f"HOL live test {host}"


def open_or_create(pane: Pane, host: str, log) -> str:
    """Open the stored disposable document for this host, or create it from a blank-document URL."""
    known = pane.cmd("docs", op="get", host=host).get("known")
    if known:
        pane.cmd("docs", op="open", host=host, timeout=120)
        return "reopened"
    if host == "onenote":
        create_onenote_notebook(pane, log)
    else:
        pane.cmd("navigate", url=HOME_URLS[host], timeout=120)
        time.sleep(6)
        pane.click_label(BLANK_LABELS[host], timeout=40, frames_like=["cloud.microsoft"])
        time.sleep(8)
        wait_editor(pane, host)
        rename_document(pane, host, log)
    pane.cmd("docs", op="remember", host=host)
    return "created"


def wait_editor(pane: Pane, host: str, timeout: float = 120):
    """Wait until the Office editor shows its ribbon (Home tab) in some frame."""
    tab = "Insert" if host == "onenote" else "Home"
    end = time.time() + timeout
    js = HELPERS + f"!!__HOL.find([{json.dumps(tab)}], '[role=tab]')"
    while time.time() < end:
        for u in ["top"] + pane.frames()[1:]:
            try:
                if pane.run_js(u, js, timeout=15):
                    return u
            except HarnessError:
                pass
        time.sleep(2)
    raise HarnessError(f"{host} editor did not load (no {tab} tab)")


def rename_document(pane: Pane, host: str, log):
    """Rename a freshly created blank document via the title-bar file name control."""
    base = DEFAULT_TITLES[host]
    pane.click_label(["Rename file", "re:^rename", f"re:^{base}\\s*\\d*(\\.\\w+)?$"], timeout=30,
                     sel="[aria-label],button,[role=button]")
    time.sleep(1.5)
    # The title becomes editable in place (contenteditable) or opens a "File name" box: select its text.
    js = r"""(() => { const a = document.activeElement; if (!a || a === document.body) return false;
        if (a.isContentEditable) { const r = document.createRange(); r.selectNodeContents(a); const s = getSelection(); s.removeAllRanges(); s.addRange(r); return 'editable' }
        if (a.tagName === 'INPUT' && /name/i.test(a.getAttribute('aria-label') || '') && !/font/i.test(a.getAttribute('aria-label') || '')) { a.select(); return 'input' }
        return false })()"""
    for u in ["top"] + pane.frames()[1:]:
        try:
            kind = pane.run_js(u, js, timeout=15)
        except HarnessError:
            continue
        if kind:
            pane.cmd("type", text=doc_name(host))
            pane.cmd("key", keyCode="Enter")
            time.sleep(3)
            log({"ev": "renamed", "host": host, "via": kind})
            return
    raise HarnessError("file-name box not found after clicking the title")


def ensure_onenote_signed_in(pane: Pane, log, timeout: float = 1200):
    """OneNote for the web has its own sign-in hop. Click its Sign in; if Microsoft then asks for
    an account (not single sign-on), show the gate and wait for the person. Never types anything."""
    time.sleep(3)
    if signin_state(pane) != "signed-out":
        return True
    pane.click_label(["Sign in"], timeout=20, sel="a,button,[role=button]")
    end = time.time() + timeout
    gated = False
    while time.time() < end:
        time.sleep(5)
        st = signin_state(pane)
        if st == "signed-in":
            if gated:
                pane.cmd("gate", text="")
            log({"ev": "onenote-signin", "state": st, "human": gated})
            return True
        if st == "signing-in" and not gated:
            gated = True
            log({"ev": "onenote-signin", "state": "waiting-for-person"})
            pane.cmd("gate", text=("ONENOTE SIGN-IN (one time): OneNote for the web asks you to sign in separately. Sign in "
                                   "here with the same Microsoft account; the harness continues by itself."))
            pane.cmd("focus")
    pane.cmd("gate", text="OneNote sign-in timed out; rerun scripts/livetest/run_macbook.sh onenote.")
    raise HarnessError("OneNote sign-in gate timed out (needs the person to sign in once)")


def create_onenote_notebook(pane: Pane, log):
    pane.cmd("navigate", url=ONENOTE_HOME, timeout=120)
    ensure_onenote_signed_in(pane, log)
    time.sleep(4)
    # Classic OneNote for the web is the host Office add-ins (OneNoteApi) run in.
    pane.click_label(["Create classic notebook", "New notebook", "^New notebook", "Create new notebook"], timeout=40)
    time.sleep(2)
    js = HELPERS + r"""(() => { const a = document.activeElement; const i = (a && a.tagName === 'INPUT') ? a : [...document.querySelectorAll('input')].find(e => __HOL.vis(e) && e.type !== 'hidden');
        if (!i) return false; i.focus(); i.select(); return true })()"""
    ok = False
    for u in ["top"] + pane.frames()[1:]:
        try:
            if pane.run_js(u, js, timeout=15):
                ok = True
                break
        except HarnessError:
            continue
    if not ok:
        raise HarnessError("notebook name box not found")
    pane.cmd("type", text=doc_name("onenote").replace(" onenote", ""))
    pane.click_label(["Create", "re:^create$"], timeout=20)
    wait_editor(pane, "onenote", timeout=180)
    log({"ev": "notebook-created"})


# OneNote for the web greys out Insert > Office Add-ins while no page is open (an empty section shows "This
# section is empty. Click here or press ENTER to create a new page."). The button keeps no disabled/aria-disabled
# attribute; only pointer-events: none, so a trusted click lands and nothing opens (A45, reproduced live for A46).
ONENOTE_PAGE_JS = HELPERS + r"""(() => {
  const ec = document.getElementById('EditorContainer')
  const empty = !!(ec && /This section is empty/i.test(ec.innerText || ''))
  const b = __HOL.find(['Office Add-ins'], 'button,[role=button]')
  let addins = 'missing'
  if (b) { const cs = getComputedStyle(b)
    addins = (b.disabled || b.getAttribute('aria-disabled') === 'true' || cs.pointerEvents === 'none') ? 'disabled' : 'enabled' }
  return { empty, addins }
})()"""


def ensure_onenote_page(pane: Pane, log, frame=None, timeout: float = 60, clock=time.monotonic, sleep=time.sleep) -> dict:
    """Make sure a page is open and Insert > Office Add-ins is enabled before it is clicked (bounded).

    In an empty section: click Add page once (a second time only if the first never took after 20 s), then wait
    for the page canvas and the enabled button. Adds at most two blank pages to whatever section OneNote reopened;
    nothing is deleted. Raises HarnessError if the button is still greyed out at the end."""
    frame = frame or wait_editor(pane, "onenote", timeout=60)
    start, added, last_add, insert_clicks, st = clock(), 0, None, 0, {}
    while True:
        try:
            st = pane.run_js(frame, ONENOTE_PAGE_JS, timeout=15) or {}
        except HarnessError as exc:
            st = {"error": str(exc)[:80]}
        if not st.get("empty") and st.get("addins") == "enabled":
            if added:
                log({"ev": "onenote-page-ready", "pagesAdded": added, "waitedS": round(clock() - start, 1)})
            return st
        waited = clock() - start
        if waited >= timeout:
            raise HarnessError(f"OneNote Insert > Office Add-ins still greyed out after {timeout:.0f} s "
                               f"(page state {st}, pages added {added})")
        if st.get("empty") and added < 2 and (last_add is None or clock() - last_add >= 20):
            try:
                how = pane.click_label(["Add page", "re:^click here or press enter to create a new page"], timeout=10,
                                       sel="[role=button],button,a")
            except HarnessError:
                how = None
            if how:
                added, last_add = added + 1, clock()
                log({"ev": "onenote-page-added", "via": how[:40], "n": added})
        elif st.get("addins") == "missing" and insert_clicks < 2:
            insert_clicks += 1
            try:
                pane.click_label(["Insert"], timeout=10, sel="[role=tab]")
            except HarnessError:
                pass
        sleep(2)


# ---- add-in ---------------------------------------------------------------------------------
def task_pane_frame(pane: Pane, host: str):
    for u in pane.frames():
        if f"/office/live/{host}/taskpane.html" in u:
            return u
    return None


def upload_addin(pane: Pane, host: str, manifest_path: str, log):
    """Add-ins > More Add-ins > MY ADD-INS > Manage My Add-ins > Upload My Add-in > Browse (CDP) > Upload."""
    if host == "onenote":
        for attempt in range(3):  # the Insert tab can take a few seconds to render its groups after a reload
            pane.click_label(["Insert"], timeout=30, sel="[role=tab]")
            time.sleep(3)
            ensure_onenote_page(pane, log)  # Office Add-ins is greyed out until a page is open (A46)
            try:
                pane.click_label(["Office Add-ins", "^Office Add-ins", "Add-ins"], timeout=15)
                break
            except HarnessError:
                if attempt == 2:
                    raise
                time.sleep(5)
    else:
        pane.click_label(["Home"], timeout=30, sel="[role=tab]")
        for attempt in range(4):  # right after load the button only shows its tooltip; retry
            pane.click_label(["Add-ins", "Office Add-ins"], timeout=30)
            try:
                pane.click_label(["More Add-ins", "^More add-ins", "Advanced", "More Settings"], timeout=8)
                break
            except HarnessError:
                if attempt == 3:
                    raise
                pane.cmd("key", keyCode="Escape")
                time.sleep(4)
    dialog = ["res.public.onecdn", "officeonline", "store.office", "WefGallery", "Resources/"]
    time.sleep(5)
    pane.click_label(["MY ADD-INS", "My add-ins"], timeout=40, frames_like=dialog, sel="[role=tab],a,button")
    time.sleep(6)  # the My Add-ins list loads after the tab switch; Manage appears with it
    try:
        pane.click_label(["Upload My Add-in", "re:upload my add-?in"], timeout=6, frames_like=dialog)
    except HarnessError:
        pane.click_label(["Manage My Add-ins", "re:manage my add-?ins"], timeout=45, frames_like=dialog)
        time.sleep(2)
        try:
            pane.click_label(["Upload My Add-in", "re:upload my add-?in"], timeout=6, frames_like=dialog)
        except HarnessError:
            pass  # this build opens the upload dialog straight from Manage My Add-ins
    time.sleep(2)
    armed = pane.cmd("armFiles", paths=[manifest_path], timeoutMs=30000)
    log({"ev": "file-chooser-armed", "sessions": armed.get("sessions")})
    pane.click_label(["Browse...", "Browse"], timeout=20, frames_like=dialog, sel="input[type=button],button,[role=button]")
    got = pane.cmd("awaitFiles", timeout=40)
    if not got.get("ok"):
        raise HarnessError(f"CDP file chooser: {got.get('error')}")
    log({"ev": "file-chooser-filled", "files": got.get("files")})
    time.sleep(1)
    pane.click_label(["Upload"], timeout=20, frames_like=dialog, sel="input[type=button],input[type=submit],button")
    log({"ev": "uploaded", "host": host})
    time.sleep(6)


# The add-in's ribbon button label (manifest Hermes.Button.Label); tests pin it to the built manifests.
RIBBON_LABELS = ["Workbench"]


def open_task_pane(pane: Pane, host: str, log, timeout: float = 60) -> str:
    """Click the ribbon's Workbench button (via the ribbon overflow when the window is narrow)."""
    end = time.time() + timeout
    while time.time() < end:
        f = task_pane_frame(pane, host)
        if f:
            return f
        try:
            pane.click_label(RIBBON_LABELS, timeout=5)
        except HarnessError:
            try:
                pane.click_label(["More Options", "re:^more commands", "Ribbon Display Options"], timeout=5)
                time.sleep(1.5)
                pane.click_label(RIBBON_LABELS, timeout=5)
            except HarnessError:
                pass
        time.sleep(4)
    raise HarnessError("Workbench task pane did not open")


def pane_state(pane: Pane, frame: str):
    return pane.run_js(frame, "({ state: globalThis.__hermesPaneState || null, status: (document.getElementById('status') || {}).textContent || '', autoShow: globalThis.__hermesAutoShow ?? null })")


OFFICE_JS_FAILED = "Could not load Office.js"  # the pane's own load-failure text (addin/shared/taskpane.js)


def wait_pane_state(pane: Pane, host: str, want, timeout: float = 60, clock=time.time, sleep=time.sleep):
    """Wait until the pane reports one of ``want``. A pane whose Office.js script failed to load says so and asks
    to be reloaded; the harness does what a person would, once per wait (A8: the first pane load in a freshly
    started harness window failed once, the reload loaded it). A second failure is reported, not retried."""
    end = clock() + timeout
    last = None
    reloaded = False
    while clock() < end:
        f = task_pane_frame(pane, host)
        if f:
            try:
                last = pane_state(pane, f)
                if last and last.get("state") in want:
                    return last
                if not reloaded and last and OFFICE_JS_FAILED in str(last.get("status") or ""):
                    reloaded = True
                    pane.log({"ev": "pane-reload-office-js-failed", "host": host})
                    pane.cmd("reloadFrame", frame=f)
                    sleep(3)
                    continue
            except HarnessError:
                pass
        sleep(1)
    raise HarnessError(f"pane never reached {want} (last {last})")


# Office for the web's own banner when its editing session ends: "Sorry, your session expired. Please refresh the
# page to continue." Every Office.js call fails until the page is reloaded (Word: RichApi "... reading 'p5d'").
SESSION_EXPIRED_JS = r"""(() => /your session (has )?expired/i.test((document.body && document.body.innerText) || ''))()"""


def expired_session_frame(pane: Pane, host: str):
    """The host frame showing the session-expired banner, or None. Read-only (page text only)."""
    for u in ["top"] + [f for f in pane.frames()[1:] if f"/office/live/{host}/" not in f]:
        try:
            if pane.run_js(u, SESSION_EXPIRED_JS, timeout=15) is True:
                return u
        except HarnessError:
            pass
    return None


def refresh_expired_session(pane: Pane, host: str, log, manifest_path=None, sleep=time.sleep):
    """What a person does at the banner: reload the same document (no write), then bring the add-in pane back the
    way prepare_host does and wait for it to resume its pairing. Raises if the pane does not come back paired."""
    log({"ev": "session-expired", "host": host})
    pane.cmd("docs", op="open", host=host, timeout=120)
    sleep(5)
    wait_editor(pane, host, timeout=180)
    # A19 (seen live twice, Word run 3 of hol-a19-77b5652-all-3 and its rerun): after the reload the add-in is still on
    # the ribbon, and Word's add-in dialog then showed no "Manage My Add-ins", so the re-upload failed. A person opens
    # the pane first; only if it does not come back is the add-in uploaded again (with one document reload, as in
    # prepare_host).
    try:
        open_task_pane(pane, host, log, timeout=40)
        log({"ev": "session-refresh-pane", "host": host, "how": "ribbon"})
    except HarnessError:
        if not manifest_path:
            raise
        try:
            upload_addin(pane, host, manifest_path, log)
        except HarnessError as exc:
            log({"ev": "session-refresh-upload-retry", "host": host, "why": str(exc)[:120]})
            pane.cmd("docs", op="open", host=host, timeout=120)
            wait_editor(pane, host, timeout=180)
            sleep(5)
            upload_addin(pane, host, manifest_path, log)
        open_task_pane(pane, host, log, timeout=90)
        log({"ev": "session-refresh-pane", "host": host, "how": "uploaded"})
    st = wait_pane_state(pane, host, ("paired", "unpaired", "rejected", "ready", "no-bridge"), timeout=90)
    log({"ev": "session-refreshed", "host": host, "pane": st.get("state")})
    if st.get("state") != "paired":
        raise HarnessError(f"after the reload the pane is {st.get('state')!r}, not paired")
    return st


# Pane states in which a person can pair (the code box is shown) or is already paired. "no-bridge": the profile's
# bridge starts with its first office_connect; the pane looks again on Pair. "waiting-bridge" (A14): a reloaded pane
# whose bridge gave its idle port back; it keeps looking for it, and a new code pairs it anew (seen live, A15).
PAIRABLE_STATES = ("ready", "unpaired", "rejected", "paired", "no-bridge", "waiting-bridge")


def type_code(pane: Pane, host: str, code: str, log):
    """The harness (standing in for the person) types the code it read in the chat into the pane."""
    frame = task_pane_frame(pane, host)
    st = wait_pane_state(pane, host, PAIRABLE_STATES, timeout=60)
    if st.get("state") == "paired":
        # The pane resumed an earlier run's chat (tab-scoped token). A fresh pairing per run: the
        # person clicks Disconnect first, then enters the new chat's code.
        pane.cmd("clickIn", frame=frame, js="(() => { const b = document.getElementById('disconnect'); if (!b || b.hidden) return null;"
                 " const r = b.getBoundingClientRect(); return { x: r.x + r.width / 2, y: r.y + r.height / 2, label: 'Disconnect' } })()")
        log({"ev": "pane-disconnected-previous-chat"})
        wait_pane_state(pane, host, ("ready", "unpaired"), timeout=30)
        frame = task_pane_frame(pane, host)
    # Focus the code box with a real click, as a person would.
    pane.cmd("clickIn", frame=frame, js="(() => { const i = document.getElementById('code'); const r = i.getBoundingClientRect();"
             " return { x: r.x + r.width / 2, y: r.y + r.height / 2, label: 'code' } })()")
    pane.run_js(frame, "(() => { const i = document.getElementById('code'); i.value = ''; i.focus(); return document.activeElement === i })()")
    pane.cmd("type", text=code.replace("-", ""))
    pane.cmd("key", keyCode="Enter")
    st = wait_pane_state(pane, host, ("paired", "rejected", "no-bridge"), timeout=60)
    log({"ev": "pane-after-code", "state": st.get("state")})
    return st


# ---- Outlook on the web (A10b) ----------------------------------------------------------------------
# The Outlook tab signs a different throwaway account in (one with an Outlook.com mailbox), in a partition of its
# own (tabscript.js partitionFor: outlook* tabs -> persist:hol-outlook). The account's address never leaves the
# page: where a step must know it (the one self-addressed fixture email, the reply's recipient check), the page
# compares it with itself and the harness only ever sees a boolean.
OUTLOOK_TAB = "outlook"
OUTLOOK_URL = "https://outlook.live.com/mail/"
OUTLOOK_SIGNIN_URL = ("https://login.live.com/login.srf?wa=wsignin1.0&wreply=https%3A%2F%2Foutlook.live.com%2Fowa%2F"
                      "%3Fnlp%3D1&id=292841")
# A19 (FRIDAY): the fixture's second decision is neutral ("the current pilot scope", no price). It is a new message
# with its own subject, sent once like the first (its own state file); the A10b message stays in Sent Items unused.
# docs/demo-script.md quotes this subject and body; tests/test_livetest_story.py keeps the two in sync.
OUTLOOK_FIXTURE_SUBJECT = "HOL live test: pilot sync"
OUTLOOK_FIXTURE_TOPIC = "pilot sync"  # names the run's decision page and must be in the brain_why answer
OUTLOOK_FIXTURE_BODY = ("Pilot sync, 6 November.\nDecision: we ship the beta on Friday.\nDecision: we keep the current "
                        "pilot scope.\nAction: Sam writes the FAQ by 12 November.")
OUTLOOK_FIXTURE_DECISIONS = ("beta", "pilot scope")  # each must be in one decision the story's page records
GYRO = "inclient.store.office.com"  # the Get add-ins dialog frame
# The ribbon's "More apps" (reading) / "Apps" (draft) button; the left app rail has another "More apps" (app launcher).
RIBBON_APPS_SEL = "[role=group] button,[role=toolbar] button,[role=menubar] button"
OWN_ADDRESS_JS = r"""((document.querySelector('[id^="primaryMailboxRoot_"]') || {}).title || '')"""
# The compose button: "New mail" before Nov 2026; Outlook on the web now labels it just "New" (beside "Expand to see
# more New options"). Exact match on "New" so the expander and other "New ..." controls never count.
NEW_MAIL_BUTTON = 'button[aria-label*="New mail" i], button[aria-label="New" i]'
OUTLOOK_MAILBOX_JS = r"""(() => ({ mailbox: location.hostname === 'outlook.live.com' && location.pathname.startsWith('/mail')
  && !!document.querySelector('%s') }))()""" % NEW_MAIL_BUTTON
# Privacy mask for screenshots: blur the folder pane, message list, reading pane and address fields; the task pane
# (an iframe) stays readable. Applied around every Outlook screenshot, on top of the window's account-area mask.
OUTLOOK_MASK_JS = r"""((on) => { let s = document.getElementById('__holmask'); if (!s) { s = document.createElement('style');
  s.id = '__holmask'; document.head.appendChild(s) }
  s.textContent = on ? '[role=listbox], [role=navigation], [role=tree], [aria-label*="Message list" i], [aria-label="Reading Pane" i], '
    + '[role=main] [role=region]:not(:has(iframe)), [aria-label="Message body" i], div[aria-label^="To"], div[aria-label^="Cc"], '
    + 'div[aria-label^="Bcc"], [aria-label^="From" i] { filter: blur(9px) !important }' : ''; return true })"""


def outlook_mailbox(pane: Pane) -> bool:
    try:
        return bool((pane.run_js("top", OUTLOOK_MAILBOX_JS) or {}).get("mailbox"))
    except HarnessError:
        return False


def outlook_shot(pane: Pane, name: str) -> str:
    """A window screenshot with the Outlook page's personal areas blurred (and the account area masked)."""
    try:
        pane.run_js("top", f"({OUTLOOK_MASK_JS})(true)")
        time.sleep(0.4)
        return pane.shot(name)
    finally:
        try:
            pane.run_js("top", f"({OUTLOOK_MASK_JS})(false)")
        except HarnessError:
            pass


def ensure_outlook_signed_in(pane: Pane, log, timeout: float = 10800, poll: float = 30) -> bool:
    """Open the Outlook tab and wait (bounded) for its mailbox; a person signs in by hand. Never types anything."""
    pane.open(OUTLOOK_URL)
    pane.show()
    pane.cmd("navigate", url=OUTLOOK_URL, timeout=90)  # an existing tab may be anywhere in the mailbox
    end = time.time() + 45  # Outlook takes a while to draw the mailbox; only then is the sign-in page needed
    while time.time() < end:
        time.sleep(3)
        if outlook_mailbox(pane):
            return True
    pane.cmd("navigate", url=OUTLOOK_SIGNIN_URL, timeout=90)
    pane.cmd("gate", text=("ONE-TIME OUTLOOK SIGN-IN (its own cookie jar): sign the throwaway Outlook account in on this tab. "
                           "The harness only checks every 30 s whether the mailbox appeared; it never types or reads what you enter."))
    pane.cmd("focus")
    end = time.time() + timeout
    while time.time() < end:
        if outlook_mailbox(pane):
            pane.cmd("gate", text="")
            log({"ev": "outlook-signed-in"})
            return True
        time.sleep(poll)
    return False


def _nav_pane(pane: Pane):
    try:
        pane.click_label(["Show navigation pane"], timeout=2)
        time.sleep(1.5)
    except HarnessError:
        pass


OUTLOOK_FOLDERS = {"Inbox": "inbox", "Sent Items": "sentitems"}


def outlook_folder(pane: Pane, name: str):
    """Show a folder by its address (the folder pane may be collapsed in a narrow window)."""
    pane.cmd("navigate", url=f"https://outlook.live.com/mail/0/{OUTLOOK_FOLDERS[name]}", timeout=90)
    end = time.time() + 40
    while time.time() < end:
        time.sleep(2)
        if outlook_mailbox(pane):
            break
    time.sleep(4)


FIXTURE_COUNT_JS = r"""((subject) => [...document.querySelectorAll('[role=listbox] [role=option]')]
  .filter(o => (o.getAttribute('aria-label') || o.innerText || '').includes(subject)).length)"""


def outlook_fixture_count(pane: Pane, folder: str = "Inbox") -> int:
    outlook_folder(pane, folder)
    n = pane.run_js("top", f"({FIXTURE_COUNT_JS})({json.dumps(OUTLOOK_FIXTURE_SUBJECT)})")
    return int(n or 0)


def open_fixture_message(pane: Pane, log, folder: str = "Inbox"):
    """Open the newest 'HOL live test' message in the folder (read mode)."""
    outlook_folder(pane, folder)
    js = (f"(() => {{ const o = [...document.querySelectorAll('[role=listbox] [role=option]')].find(o => (o.getAttribute('aria-label') "
          f"|| o.innerText || '').includes({json.dumps(OUTLOOK_FIXTURE_SUBJECT)})); if (!o) return null; o.scrollIntoView({{block: 'center'}});"
          " const r = o.getBoundingClientRect(); return { x: r.x + r.width / 2, y: r.y + r.height / 2, label: 'fixture message' } })()")
    hit = pane.cmd("clickIn", frame="top", js=js)
    if not hit.get("ok"):
        raise HarnessError(f"no 'HOL live test' message in {folder}")
    log({"ev": "outlook-fixture-opened", "folder": folder})
    time.sleep(4)


# The open message's own address: Outlook on the web shows a message being read at /mail/0/<folder>/id/<id>, where
# <id> is the conversation's stable, opaque id (the list row's data-convid; no subject or address in it). A cold
# load of that address drops the id (seen live, A19), so a clean load opens the folder and selects the row with that
# id. The harness keeps the id only to find the same message again, and never logs it.
ITEM_URL_RE = re.compile(r"^https://outlook\.live\.com/mail/0/(inbox|sentitems)/id/([A-Za-z0-9%._~=+-]{8,})$")


def fixture_item_url(pane: Pane):
    """The address of the message open now (its stable id), or None if the page is not showing one."""
    try:
        href = pane.run_js("top", "location.origin + location.pathname")
    except HarnessError:
        return None
    return href if isinstance(href, str) and ITEM_URL_RE.match(href) else None


# The row whose stable id is ``cid`` (scrolled into view), as a point to click.
ROW_BY_ID_JS = r"""((cid) => { const o = [...document.querySelectorAll('[role=listbox] [role=option]')].find(o =>
  o.getAttribute('data-convid') === cid || !!o.querySelector(`[data-convid="${CSS.escape(cid)}"]`)); if (!o) return null;
  o.scrollIntoView({block: 'center'}); const r = o.getBoundingClientRect();
  return { x: r.x + r.width / 2, y: r.y + r.height / 2, label: 'fixture message (by id)' } })"""
READ_RIBBON_JS = r"""((cid) => ({ mailbox: location.hostname === 'outlook.live.com'
  && !!document.querySelector('""" + NEW_MAIL_BUTTON + r"""'),
  selected: decodeURIComponent(location.pathname).endsWith('/id/' + cid) || location.pathname.endsWith('/id/' + cid),
  moreApps: !!__HOL.find(['More apps'], '[role=group] button,[role=toolbar] button,[role=menubar] button'),
  reading: !!__HOL.find(['Reply', 'Reply all', 'Forward']),
  form: !!document.querySelector('input[placeholder="Add title"]') }))"""


def open_mail_item_clean(pane: Pane, url: str, log, timeout: float = 60, clock=time.time, sleep=time.sleep):
    """A clean mail view of ONE message (A19): unload Outlook (about:blank), load its folder fresh, select the row
    with the message's stable id (never "the newest with that subject"), then wait until the address names that id,
    the reading ribbon (More apps) and the message (Reply/Forward) are drawn, and no event form is left. Seen live
    (A13-A15): after an event form, a message opened in the same Outlook session often offered no Workbench under
    More apps."""
    m = ITEM_URL_RE.match(url) if isinstance(url, str) else None
    if not m:
        raise HarnessError("no stable item address for the fixture message")
    folder_url, cid = f"https://outlook.live.com/mail/0/{m.group(1)}", m.group(2)
    t0 = clock()
    end = t0 + timeout
    pane.cmd("navigate", url="about:blank", timeout=30)
    pane.cmd("navigate", url=folder_url, timeout=90)
    clicked, last = False, {}
    while clock() < end:
        sleep(2)
        try:
            if not clicked:
                hit = pane.cmd("clickIn", frame="top", js=f"({ROW_BY_ID_JS})({json.dumps(cid)})")
                clicked = bool((hit or {}).get("ok"))
                continue
            last = pane.run_js("top", HELPERS + f"({READ_RIBBON_JS})({json.dumps(cid)})") or {}
        except HarnessError:
            last = {}
        if last.get("mailbox") and last.get("selected") and last.get("moreApps") and last.get("reading") and not last.get("form"):
            sleep(2)  # the ribbon's add-in list fills in just after the ribbon draws
            log({"ev": "outlook-item-clean-load", "s": round(clock() - t0, 1)})
            return True
    raise HarnessError(f"the fixture message did not load into a clean read view (row clicked {clicked}, last {last})")


# The fixture is addressed to the signed-in account itself, filled and checked inside the page: the address is read
# from the folder pane and inserted into To there; the harness sees only booleans.
FILL_TO_SELF_JS = r"""(() => { const own = %s; if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(own)) return { ok: false, why: 'own address not found' };
  const to = document.querySelector('div[contenteditable="true"][aria-label="To"]'); if (!to) return { ok: false, why: 'no To field' };
  if ((to.innerText || '').trim()) return { ok: false, why: 'To is not empty' };
  to.focus(); return { ok: document.execCommand('insertText', false, own) && document.activeElement === to } })()""" % OWN_ADDRESS_JS
RECIPIENTS_JS = r"""(() => { const own = (%s).toLowerCase(); const re = /[A-Za-z0-9._%%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}/g;
  const field = name => document.querySelector(`div[contenteditable="true"][aria-label="${name}"]`);
  const well = el => { let w = el; for (let i = 0; i < 4 && w && w.parentElement; i++) w = w.parentElement; return w };
  const addrs = el => { if (!el) return []; const w = well(el);
    const t = [...w.querySelectorAll('*')].flatMap(e => ['aria-label', 'title'].map(a => e.getAttribute(a) || '')).join(' ') + ' ' + w.innerText;
    return [...new Set((t.match(re) || []).map(x => x.toLowerCase()))] };
  const to = addrs(field('To')), cc = addrs(field('Cc')), bcc = addrs(field('Bcc'));
  // The draft's own From (the account's sending address; A10b: a reply to the self-sent fixture is addressed to it,
  // and it differs from the folder pane's mailbox address): together they are the account's own addresses.
  const from = [...new Set([...document.querySelectorAll('[aria-label^="From" i]')].flatMap(e => (((e.getAttribute('aria-label') || '')
    + ' ' + (e.getAttribute('title') || '') + ' ' + (e.innerText || '')).match(re) || []).map(x => x.toLowerCase())))];
  const mine = new Set([own, ...from].filter(Boolean));
  return { toIsOnlySelf: !!own && to.length === 1 && to[0] === own, toIsOnlyAccount: to.length === 1 && mine.has(to[0]),
           toCount: to.length, ccCount: cc.length, bccCount: bcc.length, ccEmpty: cc.length === 0, bccEmpty: bcc.length === 0 } })()""" % OWN_ADDRESS_JS


def outlook_recipients(pane: Pane) -> dict:
    """Booleans and counts only: is To exactly the signed-in account, are Cc and Bcc empty."""
    return pane.run_js("top", RECIPIENTS_JS) or {}


def outlook_new_mail(pane: Pane):
    try:
        pane.click_label(["New mail", "New"], timeout=15)  # exact labels (see NEW_MAIL_BUTTON)
    except HarnessError:  # A10b: right after a discarded reply the ribbon can lack New mail; start from the Inbox
        outlook_folder(pane, "Inbox")
        pane.click_label(["New mail", "New"], timeout=20)
    time.sleep(5)


def outlook_discard(pane: Pane, log):
    """Discard the open draft (reply or new): Discard, then confirm if Outlook asks."""
    pane.click_label(["Discard"], timeout=15)
    time.sleep(2)
    try:
        pane.click_label(["OK", "Discard"], timeout=4, sel="button")
    except HarnessError:
        pass
    log({"ev": "outlook-draft-discarded"})
    time.sleep(2)


def send_fixture_once(pane: Pane, state_file, log) -> dict:
    """The ONE approved fixture email: the harness (never the plugin) sends it from the test account to itself.

    Reuses an Inbox message with the fixture subject if there is one. A state file records that the send was made
    (written BEFORE Send is clicked), and a copy in Sent Items proves it too: then a run never sends again. If the
    sent message has not reached the Inbox (seen live, A10b: a new Outlook.com account's self-addressed mail stayed
    out of the Inbox), the read-mode steps run on its Sent Items copy, the same message. Returns the folder to use."""
    from pathlib import Path as _P
    state_file = _P(state_file)
    if outlook_fixture_count(pane) > 0:
        return {"fixture": "reused", "sent": False, "folder": "Inbox"}
    if state_file.exists() or outlook_fixture_count(pane, "Sent Items") > 0:
        for _ in range(3):  # sent before (maybe by an interrupted run): look again, never send twice
            if outlook_fixture_count(pane) > 0:
                return {"fixture": "arrived", "sent": False, "folder": "Inbox"}
            time.sleep(10)
        if outlook_fixture_count(pane, "Sent Items") > 0:
            return {"fixture": "sent-items copy", "sent": False, "folder": "Sent Items"}
        raise HarnessError("the fixture email was sent once already and is in neither Inbox nor Sent Items; not sending another")
    outlook_folder(pane, "Inbox")
    outlook_new_mail(pane)
    filled = pane.run_js("top", FILL_TO_SELF_JS)
    if not (filled or {}).get("ok"):
        outlook_discard(pane, log)
        raise HarnessError(f"could not address the fixture to the test account: {(filled or {}).get('why')}")
    time.sleep(1.5)
    pane.cmd("key", keyCode="Enter")  # resolve the typed address into a recipient
    time.sleep(2.5)
    pane.click_label(["Subject"], timeout=10, sel="input")
    pane.cmd("type", text=OUTLOOK_FIXTURE_SUBJECT)
    pane.click_label(["Message body"], timeout=10, sel="[role=textbox]")
    for i, line in enumerate(OUTLOOK_FIXTURE_BODY.split("\n")):
        if i:
            pane.cmd("key", keyCode="Enter")
        pane.cmd("type", text=line)
    time.sleep(1.5)
    rc = outlook_recipients(pane)
    log({"ev": "outlook-fixture-recipients", **{k: rc.get(k) for k in ("toIsOnlySelf", "toCount", "ccEmpty", "bccEmpty")}})
    if not (rc.get("toIsOnlySelf") and rc.get("ccEmpty") and rc.get("bccEmpty")):
        outlook_discard(pane, log)
        raise HarnessError(f"fixture not sent: To is not exactly the test account or Cc/Bcc is not empty ({rc})")
    state_file.parent.mkdir(parents=True, exist_ok=True)
    state_file.write_text(json.dumps({"sent": True, "at": time.strftime("%Y-%m-%dT%H:%M:%S"), "subject": OUTLOOK_FIXTURE_SUBJECT}))
    pane.click_label(["Send"], timeout=10, sel="button")
    log({"ev": "outlook-fixture-sent"})
    for _ in range(9):
        time.sleep(10)
        if outlook_fixture_count(pane) > 0:
            return {"fixture": "sent", "sent": True, "folder": "Inbox"}
    if outlook_fixture_count(pane, "Sent Items") > 0:
        return {"fixture": "sent (Sent Items copy)", "sent": True, "folder": "Sent Items"}
    raise HarnessError("the fixture email was sent but is in neither Inbox nor Sent Items")


def outlook_addin_installed(pane: Pane, name: str = "Workbench for Outlook") -> bool:
    f = next((u for u in pane.frames() if GYRO in u), None)
    if not f:
        return False
    t = pane.run_js(f, "(document.body && document.body.innerText) || ''") or ""
    i = t.find("Custom Addins")
    return i >= 0 and name in t[i:]


def _open_get_addins(pane: Pane):
    # Nov 2026: right after a message opens, Outlook can swallow the first "More apps" click (no flyout); click again,
    # bounded, until the flyout offers Get add-ins.
    for attempt in range(3):
        pane.click_label(["More apps", "Apps"], timeout=20, sel=RIBBON_APPS_SEL)
        time.sleep(2.5)
        try:
            pane.click_label(["re:^get add-ins"], timeout=8 if attempt < 2 else 15)
            break
        except HarnessError:
            if attempt == 2:
                raise
            pane.cmd("key", keyCode="Escape")
            time.sleep(2)
    for _ in range(20):
        if any(GYRO in u for u in pane.frames()):
            break
        time.sleep(1)
    time.sleep(5)
    pane.click_label(["My add-ins"], timeout=30, frames_like=[GYRO], sel="[role=tab],a,button,span")
    time.sleep(5)


def install_outlook_addin(pane: Pane, manifest_path: str, log, force: bool = False):
    """Get add-ins > My add-ins > Custom Addins > Add a custom add-in > Add from File... (CDP chooser) > Install.
    Skipped when 'Workbench for Outlook' is already listed (same add-in id), unless force."""
    _open_get_addins(pane)
    if outlook_addin_installed(pane) and not force:
        log({"ev": "outlook-addin", "installed": "already"})
        pane.click_label(["Close"], timeout=10)
        time.sleep(2)
        return "already"
    pane.click_label(["Add a custom add-in"], timeout=20, frames_like=[GYRO], sel="[role=combobox],button,[role=button]")
    time.sleep(2)
    armed = pane.cmd("armFiles", paths=[manifest_path], timeoutMs=30000)
    log({"ev": "file-chooser-armed", "sessions": armed.get("sessions")})
    pane.click_label(["Add from File...", "re:^add from file"], timeout=15, frames_like=[GYRO], sel="[role=option],[role=menuitem],button")
    got = pane.cmd("awaitFiles", timeout=40)
    if not got.get("ok"):
        raise HarnessError(f"CDP file chooser: {got.get('error')}")
    time.sleep(4)
    pane.click_label(["Install"], timeout=20, frames_like=[GYRO], sel="button,[role=button]")
    for _ in range(20):
        time.sleep(2)
        if outlook_addin_installed(pane):
            break
    else:
        raise HarnessError("Workbench for Outlook did not appear under Custom Addins after Install")
    log({"ev": "outlook-addin", "installed": "now"})
    pane.click_label(["Close"], timeout=10)
    time.sleep(2)
    return "installed"


def open_outlook_pane(pane: Pane, log, timeout: float = 60) -> str:
    """Open the Workbench task pane from the message (More apps) or the draft (Apps) ribbon.

    Each try opens the flyout from closed (Escape first: clicking More apps on an open flyout closes it, so a try
    that failed used to leave every other click toggling the flyout shut). A failed open logs the flyout's
    structure once (labels only, account/draft/fixture text dropped)."""
    end = time.time() + timeout
    tries = 0
    while time.time() < end:
        f = task_pane_frame(pane, "outlook")
        if f:
            return f
        if tries:
            try:
                if pane.run_js("top", FLYOUT_OPEN_JS):  # only a flyout left open; a draft never gets a stray key
                    pane.cmd("key", keyCode="Escape")
                    time.sleep(1)
            except HarnessError:
                pass
        tries += 1
        try:
            pane.click_label(["More apps", "Apps"], timeout=8, sel=RIBBON_APPS_SEL)
            time.sleep(2)
            pane.click_label(RIBBON_LABELS + ["Workbench for Outlook"], timeout=8)
        except HarnessError:
            pass
        time.sleep(5)
    try:
        log({"ev": "outlook-pane-open-failed", "tries": tries, "flyout": flyout_items(pane)})
    except HarnessError:
        pass
    raise HarnessError("Workbench task pane did not open in Outlook")


# What the More apps flyout holds (diagnostics only): visible popovers and the labels of their items.
FLYOUT_OPEN_JS = r"""[...document.querySelectorAll('[class*=PopoverSurface]')].some(p => p.getBoundingClientRect().width > 1)"""
FLYOUT_JS = r"""(() => { const v = e => { const r = e.getBoundingClientRect(); return r.width > 1 && r.height > 1 };
  return [...document.querySelectorAll('[role=menu],[role=dialog],[class*=PopoverSurface],.ms-Callout')].filter(v).slice(0, 4)
    .map(p => [...p.querySelectorAll('button,[role=menuitem],[role=option],[role=listitem]')].filter(v).slice(0, 15)
      .map(e => (e.getAttribute('aria-label') || e.innerText || '').trim().replace(/\s+/g, ' ').slice(0, 40))) })()"""


def flyout_items(pane: Pane) -> list:
    out = pane.run_js("top", FLYOUT_JS) or []
    drop = re.compile(r"account|@|profile|signed in|draft|HOL live test", re.I)
    return [[x for x in grp if isinstance(x, str) and not drop.search(x)][:15] for grp in out if isinstance(grp, list)][:4]


def close_outlook_pane(pane: Pane):
    try:
        pane.click_label(["Close add-in pane", "re:^close .*pane"], timeout=6)
        time.sleep(2)
    except HarnessError:
        pass


def open_pane_with_retry(open_fn, reset_fn, log, what: str = "Workbench task pane"):
    """open_fn(); if it raises HarnessError, reset_fn() (best effort) and open_fn() ONCE more; then fail.

    Only a harness failure (the pane did not appear) is retried; any other error is a bug and propagates as is."""
    try:
        return open_fn()
    except HarnessError as first:
        log({"ev": "task-pane-retry", "what": what, "first": str(first)[:160]})
        try:
            reset_fn()
        except HarnessError as exc:
            log({"ev": "task-pane-retry-reset-failed", "error": str(exc)[:160]})
        try:
            return open_fn()
        except HarnessError as second:
            raise HarnessError(f"{second} (after one retry via the ribbon)") from first


# Visible menu-item labels while a flyout is open (diagnostics for a pane that did not open); account-like, draft and
# fixture labels are dropped, so nothing about the mailbox's content is logged.
MENU_ITEMS_JS = r"""[...document.querySelectorAll('[role=menuitem],[role=menu] button,[role=dialog] button')].filter(b => b.getBoundingClientRect().width > 1)
  .map(b => (b.getAttribute('aria-label') || b.innerText || '').trim().slice(0, 40)).filter(Boolean).slice(0, 20)"""
OMNIA_ITEM_JS = HELPERS + "!!__HOL.find(['Workbench', 'Workbench for Outlook'])"


def menu_items(pane: Pane) -> list:
    out = pane.run_js("top", MENU_ITEMS_JS) or []
    return [x for x in out if isinstance(x, str) and not re.search(r"account|@|profile|signed in|draft|HOL live test", x, re.I)][:10]


def open_outlook_pane_retry(pane: Pane, log, timeout: float = 90, retry_timeout: float = 60, reopen=None) -> str:
    """open_outlook_pane, retried once: Escape (a half-open More apps flyout), close any pane, re-open the message
    (``reopen``: a fresh load of the same item; seen live, A15: after the event form, More apps kept failing to show
    Workbench for 90 s, while a freshly opened message showed it at once), then the ribbon again."""
    def reset():
        try:
            log({"ev": "task-pane-retry-menu", "items": menu_items(pane),
                 "omniaItemVisible": bool(pane.run_js("top", OMNIA_ITEM_JS))})
        except HarnessError:
            pass
        pane.cmd("key", keyCode="Escape")
        time.sleep(1)
        close_outlook_pane(pane)
        if reopen:
            reopen()
        time.sleep(2)
    budgets = iter((timeout, retry_timeout))
    return open_pane_with_retry(lambda: open_outlook_pane(pane, log, next(budgets)), reset, log, "Outlook mail pane")


def open_fixture_pane_clean(pane: Pane, log, folder: str, timeout: float = 90, retry_timeout: float = 60) -> str:
    """The calendar host's way from an event form to the mail pane (A19), the same every time: find the fixture
    once in its folder to learn its stable item address, load THAT address into a clean tab (open_mail_item_clean:
    about:blank, the item, the reading ribbon drawn), then the add-in's own read surface (More apps > Workbench),
    with the one retry, whose reset loads the item clean again."""
    open_fixture_message(pane, log, folder)
    url = fixture_item_url(pane)
    if not url:
        raise HarnessError("the fixture message opened, but its stable item address was not in the tab's address")
    open_mail_item_clean(pane, url, log)
    return open_outlook_pane_retry(pane, log, timeout, retry_timeout, reopen=lambda: open_mail_item_clean(pane, url, log))


# Outlook on the web opens display*Form drafts as a popout (window.open('about:blank', '_blank')); the preview pane
# (a <webview> without allowpopups, like Hermes Desktop's) blocks it. The recorder counts the attempts, nothing else.
POPOUT_RECORDER_JS = r"""(() => { if (!window.__holOpens) { window.__holOpens = 0; const o = window.open;
  window.open = function () { window.__holOpens++; return o.apply(this, arguments) } } window.__holOpens0 = window.__holOpens; return true })()"""
POPOUT_COUNT_JS = r"""((window.__holOpens || 0) - (window.__holOpens0 || 0))"""
DRAFT_OPEN_JS = r"""!!document.querySelector('[role=textbox][aria-label="Message body" i]')"""
REPLY_DRAFT_JS = r"""((marker) => { const bodies = [...document.querySelectorAll('[role=textbox][aria-label="Message body" i]')];
  const flat = s => String(s || '').split(/\s+/).filter(Boolean).join(' ');  // Outlook types non-breaking spaces
  return { drafts: bodies.length, hasMarker: bodies.some(b => flat(b.innerText).includes(flat(marker))) } })"""


def outlook_reply(pane: Pane, log):
    """The person clicks Reply on the open message (the reading pane's own button), as the fallback for a blocked popout."""
    close_outlook_pane(pane)
    pane.click_label(["Reply"], timeout=15, sel="button")
    for _ in range(10):
        time.sleep(1.5)
        if pane.run_js("top", DRAFT_OPEN_JS):
            log({"ev": "outlook-reply-opened"})
            return True
    raise HarnessError("Outlook did not open a reply when Reply was clicked")


def reply_draft_state(pane: Pane, marker: str) -> dict:
    """Is a reply draft open whose body shows the marker, and who is it addressed to (booleans only)."""
    out = pane.run_js("top", f"({REPLY_DRAFT_JS})({json.dumps(marker)})") or {}
    out.update(outlook_recipients(pane))
    return out


# ---- calendar (A13): the same Outlook tab, account and add-in ---------------------------------------------------
# The calendar runs on ONE harness-created appointment ("HOL live test event", no attendees), found by its title and
# reused by every run; a state file (written BEFORE its Save is clicked) makes sure it is never created twice. Every
# other calendar item is left alone: nothing is listed, opened or read but our own event and the forms Hermes asks
# for. Screenshots blur all of Outlook except the task pane (CALENDAR_MASK_JS), on top of the account-area mask.
CAL_URL = "https://outlook.live.com/calendar/0/view/week"
CAL_EVENT_TITLE = "HOL live test event"
CALENDAR_MASK_JS = r"""((on) => { let s = document.getElementById('__holcalmask'); if (!s) { s = document.createElement('style');
  s.id = '__holcalmask'; document.head.appendChild(s) }
  s.textContent = on ? 'body *:not(:has(iframe)):not(iframe) { filter: blur(10px) !important }' : ''; return true })"""
# In the task pane (Office.js): the open item's calendar fields. The subject comes back only when it is our fixture;
# attendees are counted, never named or addressed.
CAL_PANE_JS = r"""new Promise(res => { const it = Office.context.mailbox.item; const type = String((it && it.itemType) || '');
  if (type.toLowerCase() !== 'appointment') return res({ itemType: type || 'none' });
  const g = p => new Promise(r => { const v = it[p]; if (v && typeof v.getAsync === 'function') v.getAsync(x => r(x.value)); else r(v) });
  Promise.all([g('subject'), g('start'), g('end'), g('location'), g('requiredAttendees'), g('optionalAttendees'),
    new Promise(r => it.body.getAsync('text', x => r(x.value)))]).then(([s, st, en, loc, req, opt, body]) => {
    const mine = String(s || '').includes('HOL live test');
    res({ itemType: 'appointment', mode: typeof it.subject === 'string' ? 'read' : 'compose', subjectIsFixture: mine,
      subject: mine ? String(s) : null, location: !loc || /^HOL /.test(String(loc)) ? String(loc || '') : '(other)',
      startUtc: st ? new Date(st).toISOString() : null, endUtc: en ? new Date(en).toISOString() : null,
      offsetMinutes: (() => { try { return Office.context.mailbox.convertToLocalClientTime(new Date()).timezoneOffset } catch (e) { return null } })(),
      required: (req || []).length, optional: (opt || []).length, body: mine ? String(body || '').slice(0, 2000) : null }) }) })"""
EVENT_COUNT_JS = r"""((title) => [...document.querySelectorAll('[aria-label]')].filter(e => (e.getAttribute('aria-label') || '').includes(title)
  && (e.getAttribute('role') === 'button' || e.tagName === 'BUTTON' || e.closest('[role=gridcell],[role=main]'))).length)"""
EVENT_FORM_OPEN_JS = r"""!!document.querySelector('input[placeholder="Add title"]')"""
DISCARD_DIALOG_JS = r"""[...document.querySelectorAll('[role=dialog],[role=alertdialog]')].some(x => /^\s*Discard changes/i.test(x.innerText || ''))"""
# The new-event form Outlook opened (if it did): booleans for our own requested values only.
NEW_EVENT_FORM_JS = r"""((want) => { const t = [...document.querySelectorAll('input[placeholder="Add title"]')].find(i => i.value === want.subject);
  if (!t) return { form: !!document.querySelector('input[placeholder="Add title"]'), titleMatches: false };
  const flat = s => String(s || '').split(/\s+/).filter(Boolean).join(' ');
  const body = document.querySelector('[aria-label="Event body" i]');
  const loc = document.querySelector('input[aria-label="Search for a location" i]');
  const locText = (loc && loc.value) || [...document.querySelectorAll('[aria-label^="Location" i], [title]')].map(e => e.getAttribute('title') || '').join(' ');
  const when = [...document.querySelectorAll('div[role=button], button')].map(e => e.innerText || '').find(x => /\d{1,2}\/\d{1,2}\/\d{4} .*(AM|PM)/.test(x)) || '';
  return { form: true, titleMatches: true, bodyHasMarker: !!body && flat(body.innerText).includes(flat(want.marker)),
    locationMatches: flat(locText).includes(flat(want.location)), timeMatches: flat(when).includes(want.when), when: flat(when).slice(0, 60) } })"""


def calendar_shot(pane: Pane, name: str) -> str:
    """A window screenshot with all of Outlook blurred except the task pane (and the account area masked)."""
    try:
        pane.run_js("top", f"({CALENDAR_MASK_JS})(true)")
        time.sleep(0.5)
        return pane.shot(name)
    finally:
        try:
            pane.run_js("top", f"({CALENDAR_MASK_JS})(false)")
        except HarnessError:
            pass


def open_calendar(pane: Pane):
    pane.cmd("navigate", url=CAL_URL, timeout=90)
    end = time.time() + 45
    while time.time() < end:
        time.sleep(3)
        try:
            if pane.run_js("top", "!!document.querySelector('button[aria-label=\"New event\" i]') || [...document.querySelectorAll('button')].some(b => (b.innerText || '').trim() === 'New event')"):
                time.sleep(3)
                return True
        except HarnessError:
            pass
    raise HarnessError("Outlook calendar did not load")


def test_event_count(pane: Pane) -> int:
    return int(pane.run_js("top", f"({EVENT_COUNT_JS})({json.dumps(CAL_EVENT_TITLE)})") or 0)


def _click_js(pane: Pane, js: str, what: str):
    hit = pane.cmd("clickIn", frame="top", js=js)
    if not hit.get("ok"):
        raise HarnessError(f"no {what} to click")
    return hit


def create_test_event_once(pane: Pane, state_file, log) -> dict:
    """The ONE approved test appointment: title CAL_EVENT_TITLE, Outlook's default slot, no attendees. Reused when it
    is on the calendar; never created again once the state file exists (an interrupted run stops instead)."""
    from pathlib import Path as _P
    state_file = _P(state_file)
    open_calendar(pane)
    if test_event_count(pane) > 0:
        return {"event": "reused", "created": False}
    if state_file.exists():
        raise HarnessError("the test event was created once already and is not in this week's view; not creating another "
                           "(open its week, or delete the state file after checking the calendar by hand)")
    pane.click_label(["New event"], timeout=15, sel="button")
    time.sleep(4)
    _click_js(pane, r"""(() => { const el = document.querySelector('input[placeholder="Add title"]'); if (!el) return null;
      const r = el.getBoundingClientRect(); return { x: r.x + 20, y: r.y + r.height / 2, label: 'title' } })()""", "title field")
    pane.cmd("type", text=CAL_EVENT_TITLE)
    time.sleep(1.5)
    chk = pane.run_js("top", r"""(() => { const t = document.querySelector('input[placeholder="Add title"]');
      const a = document.querySelector('[aria-label="Invite required attendees"]');
      return { title: t ? t.value : null, attendeesEmpty: !!a && !(a.innerText || '').trim(),
               removeButtons: document.querySelectorAll('[aria-label^="Remove "]').length } })()""") or {}
    log({"ev": "calendar-fixture-check", **chk})
    if chk.get("title") != CAL_EVENT_TITLE or not chk.get("attendeesEmpty") or chk.get("removeButtons"):
        discard_event_form(pane, log)
        raise HarnessError(f"test event not saved: title or attendees not as expected ({chk})")
    state_file.parent.mkdir(parents=True, exist_ok=True)
    state_file.write_text(json.dumps({"created": True, "at": time.strftime("%Y-%m-%dT%H:%M:%S"), "title": CAL_EVENT_TITLE}))
    pane.click_label(["Save"], timeout=10, sel="button")  # the only Save the harness ever clicks (no attendees: no invitation)
    time.sleep(5)
    log({"ev": "calendar-fixture-created"})
    return {"event": "created", "created": True}


TEST_EVENT_POINT_JS = r"""((title, scroll) => { const el = [...document.querySelectorAll('[aria-label]')].find(e => (e.getAttribute('aria-label') || '')
  .includes(title)); if (!el) return null; if (scroll) el.scrollIntoView({block: 'center'}); const r = el.getBoundingClientRect();
  return { x: Math.round(r.x + Math.min(20, r.width / 2)), y: Math.round(r.y + Math.min(8, r.height / 2)), label: 'test event' } })"""
PEEK_BUTTONS_JS = r"""[...document.querySelectorAll('button')].filter(b => b.getBoundingClientRect().width > 1)
  .map(b => (b.getAttribute('aria-label') || b.innerText || '').trim()).filter(l => /^(View event|Edit|Delete)$/i.test(l))"""


def open_test_event_peek(pane: Pane, log, tries: int = 3, sleep=time.sleep) -> list:
    """Click our ONE test event until its peek card shows (View event / Delete). Seen live (A15): right after a fresh
    calendar load Outlook still scrolls the grid, so a point computed at once lands beside the event and no peek
    opens; the point is re-read until it holds still, and a click that opened no peek is retried (Escape first)."""
    title = json.dumps(CAL_EVENT_TITLE)
    for attempt in range(1, tries + 1):
        last = pane.run_js("top", f"({TEST_EVENT_POINT_JS})({title}, true)")
        for _ in range(6):  # the point holds still for one second
            sleep(1)
            now = pane.run_js("top", f"({TEST_EVENT_POINT_JS})({title}, false)")
            if now and last and (now.get("x"), now.get("y")) == (last.get("x"), last.get("y")):
                break
            last = now
        if not last:
            raise HarnessError("no test event to click")
        _click_js(pane, f"({TEST_EVENT_POINT_JS})({title}, false)", "test event")
        sleep(2.5)
        buttons = pane.run_js("top", PEEK_BUTTONS_JS) or []
        if "Delete" in buttons or "View event" in buttons:
            return list(buttons)
        log({"ev": "calendar-peek-retry", "attempt": attempt})
        pane.cmd("key", keyCode="Escape")  # whatever the click opened instead (never saved)
        sleep(2)
        if pane.run_js("top", EVENT_FORM_OPEN_JS):
            discard_event_form(pane, log)
    raise HarnessError(f"the test event's peek card did not open after {tries} clicks")


def open_test_event(pane: Pane, log):
    """Open our event's own form: its peek card, then View event (the organizer's edit form)."""
    open_calendar(pane)
    open_test_event_peek(pane, log)
    pane.click_label(["View event"], timeout=10, sel="button")
    for _ in range(10):
        time.sleep(1.5)
        if pane.run_js("top", EVENT_FORM_OPEN_JS):
            log({"ev": "calendar-event-opened"})
            return True
    raise HarnessError("the test event's form did not open")


def open_calendar_pane(pane: Pane, log, timeout: float = 60) -> str:
    """Apps > Workbench in the open event form (the same add-in as mail)."""
    end = time.time() + timeout
    while time.time() < end:
        f = task_pane_frame(pane, "outlook")
        if f:
            return f
        try:
            pane.click_label(["Apps"], timeout=8, sel="button")
            time.sleep(2)
            pane.click_label(RIBBON_LABELS + ["Workbench for Outlook"], timeout=8)
        except HarnessError:
            pass
        time.sleep(5)
    raise HarnessError("Workbench task pane did not open in the event form")


def discard_event_form(pane: Pane, log):
    """Close the open event form WITHOUT saving: Close, then Discard if Outlook asks. Never Save, never Send."""
    if not pane.run_js("top", EVENT_FORM_OPEN_JS):
        return False
    # The form's own Close (it is a dialog holding the title field), then OK in Outlook's "Discard changes" dialog
    # (seen live, A13: "Are you sure you want to discard changes to this event?" OK / Cancel).
    _click_js(pane, r"""(() => { const d = [...document.querySelectorAll('[role=dialog]')].find(x => x.querySelector('input[placeholder="Add title"]'));
      const el = [...((d || document).querySelectorAll('button[aria-label="Close"]'))].find(b => b.getBoundingClientRect().width);
      if (!el) return null; const r = el.getBoundingClientRect(); return { x: r.x + r.width / 2, y: r.y + r.height / 2, label: 'Close' } })()""",
              "Close")
    time.sleep(2)
    if pane.run_js("top", DISCARD_DIALOG_JS):
        _click_js(pane, r"""(() => { const d = [...document.querySelectorAll('[role=dialog],[role=alertdialog]')].find(x => /^\s*Discard changes/i.test(x.innerText || ''));
          const el = d && [...d.querySelectorAll('button')].find(b => /^(OK|Discard)$/i.test((b.innerText || b.getAttribute('aria-label') || '').trim()));
          if (!el) return null; const r = el.getBoundingClientRect(); return { x: r.x + r.width / 2, y: r.y + r.height / 2, label: 'OK' } })()""",
                  "Discard changes OK")
    time.sleep(2)
    log({"ev": "calendar-form-closed", "stillOpen": bool(pane.run_js("top", EVENT_FORM_OPEN_JS))})
    return True


def new_event_form_state(pane: Pane, want: dict) -> dict:
    return pane.run_js("top", f"({NEW_EVENT_FORM_JS})({json.dumps(want)})") or {}


def delete_test_event(pane: Pane, state_file, log) -> dict:
    """End of the stage: delete our ONE test event (peek > Delete > confirm), then check it is gone."""
    from pathlib import Path as _P
    open_calendar(pane)
    if test_event_count(pane) == 0:
        return {"deleted": False, "found": False}
    open_test_event_peek(pane, log)
    pane.click_label(["Delete"], timeout=10, sel="button")
    time.sleep(2)
    try:
        pane.click_label(["Delete", "OK", "Yes"], timeout=5, sel="[role=dialog] button,[role=alertdialog] button")
    except HarnessError:
        pass
    time.sleep(4)
    gone = test_event_count(pane) == 0
    log({"ev": "calendar-fixture-deleted", "gone": gone})
    if gone:  # keep the record: a later run must not quietly create another event in this stage
        _P(state_file).write_text(json.dumps({"created": True, "deleted": True, "at": time.strftime("%Y-%m-%dT%H:%M:%S"),
                                              "title": CAL_EVENT_TITLE}))
    return {"deleted": gone, "found": True}
