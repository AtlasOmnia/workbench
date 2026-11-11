// SPDX-License-Identifier: Apache-2.0
// node --test tests/js/ : live-test harness tab start-up (scripts/livetest/macbook/tabscript.js).
// Real Electron 44.9.3 throws from <webview>.getWebContentsId() on did-attach ("... and the dom-ready event
// emitted before this method can be called"); the id only exists from dom-ready. The fake webview below does
// the same, and the open-tab script must resolve on dom-ready, reject on its deadline, and never hang.
import assert from 'node:assert/strict'
import crypto from 'node:crypto'
import fs from 'node:fs'
import { createRequire } from 'node:module'
import path from 'node:path'
import { test } from 'node:test'
import { fileURLToPath } from 'node:url'
import vm from 'node:vm'

const require = createRequire(import.meta.url)
const DIR = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '../../scripts/livetest/macbook')
const { HARNESS_FILES, bare, harnessHash, openTabScript } = require(path.join(DIR, 'tabscript.js'))

const NOT_READY = 'The WebView must be attached to the DOM and the dom-ready event emitted before this method can be called.'

// A host page with #pane and #tabs; appended webviews fire did-attach at once and dom-ready after `domReadyMs`
// (null: never), like Electron's guest view.
function fakeHost({ domReadyMs = 5, id = 7, attach = true } = {}) {
  const made = []
  class El {
    constructor(tag) { this.tag = tag; this.attrs = {}; this.dataset = {}; this.listeners = {}; this.ready = false }
    setAttribute(k, v) { this.attrs[k] = v }
    addEventListener(ev, fn) { (this.listeners[ev] ||= []).push(fn) }
    fire(ev) { for (const fn of this.listeners[ev] || []) fn() }
    getWebContentsId() { if (!this.ready) throw new Error(NOT_READY); return id }
  }
  const pane = { appendChild(w) {
    made.push(w)
    if (!attach) return
    setTimeout(() => w.fire('did-attach'), 0)
    if (domReadyMs != null) setTimeout(() => { w.ready = true; w.fire('dom-ready') }, domReadyMs)
  } }
  const tabs = { children: [], appendChild(s) { this.children.push(s) } }
  const document = { createElement: tag => new El(tag), getElementById: k => ({ pane, tabs })[k] }
  return { ctx: vm.createContext({ document, setTimeout, clearTimeout, Error }), made, tabs }
}
const run = (host, opts) => vm.runInContext(openTabScript({ el: 'w-main', name: 'main', partition: 'persist:x', url: 'https://excel.cloud.microsoft/', ...opts }), host.ctx)

test('tab script: a getWebContentsId throw on did-attach is "not yet"; the id comes on dom-ready', async () => {
  const host = fakeHost({ domReadyMs: 20, id: 42 })
  assert.equal(await run(host), 42)
  const w = host.made[0]
  assert.equal(w.id, 'w-main')
  assert.equal(w.attrs.partition, 'persist:x')
  assert.equal(w.attrs.src, 'https://excel.cloud.microsoft/')
  assert.equal(w.attrs.webpreferences, 'contextIsolation=yes,nodeIntegration=no,sandbox=yes')
  assert.equal(host.tabs.children[0].dataset.tab, 'main')
})

test('tab script: a guest that never reaches dom-ready rejects at the deadline instead of hanging', async () => {
  const host = fakeHost({ domReadyMs: null })
  await assert.rejects(run(host, { timeoutMs: 60 }), /guest id not available/)
})

test('tab script: a webview that never attaches rejects at the deadline', async () => {
  const host = fakeHost({ attach: false })
  await assert.rejects(run(host, { timeoutMs: 40 }), /guest id not available/)
})

test('tab script: a host page without #pane rejects at once (no silent hang)', async () => {
  const host = fakeHost()
  host.ctx.document.getElementById = () => null
  await assert.rejects(run(host, { timeoutMs: 5000 }), /appendChild|null/)
})

test('tab script: names and urls are embedded as JSON literals', async () => {
  const host = fakeHost()
  await run(host, { el: 'w-a_b', name: "a'b\"`${x}`", url: "https://x/?q='\"" })
  assert.equal(host.tabs.children[0].textContent, "a'b\"`${x}`")
  assert.equal(host.made[0].attrs.src, "https://x/?q='\"")
})

test('harness hash: sha256 over main.js then tabscript.js, the same bytes run_macbook.sh hashes', () => {
  assert.deepEqual(HARNESS_FILES, ['main.js', 'tabscript.js'])
  const h = crypto.createHash('sha256')
  for (const f of HARNESS_FILES) h.update(fs.readFileSync(path.join(DIR, f)))
  assert.equal(harnessHash(DIR), h.digest('hex').slice(0, 12))
  const sh = fs.readFileSync(path.join(DIR, '..', 'run_macbook.sh'), 'utf8')
  assert.match(sh, /cat "\$HERE\/macbook\/main\.js" "\$HERE\/macbook\/tabscript\.js" \| shasum -a 256 \| cut -c1-12/)
  assert.match(sh, /"\$HERE\/macbook\/tabscript\.js"[^\n]*"\$MAC:hermes-office-livetest\/app\/"/)
})

test('bare(): tab and navigation URLs keep no query, fragment or OneDrive personal path (A8 scrub finding)', () => {
  assert.equal(bare('https://excel.cloud.microsoft/open/onedrive/?docId=ABC&driveId=0123456789abcdef#x'), 'https://excel.cloud.microsoft/open/onedrive/')
  for (const u of ['https://onedrive.live.com/personal/0123456789abcdef/_layouts/15/Doc.aspx?resid=1',
    'https://d.docs.live.net/0123456789abcdef/Documents/HOL%20live%20test.docx',
    'https://1drv.ms/w/s!Abc123', 'https://contoso-my.sharepoint.com/personal/jane_contoso_com/Documents/x.docx']) {
    const b = bare(u)
    assert.equal(b, new URL(u).origin + '/<path>', u)
  }
  assert.equal(bare('https://word.cloud.microsoft/personal/0123456789ABCDEF/x'), 'https://word.cloud.microsoft/personal/<cid>/x')
  assert.equal(bare('https://onenote.cloud.microsoft/en-us/0123456789abcdef/notebook'), 'https://onenote.cloud.microsoft/en-us/<cid>/notebook')
  assert.equal(bare('not a url'), 'not a url')
  const main = fs.readFileSync(path.join(DIR, 'main.js'), 'utf8')
  assert.match(main, /const \{[^}]*\bbare\b[^}]*\} = require\('\.\/tabscript(\.js)?'\)/)
  assert.doesNotMatch(main, /const bare = /)
})

// A10b: the Outlook tab signs a DIFFERENT throwaway account in, so it must never share the cookie jar of the
// Excel/Word/PowerPoint/OneNote tabs (signing in there would replace their account session).
test('partitionFor(): Outlook gets its own persistent partition; every other tab keeps the shared one', () => {
  const { partitionFor, DEFAULT_PARTITION, TAB_PARTITIONS } = require(path.join(DIR, 'tabscript.js'))
  assert.equal(DEFAULT_PARTITION, 'persist:hol-livetest')
  assert.equal(TAB_PARTITIONS.outlook, 'persist:hol-outlook')
  for (const n of ['main', 'excel', 'word', 'powerpoint', 'onenote', 'x']) assert.equal(partitionFor(n), 'persist:hol-livetest', n)
  assert.equal(partitionFor('outlook'), 'persist:hol-outlook')
  assert.equal(partitionFor('outlook-compose'), 'persist:hol-outlook')   // any outlook* tab shares the Outlook jar
  // an explicit per-open partition wins, but only a persist:hol-* name (no in-memory or foreign jar)
  assert.equal(partitionFor('excel', 'persist:hol-outlook'), 'persist:hol-outlook')
  assert.equal(partitionFor('outlook', 'persist:hol-livetest'), 'persist:hol-livetest')
  for (const bad of ['hol-outlook', 'persist:other', 'persist:hol-', 'persist:hol-a b', 'persist:hol-../x'])
    assert.throws(() => partitionFor('excel', bad), /partition/, bad)
  // env-style overrides (HOL_TAB_PARTITIONS="name=persist:hol-y,...") are validated the same way
  assert.equal(partitionFor('word', undefined, 'word=persist:hol-word2'), 'persist:hol-word2')
  assert.equal(partitionFor('excel', undefined, 'word=persist:hol-word2'), 'persist:hol-livetest')
  assert.throws(() => partitionFor('word', undefined, 'word=persist:evil'), /partition/)
})

test('main.js: tabs open with partitionFor(), clearCache covers every partition in use, no single hard-coded jar', () => {
  const main = fs.readFileSync(path.join(DIR, 'main.js'), 'utf8')
  assert.match(main, /partitionFor\(/)
  assert.doesNotMatch(main, /openTabScript\(\{[^}]*partition: PARTITION[,}]/)
  assert.match(main, /clearCache\(\)[\s\S]{0,200}partitions/)
  assert.match(main, /partition: t\.partition|partition: part/)
})

test('bare(): Outlook on the web URLs keep no message, folder or conversation id', () => {
  assert.equal(bare('https://outlook.live.com/mail/0/inbox/id/AQMkADAwATM0MDAAMS1iZjAy%3D'), 'https://outlook.live.com/mail/<path>')
  assert.equal(bare('https://outlook.live.com/mail/'), 'https://outlook.live.com/mail/')
  assert.equal(bare('https://outlook.office.com/mail/deeplink/compose?x=1'), 'https://outlook.office.com/mail/<path>')
  assert.equal(bare('https://outlook.cloud.microsoft/mail/inbox/id/ABC'), 'https://outlook.cloud.microsoft/mail/<path>')
})

test('main.js: a frame offset counts the iframe padding (A10b: Outlook add-in dialog clicks landed 17 px high)', () => {
  const main = fs.readFileSync(path.join(DIR, 'main.js'), 'utf8')
  assert.match(main, /el\.clientTop \+ \(parseFloat\(cs\.paddingTop\) \|\| 0\)/)
  assert.match(main, /el\.clientLeft \+ \(parseFloat\(cs\.paddingLeft\) \|\| 0\)/)
})
