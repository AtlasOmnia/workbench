// SPDX-License-Identifier: Apache-2.0
// node --test: the OneNote pane's read-only Links/Backlinks panel (rendering, escaping, click-to-open).
import assert from 'node:assert/strict'
import { test } from 'node:test'
import { MockNotebook, installOneNoteMock } from '../../addin/dev/onenote-mock.js'
import { clearJournal, execute } from '../../addin/onenote/executor.js'
import { createLinksPanel, panelModel, renderLinksPanel } from '../../addin/onenote/panel.js'

/** Minimal DOM: elements keep children and textContent; assigning innerHTML/outerHTML throws. */
class FakeEl {
  constructor(tag) {
    this.tagName = tag.toUpperCase()
    this.children = []
    this.dataset = {}
    this.listeners = {}
    this.hidden = false
    this._text = ''
  }
  set innerHTML(_) { throw new Error('innerHTML is never used for note text') }
  set outerHTML(_) { throw new Error('outerHTML is never used for note text') }
  insertAdjacentHTML() { throw new Error('insertAdjacentHTML is never used for note text') }
  set textContent(v) { this._text = String(v); this.children = [] }
  get textContent() { return this._text + this.children.map(c => c.textContent).join('') }
  append(...xs) { for (const x of xs) { if (!(x instanceof FakeEl)) throw new Error('only elements are appended'); this.children.push(x) } }
  replaceChildren(...xs) { this.children = []; this._text = ''; this.append(...xs) }
  addEventListener(k, fn) { (this.listeners[k] ||= []).push(fn) }
  click() { for (const fn of this.listeners.click || []) fn() }
  all(tag) { const out = []; const walk = e => { for (const c of e.children) { if (c.tagName === tag) out.push(c); walk(c) } }; walk(this); return out }
}
const doc = { createElement: t => new FakeEl(t) }

const XSS = '<img src=x onerror=alert(1)>'
const PANEL = {
  pageId: 'p-hub', title: 'Q4 Launch',
  links: [{ pageId: 'p-kick', title: 'Kickoff 2026-11-02', label: 'the kickoff' }, { pageId: 'p-x', title: XSS, label: XSS }],
  backlinks: [{ pageId: 'p-kick', title: 'Kickoff 2026-11-02', label: 'Kickoff 2026-11-02' }],
  unresolved: ['Budget'], ambiguous: ['Risks'], more: { links: 0, backlinks: 3 },
  coverage: { indexed: 4, discovered: 9, label: 'indexed 4 of 9 pages' },
  hint: 'Ask Hermes to run brain_sync (with crawl) to index more of the notebook.'
}

test('panel shows Links and Backlinks as clickable entries with coverage and the brain_sync hint', () => {
  const root = new FakeEl('section')
  root.hidden = true
  const opened = []
  renderLinksPanel(root, PANEL, { onOpen: id => opened.push(id), doc })
  assert.equal(root.hidden, false)
  assert.deepEqual(root.all('H2').map(h => h.textContent), ['Links', 'Backlinks'])
  const buttons = root.all('BUTTON')
  assert.deepEqual(buttons.map(b => b.textContent), ['the kickoff', XSS, 'Kickoff 2026-11-02'])
  assert.deepEqual(buttons.map(b => b.dataset.pageId), ['p-kick', 'p-x', 'p-kick'])
  buttons[2].click()
  assert.deepEqual(opened, ['p-kick'])
  const all = root.textContent
  assert.match(all, /indexed 4 of 9 pages\. Ask Hermes to run brain_sync/)
  assert.match(all, /and 3 more/)
  assert.match(all, /Not created yet: Budget/)
  assert.match(all, /Ambiguous \(several pages share this title\): Risks/)
})

test('note text is data: markup stays literal text, control characters and overlong labels are neutralized', () => {
  const root = new FakeEl('section')
  const long = 'L'.repeat(500)
  renderLinksPanel(root, { ...PANEL, links: [{ pageId: 'p1', title: long, label: `a\u0000b${XSS}` },
    { pageId: 42, title: 'bad id' }, { pageId: 'x'.repeat(301), title: 'too long id' }, null], backlinks: 'nope' }, { doc })
  const [b] = root.all('BUTTON')
  assert.equal(root.all('BUTTON').length, 1) // non-string / oversized ids are dropped, not rendered
  assert.equal(b.textContent, `a b${XSS}`)
  assert.ok(b.title.length < 220)
  assert.equal(panelModel({ ...PANEL, links: [{ pageId: 'p', title: long }] }).links[0].label.length, 120)
  assert.match(root.textContent, /Backlinks.*none/)
  assert.equal(root.all('IMG').length, 0)
})

test('unusable payloads hide the panel instead of rendering stale content', () => {
  const root = new FakeEl('section')
  renderLinksPanel(root, PANEL, { doc })
  assert.equal(renderLinksPanel(root, null, { doc }), null)
  assert.equal(root.hidden, true)
  assert.equal(root.children.length, 0)
  assert.equal(panelModel({ title: 'no id' }), null)
})

test('clicking an entry opens the page through open_page (no write) and refreshes; failures are shown, not retried', async () => {
  clearJournal()
  const nb = installOneNoteMock(new MockNotebook({ name: 'N', sections: [{ name: 'S', pages: [
    { title: 'Hub', html: ['<p>[[Target]]</p>'] }, { title: 'Target', html: ['<p>t</p>'] }] }] }))
  const target = nb.pageByTitle('Target')
  const calls = []
  const exec = async (tool, args) => { calls.push(tool); return execute(tool, args) }
  let refreshed = 0
  const root = new FakeEl('section')
  const panel = createLinksPanel({ root, execute: exec, refresh: () => { refreshed++ }, doc })
  assert.equal(panel.show({ accepted: false }), null)
  panel.show({ accepted: true, panel: { ...PANEL, links: [{ pageId: target.id, title: 'Target' }] } })
  const before = JSON.stringify(nb.allSections().map(s => s.pages.map(p => p.outlines.length)))
  root.all('BUTTON')[0].click()
  await new Promise(r => setTimeout(r, 0))
  assert.equal(nb.activePage().id, target.id)
  assert.equal(refreshed, 1)
  assert.equal(JSON.stringify(nb.allSections().map(s => s.pages.map(p => p.outlines.length))), before)
  assert.equal(await panel.open('missing-page'), false)
  assert.deepEqual(calls, ['open_page', 'open_page'])
  assert.match(root.textContent, /Could not open that page/)
  assert.equal(refreshed, 1)
})
