// SPDX-License-Identifier: Apache-2.0
// node --test tests/js/ : Outlook (Mailbox API) executor against the in-memory mock.
// Hard rules under test: no send call, no recipient change, the open item only, attachment names/sizes only.
import assert from 'node:assert/strict'
import fs from 'node:fs'
import path from 'node:path'
import { test } from 'node:test'
import { fileURLToPath } from 'node:url'
import { MockMailbox, installOutlookMock } from '../../addin/dev/outlook-mock.js'
import { RECIPIENT_ARGS, SUPPORTED_TOOLS, execute, executor, itemMode } from '../../addin/outlook/executor.js'

const HERE = path.dirname(fileURLToPath(import.meta.url))
const SRC = fs.readFileSync(path.join(HERE, '../../addin/outlook/executor.js'), 'utf8')
const NOTE = 'Weekly sync.\nDecision: ship on Friday.\nDecision: keep the price.\nAction: Sam drafts the FAQ.'

function fresh(mode = 'read', opts = {}) {
  executor.clearUndo()
  const mb = installOutlookMock(new MockMailbox())
  if (mode === 'read') mb.openRead({ body: NOTE, attachments: [{ name: 'agenda.pdf', size: 1234, attachmentType: 'file', isInline: false }], ...opts })
  else mb.openCompose({ body: '<p>Hi team,</p>', ...opts })
  return mb
}
async function ok(tool, args = {}) { const r = await execute(tool, args); assert.equal(r.ok, true, `${tool}: ${JSON.stringify(r)}`); return r.data }
async function err(tool, args = {}) { const r = await execute(tool, args); assert.equal(r.ok, false, `${tool} should fail`); return r.error }

test('pane advertises exactly the four mail tools, the three calendar commands and undo_last (no send tool exists)', () => {
  assert.deepEqual([...SUPPORTED_TOOLS].sort(), ['draft_event', 'draft_new', 'draft_reply', 'get_event', 'get_item', 'insert',
    'undo_last', 'update_event'])
})

test('executor source calls no send API, sets no recipient, never reads attachment contents or other items', () => {
  for (const bad of [/\bsendAsync\b/, /\.send\(/, /\b(to|cc|bcc)\.(setAsync|addAsync)\b/, /getAttachmentContentAsync\s*\(/,
    /displayReplyAllForm\s*\(/, /toRecipients\s*:/, /ccRecipients\s*:/, /bccRecipients\s*:/, /makeEwsRequestAsync/, /getCallbackTokenAsync/,
    /\bfetch\(/, /XMLHttpRequest/]) {
    assert.doesNotMatch(SRC, bad, String(bad))
  }
})

test('get_item (read): subject, from, to/cc, date, body text, attachment names and sizes only, plus a stable key', async () => {
  fresh()
  const d = await ok('get_item')
  assert.equal(d.mode, 'read')
  assert.equal(d.subject, 'HOL live test: weekly sync')
  assert.deepEqual(d.from, { name: 'Test Account', email: 'test@example.invalid' })
  assert.equal(d.to.length, 1)
  assert.deepEqual(d.cc, [])
  assert.equal(d.date, '2026-11-06T12:00:00.000Z')
  assert.equal(d.body.text, NOTE)
  assert.deepEqual(d.attachments, [{ name: 'agenda.pdf', size: 1234, type: 'file', inline: false }])
  assert.match(d.key, /^[0-9a-f]{16}$/)
  assert.equal((await ok('get_item')).key, d.key)
  const t = await ok('get_item', { maxBodyLength: 12 })
  assert.deepEqual([t.body.text, t.body.truncated, t.body.totalLength], ['Weekly sync.', true, NOTE.length])
  assert.deepEqual(globalThis.__hermesMockMailbox.forbidden, [])
})

test('get_item (compose) reads the draft through the async getters', async () => {
  fresh('compose', { subject: 'Draft', to: [{ displayName: 'A', emailAddress: 'a@example.invalid' }] })
  const d = await ok('get_item')
  assert.equal(d.mode, 'compose')
  assert.equal(d.subject, 'Draft')
  assert.equal(d.body.text, 'Hi team,')
  assert.deepEqual(d.to, [{ name: 'A', email: 'a@example.invalid' }])
})

test('get_item with no open item is NotFound', async () => {
  const mb = fresh()
  mb.item = null
  assert.equal((await err('get_item')).category, 'NotFound')
})

test('insert into the body at the cursor; undo restores the prior body exactly', async () => {
  const mb = fresh('compose')
  const before = mb.item._state.body
  const d = await ok('insert', { target: 'body', text: ' HOL marker 1.' })
  assert.deepEqual([d.inserted, d.target, d.undo.available], [14, 'body', true])
  assert.equal(mb.item._state.body, before + ' HOL marker 1.')
  const u = await ok('undo_last')
  assert.equal(u.undone, 'insert')
  assert.equal(mb.item._state.body, before)
  assert.deepEqual(mb.forbidden, [])
})

test('insert into the subject replaces it; undo restores the prior subject', async () => {
  const mb = fresh('compose', { subject: 'Old subject' })
  await ok('insert', { target: 'subject', text: 'HOL live test reply' })
  assert.equal(mb.item._state.subject, 'HOL live test reply')
  await ok('undo_last')
  assert.equal(mb.item._state.subject, 'Old subject')
})

test('insert refuses read mode, bad targets and oversize text; nothing changes', async () => {
  fresh('read')
  assert.equal((await err('insert', { target: 'body', text: 'x' })).category, 'Validation')
  const mb = fresh('compose')
  const before = JSON.stringify(mb.item._state)
  assert.equal((await err('insert', { target: 'to', text: 'x' })).category, 'Validation')
  assert.equal((await err('insert', { target: 'body', text: 'x'.repeat(5001) })).category, 'Validation')
  assert.equal((await err('insert', { target: 'subject', text: 'x'.repeat(256) })).category, 'Validation')
  assert.equal(JSON.stringify(mb.item._state), before)
})

test('any recipient or send argument is refused on every tool before anything happens', async () => {
  for (const tool of ['get_item', 'insert', 'draft_reply', 'draft_new']) {
    for (const k of ['to', 'cc', 'bcc', 'toRecipients', 'send']) {
      const mb = fresh(tool === 'insert' ? 'compose' : 'read')
      const e = await err(tool, { text: 'hi', [k]: ['someone@example.invalid'] })
      assert.equal(e.category, 'Validation', `${tool}.${k}`)
      assert.match(e.message, /never sends mail and never adds, removes or changes recipients/)
      assert.deepEqual([mb.forms, mb.forbidden], [[], []])
    }
  }
  assert.ok(RECIPIENT_ARGS.includes('bcc') && RECIPIENT_ARGS.includes('attachments'))
})

test('undo only touches the item it was recorded on', async () => {
  const mb = fresh('compose')
  await ok('insert', { target: 'body', text: 'X' })
  const first = mb.item
  mb.openCompose({ body: '<p>Other draft</p>' })
  const e = await err('undo_last')
  assert.equal(e.category, 'NotFound')
  assert.equal(mb.item._state.body, '<p>Other draft</p>')
  assert.equal(first._state.body.endsWith('X'), true)
})

test('a failed insert leaves nothing on the undo stack and reports the Office error', async () => {
  const mb = fresh('compose')
  mb.failNext = 'setSelectedDataAsync'
  const e = await err('insert', { target: 'body', text: 'X' })
  assert.equal(e.category, 'Office')
  assert.equal(executor.undoDepth(), 0)
})

test('draft_reply opens a reply form (never reply-all) with escaped text; nothing is sent', async () => {
  const mb = fresh('read')
  const d = await ok('draft_reply', { text: 'Thanks <all>.\nD1 & D2 agreed.' })
  assert.deepEqual([d.requested, d.sent, d.undo.available, 'opened' in d], ['reply', false, false, false])
  assert.equal(mb.forms.length, 1)
  assert.equal(mb.forms[0].kind, 'reply')
  assert.equal(mb.forms[0].html, '<p>Thanks &lt;all&gt;.</p><p>D1 &amp; D2 agreed.</p>')
  assert.deepEqual(mb.forbidden, [])
  fresh('compose')
  assert.equal((await err('draft_reply', { text: 'x' })).category, 'Validation')
})

test('draft_new opens a new message with subject and body only (no recipients, no attachments)', async () => {
  const mb = fresh('read')
  const d = await ok('draft_new', { subject: 'Follow-up', text: 'Hello' })
  assert.equal(d.requested, 'new')
  assert.deepEqual(Object.keys(mb.forms[0]).sort(), ['htmlBody', 'kind', 'subject'])
  assert.deepEqual(mb.forbidden, [])
  fresh('compose')
  assert.equal((await err('draft_new', { text: 'x' })).category, 'Validation')
})

test('itemMode tells read from compose', () => {
  const mb = fresh('read')
  assert.equal(itemMode(mb.item), 'read')
  mb.openCompose()
  assert.equal(itemMode(mb.item), 'compose')
})
