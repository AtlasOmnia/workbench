// SPDX-License-Identifier: Apache-2.0
// Canonical Outlook tool -> Office.js (Mailbox API) executor for the Workbench task pane.
// Same contract as the other hosts: plain JSON in, bridge envelope out.
//
// Hard limits, enforced here and again in the plugin:
//   - never sends mail: no send API is called anywhere in this file;
//   - never adds, removes or changes recipients (no to/cc/bcc argument is accepted, none is ever set);
//   - reads only the item open in this pane (Office.context.mailbox.item), never other mailbox items;
//   - attachments: names, types and sizes only, never their contents.
// Undo: compose edits capture the prior body (or subject) and undo puts it back on the same open item.
// A reply or new-message draft is REQUESTED from Outlook (display*Form), a separate unsent draft with no undo here.
// Outlook on the web opens those forms as a popout window (window.open), which a <webview> without allowpopups (the
// Hermes Desktop preview pane, A10b live) blocks: the call succeeds but no draft appears, so results never say opened.
import { ExecError, createExecutor, str } from '../shared/runtime.js'
import { CALENDAR_MUTATING, calendarTools } from './calendar.js'

export const MAX_BODY = 20000
export const MAX_INSERT = 5000
export const SNAPSHOT_MAX_CHARS = 1024 * 1024
const MAX_SUBJECT = 255
// Argument names that would address mail. Rejected by name on every tool, whatever it is.
export const RECIPIENT_ARGS = Object.freeze(['to', 'cc', 'bcc', 'from', 'toRecipients', 'ccRecipients', 'bccRecipients',
  'recipients', 'requiredAttendees', 'optionalAttendees', 'replyTo', 'sender', 'attachments', 'send', 'sendNow'])

function office() {
  const o = globalThis.Office
  if (!o || !o.context) throw new ExecError('Unavailable', 'Office.js is not ready in this pane')
  return o
}

function mailbox() {
  const m = office().context.mailbox
  if (!m) throw new ExecError('Unavailable', 'This pane is not running inside Outlook (no mailbox)')
  return m
}

function openItem() {
  const it = mailbox().item
  if (!it) throw new ExecError('NotFound', 'No message or event is open in Outlook. Open (or start) one, then ask again.')
  return it
}

/** Compose forms expose async getters (subject.getAsync); read forms expose plain values. */
export function itemMode(it) {
  return it && it.subject && typeof it.subject === 'object' && typeof it.subject.getAsync === 'function' ? 'compose' : 'read'
}

function needMode(it, want, what) {
  const mode = itemMode(it)
  if (mode !== want) {
    throw new ExecError('Validation', want === 'compose'
      ? `${what} works only in a message you are writing (compose). Open a draft or start a reply, then ask again.`
      : `${what} works only on a message you are reading. Open the received message, then ask again.`)
  }
}

function only(args, allowed) {
  const a = args || {}
  if (Object.keys(a).some(k => RECIPIENT_ARGS.includes(k))) {
    throw new ExecError('Validation', 'Hermes never sends mail and never adds, removes or changes recipients')
  }
  for (const k of Object.keys(a)) if (!allowed.includes(k)) throw new ExecError('Validation', `unexpected argument ${k}`)
  return a
}

function succeeded(r) {
  const S = (office().AsyncResultStatus || {}).Succeeded
  return !!r && (r.status === 'succeeded' || (S !== undefined && r.status === S))
}

/** Promise around one Office callback API: fn(cb) where cb receives an AsyncResult. */
function call(fn, what) {
  return new Promise((resolve, reject) => {
    try {
      fn(r => {
        if (succeeded(r)) return resolve(r.value)
        const e = (r && r.error) || {}
        reject(new ExecError('Office', `${what}: ${String(e.message || e.name || 'Outlook refused the request').slice(0, 200)}`, String(e.name || e.code || 'OfficeError')))
      })
    } catch (err) {
      reject(err instanceof ExecError ? err : new ExecError('Office', `${what}: ${String((err && err.message) || err).slice(0, 200)}`))
    }
  })
}

function coercion(kind) {
  const c = office().CoercionType || {}
  return kind === 'html' ? (c.Html || 'html') : (c.Text || 'text')
}

function clip(text, max) {
  const s = String(text ?? '')
  return s.length > max ? { text: s.slice(0, max), truncated: true, totalLength: s.length } : { text: s, truncated: false, totalLength: s.length }
}

function person(p) {
  if (!p) return null
  return { name: String(p.displayName || '').slice(0, 200), email: String(p.emailAddress || '').slice(0, 320) }
}

function people(list) {
  return (Array.isArray(list) ? list : []).slice(0, 100).map(person).filter(Boolean)
}

function attachmentsOf(list) {
  // Names, types and sizes only. Contents are never read (no getAttachmentContentAsync anywhere).
  return (Array.isArray(list) ? list : []).slice(0, 100).map(a => ({
    name: String(a.name || '').slice(0, 255), size: Number.isFinite(a.size) ? a.size : null,
    type: String(a.attachmentType || ''), inline: !!a.isInline
  }))
}

/** A stable, non-reversible key for the open item (plugin: index page id, reply-draft source check). */
export async function itemKey(it) {
  let id = it.itemId
  if (!id && typeof it.getItemIdAsync === 'function') {  // compose form of a saved item (Mailbox 1.8); none if unsaved
    try { id = await call(cb => it.getItemIdAsync(cb), 'Reading the item id') } catch { id = null }
  }
  const basis = id || [it.conversationId, typeof it.subject === 'string' ? it.subject : '',
    it.dateTimeCreated ? new Date(it.dateTimeCreated).toISOString() : ''].join('|')
  if (!basis || basis === '||') return null
  const subtle = globalThis.crypto && globalThis.crypto.subtle
  if (!subtle) return null
  const buf = await subtle.digest('SHA-256', new TextEncoder().encode('omnia-outlook:' + basis))
  return [...new Uint8Array(buf)].slice(0, 8).map(b => b.toString(16).padStart(2, '0')).join('')
}

function textToHtml(text) {
  return String(text).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;')
    .split(/\r?\n/).map(l => (l ? `<p>${l}</p>` : '<p><br></p>')).join('')
}

// The item an undo belongs to: an undo may only touch the same open item it was recorded on.
function sameItem(at, what) {
  if (mailbox().item !== at) throw new ExecError('NotFound', `${what}: the message Hermes edited is no longer the open item; nothing was changed`)
  return at
}

const tools = {
  async get_item(_ctx, args) {
    const { maxBodyLength } = only(args, ['maxBodyLength'])
    const max = maxBodyLength === undefined ? MAX_BODY : Math.min(MAX_BODY, Math.max(1, Math.floor(maxBodyLength)))
    const it = openItem()
    const mode = itemMode(it)
    const body = clip(await call(cb => it.body.getAsync(coercion('text'), cb), 'Reading the message body'), max)
    const base = { mode, itemType: String(it.itemType || 'message'), key: await itemKey(it), body }
    if (mode === 'read') {
      return {
        ...base, subject: String(it.subject || ''), from: person(it.from), to: people(it.to), cc: people(it.cc),
        date: it.dateTimeCreated ? new Date(it.dateTimeCreated).toISOString() : null, attachments: attachmentsOf(it.attachments)
      }
    }
    const subject = await call(cb => it.subject.getAsync(cb), 'Reading the subject')
    const to = it.to && typeof it.to.getAsync === 'function' ? await call(cb => it.to.getAsync(cb), 'Reading To') : []
    const cc = it.cc && typeof it.cc.getAsync === 'function' ? await call(cb => it.cc.getAsync(cb), 'Reading Cc') : []
    let attachments = []
    if (typeof it.getAttachmentsAsync === 'function') {
      try { attachments = attachmentsOf(await call(cb => it.getAttachmentsAsync(cb), 'Listing attachments')) } catch { attachments = [] }
    }
    return { ...base, subject: String(subject || ''), from: null, to: people(to), cc: people(cc), date: null, attachments }
  },

  async insert(_ctx, args) {
    const { target, text } = only(args, ['target', 'text'])
    const t = str(text, 'text', MAX_INSERT)
    const where = target || 'body'
    if (!['body', 'subject'].includes(where)) throw new ExecError('Validation', 'target must be body or subject')
    const it = openItem()
    needMode(it, 'compose', 'Inserting text')
    if (where === 'subject') {
      if (t.length > MAX_SUBJECT) throw new ExecError('Validation', `a subject holds at most ${MAX_SUBJECT} characters`)
      const prior = String(await call(cb => it.subject.getAsync(cb), 'Reading the subject') || '')
      await call(cb => it.subject.setAsync(t, cb), 'Setting the subject')
      return {
        data: { inserted: t.length, target: 'subject', replacedLength: prior.length },
        undo: async () => { const at = sameItem(it, 'Undo'); await call(cb => at.subject.setAsync(prior, cb), 'Restoring the subject') },
        undoLabel: 'set subject'
      }
    }
    const kind = typeof it.body.getTypeAsync === 'function'
      ? (String(await call(cb => it.body.getTypeAsync(cb), 'Reading the body type')).toLowerCase() === 'html' ? 'html' : 'text')
      : 'text'
    const prior = String(await call(cb => it.body.getAsync(coercion(kind), cb), 'Reading the body') || '')
    await call(cb => it.body.setSelectedDataAsync(t, { coercionType: coercion('text') }, cb), 'Inserting at the cursor')
    if (prior.length > SNAPSHOT_MAX_CHARS) {
      return { data: { inserted: t.length, target: 'body' }, undoUnavailable: 'the message body is too large to snapshot for undo' }
    }
    return {
      data: { inserted: t.length, target: 'body', location: 'cursor' },
      undo: async () => { const at = sameItem(it, 'Undo'); await call(cb => at.body.setAsync(prior, { coercionType: coercion(kind) }, cb), 'Restoring the body') },
      undoLabel: 'insert text in the body'
    }
  },

  async draft_reply(_ctx, args) {
    const { text } = only(args, ['text'])
    const t = str(text, 'text', MAX_BODY)
    const it = openItem()
    needMode(it, 'read', 'A reply draft')
    if (typeof it.displayReplyForm !== 'function') throw new ExecError('Unsupported', 'This Outlook cannot open a reply form from an add-in')
    // Reply (not reply-all): Outlook addresses it to the sender as usual; Hermes passes no recipients.
    it.displayReplyForm(textToHtml(t))
    return { data: { requested: 'reply', characters: t.length, sent: false }, undoUnavailable: 'the reply is a separate unsent draft: review it, then send or discard it yourself in Outlook' }
  },

  async draft_new(_ctx, args) {
    const { subject, text } = only(args, ['subject', 'text'])
    const t = str(text, 'text', MAX_BODY)
    const s = subject === undefined ? '' : str(subject, 'subject', MAX_SUBJECT)
    const m = mailbox()
    if (m.item && itemMode(m.item) === 'compose') throw new ExecError('Validation', 'A new draft opens from a message you are reading (or the mailbox), not from inside another draft')
    if (typeof m.displayNewMessageForm !== 'function') throw new ExecError('Unsupported', 'This Outlook cannot open a new message form from an add-in')
    // Subject and body only: no toRecipients, ccRecipients, bccRecipients or attachments are ever passed.
    m.displayNewMessageForm({ subject: s, htmlBody: textToHtml(t) })
    return { data: { requested: 'new', characters: t.length, sent: false }, undoUnavailable: 'the new message is a separate unsent draft: review it, then send or discard it yourself in Outlook' }
  }
}

// Calendar commands (calendar.js) share this pane, its pairing, its undo stack and its never-send rules.
Object.assign(tools, calendarTools({ mailbox, openItem, itemMode, only, call, coercion, clip, itemKey, sameItem, ExecError, str }))

export const MUTATING = new Set(['insert', 'draft_reply', 'draft_new', ...CALENDAR_MUTATING])

// No batch context in the Mailbox API: tools reach the open item themselves.
export const executor = createExecutor({ tools, mutating: MUTATING, run: fn => fn({}) })
export const { execute, SUPPORTED_TOOLS } = executor
