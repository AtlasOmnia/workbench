// SPDX-License-Identifier: Apache-2.0
// In-memory stand-in for Outlook's Mailbox API (DEV/TEST ONLY; never deployed). Not Microsoft code; it models only
// what addin/outlook/executor.js uses: one open item in read or compose mode, async getters/setters with
// AsyncResult callbacks, the reply / new-message forms (recorded, never sent), and the ItemChanged event.
// Every recipient setter and send entry point exists here only to COUNT calls, so tests can prove none is made.

const S = { Succeeded: 'succeeded', Failed: 'failed' }
const ok = (cb, value) => setTimeout(() => cb({ status: S.Succeeded, value }), 0)
const fail = (cb, name, message) => setTimeout(() => cb({ status: S.Failed, error: { name, message } }), 0)

// ---- Outlook on the web's appointment start/end setters, as observed live (A25 probe, 2026-11-08) -----------------
const DAY = 86400000
const localDay = (ms, off) => Math.floor((ms + off * 60000) / DAY)
const timeOfDay = (ms, off) => ((ms + off * 60000) % DAY + DAY) % DAY
/** start.setAsync: the end moves by the start's change in whole LOCAL days (keeping its wall-clock time); if that
 * leaves end <= start, Outlook recomputes end = start + old duration, and when that end falls on another local date
 * (the event would span midnight, an end at 00:00 included) the call reports success and changes NOTHING. */
export function owaSetStart(state, v, off = -300) {
  const s = v.getTime()
  let end = state.end.getTime() + (localDay(s, off) - localDay(state.start.getTime(), off)) * DAY
  if (end <= s) {
    end = s + (state.end - state.start)
    if (localDay(end, off) !== localDay(s, off)) return false   // acknowledged, silently dropped
  }
  state.start = new Date(s); state.end = new Date(end)
  return true
}
/** end.setAsync: an end on a local date from the start's date to the current end's date is set exactly; on any other
 * date the end becomes that date at the START's time of day. An end before the start is refused. */
export function owaSetEnd(state, v, off = -300) {
  const e = v.getTime()
  const day = localDay(e, off)
  const inside = day >= localDay(state.start.getTime(), off) && day <= localDay(state.end.getTime(), off)
  const end = inside ? e : day * DAY - off * 60000 + timeOfDay(state.start.getTime(), off)
  if (end < state.start.getTime()) return false
  state.end = new Date(end)
  return true
}

export class MockMailbox {
  constructor() {
    this.forms = []           // displayReplyForm / displayNewMessageForm calls
    this.forbidden = []       // any send / recipient-changing call (must stay empty)
    this.handlers = []
    this.failNext = null      // name of the next API to fail, e.g. 'setSelectedDataAsync'
    this.item = null
    this.userProfile = { timeZone: 'Eastern Standard Time' }
    this.offsetMinutes = -300  // Eastern standard time, as convertToLocalClientTime below
  }

  _guard(name, cb) {
    if (this.failNext === name) { this.failNext = null; fail(cb, 'GenericResponseError', `${name} failed (mock)`); return false }
    return true
  }

  /** A received message open in the reading pane. */
  openRead({ subject = 'HOL live test: weekly sync', from = { displayName: 'Test Account', emailAddress: 'test@example.invalid' },
    to = [{ displayName: 'Test Account', emailAddress: 'test@example.invalid' }], cc = [], body = 'Notes.', attachments = [],
    itemId = 'AAMk-mock-1', created = '2026-11-06T12:00:00Z' } = {}) {
    const mb = this
    const it = {
      itemType: 'message', itemId, conversationId: 'conv-1', subject, from, to, cc, dateTimeCreated: new Date(created), attachments,
      body: { getAsync(kind, cb) { if (mb._guard('getAsync', cb)) ok(cb, kind === 'html' ? `<p>${body}</p>` : body) } },
      displayReplyForm(html) { mb.forms.push({ kind: 'reply', html, itemId }) },
      displayReplyAllForm(html) { mb.forbidden.push('displayReplyAllForm'); mb.forms.push({ kind: 'replyAll', html }) },
      getAttachmentContentAsync() { mb.forbidden.push('getAttachmentContentAsync') }
    }
    return this._open(it)
  }

  /** A message being written (new, reply or forward), with a cursor at the end of the body. */
  openCompose({ subject = '', body = '', bodyType = 'html', to = [], cc = [], bcc = [] } = {}) {
    const mb = this
    const state = { subject, body, bodyType, to: [...to], cc: [...cc], bcc: [...bcc], cursor: body.length }
    const recip = key => ({
      getAsync(cb) { ok(cb, state[key].map(r => ({ ...r }))) },
      setAsync() { mb.forbidden.push(`${key}.setAsync`) },
      addAsync() { mb.forbidden.push(`${key}.addAsync`) }
    })
    const it = {
      itemType: 'message', _state: state,
      subject: {
        getAsync(cb) { if (mb._guard('subject.getAsync', cb)) ok(cb, state.subject) },
        setAsync(v, opt, cb) { cb = typeof opt === 'function' ? opt : cb; if (!mb._guard('subject.setAsync', cb)) return; state.subject = String(v); ok(cb) }
      },
      to: recip('to'), cc: recip('cc'), bcc: recip('bcc'),
      body: {
        getTypeAsync(cb) { ok(cb, state.bodyType) },
        getAsync(kind, opt, cb) {
          cb = typeof opt === 'function' ? opt : cb
          if (!mb._guard('getAsync', cb)) return
          ok(cb, kind === 'html' || state.bodyType === 'text' ? state.body : state.body.replace(/<[^>]*>/g, ''))
        },
        setAsync(v, opt, cb) { cb = typeof opt === 'function' ? opt : cb; if (!mb._guard('setAsync', cb)) return; state.body = String(v); state.cursor = state.body.length; ok(cb) },
        setSelectedDataAsync(v, opt, cb) {
          cb = typeof opt === 'function' ? opt : cb
          if (!mb._guard('setSelectedDataAsync', cb)) return
          const t = String(v)
          state.body = state.body.slice(0, state.cursor) + t + state.body.slice(state.cursor)
          state.cursor += t.length
          ok(cb)
        }
      },
      getAttachmentsAsync(cb) { ok(cb, []) },
      sendAsync() { mb.forbidden.push('sendAsync') },
      saveAsync(cb) { ok(cb, 'draft-id') }
    }
    return this._open(it)
  }

  /** An appointment the user was invited to, open in the reading form (attendee view). */
  openAppointmentRead({ subject = 'HOL live test: planning', start = '2026-11-09T15:00:00Z', end = '2026-11-09T15:30:00Z',
    location = 'Room 1', organizer = { displayName: 'Organizer Test', emailAddress: 'org@example.invalid', appointmentResponse: 'organizer' },
    required = [{ displayName: 'Attendee A', emailAddress: 'a@example.invalid', appointmentResponse: 'accepted' }],
    optional = [{ displayName: 'Attendee B', emailAddress: 'b@example.invalid', appointmentResponse: 'tentative' }],
    body = 'Agenda: beta launch.', recurrence = null, seriesId = '', itemId = 'AAMk-appt-1' } = {}) {
    const mb = this
    const it = {
      itemType: 'appointment', itemId, seriesId, subject, start: new Date(start), end: new Date(end), location, organizer,
      requiredAttendees: required, optionalAttendees: optional, recurrence, dateTimeCreated: new Date('2026-11-01T00:00:00Z'),
      body: { getAsync(kind, cb) { if (mb._guard('getAsync', cb)) ok(cb, kind === 'html' ? `<p>${body}</p>` : body) } },
      displayReplyForm(html) { mb.forms.push({ kind: 'reply', html, itemId }) },
      displayReplyAllForm() { mb.forbidden.push('displayReplyAllForm') }
    }
    return this._open(it)
  }

  /** An appointment being created or edited by its organizer (compose form). Start and end setters behave as
   * Outlook on the web was seen to (owaSetStart / owaSetEnd above), so the add-in must verify what each set did. */
  openAppointmentCompose({ subject = 'HOL live test event', start = '2026-11-09T15:00:00Z', end = '2026-11-09T15:30:00Z',
    location = '', body = '', bodyType = 'html', required = [], optional = [], recurrence = null, itemId = 'AAMk-appt-edit-1' } = {}) {
    const mb = this
    const state = { subject, start: new Date(start), end: new Date(end), location, body, bodyType, cursor: body.length,
      required: [...required], optional: [...optional], recurrence }
    const attendees = key => ({
      getAsync(cb) { ok(cb, state[key].map(r => ({ ...r }))) },
      setAsync() { mb.forbidden.push(`${key}.setAsync`) },
      addAsync() { mb.forbidden.push(`${key}.addAsync`) },
      removeAsync() { mb.forbidden.push(`${key}.removeAsync`) }
    })
    const cbOf = (opt, cb) => (typeof opt === 'function' ? opt : cb)
    const it = {
      itemType: 'appointment', _state: state,
      subject: {
        getAsync(cb) { if (mb._guard('subject.getAsync', cb)) ok(cb, state.subject) },
        setAsync(v, opt, cb) { cb = cbOf(opt, cb); if (!mb._guard('subject.setAsync', cb)) return; state.subject = String(v); ok(cb) }
      },
      // Start and end follow what Outlook on the web was SEEN doing (A25 live probe, Eastern time): see owaSetStart /
      // owaSetEnd below. In short: a start move shifts the end by whole local days; a start move that would make
      // Outlook recompute an end on another local date is acknowledged and silently dropped; an end on another local
      // date than the current end's takes that date at the START's time of day.
      start: {
        getAsync(cb) { ok(cb, new Date(state.start)) },
        setAsync(v, opt, cb) {
          cb = cbOf(opt, cb); if (!mb._guard('start.setAsync', cb)) return
          if (!(v instanceof Date) || isNaN(v)) return fail(cb, 'InvalidArgument', 'start must be a Date')
          owaSetStart(state, v, mb.offsetMinutes); ok(cb)
        }
      },
      end: {
        getAsync(cb) { ok(cb, new Date(state.end)) },
        setAsync(v, opt, cb) {
          cb = cbOf(opt, cb); if (!mb._guard('end.setAsync', cb)) return
          if (!(v instanceof Date) || isNaN(v)) return fail(cb, 'InvalidArgument', 'end must be a Date')
          if (!owaSetEnd(state, v, mb.offsetMinutes)) return fail(cb, 'InvalidEndTime', 'The end time cannot be before the start time.')
          ok(cb)
        }
      },
      location: {
        getAsync(cb) { ok(cb, state.location) },
        setAsync(v, opt, cb) { cb = cbOf(opt, cb); if (!mb._guard('location.setAsync', cb)) return; state.location = String(v); ok(cb) }
      },
      organizer: { getAsync(cb) { ok(cb, { displayName: 'Test Account', emailAddress: 'test@example.invalid' }) } },
      requiredAttendees: attendees('required'), optionalAttendees: attendees('optional'),
      recurrence: { getAsync(cb) { ok(cb, state.recurrence) }, setAsync() { mb.forbidden.push('recurrence.setAsync') } },
      body: {
        getTypeAsync(cb) { ok(cb, state.bodyType) },
        getAsync(kind, opt, cb) {
          cb = cbOf(opt, cb)
          if (!mb._guard('getAsync', cb)) return
          ok(cb, kind === 'html' || state.bodyType === 'text' ? state.body : state.body.replace(/<[^>]*>/g, ''))
        },
        setAsync(v, opt, cb) { cb = cbOf(opt, cb); if (!mb._guard('setAsync', cb)) return; state.body = String(v); state.cursor = state.body.length; ok(cb) },
        setSelectedDataAsync(v, opt, cb) {
          cb = cbOf(opt, cb)
          if (!mb._guard('setSelectedDataAsync', cb)) return
          const t = String(v)
          state.body = state.body.slice(0, state.cursor) + t + state.body.slice(state.cursor)
          state.cursor += t.length
          ok(cb)
        }
      },
      // A saved event opened for editing has an id; a new one does not until it is saved (never by Hermes).
      getItemIdAsync(cb) { if (itemId) ok(cb, itemId); else fail(cb, 'ItemNotFound', 'not saved yet') },
      saveAsync() { mb.forbidden.push('saveAsync') },
      sendAsync() { mb.forbidden.push('sendAsync') },
      close() { mb.forbidden.push('close') }
    }
    return this._open(it)
  }

  displayNewAppointmentForm(params) {
    for (const k of Object.keys(params || {})) {
      if (!['subject', 'start', 'end', 'location', 'body'].includes(k)) this.forbidden.push(`displayNewAppointmentForm.${k}`)
    }
    this.forms.push({ kind: 'appointment', ...params })
  }

  /** The mailbox's own time zone (Eastern, standard time in the mock): wall-clock parts for a Date. */
  convertToLocalClientTime(date) {
    const d = new Date(date.getTime() - 300 * 60000)
    return { year: d.getUTCFullYear(), month: d.getUTCMonth(), date: d.getUTCDate(), hours: d.getUTCHours(),
      minutes: d.getUTCMinutes(), seconds: d.getUTCSeconds(), milliseconds: d.getUTCMilliseconds(), timezoneOffset: -300 }
  }

  _open(it) {
    this.item = it
    for (const h of this.handlers) h({ type: 'olkItemSelectedChanged' })
    return it
  }

  displayNewMessageForm(params) {
    for (const k of Object.keys(params || {})) if (!['subject', 'htmlBody'].includes(k)) this.forbidden.push(`displayNewMessageForm.${k}`)
    this.forms.push({ kind: 'new', ...params })
  }

  addHandlerAsync(type, handler, opt, cb) { this.handlers.push(handler); cb = typeof opt === 'function' ? opt : cb; if (cb) ok(cb) }
}

export function installOutlookMock(mailbox) {
  const mb = mailbox || new MockMailbox()
  if (!mb.item) mb.openRead({ body: 'HOL live test (mock): we decided to ship on Friday.' })
  globalThis.Office = {
    HostType: { Outlook: 'Outlook' },
    AsyncResultStatus: S,
    CoercionType: { Text: 'text', Html: 'html' },
    EventType: { ItemChanged: 'olkItemSelectedChanged' },
    onReady: cb => Promise.resolve({ host: 'Outlook', platform: 'OfficeOnline' }).then(info => { cb && cb(info); return info }),
    context: { mailbox: mb, diagnostics: { version: 'mock', platform: 'OfficeOnline' }, requirements: { isSetSupported: () => true } }
  }
  globalThis.__hermesMockMailbox = mb
  return mb
}
