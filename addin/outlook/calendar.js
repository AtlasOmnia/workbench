// SPDX-License-Identifier: Apache-2.0
// Outlook calendar commands for the Workbench Outlook pane: the ONE appointment open in Outlook on the web.
//
// Hard limits (the plugin enforces them again):
//   - only the open appointment (Office.context.mailbox.item) is read: the calendar is never listed or searched;
//   - attendees are never added, removed or changed, and no attendee / resource argument is accepted;
//   - nothing is saved and no invitation or update is sent: the person saves or sends in Outlook themselves;
//   - people are returned by display name and response only, never by address.
// Edits happen in an appointment the person is editing (compose form); one undo restores every prior value.
// A new appointment is REQUESTED through displayNewAppointmentForm (subject, start, end, location, body only).

// Arguments that would address people, save or send. Rejected by name on every calendar command.
export const ATTENDEE_ARGS = Object.freeze(['requiredAttendees', 'optionalAttendees', 'attendees', 'resources', 'organizer',
  'to', 'cc', 'bcc', 'recipients', 'save', 'saveAsync', 'send', 'sendNow', 'sendInvite', 'invite', 'notify'])
const NEVER = 'Hermes never adds, removes or changes attendees, never saves the appointment and never sends invitations'
const MAX_SUBJECT = 255
const MAX_LOCATION = 255
const MAX_TEXT = 5000
const MAX_BODY = 20000
const DAY = 86400000
// An instant: date, time and an explicit offset (Z or ±hh:mm), so a time can never be read in the wrong zone.
const INSTANT = /^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}(:\d{2}(\.\d{1,3})?)?(Z|[+-]\d{2}:\d{2})$/
const DAY_NAMES = { sun: 'Sun', mon: 'Mon', tue: 'Tue', wed: 'Wed', thu: 'Thu', fri: 'Fri', sat: 'Sat', weekday: 'weekdays', weekendDay: 'weekend days', day: 'every day' }

function pad(n, w = 2) { return String(n).padStart(w, '0') }

function seriesDate(d) {
  if (!d) return null
  if (d instanceof Date) return d.toISOString().slice(0, 10)
  if (typeof d === 'object' && d.year) return `${d.year}-${pad(d.month)}-${pad(d.day)}`
  return String(d).slice(0, 10)
}

/** A one-line, human summary of an Office.js Recurrence object (or null). */
export function recurrenceSummary(r) {
  if (!r || !r.recurrenceType) return 'not recurring'
  const p = r.recurrenceProperties || {}
  const n = Number.isInteger(p.interval) && p.interval > 1 ? p.interval : 1
  const unit = { daily: 'day', weekday: 'weekday', weekly: 'week', monthly: 'month', yearly: 'year' }[r.recurrenceType] || r.recurrenceType
  let s = n > 1 ? `every ${n} ${unit}s` : ({ daily: 'daily', weekday: 'every weekday', weekly: 'weekly', monthly: 'monthly', yearly: 'yearly' }[r.recurrenceType] || r.recurrenceType)
  if (Array.isArray(p.days) && p.days.length) s += ' on ' + p.days.map(d => DAY_NAMES[d] || d).join(', ')
  if (p.dayOfMonth) s += ` on day ${p.dayOfMonth}`
  if (p.weekNumber && p.dayOfWeek) s += ` on the ${p.weekNumber} ${DAY_NAMES[p.dayOfWeek] || p.dayOfWeek}`
  if (p.month && r.recurrenceType === 'yearly') s += ` in ${p.month}`
  const st = r.seriesTime
  if (st && typeof st.getStartDate === 'function') {
    const from = seriesDate(st.getStartDate())
    const to = typeof st.getEndDate === 'function' ? seriesDate(st.getEndDate()) : null
    if (from) s += `, from ${from}`
    s += to ? ` until ${to}` : ', no end date'
  }
  return s
}

export function calendarTools(h) {
  const { mailbox, openItem, itemMode, only, call, coercion, clip, itemKey, sameItem, ExecError, str } = h

  function guard(args, allowed) {
    const a = args || {}
    if (Object.keys(a).some(k => ATTENDEE_ARGS.includes(k))) throw new ExecError('Validation', NEVER)
    return only(a, allowed)
  }

  function appointment(what) {
    const it = openItem()
    if (String(it.itemType || '').toLowerCase() !== 'appointment') {
      throw new ExecError('Validation', `${what} works only on an appointment open in Outlook's calendar. Open the event, then ask again.`)
    }
    return it
  }

  function instant(v, name) {
    if (typeof v !== 'string' || !INSTANT.test(v)) {
      throw new ExecError('Validation', `${name} must be a date and time with an offset, such as 2026-11-12T10:00:00-05:00`)
    }
    const d = new Date(v)
    if (isNaN(d)) throw new ExecError('Validation', `${name} is not a valid date and time`)
    return d
  }

  /** The time in UTC, as wall-clock time in the mailbox's own time zone (when Outlook can convert it), and that zone. */
  function when(d) {
    if (!(d instanceof Date) || isNaN(d)) return null
    const out = { utc: d.toISOString(), local: null, timeZone: null }
    const mb = mailbox()
    try { out.timeZone = (mb.userProfile && mb.userProfile.timeZone) ? String(mb.userProfile.timeZone) : null } catch { /* older Outlook */ }
    try {
      const t = typeof mb.convertToLocalClientTime === 'function' ? mb.convertToLocalClientTime(d) : null
      if (t && Number.isFinite(t.year)) {
        const off = Number(t.timezoneOffset) || 0
        const sign = off < 0 ? '-' : '+'
        out.local = `${t.year}-${pad(t.month + 1)}-${pad(t.date)}T${pad(t.hours)}:${pad(t.minutes)}:${pad(t.seconds || 0)}`
          + `${sign}${pad(Math.floor(Math.abs(off) / 60))}:${pad(Math.abs(off) % 60)}`
      }
    } catch { /* leave local null: utc is still exact */ }
    return out
  }

  const name = p => String((p && p.displayName) || '').slice(0, 200)
  const people = (list, required) => (Array.isArray(list) ? list : []).slice(0, 200)
    .map(p => ({ name: name(p), response: String((p && p.appointmentResponse) || 'none'), required }))

  async function get(it, prop, what) {
    const v = it[prop]
    if (v && typeof v === 'object' && typeof v.getAsync === 'function') return call(cb => v.getAsync(cb), what)
    return v
  }

  async function snapshot(it, maxBody) {
    const mode = itemMode(it)
    const body = clip(await call(cb => it.body.getAsync(coercion('text'), cb), 'Reading the event body'), maxBody)
    const [subject, start, end, location] = mode === 'read' ? [it.subject, it.start, it.end, it.location]
      : [await get(it, 'subject', 'Reading the subject'), await get(it, 'start', 'Reading the start'),
        await get(it, 'end', 'Reading the end'), await get(it, 'location', 'Reading the location')]
    let organizer = null
    try { organizer = await get(it, 'organizer', 'Reading the organizer') } catch { organizer = null }
    let req = [], opt = []
    try { req = await get(it, 'requiredAttendees', 'Reading attendees') || [] } catch { req = [] }
    try { opt = await get(it, 'optionalAttendees', 'Reading attendees') || [] } catch { opt = [] }
    let rec = null
    try { rec = await get(it, 'recurrence', 'Reading the recurrence') } catch { rec = null }
    const s = start instanceof Date ? start : (start ? new Date(start) : null)
    const e = end instanceof Date ? end : (end ? new Date(end) : null)
    return {
      mode, itemType: 'appointment', key: await itemKey(it), subject: String(subject || ''), start: when(s), end: when(e),
      durationMinutes: s && e ? Math.round((e - s) / 60000) : null, location: String(location || '').slice(0, 500),
      organizer: organizer ? { name: name(organizer) } : null, attendees: [...people(req, true), ...people(opt, false)],
      recurrence: { recurring: !!(rec && rec.recurrenceType) || !!it.seriesId, summary: rec && rec.recurrenceType ? recurrenceSummary(rec)
        : (it.seriesId ? 'an occurrence of a recurring series' : 'not recurring') },
      body
    }
  }

  /**
   * Put the open appointment at exactly want.start-want.end, checking what Outlook did after every step. Outlook on
   * the web (seen live, A25) moves the end by whole days when the start moves, and when a start move would make it
   * recompute an end on another date (an event ending at or after midnight) it reports success and keeps the old
   * start. So: set the start, read it back, and if it was not taken give the end room first (a later date), then the
   * start again, then the exact end. Throws (code TimesNotApplied) if the form still does not show the times.
   */
  async function moveTo(at, want, verb) {
    const read = async () => ({
      start: +new Date(await call(cb => at.start.getAsync(cb), `${verb} the start`)),
      end: +new Date(await call(cb => at.end.getAsync(cb), `${verb} the end`))
    })
    const ws = +want.start, we = +want.end
    let cur = await read()
    for (let round = 0; round < 4 && !(cur.start === ws && cur.end === we); round++) {
      if (cur.start !== ws) {
        await call(cb => at.start.setAsync(new Date(ws), cb), `${verb} the start`)
        const next = await read()
        if (next.start !== ws) {
          const room = Math.max(next.end, ws, we) + (Math.ceil(Math.abs(ws - next.start) / DAY) + 1) * DAY
          await call(cb => at.end.setAsync(new Date(room), cb), `${verb} the end`)
          await call(cb => at.start.setAsync(new Date(ws), cb), `${verb} the start`)
        }
      } else {
        await call(cb => at.end.setAsync(new Date(we), cb), `${verb} the end`)
      }
      cur = await read()
    }
    if (cur.start !== ws || cur.end !== we) {
      const which = cur.start !== ws ? 'the start' : 'the end'
      throw new ExecError('Office', `${verb} the times: Outlook did not take ${which}; the form shows ${span(cur.start, cur.end)} instead of `
        + span(ws, we), 'TimesNotApplied')
    }
  }

  const stamp = ms => { const w = when(new Date(ms)); return (w && (w.local || w.utc)) || String(ms) }
  const span = (s, e) => `${stamp(s)} – ${stamp(e)}`

  return {
    async get_event(_ctx, args) {
      const { maxBodyLength } = guard(args, ['maxBodyLength'])
      const max = maxBodyLength === undefined ? MAX_BODY : Math.min(MAX_BODY, Math.max(1, Math.floor(maxBodyLength)))
      return snapshot(appointment('Reading an event'), max)
    },

    async update_event(_ctx, args) {
      const a = guard(args, ['subject', 'start', 'end', 'location', 'text'])
      const fields = ['subject', 'start', 'end', 'location', 'text'].filter(k => a[k] !== undefined)
      if (!fields.length) throw new ExecError('Validation', 'give at least one of subject, start, end, location or text')
      const subject = a.subject === undefined ? undefined : str(a.subject, 'subject', MAX_SUBJECT)
      const location = a.location === undefined ? undefined : (typeof a.location === 'string' && a.location.length <= MAX_LOCATION
        ? a.location : (() => { throw new ExecError('Validation', `location holds at most ${MAX_LOCATION} characters`) })())
      const text = a.text === undefined ? undefined : str(a.text, 'text', MAX_TEXT)
      const start = a.start === undefined ? undefined : instant(a.start, 'start')
      const end = a.end === undefined ? undefined : instant(a.end, 'end')
      const it = appointment('Changing an event')
      if (itemMode(it) !== 'compose') {
        throw new ExecError('Validation', 'Changing an event works only in an appointment you are editing (its edit form). Open the event for editing, then ask again.')
      }
      const prior = {
        subject: String(await call(cb => it.subject.getAsync(cb), 'Reading the subject') || ''),
        start: await call(cb => it.start.getAsync(cb), 'Reading the start'),
        end: await call(cb => it.end.getAsync(cb), 'Reading the end'),
        location: String(await call(cb => it.location.getAsync(cb), 'Reading the location') || '')
      }
      // The resulting times: the duration is kept when only the start moves.
      const newStart = start || prior.start
      const newEnd = end || (start ? new Date(start.getTime() + (prior.end - prior.start)) : prior.end)
      if (!(newEnd > newStart)) throw new ExecError('Validation', 'the event must end after it starts')
      let kind = 'text', body = null
      if (text !== undefined) {
        kind = typeof it.body.getTypeAsync === 'function'
          ? (String(await call(cb => it.body.getTypeAsync(cb), 'Reading the body type')).toLowerCase() === 'html' ? 'html' : 'text') : 'text'
        body = String(await call(cb => it.body.getAsync(coercion(kind), cb), 'Reading the body') || '')
      }
      // Every touched field is put back independently; the result names what was and was not restored.
      const restore = async (at, touched) => {
        const steps = [
          ['subject', () => call(cb => at.subject.setAsync(prior.subject, cb), 'Restoring the subject')],
          ['start/end', () => moveTo(at, { start: prior.start, end: prior.end }, 'Restoring')],
          ['location', () => call(cb => at.location.setAsync(prior.location, cb), 'Restoring the location')],
          ['body', () => call(cb => at.body.setAsync(body, { coercionType: coercion(kind) }, cb), 'Restoring the body')]
        ]
        const restored = [], missed = []
        for (const [name, run] of steps) {
          if (!touched.includes(name) || (name === 'body' && body === null)) continue
          try { await run(); restored.push(name) } catch (err) { missed.push({ name, why: String((err && err.message) || err).slice(0, 200) }) }
        }
        return { restored, missed }
      }
      const partial = (r, what) => {
        const hand = r.missed.some(m => m.name === 'start/end') ? ` Set the times by hand to ${span(+prior.start, +prior.end)}.` : ''
        return `${what}: restored: ${r.restored.join(', ') || 'nothing'}; NOT restored: ${r.missed.map(m => m.name).join(', ')} `
          + `(${r.missed.map(m => m.why).join('; ')}).${hand} The form is not saved.`
      }
      const done = [], touched = []
      try {
        if (subject !== undefined) { touched.push('subject'); await call(cb => it.subject.setAsync(subject, cb), 'Setting the subject'); done.push('subject') }
        if (start || end) {
          touched.push('start/end')
          await moveTo(it, { start: newStart, end: newEnd }, 'Setting')
          if (start) done.push('start')
          if (end) done.push('end')
        }
        if (location !== undefined) { touched.push('location'); await call(cb => it.location.setAsync(location, cb), 'Setting the location'); done.push('location') }
        if (text !== undefined) { touched.push('body'); await call(cb => it.body.setSelectedDataAsync(text, { coercionType: coercion('text') }, cb), 'Inserting at the cursor'); done.push('body') }
      } catch (err) {
        // Put back what this step touched, so a refused step leaves the form as it was; say so if it could not.
        const r = await restore(it, touched)
        if (r.missed.length) throw new ExecError('Partial', partial(r, `${(err && err.message) || err}. Putting the event back was partial`))
        throw err
      }
      return {
        data: { changed: done, saved: false, sent: false, attendeesChanged: false },
        undo: async () => {
          const at = sameItem(it, 'Undo')
          const r = await restore(at, touched)
          if (r.missed.length) throw new ExecError('Partial', partial(r, 'partial'))
        },
        undoLabel: 'change the event'
      }
    },

    async draft_event(_ctx, args) {
      const a = guard(args, ['subject', 'start', 'end', 'location', 'text'])
      const subject = str(a.subject, 'subject', MAX_SUBJECT)
      if (a.start === undefined || a.end === undefined) throw new ExecError('Validation', 'a new event needs start and end')
      const start = instant(a.start, 'start')
      const end = instant(a.end, 'end')
      if (!(end > start)) throw new ExecError('Validation', 'the event must end after it starts')
      const location = a.location === undefined ? '' : (typeof a.location === 'string' && a.location.length <= MAX_LOCATION
        ? a.location : (() => { throw new ExecError('Validation', `location holds at most ${MAX_LOCATION} characters`) })())
      const text = a.text === undefined ? '' : str(a.text, 'text', MAX_BODY)
      const m = mailbox()
      if (m.item && itemMode(m.item) === 'compose') {
        throw new ExecError('Validation', 'A new event opens from a message or event you are reading, not from inside a form you are editing')
      }
      if (typeof m.displayNewAppointmentForm !== 'function') throw new ExecError('Unsupported', 'This Outlook cannot open a new event form from an add-in')
      // Subject, times, location and body only: no requiredAttendees, optionalAttendees or resources are ever passed.
      m.displayNewAppointmentForm({ subject, start, end, location, body: text })
      return {
        data: { requested: 'appointment', subject, start: when(start), end: when(end), location, characters: text.length, saved: false, sent: false },
        undoUnavailable: 'the new event is a separate unsaved form: review it, then save or discard it yourself in Outlook'
      }
    }
  }
}

export const CALENDAR_MUTATING = ['update_event', 'draft_event']
