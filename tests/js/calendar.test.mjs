// SPDX-License-Identifier: Apache-2.0
// node --test tests/js/ : Outlook calendar (appointment) commands of the Outlook executor, against the in-memory mock.
// Hard rules under test: the open appointment only, no attendee change, no save, no invitation sent, undo restores.
import assert from 'node:assert/strict'
import fs from 'node:fs'
import path from 'node:path'
import { test } from 'node:test'
import { fileURLToPath } from 'node:url'
import { MockMailbox, installOutlookMock, owaSetEnd, owaSetStart } from '../../addin/dev/outlook-mock.js'
import { SUPPORTED_TOOLS, execute, executor } from '../../addin/outlook/executor.js'
import { recurrenceSummary } from '../../addin/outlook/calendar.js'

const HERE = path.dirname(fileURLToPath(import.meta.url))
const SRC = fs.readFileSync(path.join(HERE, '../../addin/outlook/calendar.js'), 'utf8')

function fresh(kind = 'apptCompose', opts = {}) {
  executor.clearUndo()
  const mb = installOutlookMock(new MockMailbox())
  if (kind === 'apptCompose') mb.openAppointmentCompose({ body: '<p>Agenda</p>', ...opts })
  else if (kind === 'apptRead') mb.openAppointmentRead(opts)
  else if (kind === 'mailRead') mb.openRead(opts)
  else mb.openCompose(opts)
  return mb
}
async function ok(tool, args = {}) { const r = await execute(tool, args); assert.equal(r.ok, true, `${tool}: ${JSON.stringify(r)}`); return r.data }
async function err(tool, args = {}) { const r = await execute(tool, args); assert.equal(r.ok, false, `${tool} should fail`); return r.error }

test('the pane advertises the three calendar commands next to the mail ones', () => {
  for (const t of ['get_event', 'update_event', 'draft_event']) assert.ok(SUPPORTED_TOOLS.includes(t), t)
})

test('calendar source never saves, sends, closes, changes attendees or reads other items', () => {
  for (const bad of [/\.saveAsync\s*\(/, /\.sendAsync\s*\(/, /\.close\(/, /(requiredAttendees|optionalAttendees|resources)\.(setAsync|addAsync|removeAsync)/,
    /requiredAttendees\s*:/, /optionalAttendees\s*:/, /resources\s*:/, /recurrence\.setAsync/, /makeEwsRequestAsync/,
    /getCallbackTokenAsync/, /\bfetch\(/, /XMLHttpRequest/]) {
    assert.doesNotMatch(SRC, bad, String(bad))
  }
})

test('get_event (attendee read form): subject, times with zone, location, organizer, attendee names and responses, body text', async () => {
  const mb = fresh('apptRead')
  const d = await ok('get_event')
  assert.equal(d.mode, 'read')
  assert.equal(d.subject, 'HOL live test: planning')
  assert.deepEqual(d.start, { utc: '2026-11-09T15:00:00.000Z', local: '2026-11-09T10:00:00-05:00', timeZone: 'Eastern Standard Time' })
  assert.equal(d.end.local, '2026-11-09T10:30:00-05:00')
  assert.equal(d.durationMinutes, 30)
  assert.equal(d.location, 'Room 1')
  assert.deepEqual(d.organizer, { name: 'Organizer Test' })
  assert.deepEqual(d.attendees, [{ name: 'Attendee A', response: 'accepted', required: true },
    { name: 'Attendee B', response: 'tentative', required: false }])
  assert.ok(!JSON.stringify(d).includes('@example.invalid'), 'names and responses only: no addresses')
  assert.equal(d.body.text, 'Agenda: beta launch.')
  assert.deepEqual(d.recurrence, { recurring: false, summary: 'not recurring' })
  assert.match(d.key, /^[0-9a-f]{16}$/)
  assert.deepEqual(mb.forbidden, [])
})

test('get_event (organizer compose form) reads through the async getters', async () => {
  fresh('apptCompose', { subject: 'HOL live test event', location: 'Studio', required: [] })
  const d = await ok('get_event', { maxBodyLength: 3 })
  assert.equal(d.mode, 'compose')
  assert.equal(d.subject, 'HOL live test event')
  assert.equal(d.location, 'Studio')
  assert.deepEqual(d.organizer, { name: 'Test Account' })
  assert.deepEqual(d.attendees, [])
  assert.deepEqual([d.body.text, d.body.truncated], ['Age', true])
  assert.equal(d.start.utc, '2026-11-09T15:00:00.000Z')
  // A saved event being edited has an id (getItemIdAsync): it can be indexed. A new, unsaved one has none.
  assert.match(d.key, /^[0-9a-f]{16}$/)
  const mb = fresh('apptCompose', { itemId: null })
  assert.equal((await ok('get_event')).key, null)
  assert.deepEqual(mb.forbidden, [])
})

test('get_event refuses a message and an empty pane; it never enumerates the calendar', async () => {
  fresh('mailRead')
  assert.equal((await err('get_event')).category, 'Validation')
  const mb = fresh('apptRead')
  mb.item = null
  assert.equal((await err('get_event')).category, 'NotFound')
})

test('recurrence summary in words', () => {
  assert.equal(recurrenceSummary(null), 'not recurring')
  assert.equal(recurrenceSummary({ recurrenceType: 'weekly', recurrenceProperties: { interval: 1, days: ['mon', 'wed'] },
    seriesTime: { getStartDate: () => ({ year: 2026, month: 11, day: 2 }), getEndDate: () => null } }),
  'weekly on Mon, Wed, from 2026-11-02, no end date')
  assert.equal(recurrenceSummary({ recurrenceType: 'daily', recurrenceProperties: { interval: 2 },
    seriesTime: { getStartDate: () => ({ year: 2026, month: 11, day: 2 }), getEndDate: () => ({ year: 2026, month: 12, day: 1 }) } }),
  'every 2 days, from 2026-11-02 until 2026-12-01')
  assert.equal(recurrenceSummary({ recurrenceType: 'monthly', recurrenceProperties: { interval: 1, dayOfMonth: 15 } }), 'monthly on day 15')
})

test('update_event (compose): subject, start/end, location, body at cursor in one step; one undo restores all prior values', async () => {
  const mb = fresh('apptCompose', { location: 'Old room' })
  const before = { ...mb.item._state, start: +mb.item._state.start, end: +mb.item._state.end }
  const d = await ok('update_event', { subject: 'HOL live test event (moved)', start: '2026-11-10T16:00:00Z',
    end: '2026-11-10T17:00:00Z', location: 'Studio B', text: ' HOL cal marker' })
  assert.deepEqual(d.changed, ['subject', 'start', 'end', 'location', 'body'])
  assert.equal(d.undo.available, true)
  const s = mb.item._state
  assert.deepEqual([s.subject, s.start.toISOString(), s.end.toISOString(), s.location], ['HOL live test event (moved)',
    '2026-11-10T16:00:00.000Z', '2026-11-10T17:00:00.000Z', 'Studio B'])
  assert.ok(s.body.endsWith(' HOL cal marker'))
  const u = await ok('undo_last')
  assert.equal(u.undone, 'update_event')
  assert.deepEqual([s.subject, +s.start, +s.end, s.location, s.body], [before.subject, before.start, before.end, before.location, before.body])
  assert.deepEqual(mb.forbidden, [])
})

test('update_event with only start keeps Outlook\'s own end shift; undo puts both back', async () => {
  const mb = fresh('apptCompose')
  const d = await ok('update_event', { start: '2026-11-09T18:00:00Z' })
  assert.deepEqual(d.changed, ['start'])
  assert.equal(mb.item._state.end.toISOString(), '2026-11-09T18:30:00.000Z')
  await ok('undo_last')
  assert.deepEqual([mb.item._state.start.toISOString(), mb.item._state.end.toISOString()], ['2026-11-09T15:00:00.000Z', '2026-11-09T15:30:00.000Z'])
})

test('update_event refuses read forms, messages, bad times and every attendee / save / send argument; nothing changes', async () => {
  fresh('apptRead')
  assert.equal((await err('update_event', { subject: 'x' })).category, 'Validation')
  fresh('mailCompose')
  assert.equal((await err('update_event', { subject: 'x' })).category, 'Validation')
  const mb = fresh('apptCompose')
  const snap = () => JSON.stringify(mb.item._state)
  const before = snap()
  for (const bad of [{}, { start: 'tomorrow' }, { start: '2026-11-09T10:00:00' }, { start: '2026-11-09T18:00:00Z', end: '2026-11-09T17:00:00Z' },
    { end: '2026-11-09T14:00:00Z' }, { subject: 'x'.repeat(256) }, { location: 'x'.repeat(256) }, { text: 'x'.repeat(5001) }]) {
    assert.equal((await err('update_event', bad)).category, 'Validation', JSON.stringify(bad))
  }
  for (const k of ['requiredAttendees', 'optionalAttendees', 'attendees', 'resources', 'organizer', 'to', 'save', 'send', 'sendInvite']) {
    const e = await err('update_event', { subject: 'x', [k]: ['someone@example.invalid'] })
    assert.equal(e.category, 'Validation', k)
    assert.match(e.message, /never/)
  }
  assert.equal(snap(), before)
  assert.equal(executor.undoDepth(), 0)
  assert.deepEqual(mb.forbidden, [])
})

test('a failed field set rolls back the fields already set and leaves no undo entry', async () => {
  const mb = fresh('apptCompose', { location: 'Old room' })
  mb.failNext = 'location.setAsync'
  const e = await err('update_event', { subject: 'New', location: 'New room' })
  assert.equal(e.category, 'Office')
  assert.deepEqual([mb.item._state.subject, mb.item._state.location], ['HOL live test event', 'Old room'])
  assert.equal(executor.undoDepth(), 0)
})

test('undo refuses when another appointment is open', async () => {
  const mb = fresh('apptCompose')
  await ok('update_event', { subject: 'Changed' })
  const first = mb.item
  mb.openAppointmentCompose({ subject: 'Another' })
  assert.equal((await err('undo_last')).category, 'NotFound')
  assert.equal(first._state.subject, 'Changed')
  assert.equal(mb.item._state.subject, 'Another')
})

test('draft_event opens a new appointment form with subject, times, location, body only (no attendees, never saved)', async () => {
  const mb = fresh('mailRead')
  const d = await ok('draft_event', { subject: 'Follow-up: beta', start: '2026-11-12T15:00:00Z', end: '2026-11-12T15:30:00Z',
    location: 'Studio', text: 'Agenda:\n- Beta status' })
  assert.deepEqual([d.requested, d.saved, d.sent, d.undo.available, 'opened' in d], ['appointment', false, false, false, false])
  const f = mb.forms[0]
  assert.deepEqual(Object.keys(f).sort(), ['body', 'end', 'kind', 'location', 'start', 'subject'])
  assert.deepEqual([f.subject, f.start.toISOString(), f.end.toISOString(), f.location, f.body],
    ['Follow-up: beta', '2026-11-12T15:00:00.000Z', '2026-11-12T15:30:00.000Z', 'Studio', 'Agenda:\n- Beta status'])
  assert.deepEqual(mb.forbidden, [])
  fresh('apptRead')
  await ok('draft_event', { subject: 'x', start: '2026-11-12T15:00:00Z', end: '2026-11-12T15:30:00Z' })
})

test('draft_event refuses compose forms, bad times and attendee arguments', async () => {
  fresh('apptCompose')
  assert.equal((await err('draft_event', { subject: 'x', start: '2026-11-12T15:00:00Z', end: '2026-11-12T15:30:00Z' })).category, 'Validation')
  const mb = fresh('mailRead')
  for (const bad of [{ subject: 'x', start: '2026-11-12T15:00:00Z' }, { subject: 'x', start: '2026-11-12T16:00:00Z', end: '2026-11-12T15:00:00Z' },
    { subject: 'x', start: '2026-11-12T15:00:00Z', end: '2026-11-12T15:30:00Z', requiredAttendees: ['a@example.invalid'] },
    { subject: 'x', start: '2026-11-12T15:00:00Z', end: '2026-11-12T15:30:00Z', resources: ['room@example.invalid'] }]) {
    assert.equal((await err('draft_event', bad)).category, 'Validation', JSON.stringify(bad))
  }
  assert.deepEqual([mb.forms, mb.forbidden], [[], []])
})

// ---- A25: start/end moves as Outlook on the web really applies them (silent drops across midnight) ------------------
const PROBE = JSON.parse(fs.readFileSync(path.join(HERE, 'fixtures/owa-appointment-times-a25.json'), 'utf8'))

test('the mock reproduces every start/end step seen live in Outlook on the web (A25 probe, 17 scenarios)', () => {
  assert.equal(PROBE.scenarios.length, 17)
  for (const sc of PROBE.scenarios) {
    const state = { start: new Date(sc.before.start), end: new Date(sc.before.end) }
    sc.steps.forEach((st, i) => {
      const applied = (st.set === 'start' ? owaSetStart : owaSetEnd)(state, new Date(st.to))
      const reported = st.set === 'start' ? true : applied   // a dropped start still reports success
      assert.deepEqual([reported, state.start.toISOString(), state.end.toISOString()], [st.ok, st.start, st.end],
        `${sc.probe} ${sc.name} step ${i + 1}: ${st.set}=${st.to}`)
    })
  }
})

const iso = d => d.toISOString()
async function updateThenUndo(from, to) {
  const mb = fresh('apptCompose', { start: from[0], end: from[1], location: 'Old room' })
  const before = { ...mb.item._state, start: iso(mb.item._state.start), end: iso(mb.item._state.end) }
  const d = await ok('update_event', { subject: 'HOL live test event (moved)', start: to[0], end: to[1], location: 'Studio', text: ' marker' })
  assert.deepEqual(d.changed, ['subject', 'start', 'end', 'location', 'body'])
  const s = mb.item._state
  assert.deepEqual([iso(s.start), iso(s.end)], to.map(t => new Date(t).toISOString()), 'the update put the event at the requested times')
  const u = await ok('undo_last')
  assert.equal(u.undone, 'update_event')
  assert.deepEqual([s.subject, iso(s.start), iso(s.end), s.location, s.body], [before.subject, before.start, before.end, before.location, before.body])
  return mb
}

test('A24 shape: undo moves the start back to the day before (Sat 23:30-Sun 00:00 local -> +60 -> back) and restores every field', async () => {
  // RED on c94cd8d: Outlook drops the start (its recomputed end would cross midnight), then "Restoring the end: InvalidEndTime"
  await updateThenUndo(['2026-11-08T04:30:00Z', '2026-11-08T05:00:00Z'], ['2026-11-08T05:30:00Z', '2026-11-08T06:00:00Z'])
})

test('update moves an event EARLIER so it ends at midnight (Sun 00:30-01:00 -> Sat 23:30-Sun 00:00), and undo moves it back', async () => {
  await updateThenUndo(['2026-11-08T05:30:00Z', '2026-11-08T06:00:00Z'], ['2026-11-08T04:30:00Z', '2026-11-08T05:00:00Z'])
})

test('update moves an event LATER so it ends at midnight (Sun 02:30 -> Sun 23:30-Mon 00:00), and undo moves it back', async () => {
  await updateThenUndo(['2026-11-08T07:30:00Z', '2026-11-08T08:00:00Z'], ['2026-11-09T04:30:00Z', '2026-11-09T05:00:00Z'])
})

test('update moves an event over midnight both ways and across a week; plain same-day moves still work', async () => {
  await updateThenUndo(['2026-11-08T07:30:00Z', '2026-11-08T08:00:00Z'], ['2026-11-15T05:30:00Z', '2026-11-15T06:00:00Z'])
  await updateThenUndo(['2026-11-09T05:30:00Z', '2026-11-09T06:00:00Z'], ['2026-11-07T17:00:00Z', '2026-11-07T17:30:00Z'])
  await updateThenUndo(['2026-11-08T04:00:00Z', '2026-11-08T07:00:00Z'], ['2026-11-08T03:00:00Z', '2026-11-08T05:30:00Z'])  // spans midnight
  await updateThenUndo(['2026-11-09T15:00:00Z', '2026-11-09T15:30:00Z'], ['2026-11-09T16:00:00Z', '2026-11-09T16:30:00Z'])
})

test('a start-only move across midnight keeps the duration (end follows), and undo restores both', async () => {
  const mb = fresh('apptCompose', { start: '2026-11-08T07:30:00Z', end: '2026-11-08T08:00:00Z' })
  const d = await ok('update_event', { start: '2026-11-08T04:30:00Z' })
  assert.deepEqual(d.changed, ['start'])
  assert.deepEqual([iso(mb.item._state.start), iso(mb.item._state.end)], ['2026-11-08T04:30:00.000Z', '2026-11-08T05:00:00.000Z'])
  await ok('undo_last')
  assert.deepEqual([iso(mb.item._state.start), iso(mb.item._state.end)], ['2026-11-08T07:30:00.000Z', '2026-11-08T08:00:00.000Z'])
})

test('times Outlook will not take are an error naming what changed, never a silent success', async () => {
  const mb = fresh('apptCompose', { location: 'Old room' })
  mb.item.start.setAsync = (v, opt, cb) => { (typeof opt === 'function' ? opt : cb)({ status: 'succeeded' }) }  // always dropped
  const e = await err('update_event', { subject: 'New', start: '2026-11-09T18:00:00Z', end: '2026-11-09T18:30:00Z' })
  assert.equal(e.category, 'Office')
  assert.match(e.message, /start/)
  assert.deepEqual([mb.item._state.subject, iso(mb.item._state.start)], ['HOL live test event', '2026-11-09T15:00:00.000Z'], 'rolled back')
  assert.equal(executor.undoDepth(), 0)
})

test('a partial undo says exactly which fields were restored and which were not (no silent partial)', async () => {
  let mb = fresh('apptCompose', { location: 'Old room' })
  await ok('update_event', { subject: 'Moved', location: 'Studio', start: '2026-11-09T18:00:00Z', end: '2026-11-09T18:30:00Z' })
  mb.failNext = 'location.setAsync'
  let e = await err('undo_last')
  assert.equal(e.category, 'Partial')
  assert.match(e.message, /restored: subject, start\/end/)
  assert.match(e.message, /NOT restored: location/)
  assert.deepEqual([mb.item._state.subject, iso(mb.item._state.start), mb.item._state.location], ['HOL live test event', '2026-11-09T15:00:00.000Z', 'Studio'])

  mb = fresh('apptCompose', { location: 'Old room' })
  await ok('update_event', { subject: 'Moved', location: 'Studio', start: '2026-11-09T18:00:00Z', end: '2026-11-09T18:30:00Z' })
  mb.item.start.setAsync = (v, opt, cb) => { (typeof opt === 'function' ? opt : cb)({ status: 'succeeded' }) }  // always dropped
  e = await err('undo_last')
  assert.equal(e.category, 'Partial')
  assert.match(e.message, /restored: subject, location/)
  assert.match(e.message, /NOT restored: start\/end/)
  assert.match(e.message, /2026-11-09T10:00:00-05:00/, 'names the start to set by hand, in the mailbox time zone')
  assert.deepEqual([mb.item._state.subject, mb.item._state.location], ['HOL live test event', 'Old room'])
})

test('the live calendar run\'s sequence: update to 23:30-00:00, move past midnight, undo, undo (each exact)', async () => {
  const mb = fresh('apptCompose', { start: '2026-11-08T08:30:00Z', end: '2026-11-08T09:00:00Z', location: '' })
  const at = () => [iso(mb.item._state.start), iso(mb.item._state.end)]
  await ok('update_event', { subject: 'moved', start: '2026-11-09T04:30:00Z', end: '2026-11-09T05:00:00Z', location: 'HOL Studio', text: ' m' })
  assert.deepEqual(at(), ['2026-11-09T04:30:00.000Z', '2026-11-09T05:00:00.000Z'])
  await ok('update_event', { start: '2026-11-09T05:30:00Z', end: '2026-11-09T06:00:00Z' })
  assert.deepEqual(at(), ['2026-11-09T05:30:00.000Z', '2026-11-09T06:00:00.000Z'])
  await ok('undo_last')
  assert.deepEqual([...at(), mb.item._state.subject], ['2026-11-09T04:30:00.000Z', '2026-11-09T05:00:00.000Z', 'moved'])
  await ok('undo_last')
  assert.deepEqual([...at(), mb.item._state.subject, mb.item._state.location], ['2026-11-08T08:30:00.000Z', '2026-11-08T09:00:00.000Z', 'HOL live test event', ''])
})
