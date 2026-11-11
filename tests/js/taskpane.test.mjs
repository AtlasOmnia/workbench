// SPDX-License-Identifier: Apache-2.0
// node --test: task-pane helpers that do not need a DOM.
import assert from 'node:assert/strict'
import { test } from 'node:test'
import { AUTO_SHOW_SETTING, ERROR_TEXT, activityText, buildLabel, describeError, enableAutoShow } from '../../addin/shared/taskpane.js'
import { BUILD } from '../../addin/shared/build-info.js'

function fakeOffice({ fail = false } = {}) {
  const saved = {}
  const settings = {
    get: k => saved[k] ?? null,
    set: (k, v) => { saved[k] = v },
    saveAsync: cb => cb({ status: fail ? 'failed' : 'succeeded' })
  }
  return { saved, Office: { AsyncResultStatus: { Succeeded: 'succeeded' }, context: { document: { settings } } } }
}

test('paired documents ask Office to reopen the Hermes pane with the document', async () => {
  const { saved, Office } = fakeOffice()
  assert.equal(AUTO_SHOW_SETTING, 'Office.AutoShowTaskpaneWithDocument')
  assert.equal(await enableAutoShow(Office), true)
  assert.equal(saved[AUTO_SHOW_SETTING], true)
  assert.equal(await enableAutoShow(Office), true) // already set: no second write needed
})

test('hosts without document settings (OneNote) or a failed save report false, never throw', async () => {
  assert.equal(await enableAutoShow({ context: { document: {} } }), false)
  assert.equal(await enableAutoShow(undefined), false)
  assert.equal(await enableAutoShow(fakeOffice({ fail: true }).Office), false)
})

test('activity lines are readable and every error says what to do next (critic2 #6)', () => {
  assert.equal(activityText({ last: 'write_range', ok: true }), 'Write range ✓')
  const line = activityText({ last: 'create_chart', ok: false, error: { category: 'ExecutionUncertain' } })
  assert.match(line, /^Create chart ✗ — Unclear whether that change applied\. Look at the document/)
  for (const [cat, text] of Object.entries(ERROR_TEXT)) {
    assert.ok(text.length > 30 && /[.)]$/.test(text), `${cat}: ${text}`)
    assert.doesNotMatch(text, /token|password|https?:/i)
  }
  assert.equal(describeError({ category: 'Office', code: 'InvalidArgument' }).endsWith('(InvalidArgument)'), true)
  assert.match(describeError({ category: 'Mystery' }), /Ask Hermes what happened/)
})

test('pane footer shows the loaded build honestly (critic2 #7)', () => {
  assert.deepEqual(BUILD, { version: null, build: null, commit: null, dirty: null }) // unbuilt checkout placeholder
  assert.equal(buildLabel(), 'Workbench add-in · unbuilt source (dev)')
  assert.equal(buildLabel({ version: '0.3.0', build: 'a1b2c3d4e5f6', commit: 'e404f5b0a1b2', dirty: false }),
    'Workbench add-in 0.3.0 · build a1b2c3d4e5f6 · source e404f5b0a1b2')
  assert.equal(buildLabel({ version: '0.3.0', build: 'a1b2c3d4e5f6', commit: 'e404f5b0a1b2', dirty: true }),
    'Workbench add-in 0.3.0 · build a1b2c3d4e5f6 · source e404f5b0a1b2 + uncommitted changes')
  assert.equal(buildLabel({ version: '0.3.0', build: 'a1b2c3d4e5f6', commit: null, dirty: null }), 'Workbench add-in 0.3.0 · build a1b2c3d4e5f6 · source unknown')
})
