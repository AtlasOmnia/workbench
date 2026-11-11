// SPDX-License-Identifier: Apache-2.0
// node --test tests/js/ : Word executor against the strict Office.js mock.
import assert from 'node:assert/strict'
import { test } from 'node:test'
import { installWordMock, MockDocument } from '../../addin/dev/word-mock.js'
import { execute, executor, SUPPORTED_TOOLS, MUTATING } from '../../addin/word/executor.js'
// office_insert_image pane commands (plugin-level tool; undo cases live in images.test.mjs)
const IMAGE_COMMANDS = ['insert_image', 'image_status']

const CANONICAL = ['get_document_overview', 'read_selection', 'read_document_text', 'apply_edit_plan', 'insert_text', 'replace_selection',
  'find_replace', 'format_selection', 'get_document_structure', 'apply_style', 'insert_ooxml', 'insert_comment']
const PARAS = ['Quarterly report', 'Revenue grew in the East region.', 'Costs fell in the West region.']

function fresh() {
  executor.clearUndo()
  const doc = installWordMock(new MockDocument(PARAS))
  doc.paragraphs[0].style = 'Heading 1'
  doc.setSelection(1, 0, 7) // "Revenue"
  return doc
}
const fp = doc => JSON.stringify({ p: doc.paragraphs, c: doc.comments, m: doc.marks })
async function ok(tool, args = {}) { const r = await execute(tool, args); assert.equal(r.ok, true, `${tool}: ${JSON.stringify(r)}`); return r.data }

test('pane advertises all 12 Word tools plus undo_last', () => {
  assert.deepEqual([...SUPPORTED_TOOLS].sort(), [...CANONICAL, ...IMAGE_COMMANDS, 'undo_last'].sort())
})

test('read tools report overview, selection, text and structure', async () => {
  fresh()
  const o = await ok('get_document_overview')
  assert.equal(o.paragraphs, 3)
  assert.equal(o.selection.preview, 'Revenue')
  assert.equal((await ok('read_selection')).text, 'Revenue')
  const t = await ok('read_document_text', { maxLength: 10 })
  assert.deepEqual([t.text, t.truncated], ['Quarterly ', true])
  const s = await ok('get_document_structure', { maxParagraphs: 2 })
  assert.equal(s.paragraphCount, 3)
  assert.equal(s.paragraphs.length, 2)
  assert.deepEqual(s.headings[0], { index: 0, style: 'Heading 1', text: 'Quarterly report' })
})

const CASES = [
  ['apply_edit_plan', { edits: [{ find: 'grew', replace: 'rose sharply', highlightColor: 'Yellow' }, { find: 'East', replace: 'Eastern' }, { find: 'fell', replace: 'dropped' }] }],
  ['insert_text', { text: 'Prepared by Hermes. ', location: 'start' }],
  ['insert_text', { text: ' End.', location: 'end' }],
  ['insert_text', { text: ' (draft)', location: 'selection' }],
  ['replace_selection', { text: 'Income', highlightColor: 'Yellow' }],
  ['find_replace', { find: 'region', replace: 'territory', matchCase: true }],
  ['format_selection', { bold: true, color: '#C00000', fontSize: 14 }],
  ['apply_style', { style: 'Heading 2' }],
  ['apply_style', { style: 'My Custom Style' }],
  ['insert_ooxml', { ooxml: '<w:p><w:r><w:t>From OOXML</w:t></w:r></w:p>', location: 'end' }],
  ['insert_comment', { text: 'Source?' }]
]
for (const [tool, args] of CASES) {
  test(`${tool} applies and undo restores the document (${JSON.stringify(args).slice(0, 50)})`, async () => {
    const doc = fresh()
    const before = fp(doc)
    const d = await ok(tool, args)
    assert.equal(d.undo.available, true, JSON.stringify(d.undo))
    assert.notEqual(fp(doc), before)
    assert.equal((await ok('undo_last')).undone, tool)
    assert.equal(fp(doc), before, `${tool} undo mismatch`)
  })
}

test('every mutating Word tool has an undo case', () => {
  const covered = new Set(CASES.map(c => c[0]))
  for (const t of MUTATING) assert.ok(covered.has(t) || IMAGE_COMMANDS.includes(t), t)
})

test('edit plan is all-or-nothing: ambiguous or missing finds change nothing', async () => {
  const doc = fresh()
  const before = fp(doc)
  const d = await ok('apply_edit_plan', { edits: [{ find: 'region', replace: 'x' }, { find: 'Costs', replace: 'Spend' }, { find: 'absent', replace: 'y' }] })
  assert.equal(d.applied, 0)
  assert.deepEqual(d.results.map(r => r.status), ['rejected', 'not_applied', 'rejected'])
  assert.match(d.results[0].reason, /not unique/)
  assert.equal(fp(doc), before)
  assert.equal(d.undo.available, false)
})

test('overlapping edits are rejected', async () => {
  fresh()
  const d = await ok('apply_edit_plan', { edits: [{ find: 'Revenue grew', replace: 'a' }, { find: 'grew in', replace: 'b' }] })
  assert.equal(d.applied, 0)
  assert.match(d.results[0].reason, /overlaps/)
})

test('selection tools refuse an empty selection and long searches are bounded', async () => {
  const doc = fresh()
  doc.setSelection(1, 3, 3)
  let r = await execute('format_selection', { bold: true })
  assert.equal(r.error.category, 'Validation')
  r = await execute('insert_comment', { text: 'x' })
  assert.equal(r.error.category, 'Validation')
  r = await execute('find_replace', { find: 'x'.repeat(300), replace: 'y' })
  assert.equal(r.error.category, 'Validation')
})

test('undo trims the extra trailing paragraph Word adds when the body is replaced', async () => {
  const doc = fresh()
  await ok('insert_text', { text: 'X', location: 'end' })
  await ok('undo_last')
  assert.equal(doc.paragraphs.length, 3)
})
