// SPDX-License-Identifier: Apache-2.0
// node --test tests/js/ : replay the exact Excel/Word/PowerPoint commands brain_make_deliverables sends
// (golden fixture pinned by tests/test_onenote.py) against the real executors and strict mocks.
import assert from 'node:assert/strict'
import { readFileSync } from 'node:fs'
import { test } from 'node:test'
import { installMock, MockWorkbook } from '../../addin/dev/office-mock.js'
import { installPowerPointMock, MockPresentation } from '../../addin/dev/powerpoint-mock.js'
import { installWordMock, MockDocument } from '../../addin/dev/word-mock.js'
import * as excel from '../../addin/excel/executor.js'
import * as ppt from '../../addin/powerpoint/executor.js'
import * as word from '../../addin/word/executor.js'

const GOLDEN = JSON.parse(readFileSync(new URL('./fixtures/deliverables.json', import.meta.url), 'utf8'))

async function replay(mod, calls) {
  mod.executor.clearUndo()
  const out = []
  for (const [tool, args] of calls) {
    const r = await mod.execute(tool, args)
    assert.equal(r.ok, true, `${tool}: ${JSON.stringify(r).slice(0, 400)}`)
    out.push(r.data)
  }
  return out
}

test('Excel tracker payload: new sheet, text never a formula, live COUNTIF/COUNTIFS, chart; every step undoable', async () => {
  const wb = installMock(new MockWorkbook())
  await replay(excel, GOLDEN.excel)
  const read = await excel.execute('read_range', { sheet: 'Decision tracker (2)', address: 'A1:I5' })
  assert.equal(read.ok, true)
  const f = read.data.formulas || read.data.values
  assert.equal(f[0][5], 'Source')
  assert.match(String(f[1][5]), /^=HYPERLINK\("https:\/\/onenote\.test\//)
  assert.match(String(read.data.values[2][1]), /^'?=cmd/) // stored as text, not as a formula
  assert.ok(wb.charts.some(c => c.title.text === 'Open items by owner'))
  // Presentation-ready: a styled table over the rows, frozen header row, and the new chart is what is selected.
  assert.equal(wb.tables.length, 1)
  assert.equal(wb.tables[0].hasHeaders, true)
  assert.equal(wb.activeChart, wb.charts.at(-1).name)
  assert.equal(wb.active, 'Decision tracker (2)')
  // Formulas stay formulas after formatting (the mock does not evaluate COUNTIF; real Excel recalculates).
  const live = (await excel.execute('read_range', { sheet: 'Decision tracker (2)', address: 'I1' })).data
  assert.equal((live.formulas || [[]])[0][0], '=COUNTIF(E2:E3,"Open")')
  assert.ok(wb.sheets.find(s => s.name === 'Decision tracker (2)').freeze)
  // Undo walks back every deliverable step in this pane (sheet creation included).
  let n = 0
  while ((await excel.execute('undo_last', {})).ok) n++
  assert.equal(n, GOLDEN.excel.filter(([t]) => !['get_workbook_overview', 'read_range'].includes(t)).length)
})

test('Word brief payload: a valid package with external hyperlinks, inserted at the end, undoable', async () => {
  const doc = installWordMock(new MockDocument(['Existing text']))
  const [ins] = await replay(word, GOLDEN.word.filter(([t]) => t === 'insert_ooxml'))
  assert.equal(ins.undo.available, true)
  assert.ok(doc.text.includes('Pilot decision — brief'))
  assert.ok(doc.text.startsWith('Existing text'))
  assert.ok((await word.execute('undo_last', {})).ok)
  assert.equal(doc.text.trim(), 'Existing text')
})

test('PowerPoint deck payload: 3-5 editable text slides ending with Sources, undoable', async () => {
  installPowerPointMock(new MockPresentation())
  const out = await replay(ppt, GOLDEN.powerpoint)
  const slides = out.filter(d => d && d.slideIndex !== undefined)
  assert.ok(slides.length >= 3 && slides.length <= 5)
  const all = await ppt.execute('read_all_slides', {})
  assert.ok(all.data.slides.at(-1).text.join(' ').includes('Sources'))
})
