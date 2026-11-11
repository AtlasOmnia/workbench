// SPDX-License-Identifier: Apache-2.0
// node --test tests/js/ : change highlight after writes (E1-4). Writes select what changed; reads never
// touch the selection; a failed select never fails the write and comes back as a note.
import assert from 'node:assert/strict'
import { test } from 'node:test'
import { installMock, MockWorkbook } from '../../addin/dev/office-mock.js'
import { installPowerPointMock, MockPresentation } from '../../addin/dev/powerpoint-mock.js'
import { installWordMock, MockDocument } from '../../addin/dev/word-mock.js'
import * as excel from '../../addin/excel/executor.js'
import * as ppt from '../../addin/powerpoint/executor.js'
import * as word from '../../addin/word/executor.js'
import { createExecutor } from '../../addin/shared/runtime.js'

async function ok(mod, tool, args = {}) { const r = await mod.execute(tool, args); assert.equal(r.ok, true, `${tool}: ${JSON.stringify(r)}`); return r.data }

test('runtime: a failing highlight keeps the write ok and reports a note; reads never highlight', async () => {
  let selects = 0
  const run = fn => fn({ sync: async () => {} })
  const ex = createExecutor({
    run,
    mutating: new Set(['w']),
    tools: {
      async w() { return { data: { wrote: 1 }, undo: async () => {}, highlight: () => { throw Object.assign(new Error('no selection here'), { code: 'InvalidSelection' }) } } },
      async r() { return { value: 1, highlight: () => { selects++ } } }
    }
  })
  const w = await ex.execute('w', {})
  assert.equal(w.ok, true)
  assert.equal(w.data.wrote, 1)
  assert.equal(w.data.undo.available, true)
  assert.equal(w.data.highlight.selected, false)
  assert.match(w.data.highlight.note, /^The change applied, but it could not be highlighted \(InvalidSelection: no selection here\)/)
  const r = await ex.execute('r', {})
  assert.equal(r.ok, true)
  assert.equal(selects, 0) // a read's result is passed through untouched; nothing is selected
  const u = await ex.execute('undo_last', {})
  assert.equal(u.ok, true)
  assert.equal(u.data.highlight, undefined)
})

test('Excel: write_range / format_range select the written range on its sheet; read_range does not move the selection', async () => {
  excel.executor.clearUndo()
  const wb = installMock(new MockWorkbook())
  await ok(excel, 'create_sheet', { name: 'Tracker' })
  assert.equal(wb.active, 'Sheet1')
  const w = await ok(excel, 'write_range', { sheet: 'Tracker', address: 'A1:B2', values: [['a', 'b'], ['c', 'd']] })
  assert.deepEqual(w.highlight, { selected: true, target: 'Tracker!A1:B2' })
  assert.equal(wb.active, 'Tracker')
  assert.deepEqual(wb.selection, { sheet: 'Tracker', address: 'A1:B2' })
  const f = await ok(excel, 'format_range', { sheet: 'Tracker', address: 'A1:B1', format: { bold: true } })
  assert.equal(f.highlight.selected, true)
  assert.deepEqual(wb.selection, { sheet: 'Tracker', address: 'A1:B1' })
  const r = await ok(excel, 'read_range', { sheet: 'Sheet1', address: 'A1:A2' })
  assert.equal(r.highlight, undefined)
  assert.deepEqual(wb.selection, { sheet: 'Tracker', address: 'A1:B1' })
  assert.equal(wb.active, 'Tracker')
})

test('Excel: create_chart activates the new chart; without ExcelApi 1.9 it selects the source range instead', async () => {
  excel.executor.clearUndo()
  const wb = installMock(new MockWorkbook())
  await ok(excel, 'write_range', { sheet: 'Sheet1', address: 'A1:B3', values: [['Owner', 'Open'], ['Maya', 1], ['Unknown', 2]] })
  const c = await ok(excel, 'create_chart', { sheet: 'Sheet1', dataRange: 'A1:B3', chartType: 'ColumnClustered', title: 'Open' })
  assert.equal(c.highlight.selected, true)
  assert.equal(wb.activeChart, c.chart)
  globalThis.Office.context.requirements.isSetSupported = (set, v) => !(set === 'ExcelApi' && v === '1.9')
  const c2 = await ok(excel, 'create_chart', { sheet: 'Sheet1', dataRange: 'A1:B3', chartType: 'Line', title: 'Open 2' })
  assert.deepEqual(c2.highlight, { selected: true, target: 'Sheet1!A1:B3' })
  assert.equal(wb.activeChart, c.chart)
})

test('Word: insert_text and insert_ooxml select the inserted range; reads leave the selection alone', async () => {
  word.executor.clearUndo()
  const doc = installWordMock(new MockDocument(['Existing text']))
  const t = await ok(word, 'insert_text', { text: ' and more', location: 'end' })
  assert.equal(t.highlight.selected, true)
  const p = doc.paragraphs[doc.selection.p].text
  assert.equal(p.slice(doc.selection.start, doc.selection.end), ' and more')
  const x = await ok(word, 'insert_ooxml', { ooxml: '<w:p><w:r><w:t>Brief</w:t></w:r></w:p>', location: 'end' })
  assert.equal(x.highlight.selected, true)
  assert.equal(doc.selection.p, doc.paragraphs.length - 1)
  assert.equal(doc.paragraphs[doc.selection.p].text.slice(doc.selection.start, doc.selection.end), 'Brief')
  const before = { ...doc.selection }
  await ok(word, 'read_document_text', {})
  assert.deepEqual(doc.selection, before)
  // Undo restores the text; it never highlights.
  const u = await word.execute('undo_last', {})
  assert.equal(u.ok, true)
  assert.equal(u.data.highlight, undefined)
})

test('PowerPoint: a created slide is selected; without PowerPointApi 1.5 the slide is still created and a note says so', async () => {
  ppt.executor.clearUndo()
  const pres = installPowerPointMock(new MockPresentation())
  const s = await ok(ppt, 'create_text_slide', { title: 'What we decided', bullets: ['Pilot first'] })
  assert.equal(s.highlight.selected, true)
  assert.deepEqual(pres.selected, [s.slideId])
  globalThis.Office.context.requirements.isSetSupported = (set, v) => !(set === 'PowerPointApi' && v === '1.5')
  const n = pres.slides.length
  const s2 = await ok(ppt, 'create_text_slide', { title: 'Who does what' })
  assert.equal(pres.slides.length, n + 1)
  assert.equal(s2.undo.available, true)
  assert.equal(s2.highlight.selected, false)
  assert.match(s2.highlight.note, /could not be highlighted \(Unsupported: Selecting the new slide needs PowerPointApi 1\.5/)
  assert.deepEqual(pres.selected, [s.slideId])
})
