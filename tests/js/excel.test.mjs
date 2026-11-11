// SPDX-License-Identifier: Apache-2.0
// node --test tests/js/ : Excel executor against the strict Office.js mock.
import assert from 'node:assert/strict'
import { test } from 'node:test'
import { installMock, MockWorkbook, modelFingerprint } from '../../addin/dev/office-mock.js'
import { execute, executor, parseCsv, SUPPORTED_TOOLS, MUTATING } from '../../addin/excel/executor.js'
// office_insert_image pane commands (plugin-level tool; undo cases live in images.test.mjs)
const IMAGE_COMMANDS = ['insert_image', 'image_status']

function fresh() { executor.clearUndo(); return installMock(new MockWorkbook()) }
async function ok(tool, args) { const r = await execute(tool, args); assert.equal(r.ok, true, `${tool}: ${JSON.stringify(r)}`); return r.data }
const DATA = [['Region', 'Sales', 'Units'], ['East', 100, 3], ['West', 300, 1], ['North', 200, 2]]
async function seed() { await ok('write_range', { sheet: 'Sheet1', address: 'A1:C4', values: DATA }); executor.clearUndo() }

const CANONICAL = ['get_workbook_overview', 'read_range', 'summarize_range_stats', 'write_range', 'format_range', 'conditional_format_range',
  'freeze_panes', 'set_data_validation', 'create_sheet', 'delete_sheet', 'activate_sheet', 'insert_table', 'sort_table', 'filter_table',
  'create_chart', 'find_replace', 'clear_range', 'autofill_formula', 'rename_sheet', 'copy_sheet', 'protect_sheet', 'unprotect_sheet',
  'create_pivot_table', 'run_calculation', 'add_comment', 'clear_comments', 'import_csv_to_sheet']

test('pane advertises all 27 primitives plus undo_last', () => {
  assert.deepEqual([...SUPPORTED_TOOLS].sort(), [...CANONICAL, ...IMAGE_COMMANDS, 'undo_last'].sort())
})

test('overview reports sheets, active sheet and used range', async () => {
  fresh()
  await ok('write_range', { sheet: 'Sheet1', address: 'B2:C3', values: [[1, 2], [3, 4]] })
  const d = await ok('get_workbook_overview', {})
  assert.equal(d.activeSheet, 'Sheet1')
  assert.deepEqual(d.sheets[0], { name: 'Sheet1', visibility: 'Visible', usedRange: 'Sheet1!B2:C3', rows: 2, columns: 2 })
})

test('write then read round-trips values and live formulas', async () => {
  fresh()
  const w = await ok('write_range', { sheet: 'Sheet1', address: 'A1:B3', values: [['Item', 'Cost'], ['Rent', 1200], ['Total', '=SUM(B2:B2)+0']] })
  assert.equal(w.cellsWritten, 6)
  assert.equal(w.undo.available, true)
  let r = await ok('read_range', { sheet: 'Sheet1', address: 'A1:B3' })
  assert.deepEqual(r.values, [['Item', 'Cost'], ['Rent', 1200], ['Total', 1200]])
  assert.equal(r.formulas[2][1], '=SUM(B2:B2)+0')
  await ok('write_range', { sheet: 'Sheet1', address: 'B2', values: [[1500]] })
  r = await ok('read_range', { sheet: 'Sheet1', address: 'B3' })
  assert.deepEqual(r.values, [[1500]], 'formula recalculates after an edit')
})

test('summarize_range_stats is deterministic', async () => {
  fresh(); await seed()
  const s = await ok('summarize_range_stats', { sheet: 'Sheet1', address: 'B1:B5' })
  assert.deepEqual({ ...s, address: undefined }, { address: undefined, count: 5, numericCount: 3, sum: 600, average: 200, min: 100, max: 300, blankCount: 1 })
})

test('validation and Office errors are clean envelopes and change nothing', async () => {
  const wb = fresh()
  let r = await execute('write_range', { sheet: 'Sheet1', address: 'A1:B2', values: [[1, 2, 3]] })
  assert.equal(r.error.category, 'Validation')
  assert.equal(wb.sheets[0].usedBox(), null)
  r = await execute('read_range', { sheet: 'Nope', address: 'A1' })
  assert.equal(r.error.code, 'ItemNotFound')
  r = await execute('excel_nope', {})
  assert.equal(r.error.category, 'Unsupported')
  r = await execute('undo_last', {})
  assert.equal(r.error.category, 'NothingToUndo')
})

test('parseCsv handles quotes, escaped quotes and CRLF', () => {
  assert.deepEqual(parseCsv('a,"b,c"\r\n"d ""q""",e\n'), [['a', 'b,c'], ['d "q"', 'e']])
})

// Every mutating tool: apply, verify a change, undo, verify the workbook is exactly as before.
const MUTATION_CASES = [
  ['write_range', null, { sheet: 'Sheet1', address: 'A2:B2', values: [['South', 50]] }],
  ['format_range', null, { sheet: 'Sheet1', address: 'A1:C1', format: { bold: true, fillColor: '#107C41', fontColor: '#FFFFFF', numberFormat: '0.00', borders: true, columnAutofit: true, horizontalAlignment: 'Center' } }],
  ['conditional_format_range', null, { sheet: 'Sheet1', address: 'B2:B4', ruleType: 'cellValue', operator: 'GreaterThan', formula1: '150', fillColor: '#FFC7CE' }],
  ['conditional_format_range', null, { sheet: 'Sheet1', address: 'A2:A4', ruleType: 'containsText', text: 'st', fontColor: '#9C0006' }],
  ['freeze_panes', null, { sheet: 'Sheet1', rows: 1, columns: 1 }],
  ['set_data_validation', null, { sheet: 'Sheet1', address: 'D2:D4', type: 'list', source: 'Yes,No', promptTitle: 'Pick', promptMessage: 'Yes or No' }],
  ['create_sheet', null, { name: 'Summary' }],
  ['activate_sheet', async () => { await ok('create_sheet', { name: 'Other' }); executor.clearUndo() }, { name: 'Other' }],
  ['insert_table', null, { sheet: 'Sheet1', address: 'A1:C4', hasHeaders: true }],
  ['sort_table', async () => { await ok('insert_table', { sheet: 'Sheet1', address: 'A1:C4', hasHeaders: true }); executor.clearUndo() }, { tableName: 'Table1', columnIndex: 1, ascending: false }],
  ['filter_table', async () => { await ok('insert_table', { sheet: 'Sheet1', address: 'A1:C4', hasHeaders: true }); executor.clearUndo() }, { tableName: 'Table1', columnIndex: 1, criterion1: '>150' }],
  ['create_chart', null, { sheet: 'Sheet1', dataRange: 'A1:B4', chartType: 'ColumnClustered', title: 'Sales' }],
  ['find_replace', null, { sheet: 'Sheet1', find: 'est', replace: 'EST', matchCase: true }],
  ['clear_range', async () => { await ok('format_range', { sheet: 'Sheet1', address: 'A1:C1', format: { bold: true } }); executor.clearUndo() }, { sheet: 'Sheet1', address: 'A1:C2' }],
  ['autofill_formula', async () => { await ok('write_range', { sheet: 'Sheet1', address: 'D2', values: [['=B2*C2']] }); executor.clearUndo() }, { sheet: 'Sheet1', sourceAddress: 'D2', targetAddress: 'D2:D4' }],
  ['rename_sheet', null, { oldName: 'Sheet1', newName: 'Data' }],
  ['copy_sheet', null, { sourceName: 'Sheet1', newName: 'Copy' }],
  ['protect_sheet', null, { sheet: 'Sheet1', allowSort: true }],
  ['unprotect_sheet', async () => { await ok('protect_sheet', { sheet: 'Sheet1' }); executor.clearUndo() }, { sheet: 'Sheet1' }],
  ['create_pivot_table', async () => { await ok('create_sheet', { name: 'Pivot' }); executor.clearUndo() }, { sheet: 'Pivot', address: 'A3', sourceSheet: 'Sheet1', sourceRange: 'A1:C4', rowFields: ['Region'], dataFields: [{ name: 'Sales', summarizeBy: 'sum' }] }],
  ['add_comment', null, { sheet: 'Sheet1', address: 'B2', text: 'Check this' }],
  ['add_comment', async () => { await ok('add_comment', { sheet: 'Sheet1', address: 'B2', text: 'old' }); executor.clearUndo() }, { sheet: 'Sheet1', address: 'B2', text: 'new' }],
  ['clear_comments', async () => { await ok('add_comment', { sheet: 'Sheet1', address: 'B2', text: 'x' }); await ok('add_comment', { sheet: 'Sheet1', address: 'C3', text: 'y' }); executor.clearUndo() }, { sheet: 'Sheet1', address: 'A1:C4' }],
  ['import_csv_to_sheet', null, { sheet: 'Sheet1', address: 'E1', csv: 'h1,h2\n1,"=cmd|x"\n"a,b",3\n', hasHeaders: true }]
]

for (const [tool, setup, args] of MUTATION_CASES) {
  test(`${tool} applies and undo restores the exact prior workbook (${JSON.stringify(args).slice(0, 60)})`, async () => {
    const wb = fresh(); await seed()
    if (setup) await setup()
    const before = modelFingerprint(wb)
    const d = await ok(tool, args)
    assert.equal(d.undo.available, true, `${tool} should be undoable: ${JSON.stringify(d.undo)}`)
    assert.notEqual(modelFingerprint(wb), before, `${tool} changed nothing`)
    const u = await ok('undo_last', {})
    assert.equal(u.undone, tool)
    assert.equal(modelFingerprint(wb), before, `${tool} undo did not restore the workbook`)
  })
}

test('every mutating tool is covered by an undo case or a documented no-undo case', () => {
  const covered = new Set(MUTATION_CASES.map(c => c[0]))
  const noUndo = new Set(['delete_sheet'])
  for (const t of MUTATING) assert.ok(covered.has(t) || noUndo.has(t) || IMAGE_COMMANDS.includes(t), `no undo test for ${t}`)
})

test('delete_sheet reports undo unavailable and refuses the last sheet', async () => {
  const wb = fresh()
  let r = await execute('delete_sheet', { name: 'Sheet1' })
  assert.equal(r.error.category, 'Conflict')
  await ok('create_sheet', { name: 'Tmp' })
  const d = await ok('delete_sheet', { name: 'Tmp' })
  assert.equal(d.undo.available, false)
  assert.match(d.undo.reason, /cannot restore/)
  assert.equal(wb.sheets.length, 1)
  // The last undoable entry is create_sheet of the now-deleted sheet: undo fails safely, nothing else changes.
  r = await execute('undo_last', {})
  assert.equal(r.ok, false)
  assert.match(r.error.message, /Undo of create_sheet failed/)
  assert.equal(wb.sheets.length, 1)
})

test('csv import blocks formula injection and bolds headers', async () => {
  const wb = fresh()
  await ok('import_csv_to_sheet', { sheet: 'Sheet1', address: 'A1', csv: 'a,b\n=1+1,-5\n@x,+3', hasHeaders: true })
  const r = await ok('read_range', { sheet: 'Sheet1', address: 'A1:B3' })
  assert.deepEqual(r.values, [['a', 'b'], ['=1+1', -5], ['@x', 3]])
  assert.equal(wb.sheets[0].cell(1, 0).f, null, 'CSV text never becomes a formula')
  assert.equal(wb.sheets[0].cell(0, 0).fmt.font.bold, true)
})

test('password protection is never echoed and is not undoable', async () => {
  fresh()
  const d = await ok('protect_sheet', { sheet: 'Sheet1', password: 'pw-123' })
  assert.equal(JSON.stringify(d).includes('pw-123'), false)
  assert.equal(d.undo.available, false)
  let r = await execute('unprotect_sheet', { sheet: 'Sheet1', password: 'wrong' })
  assert.equal(r.ok, false)
  r = await ok('unprotect_sheet', { sheet: 'Sheet1', password: 'pw-123' })
  assert.equal(r.changed, true)
})

test('run_calculation recalculates without an undo entry', async () => {
  const wb = fresh()
  const d = await ok('run_calculation', {})
  assert.equal(d.calculated, true)
  assert.equal(wb.lastCalc, 'Full')
  assert.equal(executor.undoDepth(), 0)
})

test('undo stack is LIFO across several edits', async () => {
  const wb = fresh(); await seed()
  const base = modelFingerprint(wb)
  await ok('write_range', { sheet: 'Sheet1', address: 'A2', values: [['X']] })
  await ok('format_range', { sheet: 'Sheet1', address: 'A2', format: { italic: true } })
  await ok('create_sheet', { name: 'Two' })
  assert.equal(executor.undoDepth(), 3)
  for (const t of ['create_sheet', 'format_range', 'write_range']) assert.equal((await ok('undo_last', {})).undone, t)
  assert.equal(modelFingerprint(wb), base)
})
