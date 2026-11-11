// SPDX-License-Identifier: Apache-2.0
// Canonical Excel tool -> Office.js executor for the Hermes bridge task pane.
// Runs against the real Office.js `Excel` global (Excel for the web or desktop),
// or the strict in-memory mock (dev/office-mock.js) in tests. Arguments were
// already validated against the canonical schema by the Hermes plugin; this file
// re-checks only what Office.js would otherwise crash on.
//
// Coverage: the 27 primitive excel_* tools of the native schema set. The 28th,
// excel_execute_plan, is deliberately not implemented here (see README: its
// guards depend on a companion-side plan authority and host turn metadata).
import { ExecError, createExecutor, requireSet, str } from '../shared/runtime.js'
import { altText, imageBase64 } from '../shared/images.js'

const MAX_READ_CELLS = 5000
const SNAPSHOT_MAX_CELLS = 10000
const MAX_COMMENTS_SCAN = 500

function excelRun(fn) {
  const x = globalThis.Excel
  if (!x || typeof x.run !== 'function') throw new ExecError('Unavailable', 'Excel API is not available in this pane')
  return x.run(fn)
}

function sheetOf(ctx, name) {
  return ctx.workbook.worksheets.getItem(str(name, 'sheet'))
}

function local(address) {
  const a = String(address)
  const i = a.lastIndexOf('!')
  return i >= 0 ? a.slice(i + 1) : a
}

const PROPS = {
  format: {
    font: { bold: true, italic: true, color: true, name: true, size: true, underline: true },
    fill: { color: true, pattern: true },
    horizontalAlignment: true, verticalAlignment: true, wrapText: true, indentLevel: true, textOrientation: true,
    columnWidth: true, rowHeight: true,
    borders: { style: true, color: true, weight: true }
  }
}

const FONT_KEYS = ['bold', 'italic', 'color', 'name', 'size', 'underline']
const FORMAT_KEYS = ['horizontalAlignment', 'verticalAlignment', 'wrapText', 'indentLevel', 'textOrientation', 'columnWidth', 'rowHeight']
const BORDER_EDGES = ['top', 'bottom', 'left', 'right']

function pick(obj, keys) {
  const out = {}
  for (const k of keys) if (obj && obj[k] !== undefined && obj[k] !== null) out[k] = obj[k]
  return out
}

/**
 * getCellProperties returns service metadata (e.g. "@odata.type") and read-only fields that
 * setCellProperties rejects as InvalidArgument; keep only the settable properties we captured.
 */
export function cleanProps(grid) {
  return grid.map(row => row.map(cell => {
    const f = (cell && cell.format) || {}
    const fmt = pick(f, FORMAT_KEYS)
    if (f.font) fmt.font = pick(f.font, FONT_KEYS)
    // No-fill cells report { color: '', pattern: null }; restoreInto clears fills first, so only real fills are re-applied.
    if (f.fill && f.fill.color && f.fill.pattern !== 'None') fmt.fill = { color: f.fill.color }
    if (f.borders) {
      const b = {}
      for (const edge of BORDER_EDGES) {
        const v = f.borders[edge]
        if (v && v.style) b[edge] = v.style === 'None' ? { style: 'None' } : pick(v, ['style', 'color', 'weight'])
      }
      fmt.borders = b
    }
    return { format: fmt }
  }))
}

/** Capture formulas, number formats and (optionally) cell formatting of a range for undo. */
async function snapshot(ctx, sheetName, address, { props = false } = {}) {
  const range = sheetOf(ctx, sheetName).getRange(address)
  range.load('address,rowCount,columnCount')
  await ctx.sync()
  if (range.rowCount * range.columnCount > SNAPSHOT_MAX_CELLS) return null
  range.load('formulas,numberFormat')
  let cp = null
  if (props) {
    requireSet('ExcelApi', '1.9', 'Format undo')
    cp = range.getCellProperties(PROPS)
  }
  await ctx.sync()
  return { sheet: sheetName, address: local(range.address), formulas: range.formulas, numberFormat: range.numberFormat, props: cp ? cleanProps(cp.value) : null }
}

function restoreInto(ctx, snap) {
  const range = sheetOf(ctx, snap.sheet).getRange(snap.address)
  if (snap.props) {
    range.format.fill.clear()
    range.setCellProperties(snap.props)
  }
  range.numberFormat = snap.numberFormat
  range.formulas = snap.formulas
}

function rangeUndo(snap, label, partial) {
  if (!snap) return { undoUnavailable: `range is larger than ${SNAPSHOT_MAX_CELLS} cells; no undo snapshot was taken` }
  return { undo: async ctx => { restoreInto(ctx, snap); await ctx.sync() }, undoLabel: label, undoPartial: partial }
}

/** Change highlight (write paths only): show the sheet and select the range that was just written or formatted. */
function selectRange(sheet, address) {
  return ctx => {
    const ws = sheetOf(ctx, sheet)
    ws.activate()
    ws.getRange(local(address)).select()
    return `${sheet}!${local(address)}`
  }
}

function supports(set, version) {
  try { requireSet(set, version, 'highlight'); return true } catch { return false }
}

function rect(values) {
  return Array.isArray(values) && values.length > 0 && values.every(r => Array.isArray(r) && r.length > 0 && r.length === values[0].length)
}

/** RFC 4180 CSV parser (quotes, doubled quotes, CRLF/LF). */
export function parseCsv(text) {
  const rows = []
  let row = []
  let field = ''
  let q = false
  for (let i = 0; i < text.length; i++) {
    const ch = text[i]
    if (q) {
      if (ch === '"') {
        if (text[i + 1] === '"') { field += '"'; i++ } else q = false
      } else field += ch
    } else if (ch === '"' && field === '') q = true
    else if (ch === ',') { row.push(field); field = '' }
    else if (ch === '\n' || ch === '\r') {
      if (ch === '\r' && text[i + 1] === '\n') i++
      row.push(field); rows.push(row); row = []; field = ''
    } else field += ch
  }
  if (q) throw new ExecError('Validation', 'csv has an unterminated quoted field')
  if (field !== '' || row.length) { row.push(field); rows.push(row) }
  return rows
}

function csvCell(v) {
  // Formula-injection guard: CSV is data. A leading = + - @ that is not a plain number stays text.
  if (/^[=+\-@]/.test(v) && !/^[+-]?\d+(\.\d+)?$/.test(v)) return `'${v}`
  return v
}

async function findCommentsIn(ctx, ws, address) {
  requireSet('ExcelApi', '1.10', 'Comments')
  const comments = ws.comments
  comments.load('items/id,items/content')
  await ctx.sync()
  if (comments.items.length > MAX_COMMENTS_SCAN) throw new ExecError('Capacity', `Sheet has more than ${MAX_COMMENTS_SCAN} comments`)
  const target = ws.getRange(address)
  const probes = comments.items.map(c => {
    const loc = c.getLocation()
    loc.load('address')
    const hit = target.getIntersectionOrNullObject(loc)
    hit.load('address')
    return { c, loc, hit }
  })
  await ctx.sync()
  return probes.filter(p => !p.hit.isNullObject).map(p => ({ id: p.c.id, content: p.c.content, address: local(p.loc.address) }))
}

const tools = {
  async get_workbook_overview(ctx) {
    const sheets = ctx.workbook.worksheets
    sheets.load('items/name,items/visibility')
    const active = sheets.getActiveWorksheet()
    active.load('name')
    const selection = ctx.workbook.getSelectedRange()
    selection.load('address')
    await ctx.sync()
    const used = sheets.items.map(s => {
      const r = s.getUsedRangeOrNullObject(true)
      r.load('address,rowCount,columnCount')
      return r
    })
    await ctx.sync()
    return {
      activeSheet: active.name,
      selection: selection.address,
      sheets: sheets.items.map((s, i) => ({
        name: s.name,
        visibility: s.visibility,
        usedRange: used[i].isNullObject ? null : used[i].address,
        rows: used[i].isNullObject ? 0 : used[i].rowCount,
        columns: used[i].isNullObject ? 0 : used[i].columnCount
      }))
    }
  },

  async read_range(ctx, { sheet, address }) {
    const range = sheetOf(ctx, sheet).getRange(str(address, 'address'))
    range.load('address,rowCount,columnCount')
    await ctx.sync()
    const cells = range.rowCount * range.columnCount
    if (cells > MAX_READ_CELLS) {
      return { address: range.address, rows: range.rowCount, columns: range.columnCount, truncated: true,
        message: `Range has ${cells} cells; read at most ${MAX_READ_CELLS} at a time.` }
    }
    range.load('values,formulas,numberFormat')
    await ctx.sync()
    return { address: range.address, rows: range.rowCount, columns: range.columnCount,
      values: range.values, formulas: range.formulas, numberFormat: range.numberFormat, truncated: false }
  },

  async summarize_range_stats(ctx, { sheet, address }) {
    const range = sheetOf(ctx, sheet).getRange(str(address, 'address'))
    range.load('address,rowCount,columnCount')
    await ctx.sync()
    if (range.rowCount * range.columnCount > 100000) throw new ExecError('Capacity', 'Range is larger than 100000 cells')
    range.load('values')
    await ctx.sync()
    let count = 0; let numericCount = 0; let blankCount = 0; let sum = 0; let min = null; let max = null
    for (const row of range.values) {
      for (const v of row) {
        count++
        if (v === '' || v === null) { blankCount++; continue }
        if (typeof v === 'number' && Number.isFinite(v)) {
          numericCount++; sum += v
          min = min === null ? v : Math.min(min, v)
          max = max === null ? v : Math.max(max, v)
        }
      }
    }
    return { address: range.address, count, numericCount, sum, average: numericCount ? sum / numericCount : null, min, max, blankCount }
  },

  async write_range(ctx, { sheet, address, values }) {
    if (!rect(values)) throw new ExecError('Validation', 'values must be a non-empty rectangular 2D array')
    const range = sheetOf(ctx, sheet).getRange(str(address, 'address'))
    range.load('rowCount,columnCount')
    await ctx.sync()
    if (range.rowCount !== values.length || range.columnCount !== values[0].length) {
      throw new ExecError('Validation', `values are ${values.length}x${values[0].length} but ${address} is ${range.rowCount}x${range.columnCount}`)
    }
    const snap = await snapshot(ctx, sheet, address)
    // `formulas` accepts constants too: strings beginning with = become formulas.
    range.formulas = values
    range.load('address')
    await ctx.sync()
    return { data: { address: range.address, cellsWritten: values.length * values[0].length }, ...rangeUndo(snap, `write ${address}`), highlight: selectRange(sheet, address) }
  },

  async format_range(ctx, { sheet, address, format }) {
    const f = format || {}
    const snap = await snapshot(ctx, sheet, str(address, 'address'), { props: true })
    const range = sheetOf(ctx, sheet).getRange(address)
    const font = range.format.font
    if (f.bold !== undefined) font.bold = f.bold
    if (f.italic !== undefined) font.italic = f.italic
    if (f.fontColor !== undefined) font.color = f.fontColor
    if (f.fontName !== undefined) font.name = f.fontName
    if (f.fontSize !== undefined) font.size = f.fontSize
    if (f.underline !== undefined) font.underline = f.underline
    if (f.fillColor !== undefined) range.format.fill.color = f.fillColor
    if (f.horizontalAlignment !== undefined) range.format.horizontalAlignment = f.horizontalAlignment
    if (f.verticalAlignment !== undefined) range.format.verticalAlignment = f.verticalAlignment
    if (f.wrapText !== undefined) range.format.wrapText = f.wrapText
    if (f.indentLevel !== undefined) range.format.indentLevel = f.indentLevel
    if (f.textOrientation !== undefined) range.format.textOrientation = f.textOrientation
    if (f.numberFormat !== undefined) {
      range.load('rowCount,columnCount')
      await ctx.sync()
      range.numberFormat = Array.from({ length: range.rowCount }, () => Array(range.columnCount).fill(f.numberFormat))
    }
    if (f.borders) {
      for (const edge of ['EdgeTop', 'EdgeBottom', 'EdgeLeft', 'EdgeRight', 'InsideHorizontal', 'InsideVertical']) {
        const b = range.format.borders.getItem(edge)
        b.style = 'Continuous'
        if (f.borderColor) b.color = f.borderColor
        if (f.borderWeight) b.weight = f.borderWeight
      }
    }
    if (f.columnAutofit) range.format.autofitColumns()
    if (f.rowAutofit) range.format.autofitRows()
    range.load('address')
    await ctx.sync()
    return { data: { address: range.address, applied: Object.keys(f) }, ...rangeUndo(snap, `format ${address}`), highlight: selectRange(sheet, address) }
  },

  async conditional_format_range(ctx, { sheet, address, ruleType, operator, formula1, formula2, text, fillColor, fontColor }) {
    const range = sheetOf(ctx, sheet).getRange(str(address, 'address'))
    let cf
    if (ruleType === 'cellValue') {
      if (!formula1) throw new ExecError('Validation', 'cellValue rules need formula1')
      cf = range.conditionalFormats.add('CellValue')
      const rule = { formula1: String(formula1), operator: operator || 'GreaterThan' }
      if (formula2 !== undefined) rule.formula2 = String(formula2)
      cf.cellValue.rule = rule
      if (fillColor) cf.cellValue.format.fill.color = fillColor
      if (fontColor) cf.cellValue.format.font.color = fontColor
    } else if (ruleType === 'containsText') {
      if (!text) throw new ExecError('Validation', 'containsText rules need text')
      cf = range.conditionalFormats.add('ContainsText')
      cf.textComparison.rule = { operator: operator || 'Contains', text: String(text) }
      if (fillColor) cf.textComparison.format.fill.color = fillColor
      if (fontColor) cf.textComparison.format.font.color = fontColor
    } else {
      throw new ExecError('Validation', 'ruleType must be cellValue or containsText')
    }
    cf.load('id')
    range.load('address')
    await ctx.sync()
    const id = cf.id
    return {
      data: { address: range.address, ruleType, id },
      undo: async c => { sheetOf(c, sheet).getRange(address).conditionalFormats.getItem(id).delete(); await c.sync() },
      undoLabel: `conditional format on ${address}`
    }
  },

  async freeze_panes(ctx, { sheet, rows, columns }) {
    const r = Number(rows); const c = Number(columns)
    if (!Number.isInteger(r) || !Number.isInteger(c) || r < 0 || c < 0 || r > 1000 || c > 100) {
      throw new ExecError('Validation', 'rows and columns must be whole numbers (rows <= 1000, columns <= 100)')
    }
    const ws = sheetOf(ctx, sheet)
    const prior = ws.freezePanes.getLocationOrNullObject()
    prior.load('address')
    await ctx.sync()
    const priorAddress = prior.isNullObject ? null : local(prior.address)
    ws.freezePanes.unfreeze()
    if (r > 0 && c > 0) ws.freezePanes.freezeAt(ws.getRangeByIndexes(0, 0, r, c))
    else if (r > 0) ws.freezePanes.freezeRows(r)
    else if (c > 0) ws.freezePanes.freezeColumns(c)
    await ctx.sync()
    return {
      data: { sheet, frozenRows: r, frozenColumns: c, previous: priorAddress },
      undo: async x => {
        const w = sheetOf(x, sheet)
        w.freezePanes.unfreeze()
        if (priorAddress) w.freezePanes.freezeAt(priorAddress)
        await x.sync()
      },
      undoLabel: `freeze panes on ${sheet}`
    }
  },

  async set_data_validation(ctx, { sheet, address, type, source, promptTitle, promptMessage }) {
    if (type !== 'list') throw new ExecError('Validation', 'type must be list')
    const range = sheetOf(ctx, sheet).getRange(str(address, 'address'))
    const dv = range.dataValidation
    dv.load('type,rule,prompt')
    await ctx.sync()
    const priorType = dv.type
    const priorRule = dv.rule
    const priorPrompt = dv.prompt
    dv.clear()
    dv.rule = { list: { inCellDropDown: true, source: str(source, 'source', 8192) } }
    if (promptTitle || promptMessage) dv.prompt = { showPrompt: true, title: String(promptTitle || ''), message: String(promptMessage || '') }
    range.load('address')
    await ctx.sync()
    const data = { address: range.address, type: 'list', source }
    const restorable = priorType === 'None' || (priorRule && ['WholeNumber', 'Decimal', 'List', 'Date', 'Time', 'TextLength', 'Custom'].includes(priorType))
    if (!restorable) return { data, undoUnavailable: `the range had mixed validation (${priorType}) before` }
    return {
      data,
      undo: async x => {
        const d = sheetOf(x, sheet).getRange(address).dataValidation
        d.clear()
        if (priorType !== 'None') {
          const rule = {}
          for (const [k, v] of Object.entries(priorRule)) if (v) rule[k] = v
          d.rule = rule
          if (priorPrompt) d.prompt = priorPrompt
        }
        await x.sync()
      },
      undoLabel: `data validation on ${address}`
    }
  },

  async create_sheet(ctx, { name }) {
    const existing = ctx.workbook.worksheets.getItemOrNullObject(str(name, 'name', 31))
    await ctx.sync()
    if (!existing.isNullObject) throw new ExecError('Conflict', `A sheet named ${name} already exists`)
    const ws = ctx.workbook.worksheets.add(name)
    ws.load('name')
    await ctx.sync()
    const created = ws.name
    return { data: { name: created }, undo: async x => { x.workbook.worksheets.getItem(created).delete(); await x.sync() }, undoLabel: `create sheet ${created}` }
  },

  async delete_sheet(ctx, { name }) {
    const sheets = ctx.workbook.worksheets
    sheets.load('items/name,items/visibility')
    const ws = sheets.getItem(str(name, 'name'))
    ws.load('name')
    await ctx.sync()
    if (sheets.items.filter(s => s.visibility === 'Visible').length <= 1) throw new ExecError('Conflict', 'Cannot delete the only visible sheet')
    ws.delete()
    await ctx.sync()
    return { data: { deleted: ws.name }, undoUnavailable: 'Office.js cannot restore a deleted worksheet; copy it first if you may need it back' }
  },

  async activate_sheet(ctx, { name }) {
    const prior = ctx.workbook.worksheets.getActiveWorksheet()
    prior.load('name')
    const ws = ctx.workbook.worksheets.getItem(str(name, 'name'))
    ws.activate()
    ws.load('name')
    await ctx.sync()
    const before = prior.name
    return { data: { activeSheet: ws.name, previous: before }, undo: async x => { x.workbook.worksheets.getItem(before).activate(); await x.sync() }, undoLabel: `activate ${ws.name}` }
  },

  async insert_table(ctx, { sheet, address, hasHeaders }) {
    const snap = await snapshot(ctx, sheet, str(address, 'address'), { props: true })
    const table = sheetOf(ctx, sheet).tables.add(address, !!hasHeaders)
    table.load('name')
    await ctx.sync()
    const name = table.name
    const highlight = selectRange(sheet, address)
    if (!snap) return { data: { table: name }, undoUnavailable: 'table range too large to snapshot', highlight }
    return {
      data: { table: name },
      highlight,
      undo: async x => {
        x.workbook.tables.getItem(name).convertToRange()
        await x.sync()
        restoreInto(x, snap)
        await x.sync()
      },
      undoLabel: `insert table ${name}`
    }
  },

  async sort_table(ctx, { tableName, columnIndex, ascending }) {
    const table = ctx.workbook.tables.getItem(str(tableName, 'tableName'))
    const cols = table.columns
    cols.load('count')
    const body = table.getDataBodyRange()
    body.load('address,formulas,rowCount,columnCount')
    const ws = table.worksheet
    ws.load('name')
    await ctx.sync()
    const idx = Number(columnIndex)
    if (!Number.isInteger(idx) || idx < 0 || idx >= cols.count) throw new ExecError('Validation', `columnIndex must be 0..${cols.count - 1}`)
    const prior = body.formulas
    const bodyAddr = local(body.address)
    const sheetName = ws.name
    table.sort.apply([{ key: idx, ascending: ascending !== false }])
    await ctx.sync()
    return {
      data: { table: tableName, columnIndex: idx, ascending: ascending !== false, rows: body.rowCount },
      undo: async x => {
        x.workbook.tables.getItem(tableName).sort.clear()
        sheetOf(x, sheetName).getRange(bodyAddr).formulas = prior
        await x.sync()
      },
      undoLabel: `sort ${tableName}`
    }
  },

  async filter_table(ctx, { tableName, columnIndex, values, criterion1, criterion2, operator }) {
    const table = ctx.workbook.tables.getItem(str(tableName, 'tableName'))
    const cols = table.columns
    cols.load('count')
    await ctx.sync()
    const idx = Number(columnIndex)
    if (!Number.isInteger(idx) || idx < 0 || idx >= cols.count) throw new ExecError('Validation', `columnIndex must be 0..${cols.count - 1}`)
    const filter = cols.getItemAt(idx).filter
    filter.load('criteria')
    await ctx.sync()
    const prior = filter.criteria
    if (Array.isArray(values) && values.length) filter.applyValuesFilter(values.map(String))
    else if (criterion1) filter.applyCustomFilter(String(criterion1), criterion2 ? String(criterion2) : undefined, operator || 'And')
    else throw new ExecError('Validation', 'Provide values or criterion1')
    await ctx.sync()
    return {
      data: { table: tableName, columnIndex: idx, filtered: true },
      undo: async x => {
        const fl = x.workbook.tables.getItem(tableName).columns.getItemAt(idx).filter
        fl.clear()
        if (prior && prior.filterOn && prior.filterOn !== 'None' && prior.filterOn !== 'Unknown') {
          const c = {}
          for (const [k, v] of Object.entries(prior)) if (v !== null && v !== undefined && v !== '') c[k] = v
          fl.apply(c)
        }
        await x.sync()
      },
      undoLabel: `filter ${tableName}`
    }
  },

  async create_chart(ctx, { sheet, dataRange, chartType, title }) {
    const ws = sheetOf(ctx, sheet)
    const source = ws.getRange(str(dataRange, 'dataRange'))
    const chart = ws.charts.add(str(chartType, 'chartType'), source, 'Auto')
    chart.title.text = String(title || '')
    chart.load('name')
    await ctx.sync()
    const name = chart.name
    // Chart.activate needs ExcelApi 1.9; older hosts get the chart's source range selected instead.
    const highlight = c => {
      if (!supports('ExcelApi', '1.9')) return selectRange(sheet, dataRange)(c)
      const w = sheetOf(c, sheet)
      w.activate()
      w.charts.getItem(name).activate()
      return `chart ${name} on ${sheet}`
    }
    return { data: { chart: name, chartType }, undo: async x => { sheetOf(x, sheet).charts.getItem(name).delete(); await x.sync() }, undoLabel: `chart ${name}`, highlight }
  },

  async find_replace(ctx, { sheet, find, replace, matchCase }) {
    requireSet('ExcelApi', '1.9', 'Find and replace')
    const ws = sheetOf(ctx, sheet)
    const used = ws.getUsedRangeOrNullObject(true)
    used.load('address')
    await ctx.sync()
    if (used.isNullObject) return { data: { sheet, replacements: 0 }, undoUnavailable: 'nothing was changed' }
    const snap = await snapshot(ctx, sheet, local(used.address))
    const n = ws.replaceAll(str(find, 'find', 32767), String(replace ?? ''), { completeMatch: false, matchCase: !!matchCase })
    await ctx.sync()
    const out = { data: { sheet, replacements: n.value } }
    if (!n.value) return { ...out, undoUnavailable: 'nothing was changed' }
    return { ...out, ...rangeUndo(snap, `replace "${find}" on ${sheet}`) }
  },

  async clear_range(ctx, { sheet, address }) {
    const snap = await snapshot(ctx, sheet, str(address, 'address'), { props: true })
    const range = sheetOf(ctx, sheet).getRange(address)
    range.clear()
    range.load('address')
    await ctx.sync()
    return { data: { cleared: range.address }, ...rangeUndo(snap, `clear ${address}`, 'comments, conditional formats and validation in the range are not restored'), highlight: selectRange(sheet, address) }
  },

  async autofill_formula(ctx, { sheet, sourceAddress, targetAddress }) {
    requireSet('ExcelApi', '1.9', 'Autofill')
    const ws = sheetOf(ctx, sheet)
    const src = ws.getRange(str(sourceAddress, 'sourceAddress'))
    const dest = src.getBoundingRect(ws.getRange(str(targetAddress, 'targetAddress')))
    dest.load('address')
    await ctx.sync()
    const destAddr = local(dest.address)
    const snap = await snapshot(ctx, sheet, destAddr)
    src.autoFill(destAddr, 'FillDefault')
    await ctx.sync()
    return { data: { source: sourceAddress, filled: dest.address }, ...rangeUndo(snap, `autofill ${destAddr}`), highlight: selectRange(sheet, destAddr) }
  },

  async rename_sheet(ctx, { oldName, newName }) {
    str(newName, 'newName', 31)
    const clash = ctx.workbook.worksheets.getItemOrNullObject(newName)
    const ws = ctx.workbook.worksheets.getItem(str(oldName, 'oldName'))
    ws.load('name')
    await ctx.sync()
    if (!clash.isNullObject && newName.toLowerCase() !== oldName.toLowerCase()) throw new ExecError('Conflict', `A sheet named ${newName} already exists`)
    const before = ws.name
    ws.name = newName
    ws.load('name')
    await ctx.sync()
    const after = ws.name
    return { data: { oldName: before, newName: after }, undo: async x => { x.workbook.worksheets.getItem(after).name = before; await x.sync() }, undoLabel: `rename ${before}` }
  },

  async copy_sheet(ctx, { sourceName, newName }) {
    requireSet('ExcelApi', '1.10', 'Copy sheet')
    str(newName, 'newName', 31)
    const clash = ctx.workbook.worksheets.getItemOrNullObject(newName)
    const src = ctx.workbook.worksheets.getItem(str(sourceName, 'sourceName'))
    const active = ctx.workbook.worksheets.getActiveWorksheet()
    active.load('name')
    await ctx.sync()
    if (!clash.isNullObject) throw new ExecError('Conflict', `A sheet named ${newName} already exists`)
    const wasActive = active.name
    const copy = src.copy('End')
    copy.name = newName
    copy.load('name')
    await ctx.sync()
    const name = copy.name
    return {
      data: { source: sourceName, newName: name },
      undo: async x => {
        // Deleting the (active) copy would activate its neighbour; put the user back where they were.
        x.workbook.worksheets.getItem(wasActive).activate()
        x.workbook.worksheets.getItem(name).delete()
        await x.sync()
      },
      undoLabel: `copy sheet ${name}`
    }
  },

  async protect_sheet(ctx, { sheet, password, allowSort, allowFilter }) {
    const ws = sheetOf(ctx, sheet)
    ws.protection.load('protected')
    await ctx.sync()
    if (ws.protection.protected) throw new ExecError('Conflict', `${sheet} is already protected`)
    const options = {}
    if (allowSort !== undefined) options.allowSort = !!allowSort
    if (allowFilter !== undefined) options.allowAutoFilter = !!allowFilter
    if (password) ws.protection.protect(options, String(password))
    else ws.protection.protect(options)
    await ctx.sync()
    const data = { sheet, protected: true, withPassword: !!password }
    if (password) return { data, undoUnavailable: 'the password is never stored, so Hermes cannot unprotect it for you' }
    return { data, undo: async x => { sheetOf(x, sheet).protection.unprotect(); await x.sync() }, undoLabel: `protect ${sheet}` }
  },

  async unprotect_sheet(ctx, { sheet, password }) {
    const ws = sheetOf(ctx, sheet)
    ws.protection.load('protected,options')
    await ctx.sync()
    if (!ws.protection.protected) return { data: { sheet, protected: false, changed: false }, undoUnavailable: 'nothing was changed' }
    const prior = ws.protection.options
    if (password) ws.protection.unprotect(String(password))
    else ws.protection.unprotect()
    await ctx.sync()
    const data = { sheet, protected: false, changed: true }
    if (password) return { data, undoUnavailable: 'the password is never stored, so the original protection cannot be recreated' }
    return { data, undo: async x => { sheetOf(x, sheet).protection.protect(prior || {}); await x.sync() }, undoLabel: `unprotect ${sheet}` }
  },

  async create_pivot_table(ctx, { sheet, address, sourceSheet, sourceRange, rowFields, columnFields, dataFields }) {
    requireSet('ExcelApi', '1.8', 'Pivot table creation')
    const dest = sheetOf(ctx, sheet)
    const src = sheetOf(ctx, sourceSheet).getRange(str(sourceRange, 'sourceRange'))
    const existing = dest.pivotTables
    existing.load('items/name')
    await ctx.sync()
    let n = existing.items.length + 1
    const taken = new Set(existing.items.map(p => p.name))
    while (taken.has(`HermesPivot${n}`)) n++
    const name = `HermesPivot${n}`
    // A new pivot autofits its columns; capture widths of the destination row band so undo can put them back.
    const band = dest.getRange(str(address, 'address')).getCell(0, 0).getResizedRange(0, 25)
    band.load('address')
    let widths = null
    try {
      requireSet('ExcelApi', '1.9', 'Column widths')
      widths = band.getCellProperties({ format: { columnWidth: true } })
    } catch (err) { if (!(err instanceof ExecError)) throw err }
    await ctx.sync()
    const bandAddr = local(band.address)
    const priorWidths = widths ? widths.value.map(r => r.map(c => ({ format: { columnWidth: c.format.columnWidth } }))) : null
    const pivot = dest.pivotTables.add(name, src, dest.getRange(address))
    for (const f of rowFields || []) pivot.rowHierarchies.add(pivot.hierarchies.getItem(str(f, 'rowFields[]')))
    for (const f of columnFields || []) pivot.columnHierarchies.add(pivot.hierarchies.getItem(str(f, 'columnFields[]')))
    const agg = { sum: 'Sum', count: 'Count', average: 'Average', max: 'Max', min: 'Min' }
    for (const d of dataFields || []) {
      const h = pivot.dataHierarchies.add(pivot.hierarchies.getItem(str(d.name, 'dataFields[].name')))
      if (d.summarizeBy) h.summarizeBy = agg[d.summarizeBy]
    }
    pivot.load('name')
    await ctx.sync()
    return {
      data: { pivotTable: name, sheet, address },
      undo: async x => {
        const w = sheetOf(x, sheet)
        w.pivotTables.getItem(name).delete()
        if (priorWidths) w.getRange(bandAddr).setCellProperties(priorWidths)
        await x.sync()
      },
      undoLabel: `pivot ${name}`
    }
  },

  async run_calculation(ctx) {
    ctx.workbook.application.calculate('Full')
    await ctx.sync()
    return { calculated: true, calculationType: 'Full' }
  },

  async add_comment(ctx, { sheet, address, text }) {
    const ws = sheetOf(ctx, sheet)
    const cell = ws.getRange(str(address, 'address')).getCell(0, 0)
    cell.load('address')
    await ctx.sync()
    const cellAddr = local(cell.address)
    const existing = (await findCommentsIn(ctx, ws, cellAddr))[0]
    if (existing) {
      const c = ctx.workbook.comments.getItem(existing.id)
      c.content = String(text)
      await ctx.sync()
      return {
        data: { address: cell.address, updated: true },
        undo: async x => { x.workbook.comments.getItem(existing.id).content = existing.content; await x.sync() },
        undoLabel: `comment on ${cellAddr}`
      }
    }
    const c = ctx.workbook.comments.add(cell, String(text))
    c.load('id')
    await ctx.sync()
    const id = c.id
    return { data: { address: cell.address, updated: false }, undo: async x => { x.workbook.comments.getItem(id).delete(); await x.sync() }, undoLabel: `comment on ${cellAddr}` }
  },

  async clear_comments(ctx, { sheet, address }) {
    const ws = sheetOf(ctx, sheet)
    const found = await findCommentsIn(ctx, ws, str(address, 'address'))
    for (const f of found) ctx.workbook.comments.getItem(f.id).delete()
    await ctx.sync()
    const data = { address, deleted: found.length }
    if (!found.length) return { data, undoUnavailable: 'nothing was changed' }
    return {
      data,
      undo: async x => {
        const w = sheetOf(x, sheet)
        for (const f of found) x.workbook.comments.add(w.getRange(f.address), f.content)
        await x.sync()
      },
      undoLabel: `clear comments in ${address}`,
      undoPartial: 'comment text is restored; replies and original authorship are not'
    }
  },

  async import_csv_to_sheet(ctx, { sheet, address, csv, hasHeaders }) {
    if (typeof csv !== 'string' || !csv.length) throw new ExecError('Validation', 'csv must be a non-empty string')
    if (csv.length > 4 * 1024 * 1024) throw new ExecError('Capacity', 'csv is larger than 4 MB')
    const rows = parseCsv(csv)
    if (!rows.length) throw new ExecError('Validation', 'csv has no rows')
    const width = Math.max(...rows.map(r => r.length))
    if (rows.length * width > 200000) throw new ExecError('Capacity', 'csv has more than 200000 cells')
    const values = rows.map(r => Array.from({ length: width }, (_, i) => csvCell(r[i] ?? '')))
    const ws = sheetOf(ctx, sheet)
    const target = ws.getRange(str(address, 'address')).getCell(0, 0).getResizedRange(values.length - 1, width - 1)
    target.load('address')
    await ctx.sync()
    const targetAddr = local(target.address)
    const snap = await snapshot(ctx, sheet, targetAddr, { props: !!hasHeaders })
    target.values = values
    if (hasHeaders) target.getRow(0).format.font.bold = true
    await ctx.sync()
    return { data: { address: target.address, rows: values.length, columns: width, hasHeaders: !!hasHeaders }, ...rangeUndo(snap, `import CSV into ${targetAddr}`), highlight: selectRange(sheet, targetAddr) }
  },

  // ---- office_insert_image (plugin-level tool; not a canonical excel_* schema) ----
  async insert_image(ctx, { base64, alt, title, sheet, cell, width, height }) {
    requireSet('ExcelApi', '1.9', 'Pictures on a worksheet (Shapes.addImage)')
    const { b64, kind } = imageBase64(base64)
    const description = altText(alt)
    const ws = sheet ? sheetOf(ctx, sheet) : ctx.workbook.worksheets.getActiveWorksheet()
    ws.load('name')
    const anchor = cell ? ws.getRange(str(cell, 'cell', 20)).getCell(0, 0) : null
    if (anchor) anchor.load('left,top,address')
    await ctx.sync()
    const shape = ws.shapes.addImage(b64)
    shape.name = 'Hermes image'
    if (anchor) { shape.left = anchor.left; shape.top = anchor.top }
    shape.lockAspectRatio = !(width !== undefined && height !== undefined)
    if (width !== undefined) shape.width = Number(width)
    if (height !== undefined) shape.height = Number(height)
    shape.load('id')
    await ctx.sync()
    const id = shape.id
    const sheetName = ws.name
    let altTextApplied = true
    try {
      shape.altTextDescription = description
      shape.altTextTitle = String(title || description).slice(0, 255)
      await ctx.sync()
    } catch { altTextApplied = false }
    return {
      data: { sheet: sheetName, shapeId: id, cell: anchor ? local(anchor.address) : null, format: kind, altTextApplied },
      undo: async x => { x.workbook.worksheets.getItem(sheetName).shapes.getItem(id).delete(); await x.sync() },
      undoLabel: `insert image on ${sheetName}`
    }
  },

  async image_status(ctx, { sheet, shapeId }) {
    const shapes = sheetOf(ctx, sheet).shapes
    shapes.load('items/id,items/type')
    await ctx.sync()
    const hit = shapes.items.find(s => s.id === str(shapeId, 'shapeId'))
    return { exists: !!hit, ...(hit ? { type: hit.type } : {}), images: shapes.items.filter(s => s.type === 'Image').length }
  }
}

export const MUTATING = new Set([
  'write_range', 'format_range', 'conditional_format_range', 'freeze_panes', 'set_data_validation', 'create_sheet',
  'delete_sheet', 'activate_sheet', 'insert_table', 'sort_table', 'filter_table', 'create_chart', 'find_replace',
  'clear_range', 'autofill_formula', 'rename_sheet', 'copy_sheet', 'protect_sheet', 'unprotect_sheet',
  'create_pivot_table', 'add_comment', 'clear_comments', 'import_csv_to_sheet', 'insert_image'
])

export const executor = createExecutor({ tools, mutating: MUTATING, run: excelRun })
export const { execute, SUPPORTED_TOOLS } = executor
