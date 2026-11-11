// SPDX-License-Identifier: Apache-2.0
// Strict in-memory stand-in for the Office.js Excel subset the executor uses.
// DEV/TEST ONLY: loaded by taskpane.html when ?mock=1 and by node tests. It
// enforces Office.js's load()/sync() rule (reading an unloaded property throws
// PropertyNotLoaded) so executor bugs surface before a live Excel session.
// It is not Microsoft code and does not claim Office.js fidelity beyond this.
import { ClientResult, Loadable, LoadableCollection, MockContextBase, OfficeError, installOffice, makeRun, persistence } from './mock-core.js'

function colToNum(s) { let n = 0; for (const ch of s) n = n * 26 + (ch.charCodeAt(0) - 64); return n - 1 }
function numToCol(n) { let s = ''; n += 1; while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26) } return s }
function parseCell(ref) {
  const m = /^\$?([A-Z]{1,3})\$?([0-9]{1,7})$/.exec(ref.toUpperCase())
  if (!m) throw new OfficeError('InvalidArgument', `Bad address ${ref}`)
  return { c: colToNum(m[1]), r: Number(m[2]) - 1 }
}
export function parseAddress(addr) {
  const local = String(addr).includes('!') ? String(addr).split('!').pop() : String(addr)
  const [a, b] = local.replace(/\$/g, '').split(':')
  const p = parseCell(a); const q = b ? parseCell(b) : p
  return { r0: Math.min(p.r, q.r), c0: Math.min(p.c, q.c), r1: Math.max(p.r, q.r), c1: Math.max(p.c, q.c) }
}
function boxAddr({ r0, c0, r1, c1 }) {
  const a = `${numToCol(c0)}${r0 + 1}`; const b = `${numToCol(c1)}${r1 + 1}`
  return a === b ? a : `${a}:${b}`
}
function fmtAddress(sheet, box) { return `${/[^A-Za-z0-9_]/.test(sheet) ? `'${sheet}'` : sheet}!${boxAddr(box)}` }
function inBox(b, r, c) { return r >= b.r0 && r <= b.r1 && c >= b.c0 && c <= b.c1 }
function shiftFormula(f, dr, dc) {
  return f.replace(/(\$?)([A-Z]{1,3})(\$?)([0-9]{1,7})/g, (_m, d1, col, d2, row) =>
    `${d1}${d1 ? col : numToCol(colToNum(col) + dc)}${d2}${d2 ? row : Number(row) + dr}`)
}
const EDGE = { EdgeTop: 'top', EdgeBottom: 'bottom', EdgeLeft: 'left', EdgeRight: 'right' }
const DEFAULT_FMT = () => ({
  font: { bold: false, italic: false, color: '#000000', name: 'Calibri', size: 11, underline: 'None' },
  fill: { color: '', pattern: null }, // what real Excel reports for a cell with no fill
  horizontalAlignment: 'General', verticalAlignment: 'Bottom', wrapText: false, indentLevel: 0, textOrientation: 0,
  columnWidth: 48, rowHeight: 15,
  borders: { top: { style: 'None' }, bottom: { style: 'None' }, left: { style: 'None' }, right: { style: 'None' } }
})

let uid = 0
const nextId = p => `${p}-${++uid}`

class MockRange extends Loadable {
  constructor(ctx, sheet, box, isNull = false) {
    super(ctx); this._sheet = sheet; this._box = box; this.isNullObject = isNull
    const self = this
    const q = fn => ctx._queue(fn)
    this.format = {
      font: new Proxy({}, { set(_t, k, v) { q(() => self._each(c => { c.fmt.font = { ...(c.fmt.font || {}), [k]: v } })); return true } }),
      fill: new Proxy({ clear: () => q(() => self._each(c => { delete c.fmt.fill })) }, { set(_t, k, v) { if (k === 'color') q(() => self._each(c => { c.fmt.fill = { color: v, pattern: 'Solid' } })); return true } }),
      borders: { getItem: edge => new Proxy({}, { set(_t, k, v) { q(() => self._each((c, r, col) => { const e = EDGE[edge]; if (!e) return; c.fmt.borders = { ...(c.fmt.borders || {}), [e]: { ...((c.fmt.borders || {})[e] || {}), [k]: v } } })); return true } }) },
      autofitColumns() { q(() => self._each(c => { c.fmt.columnWidth = 99 })) },
      autofitRows() { q(() => self._each(c => { c.fmt.rowHeight = 30 })) }
    }
    // Range.select (ExcelApi 1.1): the UI selection; real Excel selects only on the active sheet.
    this.select = () => q(() => {
      if (sheet.wb.active !== sheet.name) throw new OfficeError('InvalidSelection', 'The range is not on the active worksheet')
      sheet.wb.selection = { sheet: sheet.name, address: boxAddr(self._box) }
    })
    for (const k of ['horizontalAlignment', 'verticalAlignment', 'wrapText', 'indentLevel', 'textOrientation']) {
      Object.defineProperty(this.format, k, { set(v) { q(() => self._each(c => { c.fmt[k] = v })) }, enumerable: true })
    }
    this.conditionalFormats = {
      add: type => {
        const cf = new (class extends Loadable { get id() { this._need('id'); return this._id } })(ctx)
        cf._id = nextId('cf')
        const spec = { format: { fill: {}, font: {} }, rule: null }
        if (type === 'CellValue') cf.cellValue = spec
        else cf.textComparison = spec
        q(() => sheet.cfs.push({ id: cf._id, box: self._box, type, spec }))
        return cf
      },
      getItem: id => ({ delete: () => q(() => { const i = sheet.cfs.findIndex(x => x.id === id); if (i < 0) throw new OfficeError('ItemNotFound', 'cf'); sheet.cfs.splice(i, 1) }) })
    }
    const dvState = {}
    this.dataValidation = new (class extends Loadable {
      get type() { this._need('type'); const s = new Set(); self._each(c => s.add(JSON.stringify(c.dv || null))); if (s.size > 1) return 'Inconsistent'; const v = JSON.parse([...s][0]); return v ? 'List' : 'None' }
      get rule() { this._need('rule'); let first; self._each(c => { if (first === undefined) first = c.dv }); return first ? { list: first.list } : { list: null, wholeNumber: null } }
      set rule(v) { q(() => self._each(c => { c.dv = { ...(c.dv || {}), list: v.list } })) }
      get prompt() { this._need('prompt'); let first; self._each(c => { if (first === undefined) first = c.dv }); return first && first.prompt ? first.prompt : { showPrompt: false, title: '', message: '' } }
      set prompt(v) { q(() => self._each(c => { c.dv = { ...(c.dv || {}), prompt: v } })) }
      clear() { q(() => self._each(c => { c.dv = null })) }
    })(ctx)
    void dvState
  }
  _each(fn) { const { r0, c0, r1, c1 } = this._box; for (let r = r0; r <= r1; r++) for (let c = c0; c <= c1; c++) fn(this._sheet.cell(r, c), r, c) }
  _grid(fn) { const { r0, c0, r1, c1 } = this._box; const out = []; for (let r = r0; r <= r1; r++) { const row = []; for (let c = c0; c <= c1; c++) row.push(fn(this._sheet.cell(r, c), r, c)); out.push(row) } return out }
  _set(matrix, fn) {
    if (matrix.length !== this._box.r1 - this._box.r0 + 1 || matrix.some(r => r.length !== this._box.c1 - this._box.c0 + 1)) {
      throw new OfficeError('InvalidArgument', 'The number of rows or columns in the input array does not match the size or dimensions of the range.')
    }
    if (this._sheet.protection.protected) throw new OfficeError('AccessDenied', 'The cell or chart you are trying to change is protected.')
    matrix.forEach((row, i) => row.forEach((v, j) => fn(this._sheet.cell(this._box.r0 + i, this._box.c0 + j), v)))
  }
  get address() { this._need('address'); return fmtAddress(this._sheet.name, this._box) }
  // Points from the sheet origin, with the mock's default 48 pt columns and 15 pt rows (ExcelApi 1.3).
  get left() { this._need('left'); return this._box.c0 * 48 }
  get top() { this._need('top'); return this._box.r0 * 15 }
  get rowCount() { this._need('rowCount'); return this._box.r1 - this._box.r0 + 1 }
  get columnCount() { this._need('columnCount'); return this._box.c1 - this._box.c0 + 1 }
  get values() { this._need('values'); return this._grid(c => (c.f ? this._sheet.wb.evaluate(this._sheet, c.f) : c.v ?? '')) }
  set values(m) {
    this._ctx._queue(() => this._set(m, (c, v) => {
      if (typeof v === 'string' && v.startsWith("'")) { c.v = v.slice(1); c.f = null } else if (typeof v === 'string' && v.startsWith('=')) { c.f = v; c.v = null } else if (typeof v === 'string' && /^[+-]?\d+(\.\d+)?$/.test(v)) { c.v = Number(v); c.f = null } else { c.v = v === '' ? null : v; c.f = null }
    }))
  }
  get formulas() { this._need('formulas'); return this._grid(c => c.f ?? c.v ?? '') }
  set formulas(m) { this._ctx._queue(() => this._set(m, (c, v) => { if (typeof v === 'string' && v.startsWith('=')) { c.f = v; c.v = null } else { c.v = v === '' ? null : v; c.f = null } })) }
  get numberFormat() { this._need('numberFormat'); return this._grid(c => c.nf ?? 'General') }
  set numberFormat(m) { this._ctx._queue(() => this._set(m, (c, v) => { c.nf = v === 'General' ? null : v })) }
  clear() { this._ctx._queue(() => this._each(c => { c.v = null; c.f = null; c.nf = null; c.fmt = {} })) }
  getCellProperties() {
    return new ClientResult(this._ctx, () => this._grid((c, r, col) => {
      const d = DEFAULT_FMT(); const f = c.fmt
      return { address: fmtAddress(this._sheet.name, { r0: r, c0: col, r1: r, c1: col }), format: {
        '@odata.type': 'Microsoft.ExcelServices.CellPropertiesFormatInternal', ...d, ...Object.fromEntries(Object.entries(f).filter(([k]) => !['font', 'fill', 'borders'].includes(k))),
        font: { '@odata.type': 'Microsoft.ExcelServices.CellPropertiesFont', ...d.font, ...(f.font || {}) }, fill: f.fill ? { ...f.fill } : d.fill, borders: { ...d.borders, ...(f.borders || {}) }
      } }
    }))
  }
  setCellProperties(grid) {
    this._ctx._queue(() => grid.forEach((row, i) => row.forEach((p, j) => {
      const c = this._sheet.cell(this._box.r0 + i, this._box.c0 + j)
      const f = (p && p.format) || {}
      const d = DEFAULT_FMT()
      // Like Office.js: only the properties supplied are set; anything else keeps its value.
      const nf = { ...c.fmt }
      for (const [k, v] of Object.entries(f)) {
        if (k.startsWith('@')) throw new OfficeError('InvalidArgument', 'The argument is invalid or missing or has an incorrect format.')
        if (k === 'font') { for (const fk of Object.keys(v)) if (!['bold', 'italic', 'color', 'name', 'size', 'underline'].includes(fk)) throw new OfficeError('InvalidArgument', 'The argument is invalid or missing or has an incorrect format.'); nf.font = { ...d.font, ...(nf.font || {}), ...v } }
        else if (k === 'fill') { if (v.color === '' || v.pattern === null) throw new OfficeError('InvalidArgument', 'The argument is invalid or missing or has an incorrect format.'); if (v.pattern === 'None') delete nf.fill; else nf.fill = { color: v.color, pattern: 'Solid' } }
        else if (k === 'borders') nf.borders = { ...d.borders, ...(nf.borders || {}), ...v }
        else nf[k] = v
      }
      c.fmt = nf
    })))
  }
  getCell(r, c) { return new MockRange(this._ctx, this._sheet, { r0: this._box.r0 + r, c0: this._box.c0 + c, r1: this._box.r0 + r, c1: this._box.c0 + c }) }
  getResizedRange(dr, dc) { return new MockRange(this._ctx, this._sheet, { ...this._box, r1: this._box.r1 + dr, c1: this._box.c1 + dc }) }
  getRow(i) { return new MockRange(this._ctx, this._sheet, { ...this._box, r0: this._box.r0 + i, r1: this._box.r0 + i }) }
  getBoundingRect(o) { const a = this._box; const b = o._box; return new MockRange(this._ctx, this._sheet, { r0: Math.min(a.r0, b.r0), c0: Math.min(a.c0, b.c0), r1: Math.max(a.r1, b.r1), c1: Math.max(a.c1, b.c1) }) }
  getIntersectionOrNullObject(o) {
    const a = this._box; const b = o._box || parseAddress(o)
    const box = { r0: Math.max(a.r0, b.r0), c0: Math.max(a.c0, b.c0), r1: Math.min(a.r1, b.r1), c1: Math.min(a.c1, b.c1) }
    const isNull = box.r0 > box.r1 || box.c0 > box.c1 || (o._sheet && o._sheet !== this._sheet)
    return new MockRange(this._ctx, this._sheet, isNull ? { r0: 0, c0: 0, r1: 0, c1: 0 } : box, isNull)
  }
  autoFill(dest) {
    const d = typeof dest === 'string' ? parseAddress(dest) : dest._box
    this._ctx._queue(() => {
      const s = this._box
      for (let r = d.r0; r <= d.r1; r++) {
        for (let c = d.c0; c <= d.c1; c++) {
          if (inBox(s, r, c)) continue
          const src = this._sheet.cell(s.r0 + ((r - s.r0) % (s.r1 - s.r0 + 1) + (s.r1 - s.r0 + 1)) % (s.r1 - s.r0 + 1), s.c0 + ((c - s.c0) % (s.c1 - s.c0 + 1) + (s.c1 - s.c0 + 1)) % (s.c1 - s.c0 + 1))
          const cell = this._sheet.cell(r, c)
          if (src.f) { cell.f = shiftFormula(src.f, r - s.r0, c - s.c0); cell.v = null } else { cell.v = src.v; cell.f = null }
        }
      }
    })
  }
}

class MockSheet extends Loadable {
  constructor(ctx, model) {
    super(ctx); this._m = model
    const q = fn => ctx._queue(fn)
    const wb = model.wb
    this.tables = { add: (address, hasHeaders) => wb.addTable(ctx, this._m, address, hasHeaders) }
    this.charts = {
      add: (type, range, seriesBy) => wb.addChart(ctx, this._m, type, range, seriesBy),
      getItem: name => ({
        delete: () => q(() => { const i = wb.charts.findIndex(c => c.name === name && c.sheet === this._m.name); if (i < 0) throw new OfficeError('ItemNotFound', 'chart'); wb.charts.splice(i, 1) }),
        activate: () => q(() => { if (!wb.charts.some(c => c.name === name && c.sheet === this._m.name)) throw new OfficeError('ItemNotFound', 'chart'); wb.activeChart = name })
      })
    }
    const sheet = this
    this.freezePanes = {
      getLocationOrNullObject: () => (model.freeze ? new MockRange(ctx, model, model.freeze) : new MockRange(ctx, model, { r0: 0, c0: 0, r1: 0, c1: 0 }, true)),
      unfreeze: () => q(() => { model.freeze = null }),
      freezeAt: r => q(() => { model.freeze = typeof r === 'string' ? parseAddress(r) : r._box }),
      freezeRows: n => q(() => { model.freeze = { r0: 0, c0: 0, r1: n - 1, c1: 16383 } }),
      freezeColumns: n => q(() => { model.freeze = { r0: 0, c0: 0, r1: 1048575, c1: n - 1 } })
    }
    this.protection = new (class extends Loadable {
      get protected() { this._need('protected'); return model.protection.protected }
      get options() { this._need('options'); return { ...model.protection.options } }
      protect(options = {}, password) { q(() => { if (model.protection.protected) throw new OfficeError('InvalidOperation', 'already protected'); model.protection = { protected: true, options, password: password || null } }) }
      unprotect(password) { q(() => { if (model.protection.password && model.protection.password !== password) throw new OfficeError('InvalidArgument', 'The password is incorrect'); model.protection = { protected: false, options: {}, password: null } }) }
    })(ctx)
    const pivotList = () => wb.pivots.filter(p => p.sheet === model.name).map(p => Object.assign(new (class extends Loadable { get name() { this._need('name'); return p.name } })(ctx)))
    this.pivotTables = Object.assign(new LoadableCollection(ctx, pivotList), {
      add: (name, src, dest) => wb.addPivot(ctx, model, name, src, dest),
      getItem: name => ({ delete: () => q(() => { const i = wb.pivots.findIndex(p => p.name === name); if (i < 0) throw new OfficeError('ItemNotFound', 'pivot'); wb.pivots.splice(i, 1) }) })
    })
    this.comments = new LoadableCollection(ctx, () => wb.comments.filter(c => c.sheet === model.name).map(c => wb.commentObj(ctx, c)))
    // Worksheet pictures (ExcelApi 1.9): addImage(base64) takes a bare base64 PNG/JPEG.
    const shapeObj = rec => {
      const o = new (class extends Loadable {
        get id() { this._need('id'); return rec.id }
        get name() { this._need('name'); return rec.name }
        set name(v) { q(() => { rec.name = String(v) }) }
        get type() { this._need('type'); return rec.type }
        get altTextDescription() { this._need('altTextDescription'); return rec.altTextDescription }
        get altTextTitle() { this._need('altTextTitle'); return rec.altTextTitle }
        get lockAspectRatio() { this._need('lockAspectRatio'); return rec.lockAspectRatio }
        delete() { q(() => { const i = model.shapes.indexOf(rec); if (i < 0) throw new OfficeError('ItemNotFound', 'shape'); model.shapes.splice(i, 1) }) }
      })(ctx)
      for (const k of ['left', 'top', 'width', 'height', 'altTextDescription', 'altTextTitle', 'lockAspectRatio']) {
        const get = Object.getOwnPropertyDescriptor(Object.getPrototypeOf(o), k)
        Object.defineProperty(o, k, { get: get ? get.get : function () { this._need(k); return rec[k] }, set: v => q(() => { rec[k] = v }) })
      }
      return o
    }
    this.shapes = Object.assign(new LoadableCollection(ctx, () => model.shapes.map(shapeObj)), {
      addImage: b64 => {
        const rec = { id: nextId('shape'), name: `Picture ${uid}`, type: 'Image', left: 0, top: 0, width: 96, height: 96, altTextDescription: '', altTextTitle: '', lockAspectRatio: true }
        if (typeof b64 !== 'string' || !/^(iVBORw0KGgo|\/9j\/)/.test(b64) || b64.startsWith('data:')) ctx._fail(new OfficeError('InvalidArgument', 'The argument is invalid or missing or has an incorrect format.'))
        else q(() => model.shapes.push(rec))
        return shapeObj(rec)
      },
      getItem: id => {
        const rec = model.shapes.find(x => x.id === id)
        if (!rec) { ctx._fail(new OfficeError('ItemNotFound', 'shape')); return shapeObj({ id }) }
        return shapeObj(rec)
      }
    })
    void sheet
  }
  get name() { this._need('name'); return this._m.name }
  set name(v) { this._ctx._queue(() => { if (this._m.wb.sheets.some(s => s !== this._m && s.name.toLowerCase() === v.toLowerCase())) throw new OfficeError('ItemAlreadyExists', 'exists'); const old = this._m.name; this._m.name = v; if (this._m.wb.active === old) this._m.wb.active = v }) }
  get visibility() { this._need('visibility'); return 'Visible' }
  getRange(address) { return new MockRange(this._ctx, this._m, parseAddress(address)) }
  getRangeByIndexes(r, c, rc, cc) { return new MockRange(this._ctx, this._m, { r0: r, c0: c, r1: r + rc - 1, c1: c + cc - 1 }) }
  getUsedRangeOrNullObject() {
    const box = this._m.usedBox()
    return box ? new MockRange(this._ctx, this._m, box) : new MockRange(this._ctx, this._m, { r0: 0, c0: 0, r1: 0, c1: 0 }, true)
  }
  activate() { this._ctx._queue(() => { this._m.wb.active = this._m.name }) }
  delete() { this._ctx._queue(() => { const wb = this._m.wb; if (wb.sheets.length <= 1) throw new OfficeError('InvalidOperation', 'last sheet'); const i = wb.sheets.indexOf(this._m); wb.sheets.splice(i, 1); if (wb.active === this._m.name) wb.active = wb.sheets[Math.max(0, i - 1)].name }) }
  copy() {
    const wb = this._m.wb
    const m = new SheetModel(wb, `${this._m.name} (2)`)
    for (const [k, c] of this._m.cells) m.cells.set(k, JSON.parse(JSON.stringify(c)))
    this._ctx._queue(() => { wb.sheets.push(m); wb.active = m.name })
    return new MockSheet(this._ctx, m)
  }
  replaceAll(find, repl, { matchCase } = {}) {
    return new ClientResult(this._ctx, () => {
      let n = 0
      const re = new RegExp(find.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'), matchCase ? 'g' : 'gi')
      for (const c of this._m.cells.values()) {
        if (c.f && re.test(c.f)) { re.lastIndex = 0; c.f = c.f.replace(re, repl); n++ } else if (typeof c.v === 'string' && re.test(c.v)) { re.lastIndex = 0; c.v = c.v.replace(re, repl); n++ }
        re.lastIndex = 0
      }
      return n
    })
  }
}

class SheetModel {
  constructor(wb, name) { this.wb = wb; this.name = name; this.cells = new Map(); this.log = []; this.cfs = []; this.shapes = []; this.freeze = null; this.protection = { protected: false, options: {}, password: null } }
  cell(r, c) { const k = `${r},${c}`; if (!this.cells.has(k)) this.cells.set(k, { v: null, f: null, nf: null, fmt: {}, dv: null }); return this.cells.get(k) }
  usedBox() {
    let box = null
    for (const [k, c] of this.cells) {
      if (c.v === null && c.f === null) continue
      const [r, col] = k.split(',').map(Number)
      box = box ? { r0: Math.min(box.r0, r), c0: Math.min(box.c0, col), r1: Math.max(box.r1, r), c1: Math.max(box.c1, col) } : { r0: r, c0: col, r1: r, c1: col }
    }
    return box
  }
}

export class MockWorkbook {
  constructor() { this.sheets = [new SheetModel(this, 'Sheet1')]; this.active = 'Sheet1'; this.tables = []; this.charts = []; this.pivots = []; this.comments = []; this.calcs = 0; this.syncs = 0 }
  find(name) { return this.sheets.find(s => s.name.toLowerCase() === String(name).toLowerCase()) }
  addTable(ctx, sheet, address, hasHeaders) {
    const t = new (class extends Loadable { get name() { this._need('name'); return this._name } })(ctx)
    ctx._queue(() => {
      const box = parseAddress(address)
      if (this.tables.some(x => x.sheet === sheet.name && !(x.box.r1 < box.r0 || x.box.r0 > box.r1 || x.box.c1 < box.c0 || x.box.c0 > box.c1))) throw new OfficeError('InvalidOperation', 'A table can\'t overlap another table.')
      t._name = `Table${this.tables.length + 1}`
      this.tables.push({ name: t._name, sheet: sheet.name, box, hasHeaders, filters: {}, sorted: null })
    })
    return t
  }
  addChart(ctx, sheet, type, range, seriesBy) {
    const c = new (class extends Loadable { get name() { this._need('name'); return this._name } })(ctx)
    c.title = {}
    ctx._queue(() => { c._name = `Chart ${this.charts.length + 1}`; this.charts.push({ name: c._name, sheet: sheet.name, type, source: range._box, seriesBy, title: c.title }) })
    return c
  }
  addPivot(ctx, sheet, name, src, dest) {
    const p = { name, sheet: sheet.name, src: src._box, srcSheet: src._sheet.name, dest: dest._box, rows: [], cols: [], data: [] }
    const headers = () => { const s = src._sheet; const out = []; for (let c = p.src.c0; c <= p.src.c1; c++) out.push(String(s.cell(p.src.r0, c).v ?? '')) ; return out }
    const hier = n => ({ _n: n })
    const pivot = new (class extends Loadable { get name() { this._need('name'); return name } })(ctx)
    pivot.hierarchies = { getItem: n => hier(n) }
    const adder = key => ({ add: h => { const o = { name: h._n, summarizeBy: 'Sum' }; ctx._queue(() => { if (!headers().includes(h._n)) throw new OfficeError('InvalidArgument', `No field ${h._n}`); p[key].push(o) }); return o } })
    pivot.rowHierarchies = adder('rows'); pivot.columnHierarchies = adder('cols'); pivot.dataHierarchies = adder('data')
    ctx._queue(() => {
      if (this.pivots.some(x => x.name === name)) throw new OfficeError('ItemAlreadyExists', 'pivot')
      this.pivots.push(p)
      for (let c = p.dest.c0; c < p.dest.c0 + 3; c++) sheet.cell(p.dest.r0, c).fmt.columnWidth = 78.75 // real pivots autofit
    })
    return pivot
  }
  commentObj(ctx, c) {
    const wb = this
    return new (class extends Loadable {
      get id() { this._need('id'); return c.id }
      get content() { this._need('content'); return c.content }
      getLocation() { return new MockRange(ctx, wb.find(c.sheet), parseAddress(c.address)) }
    })(ctx)
  }
  table(name) { const t = this.tables.find(x => x.name === name); return t }
  /** Tiny evaluator: =SUM(A1:B2), =A1+B1 style refs and numbers. Enough to show live recalculation. */
  evaluate(sheet, formula) {
    const val = ref => { const { r, c } = parseCell(ref); const cell = sheet.cell(r, c); return Number(cell.f ? this.evaluate(sheet, cell.f) : cell.v) || 0 }
    let expr = formula.slice(1).toUpperCase().replace(/SUM\(\$?([A-Z]+)\$?([0-9]+):\$?([A-Z]+)\$?([0-9]+)\)/g, (_m, a1, a2, b1, b2) => {
      const box = parseAddress(`${a1}${a2}:${b1}${b2}`); let s = 0
      for (let r = box.r0; r <= box.r1; r++) for (let c = box.c0; c <= box.c1; c++) s += val(`${numToCol(c)}${r + 1}`)
      return String(s)
    })
    expr = expr.replace(/\$?[A-Z]{1,3}\$?[0-9]{1,7}/g, ref => String(val(ref.replace(/\$/g, ''))))
    if (!/^[0-9+\-*/(). ]*$/.test(expr)) return '#NAME?'
    try { return arith(expr || '0') } catch { return '#VALUE!' }
  }
}

/** Recursive-descent arithmetic (+ - * / parentheses); no eval. */
function arith(src) {
  const s = src.replace(/\s+/g, ''); let i = 0
  const num = () => { const m = /^[0-9]*\.?[0-9]+/.exec(s.slice(i)); if (!m) throw new Error('num'); i += m[0].length; return Number(m[0]) }
  const atom = () => { if (s[i] === '(') { i++; const v = sum(); if (s[i++] !== ')') throw new Error(')'); return v } if (s[i] === '-') { i++; return -atom() } return num() }
  const prod = () => { let v = atom(); while (s[i] === '*' || s[i] === '/') { const op = s[i++]; const w = atom(); v = op === '*' ? v * w : v / w } return v }
  const sum = () => { let v = prod(); while (s[i] === '+' || s[i] === '-') { const op = s[i++]; const w = prod(); v = op === '+' ? v + w : v - w } return v }
  const v = sum(); if (i !== s.length) throw new Error('trailing'); return v
}

class MockContext extends MockContextBase {
  constructor(wb) {
    super(); this._model = wb
    const ctx = this
    const q = fn => ctx._queue(fn)
    const tableObj = name => {
      const t = wb.table(name)
      if (!t) { ctx._fail(new OfficeError('ItemNotFound', `The requested resource doesn't exist.`)); return tableObj.dummy() }
      const sheet = wb.find(t.sheet)
      const bodyBox = () => ({ ...t.box, r0: t.box.r0 + (t.hasHeaders ? 1 : 0) })
      const width = t.box.c1 - t.box.c0 + 1
      const column = i => ({
        filter: new (class extends Loadable {
          get criteria() { this._need('criteria'); return t.filters[i] ? { ...t.filters[i] } : { filterOn: 'None', values: null, criterion1: null } }
          applyValuesFilter(values) { q(() => { t.filters[i] = { filterOn: 'Values', values } }) }
          applyCustomFilter(c1, c2, op) { q(() => { t.filters[i] = { filterOn: 'Custom', criterion1: c1, criterion2: c2 || null, operator: op } }) }
          apply(c) { q(() => { t.filters[i] = { ...c } }) }
          clear() { q(() => { delete t.filters[i] }) }
        })(ctx)
      })
      return {
        columns: Object.assign(new LoadableCollection(ctx, () => Array.from({ length: width }, () => ({ _loaded: new Set() }))), { getItemAt: column }),
        getDataBodyRange: () => new MockRange(ctx, sheet, bodyBox()),
        worksheet: new MockSheet(ctx, sheet),
        sort: {
          apply: fields => q(() => {
            const b = bodyBox(); const { key, ascending } = fields[0]
            const rows = []
            for (let r = b.r0; r <= b.r1; r++) { const row = []; for (let c = b.c0; c <= b.c1; c++) row.push({ ...sheet.cell(r, c) }); rows.push(row) }
            const v = cell => (cell.f ? wb.evaluate(sheet, cell.f) : cell.v)
            rows.sort((x, y) => { const a = v(x[key]); const b2 = v(y[key]); const cmp = a < b2 ? -1 : a > b2 ? 1 : 0; return ascending === false ? -cmp : cmp })
            rows.forEach((row, i) => row.forEach((cell, j) => sheet.cells.set(`${b.r0 + i},${b.c0 + j}`, cell)))
            t.sorted = { key, ascending }
          }),
          clear: () => q(() => { t.sorted = null })
        },
        convertToRange: () => q(() => { wb.tables.splice(wb.tables.indexOf(t), 1) })
      }
    }
    tableObj.dummy = () => ({ columns: { load() {}, getItemAt: () => ({ filter: { load() {} } }) }, getDataBodyRange: () => ({ load() {} }), worksheet: { load() {} }, sort: { apply() {}, clear() {} }, convertToRange() {} })
    this.workbook = {
      worksheets: new (class extends Loadable {
        get items() { if (!this._loaded.size) throw new OfficeError('PropertyNotLoaded', "The property 'items' is not available. Call load() and sync()."); return wb.sheets.map(m => { const s = new MockSheet(ctx, m); s._loaded = new Set(this._loaded); return s }) }
        getItem(name) { const m = wb.find(name); if (!m) { ctx._fail(new OfficeError('ItemNotFound', `The requested resource doesn't exist.`)); return new MockSheet(ctx, new SheetModel(wb, name)) } return new MockSheet(ctx, m) }
        getItemOrNullObject(name) { const m = wb.find(name); const s = new MockSheet(ctx, m || new SheetModel(wb, name)); s.isNullObject = !m; return s }
        getActiveWorksheet() { return new MockSheet(ctx, wb.find(wb.active)) }
        add(name) { const m = new SheetModel(wb, name); ctx._queue(() => { if (wb.find(name)) throw new OfficeError('ItemAlreadyExists', 'exists'); wb.sheets.push(m) }); return new MockSheet(ctx, m) }
      })(ctx),
      tables: { getItem: tableObj },
      comments: {
        add: (cell, content) => {
          const rec = { id: nextId('comment'), sheet: cell._sheet.name, address: boxAddr({ r0: cell._box.r0, c0: cell._box.c0, r1: cell._box.r0, c1: cell._box.c0 }), content }
          q(() => wb.comments.push(rec))
          return wb.commentObj(ctx, rec)
        },
        getItem: id => {
          const rec = () => { const r = wb.comments.find(c => c.id === id); if (!r) throw new OfficeError('ItemNotFound', 'comment'); return r }
          return Object.defineProperty({ delete: () => q(() => { wb.comments.splice(wb.comments.indexOf(rec()), 1) }) }, 'content', { set: v => q(() => { rec().content = v }) })
        }
      },
      application: { calculate: type => q(() => { wb.calcs++; wb.lastCalc = type }) },
      getSelectedRange() { return new MockRange(ctx, wb.find(wb.active), { r0: 0, c0: 0, r1: 0, c1: 0 }) }
    }
  }
}

const excelStore = persistence('hermes-mock-excel',
  wb => ({ uid, active: wb.active, tables: wb.tables, charts: wb.charts.map(c => ({ ...c, title: { text: c.title.text } })), pivots: wb.pivots, comments: wb.comments,
    sheets: wb.sheets.map(m => ({ name: m.name, cells: [...m.cells], cfs: m.cfs, shapes: m.shapes, freeze: m.freeze, protection: m.protection })) }),
  d => {
    const wb = new MockWorkbook()
    uid = Math.max(uid, d.uid || 0)
    Object.assign(wb, { active: d.active, tables: d.tables, charts: d.charts, pivots: d.pivots, comments: d.comments })
    wb.sheets = d.sheets.map(x => Object.assign(new SheetModel(wb, x.name), { cells: new Map(x.cells), cfs: x.cfs, shapes: x.shapes || [], freeze: x.freeze, protection: x.protection }))
    return wb
  })

export function installMock(wb) {
  wb = wb || excelStore.load() || new MockWorkbook()
  globalThis.Excel = { run: makeRun(MockContext, wb, () => excelStore.save(wb)) }
  installOffice('Excel', { url: 'mock://Book1.xlsx' })
  globalThis.__hermesMockWorkbook = wb
  globalThis.__hermesMockFingerprint = () => modelFingerprint(wb)
  return wb
}

/** Normalised view of the whole mock workbook, for undo round-trip assertions in tests. */
export function modelFingerprint(wb) {
  const d = JSON.stringify(DEFAULT_FMT())
  const resolved = f => JSON.stringify({ ...DEFAULT_FMT(), ...Object.fromEntries(Object.entries(f || {}).filter(([k]) => !['font', 'fill', 'borders'].includes(k))),
    font: { ...DEFAULT_FMT().font, ...((f || {}).font || {}) }, fill: (f || {}).fill || DEFAULT_FMT().fill, borders: { ...DEFAULT_FMT().borders, ...((f || {}).borders || {}) } })
  return JSON.stringify({
    active: wb.active,
    sheets: wb.sheets.map(s => ({
      name: s.name,
      cells: [...s.cells].filter(([, c]) => c.v !== null || c.f !== null || c.nf || c.dv || resolved(c.fmt) !== d)
        .map(([k, c]) => [k, c.v, c.f, c.nf, c.dv, resolved(c.fmt)]).sort(),
      cfs: s.cfs.map(x => x.id), shapes: s.shapes.map(x => ({ ...x })), freeze: s.freeze, protection: s.protection
    })),
    tables: wb.tables.map(t => ({ ...t })), charts: wb.charts.map(c => c.name), pivots: wb.pivots.map(p => p.name),
    comments: wb.comments.map(c => [c.sheet, c.address, c.content]).sort()
  })
}
