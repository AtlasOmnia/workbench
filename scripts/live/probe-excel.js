// SPDX-License-Identifier: Apache-2.0
// Live-test probe (CDP, real Office.js): deterministic snapshot of the whole small test workbook.
Excel.run(async ctx => {
  const sheets = ctx.workbook.worksheets
  sheets.load('items/name,items/visibility,items/position')
  const active = sheets.getActiveWorksheet(); active.load('name')
  await ctx.sync()
  const per = sheets.items.map(s => {
    const u = s.getUsedRangeOrNullObject(); u.load('address,formulas,numberFormat')
    const ch = s.charts; ch.load('items/name')
    const t = s.tables; t.load('items/name')
    const pv = s.pivotTables; pv.load('items/name')
    const pr = s.protection; pr.load('protected')
    const fz = s.freezePanes.getLocationOrNullObject(); fz.load('address')
    const cm = s.comments; cm.load('items/content')
    const area = s.getRange('A1:K12')
    const cf = area.conditionalFormats.getCount()
    const dv = s.getRange('D2:D4').dataValidation; dv.load('type')
    const props = area.getCellProperties({ format: { font: { bold: true, italic: true, color: true }, fill: { color: true }, horizontalAlignment: true, columnWidth: true, borders: { style: true } } })
    return { s, u, ch, t, pv, pr, fz, cm, cf, dv, props }
  })
  await ctx.sync()
  const tables = []
  for (const p of per) for (const t of p.t.items) {
    const cols = t.columns; cols.load('items/name')
    const r = t.getRange(); r.load('address')
    tables.push({ t, cols, r })
  }
  await ctx.sync()
  const filters = tables.map(x => x.cols.items.map(c => { const f = c.filter; f.load('criteria'); return f }))
  await ctx.sync()
  return JSON.stringify({
    active: active.name,
    sheets: per.map(p => ({
      name: p.s.name, vis: p.s.visibility, pos: p.s.position,
      used: p.u.isNullObject ? null : [p.u.address, p.u.formulas, p.u.numberFormat],
      charts: p.ch.items.map(c => c.name), tables: p.t.items.length, pivots: p.pv.items.map(x => x.name), protected: p.pr.protected,
      freeze: p.fz.isNullObject ? null : p.fz.address, comments: p.cm.items.map(c => c.content).sort(), cf: p.cf.value, dv: p.dv.type,
      fmt: p.props.value.map(row => row.map(c => JSON.stringify(c.format))).flat().join('|')
    })),
    tables: tables.map((x, i) => ({ addr: x.r.address, cols: x.cols.items.map(c => c.name), filters: filters[i].map(f => JSON.stringify(f.criteria)) }))
  })
})
