// SPDX-License-Identifier: Apache-2.0
// Live-test probe (CDP, real Office.js): every slide's layout, shapes, geometry, text and font (ids excluded).
PowerPoint.run(async ctx => {
  const slides = ctx.presentation.slides
  slides.load('items/id')
  await ctx.sync()
  const per = slides.items.map(s => { const sh = s.shapes; sh.load('items/type,items/name,items/left,items/top,items/width,items/height'); const l = s.layout; l.load('name'); return { sh, l } })
  await ctx.sync()
  const texts = []
  for (const p of per) for (const s of p.sh.items) {
    if (['GeometricShape', 'TextBox', 'Placeholder'].includes(s.type)) { const tr = s.textFrame.textRange; tr.load('text,font/bold,font/italic,font/color,font/size,font/name'); texts.push(tr) } else texts.push(null)
  }
  await ctx.sync()
  let k = 0
  return JSON.stringify(per.map(p => ({ layout: p.l.name, shapes: p.sh.items.map(s => { const t = texts[k++]; return [s.type, s.name, Math.round(s.left), Math.round(s.top), Math.round(s.width), Math.round(s.height), t && [t.text, t.font.bold, t.font.italic, t.font.color, t.font.size, t.font.name]] }) })))
})
