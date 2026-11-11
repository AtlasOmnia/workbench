// SPDX-License-Identifier: Apache-2.0
// Live-test probe (CDP, real Office.js): text, styles, per-word formatting and comments of the small test document.
Word.run(async ctx => {
  const paras = ctx.document.body.paragraphs
  paras.load('items/text,items/style')
  const comments = ctx.document.body.getComments()
  comments.load('items/content')
  await ctx.sync()
  const words = paras.items.map(p => { const w = p.getTextRanges([' '], false); w.load('items/text,items/font/bold,items/font/italic,items/font/color,items/font/size,items/font/highlightColor'); return w })
  await ctx.sync()
  return JSON.stringify({
    paragraphs: paras.items.map((p, i) => [p.text, p.style, words[i].items.map(w => [w.text, w.font.bold, w.font.italic, w.font.color, w.font.size, w.font.highlightColor])]),
    comments: comments.items.map(c => c.content)
  })
})
