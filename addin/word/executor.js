// SPDX-License-Identifier: Apache-2.0
// Canonical Word tool -> Office.js executor for the Hermes bridge task pane.
// Same contract as the Excel executor: plain JSON in, bridge envelope out.
//
// Undo: Word's Office.js has no undo API. Before any text-changing tool the pane
// captures the body as OOXML (bounded) and undo puts it back; comments are undone
// by deleting the comment Hermes inserted.
import { ExecError, createExecutor, requireSet, str } from '../shared/runtime.js'
import { altText, imageBase64, pictureHtml } from '../shared/images.js'

const MAX_TEXT = 20000
const SNAPSHOT_MAX_CHARS = 5 * 1024 * 1024
const WORD_SEARCH_MAX = 255
const BUILTIN_STYLES = new Set(['Normal', 'Heading1', 'Heading2', 'Heading3', 'Heading4', 'Heading5', 'Heading6', 'Heading7',
  'Heading8', 'Heading9', 'Title', 'Subtitle', 'Quote', 'IntenseQuote', 'ListParagraph', 'NoSpacing', 'Strong', 'Emphasis'])

function wordRun(fn) {
  const w = globalThis.Word
  if (!w || typeof w.run !== 'function') throw new ExecError('Unavailable', 'Word API is not available in this pane')
  return w.run(fn)
}

function docName() {
  const url = String((globalThis.Office && globalThis.Office.context && globalThis.Office.context.document && globalThis.Office.context.document.url) || '')
  return url.split(/[\\/]/).pop().split('?')[0].slice(0, 200)
}

function clip(text, max) {
  const s = String(text ?? '')
  return s.length > max ? { text: s.slice(0, max), truncated: true, totalLength: s.length } : { text: s, truncated: false, totalLength: s.length }
}

/** Body OOXML snapshot for undo; null when the body is too large. */
async function snapshotBody(ctx) {
  const body = ctx.document.body
  const ooxml = body.getOoxml()
  const paras = body.paragraphs
  paras.load('items/text')
  await ctx.sync()
  if (!ooxml.value || ooxml.value.length > SNAPSHOT_MAX_CHARS) return null
  return { ooxml: ooxml.value, paragraphs: paras.items.length }
}

function bodyUndo(snap, label) {
  if (!snap) return { undoUnavailable: 'document is too large to snapshot for undo' }
  return {
    undo: async ctx => {
      const body = ctx.document.body
      body.insertOoxml(snap.ooxml, 'Replace')
      await ctx.sync()
      // Replacing the body can leave one extra empty trailing paragraph; trim exactly that.
      const paras = body.paragraphs
      paras.load('items/text')
      await ctx.sync()
      if (paras.items.length === snap.paragraphs + 1 && paras.items[paras.items.length - 1].text === '' && paras.items.length > 1) {
        paras.items[paras.items.length - 1].delete()
        await ctx.sync()
      }
    },
    undoLabel: label
  }
}

function highlight(range, color) {
  if (color) range.font.highlightColor = String(color)
}

/**
 * Change highlight (insert paths only): select the range the insert returned (Range.select, WordApi 1.1).
 * The user's selection moves to the new text, so a following selection tool acts on what was just inserted.
 */
function selectInserted(range, what) {
  return () => {
    if (!range || typeof range.select !== 'function') throw new ExecError('Unsupported', 'Word did not return the inserted range')
    range.select()
    return what
  }
}

const tools = {
  async get_document_overview(ctx) {
    const body = ctx.document.body
    body.load('text')
    const paras = body.paragraphs
    paras.load('items/style')
    const sel = ctx.document.getSelection()
    sel.load('text')
    const props = ctx.document.properties
    props.load('title,author')
    await ctx.sync()
    return {
      document: docName(),
      title: props.title || '',
      paragraphs: paras.items.length,
      characters: body.text.length,
      words: (body.text.match(/\S+/g) || []).length,
      selection: { length: sel.text.length, preview: sel.text.slice(0, 200) }
    }
  },

  async read_selection(ctx) {
    const sel = ctx.document.getSelection()
    sel.load('text')
    await ctx.sync()
    return clip(sel.text, MAX_TEXT)
  },

  async read_document_text(ctx, { maxLength }) {
    const max = maxLength === undefined ? MAX_TEXT : Math.min(MAX_TEXT, Math.max(1, Math.floor(maxLength)))
    const body = ctx.document.body
    body.load('text')
    await ctx.sync()
    return clip(body.text, max)
  },

  async apply_edit_plan(ctx, { edits }) {
    if (!Array.isArray(edits) || !edits.length || edits.length > 100) throw new ExecError('Validation', 'edits must contain 1-100 items')
    const body = ctx.document.body
    const searches = edits.map(e => {
      if (typeof e.find !== 'string' || !e.find) throw new ExecError('Validation', 'every edit needs find')
      if (e.find.length > WORD_SEARCH_MAX) return null
      const r = body.search(e.find, { matchCase: !!e.matchCase, matchWholeWord: !!e.matchWholeWord })
      r.load('items/text')
      return r
    })
    await ctx.sync()
    const results = edits.map((e, i) => {
      const r = searches[i]
      if (!r) return { index: i, status: 'rejected', reason: `find is longer than Word's ${WORD_SEARCH_MAX}-character search limit` }
      if (r.items.length === 0) return { index: i, status: 'rejected', reason: 'find text not found' }
      if (r.items.length > 1) return { index: i, status: 'rejected', reason: `find text is not unique (${r.items.length} matches)` }
      return { index: i, status: 'ready' }
    })
    if (results.every(x => x.status === 'ready') && edits.length > 1) {
      requireSet('WordApi', '1.3', 'Edit-plan overlap checks')
      const pairs = []
      for (let i = 0; i < edits.length; i++) {
        for (let j = i + 1; j < edits.length; j++) pairs.push([i, j, searches[i].items[0].compareLocationWith(searches[j].items[0])])
      }
      await ctx.sync()
      const ok = new Set(['Before', 'After', 'AdjacentBefore', 'AdjacentAfter'])
      for (const [i, j, rel] of pairs) {
        if (!ok.has(rel.value)) {
          results[i] = { index: i, status: 'rejected', reason: `overlaps edit ${j}` }
          results[j] = { index: j, status: 'rejected', reason: `overlaps edit ${i}` }
        }
      }
    }
    if (results.some(x => x.status !== 'ready')) {
      // All-or-nothing: nothing was changed.
      return { data: { applied: 0, results: results.map(x => (x.status === 'ready' ? { ...x, status: 'not_applied' } : x)) }, undoUnavailable: 'nothing was changed' }
    }
    const snap = await snapshotBody(ctx)
    edits.forEach((e, i) => {
      const nr = searches[i].items[0].insertText(String(e.replace ?? ''), 'Replace')
      highlight(nr, e.highlightColor)
    })
    await ctx.sync()
    return { data: { applied: edits.length, results: edits.map((e, i) => ({ index: i, status: 'applied', ...(e.reason ? { reason: e.reason } : {}) })) }, ...bodyUndo(snap, `edit plan (${edits.length} edits)`) }
  },

  async insert_text(ctx, { text, location }) {
    const loc = location || 'selection'
    const snap = await snapshotBody(ctx)
    const t = str(text, 'text', 5000)
    let r
    if (loc === 'start') r = ctx.document.body.insertText(t, 'Start')
    else if (loc === 'end') r = ctx.document.body.insertText(t, 'End')
    else if (loc === 'selection') r = ctx.document.getSelection().insertText(t, 'End')
    else throw new ExecError('Validation', 'location must be start, end or selection')
    await ctx.sync()
    return { data: { inserted: t.length, location: loc }, ...bodyUndo(snap, `insert text (${loc})`), highlight: selectInserted(r, 'inserted text') }
  },

  async replace_selection(ctx, { text, highlightColor }) {
    const sel = ctx.document.getSelection()
    sel.load('text')
    await ctx.sync()
    const snap = await snapshotBody(ctx)
    const r = sel.insertText(str(text, 'text', 5000), 'Replace')
    highlight(r, highlightColor)
    await ctx.sync()
    return { data: { replacedLength: sel.text.length, insertedLength: text.length }, ...bodyUndo(snap, 'replace selection') }
  },

  async find_replace(ctx, { find, replace, matchCase, matchWholeWord, highlightColor }) {
    if (typeof find !== 'string' || !find) throw new ExecError('Validation', 'find must be a non-empty string')
    if (find.length > WORD_SEARCH_MAX) throw new ExecError('Validation', `find is longer than Word's ${WORD_SEARCH_MAX}-character search limit`)
    const found = ctx.document.body.search(find, { matchCase: !!matchCase, matchWholeWord: !!matchWholeWord })
    found.load('items/text')
    await ctx.sync()
    if (!found.items.length) return { data: { replacements: 0 }, undoUnavailable: 'nothing was changed' }
    if (found.items.length > 500) throw new ExecError('Capacity', `${found.items.length} matches; narrow the search (limit 500)`)
    const snap = await snapshotBody(ctx)
    for (const item of found.items) highlight(item.insertText(String(replace ?? ''), 'Replace'), highlightColor)
    await ctx.sync()
    return { data: { replacements: found.items.length }, ...bodyUndo(snap, `replace "${find}"`) }
  },

  async format_selection(ctx, { bold, italic, color, highlightColor, fontSize }) {
    const sel = ctx.document.getSelection()
    sel.load('text')
    await ctx.sync()
    if (!sel.text) throw new ExecError('Validation', 'Nothing is selected in the document; select text first')
    const snap = await snapshotBody(ctx)
    const f = sel.font
    const applied = []
    if (bold !== undefined) { f.bold = !!bold; applied.push('bold') }
    if (italic !== undefined) { f.italic = !!italic; applied.push('italic') }
    if (color !== undefined) { f.color = String(color); applied.push('color') }
    if (highlightColor !== undefined) { f.highlightColor = highlightColor ? String(highlightColor) : null; applied.push('highlightColor') }
    if (fontSize !== undefined) { f.size = Number(fontSize); applied.push('fontSize') }
    await ctx.sync()
    return { data: { applied, selectionLength: sel.text.length }, ...bodyUndo(snap, 'format selection') }
  },

  async get_document_structure(ctx, { maxParagraphs }) {
    const max = maxParagraphs === undefined ? 50 : Math.min(100, Math.max(1, Math.floor(maxParagraphs)))
    const paras = ctx.document.body.paragraphs
    paras.load('items/text,items/style,items/isListItem')
    const tables = ctx.document.body.tables
    tables.load('items/rowCount')
    await ctx.sync()
    const styles = {}
    for (const p of paras.items) styles[p.style] = (styles[p.style] || 0) + 1
    return {
      paragraphCount: paras.items.length,
      tableCount: tables.items.length,
      headings: paras.items.map((p, i) => ({ i, p })).filter(x => /^(Heading|Title)/i.test(x.p.style)).slice(0, max).map(x => ({ index: x.i, style: x.p.style, text: x.p.text.slice(0, 120) })),
      styleCounts: styles,
      paragraphs: paras.items.slice(0, max).map((p, i) => ({ index: i, style: p.style, isListItem: !!p.isListItem, preview: p.text.slice(0, 160), length: p.text.length })),
      truncated: paras.items.length > max
    }
  },

  async apply_style(ctx, { style }) {
    const name = str(style, 'style', 255)
    const paras = ctx.document.getSelection().paragraphs
    paras.load('items/text')
    await ctx.sync()
    if (!paras.items.length) throw new ExecError('Validation', 'No paragraphs are selected')
    const snap = await snapshotBody(ctx)
    const builtIn = name.replace(/\s+/g, '')
    for (const p of paras.items) {
      if (BUILTIN_STYLES.has(builtIn)) p.styleBuiltIn = builtIn
      else p.style = name
    }
    await ctx.sync()
    return { data: { style: name, paragraphs: paras.items.length, builtIn: BUILTIN_STYLES.has(builtIn) }, ...bodyUndo(snap, `apply style ${name}`) }
  },

  async insert_ooxml(ctx, { ooxml, location }) {
    const x = str(ooxml, 'ooxml', 20000)
    if (!/<pkg:package|<w:document|<w:p[\s>]/.test(x)) throw new ExecError('Validation', 'ooxml must be a WordprocessingML package or fragment')
    const loc = location || 'selection'
    const snap = await snapshotBody(ctx)
    let r
    if (loc === 'start') r = ctx.document.body.insertOoxml(x, 'Start')
    else if (loc === 'end') r = ctx.document.body.insertOoxml(x, 'End')
    else if (loc === 'selection') r = ctx.document.getSelection().insertOoxml(x, 'Replace')
    else throw new ExecError('Validation', 'location must be start, end or selection')
    await ctx.sync()
    return { data: { inserted: true, location: loc }, ...bodyUndo(snap, `insert OOXML (${loc})`), highlight: selectInserted(r, 'inserted content') }
  },

  async insert_comment(ctx, { text }) {
    requireSet('WordApi', '1.4', 'Comments')
    const sel = ctx.document.getSelection()
    sel.load('text')
    await ctx.sync()
    if (!sel.text) throw new ExecError('Validation', 'Nothing is selected in the document; select the text to comment on')
    const c = sel.insertComment(str(text, 'text', 5000))
    c.load('id')
    await ctx.sync()
    const id = c.id
    return {
      data: { commentId: id, anchorPreview: sel.text.slice(0, 120) },
      undo: async x => {
        const all = x.document.body.getComments()
        all.load('items/id')
        await x.sync()
        const hit = all.items.find(i => i.id === id)
        if (!hit) throw new ExecError('NotFound', 'The comment Hermes added is no longer in the document')
        hit.delete()
        await x.sync()
      },
      undoLabel: 'insert comment'
    }
  },

  // ---- office_insert_image (plugin-level tool; not a canonical word_* schema) ----
  // The picture (or HTML block) is wrapped in a content control tagged by the plugin, so readback
  // and undo find exactly what Hermes inserted and nothing else.
  async insert_image(ctx, { base64, html, alt, title, location, width, height, tag }) {
    const t = str(tag, 'tag', 64)
    const loc = location || 'end'
    if (!['end', 'start', 'selection'].includes(loc)) throw new ExecError('Validation', 'location must be end, start or selection')
    let cc
    let kind = 'html'
    const h = html !== undefined ? pictureHtml(html) : null
    const img = h === null ? imageBase64(base64) : null
    const description = h === null ? altText(alt) : null
    requireSet('WordApi', '1.3', h !== null ? 'Inserting HTML in a content control' : 'Inserting pictures in a content control')
    // The picture (or HTML) goes into a paragraph of its own that Hermes' control wraps. Inserting at the body's end
    // instead joins whatever ends the document: after an earlier Hermes picture, Word for the web put the new HTML's
    // picture inside THAT picture's control (seen live, A9), so readback found 0 of 1 and undo left the picture.
    const slot = await ownParagraph(ctx, loc)
    if (h !== null) {
      cc = slot.insertHtml(h, 'Replace').insertContentControl()
    } else {
      kind = img.kind
      cc = slot.insertContentControl()
      const pic = cc.insertInlinePictureFromBase64(img.b64, 'Start')
      pic.altTextDescription = description
      pic.altTextTitle = String(title || description).slice(0, 255)
      pic.lockAspectRatio = !(width !== undefined && height !== undefined)
      if (width !== undefined) pic.width = Number(width)
      if (height !== undefined) pic.height = Number(height)
    }
    cc.tag = t
    cc.title = 'Hermes image'
    cc.appearance = 'BoundingBox'
    cc.load('id')
    await ctx.sync()
    const ccId = cc.id
    return {
      data: { tag: t, contentControlId: ccId, location: loc, kind, ...(html === undefined ? { altTextApplied: true } : {}) },
      undo: async x => {
        const ccs = x.document.contentControls.getByTag(t)
        ccs.load('items/id')
        await x.sync()
        if (ccs.items.length !== 1) throw new ExecError('NotFound', 'The picture Hermes inserted is no longer in the document (its content control is gone); nothing was removed')
        ccs.items[0].delete(false)
        await x.sync()
      },
      undoLabel: `insert image (${loc})`
    }
  },

  async image_status(ctx, { tag }) {
    const ccs = ctx.document.contentControls.getByTag(str(tag, 'tag', 64))
    ccs.load('items/id')
    await ctx.sync()
    if (!ccs.items.length) return { exists: false, controls: 0, pictures: 0 }
    const pics = ccs.items[0].inlinePictures
    pics.load('items')
    await ctx.sync()
    return { exists: ccs.items.length === 1, controls: ccs.items.length, pictures: pics.items.length }
  }
}

/**
 * An empty paragraph for office_insert_image that belongs to no content control and is never the document's final
 * paragraph (Word keeps the final paragraph mark, so deleting a control there leaves an empty paragraph behind).
 * end: just before the document's empty, uncontrolled last paragraph (one is added when the document does not end in
 * one; that one is then left by undo); start: before the first paragraph; selection: after the selected paragraph, or
 * after the content control that holds it.
 */
async function ownParagraph(ctx, loc) {
  const body = ctx.document.body
  if (loc === 'start') return body.paragraphs.getFirst().insertParagraph('', 'Before')
  if (loc === 'selection') {
    const at = ctx.document.getSelection().paragraphs.getLast()
    const parent = at.parentContentControlOrNullObject
    parent.load('isNullObject')
    await ctx.sync()
    return (parent.isNullObject ? at : parent).insertParagraph('', 'After')
  }
  const last = body.paragraphs.getLast()
  last.load('text')
  const pics = last.inlinePictures
  pics.load('items')
  const inner = last.contentControls
  inner.load('items')
  const parent = last.parentContentControlOrNullObject
  parent.load('isNullObject')
  await ctx.sync()
  const free = last.text === '' && !pics.items.length && !inner.items.length && parent.isNullObject
  return (free ? last : body.insertParagraph('', 'End')).insertParagraph('', 'Before')
}

export const MUTATING = new Set(['apply_edit_plan', 'insert_text', 'replace_selection', 'find_replace', 'format_selection',
  'apply_style', 'insert_ooxml', 'insert_comment', 'insert_image'])

export const executor = createExecutor({ tools, mutating: MUTATING, run: wordRun })
export const { execute, SUPPORTED_TOOLS } = executor
