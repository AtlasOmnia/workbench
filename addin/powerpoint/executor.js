// SPDX-License-Identifier: Apache-2.0
// Canonical PowerPoint tool -> Office.js executor for the Hermes bridge task pane.
// Same contract as the Excel executor. Undo deletes what Hermes created, restores
// the text/font/layout/position it changed, and re-inserts a deleted slide from a
// base64 export taken just before deletion (PowerPointApi 1.8).
import { ExecError, createExecutor, requireSet, str } from '../shared/runtime.js'
import { altText, imageBase64 } from '../shared/images.js'

const TEXT_TYPES = new Set(['GeometricShape', 'TextBox', 'Placeholder'])
const DEFAULT_MAX = 20000
const MAX_IMAGE_BASE64 = 8 * 1024 * 1024

function pptRun(fn) {
  const p = globalThis.PowerPoint
  if (!p || typeof p.run !== 'function') throw new ExecError('Unavailable', 'PowerPoint API is not available in this pane')
  return p.run(fn)
}

function docName() {
  const url = String((globalThis.Office && globalThis.Office.context && globalThis.Office.context.document && globalThis.Office.context.document.url) || '')
  return url.split(/[\\/]/).pop().split('?')[0].slice(0, 200)
}

function cap(maxLength) {
  return maxLength === undefined ? DEFAULT_MAX : Math.max(1, Math.min(DEFAULT_MAX, Math.floor(Number(maxLength))))
}

async function slideIds(ctx) {
  const s = ctx.presentation.slides
  s.load('items/id')
  await ctx.sync()
  return s.items.map(x => x.id)
}

async function selectedIds(ctx) {
  try {
    requireSet('PowerPointApi', '1.5', 'Slide selection')
    const sel = ctx.presentation.getSelectedSlides()
    sel.load('items/id')
    await ctx.sync()
    return sel.items.map(x => x.id)
  } catch (err) {
    if (err instanceof ExecError) return []
    throw err
  }
}

/** Retry ``fn`` in a fresh PowerPoint.run while PowerPoint for the web answers GeneralException (bounded). Reads only. */
async function onFreshContext(fn, tries = 5, waitMs = 200) {
  let last
  for (let i = 0; i < tries; i++) {
    if (i) await new Promise(r => setTimeout(r, waitMs * i))
    try { return await pptRun(fn) } catch (err) {
      if (!err || err.code !== 'GeneralException') throw err
      last = err
    }
  }
  throw last
}

/** Ids of a slide's title and body placeholders (PowerPointApi 1.8 placeholder types), else nulls. */
async function placeholderIds(ctx, slideId) {
  const shapes = ctx.presentation.slides.getItem(slideId).shapes
  shapes.load('items/id,items/type')
  await ctx.sync()
  const out = { title: null, body: null }
  const placeholders = shapes.items.filter(s => s.type === 'Placeholder')
  let ok18 = true
  try { requireSet('PowerPointApi', '1.8', 'Placeholder types') } catch { ok18 = false }
  if (!ok18 || !placeholders.length) return out
  const fmts = placeholders.map(s => { const f = s.placeholderFormat; f.load('type'); return f })
  await ctx.sync()
  placeholders.forEach((s, i) => {
    const ty = fmts[i].type
    if (!out.title && (ty === 'Title' || ty === 'CenterTitle')) out.title = s.id
    else if (!out.body && (ty === 'Body' || ty === 'Content' || ty === 'Subtitle')) out.body = s.id
  })
  return out
}

/** Resolve slideIndex (or the first selected slide, else slide 0) to { id, index, slide }. */
async function resolveSlide(ctx, slideIndex) {
  const ids = await slideIds(ctx)
  if (!ids.length) throw new ExecError('NotFound', 'The presentation has no slides')
  let index
  if (slideIndex === undefined) {
    const sel = await selectedIds(ctx)
    index = Math.max(0, sel.length ? ids.indexOf(sel[0]) : 0)
  } else {
    index = Number(slideIndex)
    if (!Number.isInteger(index) || index < 0 || index >= ids.length) throw new ExecError('Validation', `slideIndex must be 0..${ids.length - 1}`)
  }
  return { id: ids[index], index, ids, slide: ctx.presentation.slides.getItem(ids[index]) }
}

/** Text of every text-bearing shape on a slide; tolerant of shapes without a text frame. */
async function shapeTexts(ctx, slide) {
  const shapes = slide.shapes
  shapes.load('items/id,items/name,items/type,items/left,items/top,items/width,items/height')
  await ctx.sync()
  const out = shapes.items.map(s => ({ id: s.id, name: s.name, type: s.type, left: s.left, top: s.top, width: s.width, height: s.height, text: null }))
  const textual = shapes.items.map((s, i) => ({ s, i })).filter(x => TEXT_TYPES.has(x.s.type))
  try {
    const ranges = textual.map(x => { const tr = x.s.textFrame.textRange; tr.load('text'); return tr })
    await ctx.sync()
    textual.forEach((x, k) => { out[x.i].text = ranges[k].text })
  } catch {
    for (const x of textual) {
      try {
        const tr = x.s.textFrame.textRange
        tr.load('text')
        await ctx.sync()
        out[x.i].text = tr.text
      } catch { out[x.i].text = null }
    }
  }
  return out
}

function budget(items, max) {
  let used = 0
  const kept = []
  for (const it of items) {
    const size = JSON.stringify(it).length
    if (used + size > max) return { items: kept, truncated: true }
    used += size
    kept.push(it)
  }
  return { items: kept, truncated: false }
}

// PowerPoint for the web gives a new slide a provisional id for a moment, then its real one (probed live, A11:
// '4123571116#123571113' for about 0.4 s, then '259#1316651161'; getItem on the provisional id then fails with
// InvalidParam, and its shapes read back empty). The dev mock sets __hermesPptTiming to keep its tests fast.
// A19 (probed live again): the provisional id now lasts 0.5-2.1 s, longer than three reads 300 ms apart, so a new
// slide's id must also hold still for settleQuietMs of wall time; a write that still meets the provisional id is
// retried on the re-read id (createTextSlide).
const timing = () => ({ settleMs: 300, settleReads: 3, settleMax: 40, settleQuietMs: 2500, ...(globalThis.__hermesPptTiming || {}) })
const sleep = ms => new Promise(r => setTimeout(r, ms))

// A19: errors a fill meets on a new slide's id that was still provisional: the batch fails at its first getItem, so
// nothing in it was written and it can be run again on the id read afresh.
const FILL_RETRIES = 4
const STALE_ID_CODES = new Set(['GeneralException', 'ItemNotFound', 'InvalidArgument', 'InvalidParam'])
function staleSlideId(err) {
  if (!err || !STALE_ID_CODES.has(err.code)) return false
  const loc = err.debugInfo && err.debugInfo.errorLocation
  return !loc || /getItem/.test(String(loc))
}

/** The new slide's settled id: read in fresh contexts until it has not changed for settleReads reads (bounded). */
async function newSlideId(ctx, before) {
  if (ctx) await ctx.sync()
  const t = timing()
  let id = null
  let index = -1
  let same = 0
  let since = 0
  for (let i = 0; i < t.settleMax; i++) {
    if (i) await sleep(t.settleMs)
    const after = await pptRun(c => slideIds(c))
    const added = after.filter(x => !before.includes(x))
    if (added.length !== 1) { id = null; same = 0; continue }
    if (added[0] === id) same++
    else { id = added[0]; same = 1; since = Date.now() }
    index = after.indexOf(id)
    if (same >= t.settleReads && Date.now() - since >= t.settleQuietMs) return { id, index }
  }
  if (id) return { id, index }
  throw new ExecError('ExecutionUncertain', 'Could not identify the new slide')
}

function deleteSlideUndo(id, label) {
  return { undo: async x => { x.presentation.slides.getItem(id).delete(); await x.sync() }, undoLabel: label }
}

function deleteShapeUndo(slideId, shapeId, label) {
  return {
    undo: async x => {
      // Look the shape up by id first: getItem on an id that is gone fails with an opaque InvalidParam (seen live).
      if (!(await slideIds(x)).includes(slideId)) throw new ExecError('NotFound', 'The slide Hermes changed is no longer in the presentation; nothing was removed')
      const shapes = x.presentation.slides.getItem(slideId).shapes
      shapes.load('items/id')
      await x.sync()
      const hit = shapes.items.find(s => s.id === shapeId)
      if (!hit) throw new ExecError('NotFound', `The shape Hermes added (${shapeId}) is no longer on that slide; nothing was removed`)
      hit.delete()
      await x.sync()
    },
    undoLabel: label
  }
}

/**
 * Change highlight (slide creation only): select the new slide so PowerPoint shows it (PowerPointApi 1.5).
 * Tools that default to "the selected slide" then act on that new slide.
 */
function goToSlide(id, index) {
  return ctx => {
    requireSet('PowerPointApi', '1.5', 'Selecting the new slide')
    ctx.presentation.setSelectedSlides([id])
    return `slide ${index}`
  }
}

async function findLayout(ctx, layoutId) {
  const masters = ctx.presentation.slideMasters
  masters.load('items/id,items/name')
  await ctx.sync()
  const probes = masters.items.map(m => { const l = m.layouts.getItemOrNullObject(layoutId); l.load('id,name'); return { m, l } })
  await ctx.sync()
  const hit = probes.find(p => !p.l.isNullObject)
  if (!hit) throw new ExecError('NotFound', `No slide layout with id ${layoutId}`)
  return { layout: hit.l, masterId: hit.m.id }
}

async function defaultLayoutId(ctx, wanted) {
  const masters = ctx.presentation.slideMasters
  masters.load('items/id')
  await ctx.sync()
  if (!masters.items.length) return null
  const layouts = masters.items[0].layouts
  layouts.load('items/id,items/name')
  await ctx.sync()
  for (const name of wanted) {
    const l = layouts.items.find(x => x.name.toLowerCase() === name.toLowerCase())
    if (l) return { layoutId: l.id, slideMasterId: masters.items[0].id }
  }
  return null
}

function base64Kind(b64) {
  const head = b64.slice(0, 16)
  if (head.startsWith('iVBORw0KGgo')) return 'png'
  if (head.startsWith('/9j/')) return 'jpeg'
  return null
}

const tools = {
  async get_presentation_overview(ctx, { maxLength }) {
    const max = cap(maxLength)
    const ids = await slideIds(ctx)
    const sel = await selectedIds(ctx)
    const slides = []
    for (let i = 0; i < ids.length; i++) {
      const shapes = await shapeTexts(ctx, ctx.presentation.slides.getItem(ids[i]))
      const firstText = shapes.find(s => s.text)
      slides.push({ index: i, id: ids[i], shapes: shapes.length, title: firstText ? firstText.text.slice(0, 120) : '' })
    }
    const b = budget(slides, max)
    return { document: docName(), slideCount: ids.length, selectedSlideIndexes: sel.map(id => ids.indexOf(id)).filter(i => i >= 0), slides: b.items, truncated: b.truncated }
  },

  async read_selected_slides(ctx, { maxLength }) {
    const ids = await slideIds(ctx)
    const sel = await selectedIds(ctx)
    const out = []
    for (const id of sel) out.push({ index: ids.indexOf(id), id, shapes: await shapeTexts(ctx, ctx.presentation.slides.getItem(id)) })
    const b = budget(out, cap(maxLength))
    return { slides: b.items, truncated: b.truncated, selectedCount: sel.length }
  },

  async read_all_slides(ctx, { maxLength }) {
    const ids = await slideIds(ctx)
    const out = []
    for (let i = 0; i < ids.length; i++) {
      const shapes = await shapeTexts(ctx, ctx.presentation.slides.getItem(ids[i]))
      out.push({ index: i, id: ids[i], text: shapes.filter(s => s.text).map(s => s.text) })
    }
    const b = budget(out, cap(maxLength))
    return { slideCount: ids.length, slides: b.items, truncated: b.truncated }
  },

  async create_slide(ctx, { layoutId, slideMasterId }) {
    requireSet('PowerPointApi', '1.3', 'Creating slides')
    const before = await slideIds(ctx)
    const opts = {}
    if (layoutId) opts.layoutId = String(layoutId)
    if (slideMasterId) opts.slideMasterId = String(slideMasterId)
    ctx.presentation.slides.add(opts)
    await ctx.sync()
    const { id, index } = await newSlideId(ctx, before)
    return { data: { slideId: id, slideIndex: index }, ...deleteSlideUndo(id, 'create slide'), highlight: goToSlide(id, index) }
  },

  async create_text_slide(ctx, { title, body, bullets, layoutId, slideMasterId }) {
    requireSet('PowerPointApi', '1.4', 'Creating text slides')
    const t = str(title, 'title', 2000)
    const lines = [body ? String(body) : null, ...(Array.isArray(bullets) ? bullets.map(String) : [])].filter(x => x !== null && x !== '')
    const before = await slideIds(ctx)
    let opts = {}
    if (layoutId) opts = { layoutId: String(layoutId), ...(slideMasterId ? { slideMasterId: String(slideMasterId) } : {}) }
    else opts = (await defaultLayoutId(ctx, lines.length ? ['Title and Content', 'Title Only', 'Blank'] : ['Title Only', 'Title and Content', 'Blank'])) || {}
    ctx.presentation.slides.add(opts)
    await ctx.sync()
    let { id, index } = await newSlideId(ctx, before)
    // PowerPoint for the web fails (GeneralException) to load a new slide's shapes in the context that added it
    // (probed live, A9): read and fill them in fresh contexts. A19: an id that turns out provisional (the slide got
    // its real id after the settle) fails the fill at its getItem, before anything is written: the id is read again
    // and the fill retried (bounded). On failure the empty slide is removed again, found by its current id.
    let used
    const fill = async () => {
      const found = await onFreshContext(c => placeholderIds(c, id))
      used = !!found.title
      await pptRun(async c => {
        const slide = c.presentation.slides.getItem(id)
        if (found.title) slide.shapes.getItem(found.title).textFrame.textRange.text = t
        else slide.shapes.addTextBox(t, { left: 40, top: 28, width: 880, height: 72 })
        if (lines.length) {
          if (found.body) slide.shapes.getItem(found.body).textFrame.textRange.text = lines.join('\n')
          else slide.shapes.addTextBox(lines.map((l, i) => (i === 0 && body ? l : `• ${l}`)).join('\n'), { left: 40, top: 120, width: 880, height: 380 })
        }
        await c.sync()
      })
    }
    try {
      for (let attempt = 0; ; attempt++) {
        try { await fill(); break } catch (err) {
          if (attempt >= FILL_RETRIES || !staleSlideId(err)) throw err
          await sleep(timing().settleMs * (attempt + 1))
          ;({ id, index } = await newSlideId(null, before))
        }
      }
      // The id the undo and the highlight use: still the one the fill used, else (a later change) settled again.
      const after = await pptRun(c => slideIds(c))
      const added = after.filter(x => !before.includes(x))
      if (added.length === 1 && added[0] === id) index = after.indexOf(id)
      else ({ id, index } = await newSlideId(null, before))
    } catch (err) {
      try {
        const cur = await newSlideId(null, before)
        await pptRun(async c => { c.presentation.slides.getItem(cur.id).delete(); await c.sync() })
      } catch { /* reported below */ }
      throw err
    }

    return { data: { slideId: id, slideIndex: index, usedPlaceholders: used }, ...deleteSlideUndo(id, 'create text slide'), highlight: goToSlide(id, index) }
  },

  async add_text_box(ctx, { text, slideIndex, left, top, width, height }) {
    requireSet('PowerPointApi', '1.4', 'Text boxes')
    const { id: slideId, index, slide } = await resolveSlide(ctx, slideIndex)
    const shape = slide.shapes.addTextBox(str(text, 'text', 20000), { left: left ?? 60, top: top ?? 60, width: width ?? 400, height: height ?? 60 })
    shape.load('id')
    await ctx.sync()
    return { data: { slideIndex: index, shapeId: shape.id }, ...deleteShapeUndo(slideId, shape.id, 'add text box') }
  },

  async update_slide_text(ctx, { text, slideIndex, shapeId }) {
    const t = String(text ?? '')
    let target = null
    if (shapeId) {
      const { id, slide } = await resolveSlide(ctx, slideIndex)
      target = { slideId: id, shape: slide.shapes.getItem(String(shapeId)) }
    } else {
      requireSet('PowerPointApi', '1.5', 'Selected text')
      const shapes = ctx.presentation.getSelectedShapes()
      shapes.load('items/id')
      const sel = await selectedIds(ctx)
      await ctx.sync()
      if (shapes.items.length === 1 && sel.length) target = { slideId: sel[0], shape: ctx.presentation.slides.getItem(sel[0]).shapes.getItem(shapes.items[0].id) }
      else {
        const tr = ctx.presentation.getSelectedTextRangeOrNullObject()
        tr.load('text')
        await ctx.sync()
        if (tr.isNullObject) throw new ExecError('Validation', 'Select one shape or some text on a slide, or pass slideIndex and shapeId')
        tr.text = t
        await ctx.sync()
        return { data: { updated: 'selection' }, undoUnavailable: 'selected text has no stable anchor to restore; pass shapeId for undoable edits' }
      }
    }
    const tr = target.shape.textFrame.textRange
    tr.load('text')
    target.shape.load('id')
    await ctx.sync()
    const prior = tr.text
    tr.text = t
    await ctx.sync()
    const sid = target.shape.id
    const slideId = target.slideId
    return {
      data: { shapeId: sid, previousLength: prior.length, newLength: t.length },
      undo: async x => { x.presentation.slides.getItem(slideId).shapes.getItem(sid).textFrame.textRange.text = prior; await x.sync() },
      undoLabel: 'update slide text',
      undoPartial: 'text is restored; mixed run formatting inside the shape is not'
    }
  },

  async add_speaker_notes() {
    throw new ExecError('Unsupported', 'PowerPoint Office.js exposes no speaker-notes API in this snapshot; add notes in PowerPoint directly')
  },

  async add_shape(ctx, { shapeType, slideIndex, left, top, width, height }) {
    requireSet('PowerPointApi', '1.4', 'Shapes')
    const { id: slideId, index, slide } = await resolveSlide(ctx, slideIndex)
    const shape = slide.shapes.addGeometricShape(str(shapeType, 'shapeType', 64), { left: left ?? 100, top: top ?? 100, width: width ?? 160, height: height ?? 100 })
    shape.load('id')
    await ctx.sync()
    return { data: { slideIndex: index, shapeId: shape.id, shapeType }, ...deleteShapeUndo(slideId, shape.id, `add ${shapeType}`) }
  },

  async read_slide_content(ctx, { slideIndex, maxLength }) {
    const { id, index, slide } = await resolveSlide(ctx, slideIndex)
    const b = budget((await shapeTexts(ctx, slide)).filter(s => s.text !== null), cap(maxLength))
    return { slideIndex: index, slideId: id, shapes: b.items, truncated: b.truncated }
  },

  async delete_slide(ctx, { slideIndex }) {
    const { id, index, ids, slide } = await resolveSlide(ctx, slideIndex)
    if (ids.length <= 1) throw new ExecError('Conflict', 'Refusing to delete the last remaining slide')
    let exported = null
    try {
      requireSet('PowerPointApi', '1.8', 'Slide export')
      const b = slide.exportAsBase64()
      await ctx.sync()
      exported = b.value
    } catch (err) {
      if (!(err instanceof ExecError)) exported = null
    }
    slide.delete()
    await ctx.sync()
    const data = { deletedIndex: index, slideCount: ids.length - 1 }
    if (!exported) return { data, undoUnavailable: 'slide export (PowerPointApi 1.8) is unavailable, so the slide cannot be restored' }
    const anchor = index > 0 ? ids[index - 1] : ids[1]
    return {
      data,
      undo: async x => {
        const before = await slideIds(x)
        x.presentation.insertSlidesFromBase64(exported, { targetSlideId: anchor, formatting: 'KeepSourceFormatting' })
        await x.sync()
        const { id: nid } = await newSlideId(x, before)
        x.presentation.slides.getItem(nid).moveTo(index)
        await x.sync()
      },
      undoLabel: `delete slide ${index}`,
      undoPartial: 'the slide is re-inserted from an export; its slide id changes'
    }
  },

  async duplicate_slide(ctx, { slideIndex }) {
    requireSet('PowerPointApi', '1.8', 'Slide duplication (export + insert)')
    const { id, index, slide } = await resolveSlide(ctx, slideIndex)
    const before = await slideIds(ctx)
    const b = slide.exportAsBase64()
    await ctx.sync()
    ctx.presentation.insertSlidesFromBase64(b.value, { targetSlideId: id, formatting: 'KeepSourceFormatting' })
    await ctx.sync()
    const { id: nid, index: nindex } = await newSlideId(ctx, before)
    return { data: { sourceIndex: index, newSlideId: nid, newSlideIndex: nindex }, ...deleteSlideUndo(nid, `duplicate slide ${index}`), highlight: goToSlide(nid, nindex) }
  },

  async reorder_slide(ctx, { fromIndex, toIndex }) {
    requireSet('PowerPointApi', '1.8', 'Slide.moveTo')
    const ids = await slideIds(ctx)
    const from = Number(fromIndex); const to = Number(toIndex)
    for (const [n, v] of [['fromIndex', from], ['toIndex', to]]) {
      if (!Number.isInteger(v) || v < 0 || v >= ids.length) throw new ExecError('Validation', `${n} must be 0..${ids.length - 1}`)
    }
    const id = ids[from]
    ctx.presentation.slides.getItem(id).moveTo(to)
    await ctx.sync()
    return { data: { slideId: id, fromIndex: from, toIndex: to }, undo: async x => { x.presentation.slides.getItem(id).moveTo(from); await x.sync() }, undoLabel: `move slide ${from}->${to}` }
  },

  async find_and_replace_text(ctx, { findText, replaceText, slideIndex, shapeId }) {
    const find = str(findText, 'findText', 5000)
    const repl = String(replaceText ?? '')
    const { id, index, slide } = await resolveSlide(ctx, slideIndex)
    const shapes = (await shapeTexts(ctx, slide)).filter(s => s.text !== null && (!shapeId || s.id === String(shapeId)))
    if (shapeId && !shapes.length) throw new ExecError('NotFound', `No text shape ${shapeId} on slide ${index}`)
    const changed = []
    for (const s of shapes) {
      const n = s.text.split(find).length - 1
      if (n > 0) {
        slide.shapes.getItem(s.id).textFrame.textRange.text = s.text.split(find).join(repl)
        changed.push({ id: s.id, prior: s.text, count: n })
      }
    }
    await ctx.sync()
    const data = { slideIndex: index, replacements: changed.reduce((a, c) => a + c.count, 0), shapesChanged: changed.length }
    if (!changed.length) return { data, undoUnavailable: 'nothing was changed' }
    return {
      data,
      undo: async x => {
        const sl = x.presentation.slides.getItem(id)
        for (const c of changed) sl.shapes.getItem(c.id).textFrame.textRange.text = c.prior
        await x.sync()
      },
      undoLabel: `replace "${find}" on slide ${index}`,
      undoPartial: 'text is restored; mixed run formatting inside changed shapes is not'
    }
  },

  async list_slide_layouts(ctx) {
    const masters = ctx.presentation.slideMasters
    masters.load('items/id,items/name')
    await ctx.sync()
    const layouts = masters.items.map(m => { const l = m.layouts; l.load('items/id,items/name'); return l })
    await ctx.sync()
    return { slideMasters: masters.items.map((m, i) => ({ id: m.id, name: m.name, layouts: layouts[i].items.map(l => ({ id: l.id, name: l.name })) })) }
  },

  async set_slide_layout(ctx, { slideIndex, layoutId }) {
    requireSet('PowerPointApi', '1.8', 'Slide.applyLayout')
    const { id, index, slide } = await resolveSlide(ctx, slideIndex)
    const lay = slide.layout
    lay.load('id')
    await ctx.sync()
    const priorId = lay.id
    const { layout } = await findLayout(ctx, str(layoutId, 'layoutId', 255))
    slide.applyLayout(layout)
    await ctx.sync()
    return {
      data: { slideIndex: index, layoutId: layout.id, layoutName: layout.name, previousLayoutId: priorId },
      undo: async x => { const { layout: prev } = await findLayout(x, priorId); x.presentation.slides.getItem(id).applyLayout(prev); await x.sync() },
      undoLabel: `layout of slide ${index}`,
      undoPartial: 'the previous layout is re-applied; placeholder positions follow that layout'
    }
  },

  async format_text(ctx, { slideIndex, shapeId, font }) {
    requireSet('PowerPointApi', '1.4', 'Text formatting')
    const { id, index, slide } = await resolveSlide(ctx, slideIndex)
    const f = font || {}
    const sid = str(shapeId, 'shapeId', 255)
    const target = slide.shapes.getItem(sid).textFrame.textRange.font
    target.load('bold,italic,color,size,name')
    await ctx.sync()
    const prior = { bold: target.bold, italic: target.italic, color: target.color, size: target.size, name: target.name }
    const applied = []
    for (const k of ['bold', 'italic', 'color', 'size', 'name']) if (f[k] !== undefined) { target[k] = f[k]; applied.push(k) }
    if (!applied.length) throw new ExecError('Validation', 'font must set at least one of bold, italic, color, size, name')
    await ctx.sync()
    return {
      data: { slideIndex: index, shapeId: sid, applied },
      undo: async x => {
        const t = x.presentation.slides.getItem(id).shapes.getItem(sid).textFrame.textRange.font
        for (const k of applied) if (prior[k] !== null && prior[k] !== undefined) t[k] = prior[k]
        await x.sync()
      },
      undoLabel: `format text of ${sid}`,
      undoPartial: 'properties that were mixed before the change are left as set'
    }
  },

  async add_image(ctx, { base64Image, slideIndex, left, top, width, height }) {
    if (typeof base64Image !== 'string' || !base64Image) throw new ExecError('Validation', 'base64Image must be a non-empty string')
    const b64 = base64Image.replace(/^data:image\/(png|jpeg);base64,/, '')
    if (b64.length > MAX_IMAGE_BASE64) throw new ExecError('Capacity', 'image is larger than 8 MB (base64)')
    if (!/^[A-Za-z0-9+/]+={0,2}$/.test(b64) || !base64Kind(b64)) throw new ExecError('Validation', 'base64Image must be a base64 PNG or JPEG')
    const r = await placePicture(ctx, b64, { slideIndex, left, top, width, height })
    const data = { slideIndex: r.index, shapeId: r.shapeId, format: base64Kind(b64) }
    if (!r.shapeId) return { data, undoUnavailable: 'could not identify the inserted picture' }
    return { data, ...deleteShapeUndo(r.slideId, r.shapeId, 'add image') }
  },

  // ---- office_insert_image (plugin-level tool; same insert path as add_image, plus alt text) ----
  async insert_image(ctx, { base64, alt, title, slideIndex, left, top, width, height }) {
    const { b64, kind } = imageBase64(base64)
    const description = altText(alt)
    const r = await placePicture(ctx, b64, { slideIndex, left, top, width, height })
    const data = { slideIndex: r.index, slideId: r.slideId, shapeId: r.shapeId, format: kind, altTextApplied: false }
    if (!r.shapeId) return { data, undoUnavailable: 'could not identify the inserted picture' }
    try {
      requireSet('PowerPointApi', '1.10', 'Alt text on shapes')
      const shape = ctx.presentation.slides.getItem(r.slideId).shapes.getItem(r.shapeId)
      shape.altTextDescription = description
      shape.altTextTitle = String(title || description).slice(0, 255)
      await ctx.sync()
      data.altTextApplied = true
    } catch { data.altTextApplied = false }
    return { data, ...deleteShapeUndo(r.slideId, r.shapeId, 'insert image') }
  },

  async image_status(ctx, { slideId, shapeId }) {
    const ids = await slideIds(ctx)
    const sid = str(slideId, 'slideId')
    if (!ids.includes(sid)) return { exists: false }
    const shapes = ctx.presentation.slides.getItem(sid).shapes
    shapes.load('items/id,items/type')
    await ctx.sync()
    const hit = shapes.items.find(s => s.id === str(shapeId, 'shapeId'))
    return { exists: !!hit, ...(hit ? { type: hit.type } : {}) }
  }
}

/** Insert a picture on a slide through the Common API (ImageCoercion 1.1); returns the new shape's ids.
 *
 * Seen live in PowerPoint for the web (A9): (1) without imageLeft/imageTop it writes INTO the current selection, so
 * with a shape still selected (the previous picture is) it fails "Cannot write to the current selection"; a position
 * is therefore always passed. (2) The new picture reads back with id "0" for a moment before it gets its real id,
 * so the new shape is identified by diffing every slide's shapes until it has a real id (selection breaks a tie). */
async function placePicture(ctx, b64, { slideIndex, left, top, width, height }) {
  const doc = globalThis.Office && globalThis.Office.context && globalThis.Office.context.document
  if (!doc || typeof doc.setSelectedDataAsync !== 'function') throw new ExecError('Unsupported', 'This PowerPoint does not support inserting images from an add-in')
  requireSet('PowerPointApi', '1.5', 'Selecting the target slide')
  const { id, index } = await resolveSlide(ctx, slideIndex)
  const before = await shapesBySlide(ctx)
  ctx.presentation.setSelectedSlides([id])
  await ctx.sync()
  const pos = await defaultPosition(ctx)
  const opts = { coercionType: 'image', imageLeft: left !== undefined ? left : pos.left, imageTop: top !== undefined ? top : pos.top }
  if (width !== undefined) opts.imageWidth = width
  if (height !== undefined) opts.imageHeight = height
  await new Promise((resolve, reject) => doc.setSelectedDataAsync(b64, opts, r => (r.status === 'succeeded' ? resolve() : reject(new ExecError('Office', (r.error && r.error.message) || 'Image insert failed', (r.error && String(r.error.code)) || 'InsertFailed')))))
  for (let i = 0; i < PICTURE_ID_TRIES; i++) {
    const hit = await newPicture(ctx, id, before)
    if (hit) return hit.slideId === id ? { slideId: id, index, shapeId: hit.shapeId } : { ...hit }
    await new Promise(r => setTimeout(r, PICTURE_ID_WAIT_MS))
  }
  return { slideId: id, index, shapeId: null }
}

const PICTURE_ID_TRIES = 20
const PICTURE_ID_WAIT_MS = 150

async function shapesBySlide(ctx) {
  const ids = await slideIds(ctx)
  const cols = ids.map(sid => { const c = ctx.presentation.slides.getItem(sid).shapes; c.load('items/id'); return c })
  await ctx.sync()
  return new Map(ids.map((sid, i) => [sid, new Set(cols[i].items.map(s => s.id))]))
}

async function defaultPosition(ctx) {
  try {
    requireSet('PowerPointApi', '1.10', 'Slide size')
    const ps = ctx.presentation.pageSetup
    ps.load('slideWidth,slideHeight')
    await ctx.sync()
    return { left: Math.round(ps.slideWidth / 6), top: Math.round(ps.slideHeight / 6) }
  } catch {
    return { left: 72, top: 72 }
  }
}

/** The one new picture with a real id (target slide first, then any other slide), or null while it is not known yet. */
async function newPicture(ctx, targetId, before) {
  const ids = await slideIds(ctx)
  const order = [targetId, ...ids.filter(x => x !== targetId)].filter(x => ids.includes(x))
  const cols = order.map(sid => { const c = ctx.presentation.slides.getItem(sid).shapes; c.load('items/id,items/type'); return c })
  await ctx.sync()
  let selected = null
  for (let k = 0; k < order.length; k++) {
    const had = before.get(order[k]) || new Set()
    const added = cols[k].items.filter(s => !had.has(s.id))
    if (!added.length) continue
    if (added.some(s => s.id === '0')) return null // PowerPoint has not assigned the real id yet
    const pics = added.filter(s => s.type === 'Image')
    let pick = pics.length === 1 ? pics[0] : null
    if (!pick && pics.length > 1) {
      if (!selected) {
        const sel = ctx.presentation.getSelectedShapes()
        sel.load('items/id')
        await ctx.sync()
        selected = new Set(sel.items.map(s => s.id))
      }
      const both = pics.filter(s => selected.has(s.id))
      pick = both.length === 1 ? both[0] : null
    }
    if (pick) return { slideId: order[k], index: ids.indexOf(order[k]), shapeId: pick.id }
    if (pics.length) return null
  }
  return null
}

export const MUTATING = new Set(['create_slide', 'create_text_slide', 'add_text_box', 'update_slide_text', 'add_speaker_notes',
  'add_shape', 'delete_slide', 'duplicate_slide', 'reorder_slide', 'find_and_replace_text', 'set_slide_layout', 'format_text', 'add_image',
  'insert_image'])

export const executor = createExecutor({ tools, mutating: MUTATING, run: pptRun })
export const { execute, SUPPORTED_TOOLS } = executor
