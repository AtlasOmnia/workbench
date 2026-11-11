// SPDX-License-Identifier: Apache-2.0
// OneNote (OneNoteApi 1.1, OneNote for the web only) executor for the Hermes bridge task pane.
//
// Ground rules that come straight from the OneNote JS API, enforced here:
//  - Page CONTENT is readable/writable only for the ACTIVE page. Titles, ids and urls of every
//    page are readable without navigating. get_page therefore never reads another page; the
//    agent opens it first (open_page), which the user sees happen.
//  - Hermes never rewrites the user's text (RichText.text is read-only anyway). Writes are new
//    pages and new outlines/paragraphs that Hermes owns; the plugin's ledger records their ids,
//    and delete_owned / replace_owned only remove ids the plugin owns whose content still has the
//    fingerprint recorded at write time (an edited block is kept; links blocks also re-check the marker).
//  - There is no Page.delete: a page Hermes created can only be deleted by the user.
//  - Every mutation carries the plugin's opId and the paired notebook id. The notebook id is
//    re-checked right before writing (zero mutation on mismatch); the opId journal (tab-scoped,
//    survives a pane reload) answers op_status after a lost result, and refuses to run the same
//    opId twice. Nothing here retries.
import { ExecError, officeError, str } from '../shared/runtime.js'
import { altText, pictureHtml } from '../shared/images.js'

export const MARKER = 'Links & Backlinks'
const MAX_HTML = 60000
const MAX_TEXT = 50000
const MAX_PAGES = 2000
const MAX_LEVEL = 64 // nesting levels read below an outline's top-level paragraphs (list items, indented lines); a
// paragraph still deeper marks its outline `incomplete` (A42): never fingerprinted, deleted or read as "absent"
const JOURNAL_KEY = 'hermes-onenote-ops'
const JOURNAL_MAX = 60
const NAV_POLL_MS = 400
let listBudgetMs = 20000 // list_pages answers within this, well inside Hermes' 60 s wait (A27)
export function setListBudgetMs(ms) { listBudgetMs = ms }
const DUP_TITLE_MAX = 20 // "<title> (2)" .. "<title> (20)" for a re-run's decision page (A27)

function oneRun(fn) {
  const o = globalThis.OneNote
  if (!o || typeof o.run !== 'function') throw new ExecError('Unavailable', 'OneNote API is not available in this pane')
  return o.run(fn)
}

// ---- opId journal ----------------------------------------------------------------------------
const memJournal = new Map()
function journalStore() { try { return globalThis.sessionStorage || null } catch { return null } }
function readJournal() {
  const s = journalStore()
  if (!s) return memJournal
  try { return new Map(JSON.parse(s.getItem(JOURNAL_KEY) || '[]')) } catch { return new Map() }
}
function writeJournal(map) {
  while (map.size > JOURNAL_MAX) map.delete(map.keys().next().value)
  const s = journalStore()
  if (!s) return
  try { s.setItem(JOURNAL_KEY, JSON.stringify([...map])) } catch { /* storage full/disabled: op_status degrades to unknown */ }
}
function journalSet(opId, entry) { const j = readJournal(); j.set(opId, entry); writeJournal(j); if (j !== memJournal) memJournal.set(opId, entry) }
function journalGet(opId) { return readJournal().get(opId) || memJournal.get(opId) || null }
export function clearJournal() { memJournal.clear(); const s = journalStore(); if (s) try { s.removeItem(JOURNAL_KEY) } catch { /* ignore */ } }

// ---- helpers -----------------------------------------------------------------------------------
function html(v, name = 'html') {
  if (typeof v !== 'string' || !v.trim() || v.length > MAX_HTML) throw new ExecError('Validation', `${name} must be non-empty HTML up to ${MAX_HTML} characters`)
  // OneNote ignores scripts, but never pass them: the plugin builds HTML from untrusted note text.
  if (/<\s*(script|iframe|object|embed|style)\b|\son[a-z]+\s*=|javascript:/i.test(v)) throw new ExecError('Validation', `${name} contains markup OneNote add-ins must not insert`)
  return v
}

async function activeIds(ctx) {
  const nb = ctx.application.getActiveNotebook()
  nb.load('id,name')
  const page = ctx.application.getActivePageOrNull()
  page.load('id,title,webUrl')
  const section = ctx.application.getActiveSection()
  section.load('id,name')
  await ctx.sync()
  return { notebook: { id: nb.id, name: nb.name }, section: { id: section.id, name: section.name }, page: page.isNullObject ? null : { id: page.id, title: page.title, webUrl: page.webUrl }, pageObj: page }
}

async function expectNotebook(ctx, expected) {
  const nb = ctx.application.getActiveNotebook()
  nb.load('id')
  await ctx.sync()
  if (expected && nb.id !== expected) throw new ExecError('Conflict', 'IdentityMismatch: the pane now shows a different notebook than the one paired with this chat; nothing was changed', 'IdentityMismatch')
  return nb.id
}

/** Every section in the notebook (section groups walked, depth-bounded) with its path. */
async function allSections(ctx) {
  const nb = ctx.application.getActiveNotebook()
  const out = []
  async function walk(container, prefix, depth) {
    const secs = container.sections
    secs.load('items/id,items/name')
    const groups = container.sectionGroups
    groups.load('items/id,items/name')
    await ctx.sync()
    for (const s of secs.items) out.push({ obj: s, id: s.id, name: s.name, path: [...prefix, s.name].join(' / ') })
    if (depth < 4) for (const g of groups.items) await walk(g, [...prefix, g.name], depth + 1)
  }
  await walk(nb, [], 0)
  // Live OneNote for the web can report the SAME Section.id for several sections of one notebook
  // (seen after a reload). Give each such section a stable id of its own ("<id>#<path>") so every
  // sectionId Hermes is shown names exactly one section; `raw` keeps what OneNote reported.
  const counts = new Map()
  for (const s of out) counts.set(s.id, (counts.get(s.id) || 0) + 1)
  for (const s of out) { s.raw = s.id; if (counts.get(s.id) > 1) s.id = `${s.id}#${s.path}` }
  return out
}

/** The section a sectionId (or, without one, the open section) names; never a guess among look-alikes. */
function resolveSection(secs, sectionId, active) {
  if (sectionId) {
    const exact = secs.find(s => s.id === sectionId)
    if (exact) return exact
    if (secs.filter(s => s.raw === sectionId).length > 1) {
      throw new ExecError('Validation', 'That sectionId is shared by several sections in OneNote for the web; use an id from onenote_list_sections')
    }
    return secs.find(s => s.raw === sectionId) || null
  }
  const same = secs.filter(s => s.raw === active.id)
  return same.length > 1 ? (same.find(s => s.name === active.name) || null) : (same[0] || null)
}

async function pagesOf(ctx, section) {
  const pages = section.obj.pages
  pages.load('items/id,items/title,items/webUrl,items/clientUrl,items/pageLevel')
  await ctx.sync()
  return pages.items
}

async function findPage(ctx, pageId) {
  for (const s of await allSections(ctx)) {
    let pages
    try { pages = await pagesOf(ctx, s) } catch { continue }
    const p = pages.find(x => x.id === pageId)
    if (p) return { page: p, section: s }
  }
  return null
}

function pageInfo(p, s) {
  return { id: p.id, title: p.title, webUrl: p.webUrl, clientUrl: p.clientUrl, level: p.pageLevel, sectionId: s.id, sectionPath: s.path }
}

/** Every outline of the page with its paragraphs' type/text, nested paragraphs flattened in page order with `level`
 *  (table text row -> cell -> paragraphs). An outline with paragraphs this read did not follow (deeper than MAX_LEVEL,
 *  or under a table cell's paragraph) is marked `incomplete: true`, and so is the result (A42 P0-2). */
async function readOutlines(ctx, pageObj) {
  const contents = pageObj.contents
  contents.load('items/id,items/type,items/left,items/top')
  await ctx.sync()
  const outlines = []
  const omitted = { images: 0, ink: 0, other: 0 }
  const pending = []
  for (const c of contents.items) {
    if (c.type !== 'Outline') { if (c.type === 'Image') omitted.images++; else if (c.type === 'Ink') omitted.ink++; else omitted.other++; continue }
    const o = c.outline
    o.load('id')
    const paras = o.paragraphs
    paras.load('items/id,items/type')
    pending.push({ c, o, paras })
  }
  await ctx.sync()
  const reads = []
  const tables = []
  let below = [] // F8 D1: list items are CHILD paragraphs of the paragraph above the list (OneNote for the web)
  const visit = (p, level, into, entry) => {
    const row = level ? { id: p.id, type: p.type, level } : { id: p.id, type: p.type }
    into.push(row)
    if (p.type === 'RichText') {
      const rt = p.richText
      rt.load('text')
      reads.push({ row, rt, h: rt.getHtml() })
      const kids = p.paragraphs // loaded at every level: past MAX_LEVEL only to see whether anything is there
      kids.load('items/id,items/type')
      below.push({ row, kids, level, entry })
    } else if (p.type === 'Table') {
      const t = p.table
      t.load('rowCount,columnCount')
      const rows = t.rows
      rows.load('items/id')
      tables.push({ row, t, rows, entry })
    } else if (p.type === 'Image') omitted.images++
    else if (p.type === 'Ink') omitted.ink++
    else omitted.other++
  }
  for (const { c, o, paras } of pending) {
    const entry = { id: o.id, contentId: c.id, left: c.left, top: c.top, paragraphs: [] }
    outlines.push(entry)
    for (const p of paras.items) visit(p, 0, entry.paragraphs, entry)
  }
  await ctx.sync()
  while (below.length) { // one batch per nesting level
    const level = below
    below = []
    for (const b of level) {
      b.row.children = []
      if (!b.kids.items.length) continue
      if (b.level >= MAX_LEVEL) { b.entry.incomplete = true; continue } // never silently cut: the read says so
      for (const k of b.kids.items) visit(k, b.level + 1, b.row.children, b.entry)
    }
    await ctx.sync()
  }
  // Page order: each paragraph, then the paragraphs under it (depth-first), flat; `level` says how deep (absent = 0).
  const flat = rows => rows.flatMap(r => { const kids = r.children || []; delete r.children; return [r, ...flat(kids)] })
  for (const entry of outlines) entry.paragraphs = flat(entry.paragraphs)
  for (const r of reads) { r.row.text = r.rt.text; r.row.html = r.h.value }
  // OneNote.Table exposes no values: walk rows -> cells -> paragraphs, one batch per level.
  const cellsOf = tables.map(tb => tb.rows.items.map(rw => { const cs = rw.cells; cs.load('items/id'); return cs }))
  if (tables.length) await ctx.sync()
  const parasOf = cellsOf.map(rs => rs.map(cs => cs.items.map(cell => { const ps = cell.paragraphs; ps.load('items/type'); return ps })))
  if (tables.length) await ctx.sync()
  const cellKids = []
  const textsOf = parasOf.map((rs, i) => rs.map(cs => cs.map(ps => ps.items.filter(x => x.type === 'RichText').map(x => {
    const rt = x.richText
    rt.load('text')
    const kids = x.paragraphs // a cell's nested paragraphs are not read: their presence makes the outline incomplete
    kids.load('items/id')
    cellKids.push({ kids, entry: tables[i].entry })
    return rt
  }))))
  if (tables.length) await ctx.sync()
  for (const { kids, entry } of cellKids) if (kids.items.length) entry.incomplete = true
  tables.forEach((tb, i) => {
    tb.row.rows = textsOf[i].map(cs => cs.map(rts => rts.map(rt => rt.text).join('\n')))
    tb.row.text = tb.row.rows.map(x => x.join('\t')).join('\n')
  })
  for (const o of outlines) if (!o.incomplete) delete o.incomplete
  return { outlines, omitted, incomplete: outlines.some(o => o.incomplete) }
}

// ---- integrity fingerprints (A6 #6) ------------------------------------------------------------
// sha256 of the outline's canonical text: per paragraph "type<US>text" (NFC, whitespace collapsed),
// joined by <RS>. tests/test_onenote.py and brain.outline_fingerprint compute the same value. The plugin
// stores it in its ledger at write time; a delete/replace runs only when the outline still matches,
// so a block a person edited (even with the 'Links & Backlinks' heading intact) is never removed.
const WS = /[\t\n\v\f\r \u00a0\u1680\u2000-\u200a\u2028\u2029\u202f\u205f\u3000\ufeff]+/
export async function outlineFingerprint(paragraphs) {
  const canon = (paragraphs || []).map(p => `${(p && p.type) || 'RichText'}\u001f${String((p && p.text) || '').normalize('NFC').split(WS).filter(Boolean).join(' ')}`).join('\u001e')
  const bytes = await globalThis.crypto.subtle.digest('SHA-256', new TextEncoder().encode(canon))
  return [...new Uint8Array(bytes)].map(b => b.toString(16).padStart(2, '0')).join('')
}

/** id -> fingerprint for the page's outlines and their paragraphs. A paragraph's covers it and the paragraphs
 *  under it (deleting it deletes them; F8): rows that follow it with a deeper `level`. brain.page_fingerprints agrees.
 *  An incomplete outline (A42 P0-2) gets no fingerprints: its id and its paragraphs' ids are in `incomplete`. */
async function fingerprintsOf(ctx, pageObj) {
  const { outlines } = await readOutlines(ctx, pageObj)
  const out = new Map()
  out.incomplete = new Set()
  for (const o of outlines) {
    if (o.incomplete) { out.incomplete.add(o.id); for (const p of o.paragraphs) out.incomplete.add(p.id); continue }
    out.set(o.id, await outlineFingerprint(o.paragraphs))
    const ps = o.paragraphs
    for (let i = 0; i < ps.length; i++) {
      let j = i + 1
      while (j < ps.length && (ps[j].level || 0) > (ps[i].level || 0)) j++
      out.set(ps[i].id, await outlineFingerprint(ps.slice(i, j)))
    }
  }
  return out
}

/** A42 P0-2: a delete takes everything under the block with it, so it needs a read of all of it; nothing is changed. */
function refuseIncomplete(fps, ids) {
  if (ids.some(id => fps.incomplete.has(id))) {
    throw new ExecError('Conflict', 'Part of this block is nested deeper than Hermes can read, so Hermes cannot check that ' +
      'it is unchanged; nothing was deleted or written. If it is unwanted, delete it by hand in OneNote.', 'IncompleteRead')
  }
}

/** Ids of the outlines on the page now (for the opId journal, before a write adds one). */
async function outlineIdsOn(ctx, pageObj) {
  const contents = pageObj.contents
  contents.load('items/id,items/type')
  await ctx.sync()
  const list = contents.items.filter(c => c.type === 'Outline').map(c => { const o = c.outline; o.load('id'); return o })
  await ctx.sync()
  return list.map(o => o.id)
}

// The page proxy a write used has a stale `contents` once loaded: OneNote for the web does not add outlines
// created since to that proxy's collection (observed live, A8), so the written ids were never found and the
// ledger got no fingerprints (every later refresh/undo then kept Hermes' own block as "edited"). Read through
// a fresh proxy of the active page (the write already checked the active page is the target).
async function writtenFingerprints(ctx, _pageObj, ids) {
  const all = await fingerprintsOf(ctx, ctx.application.getActivePage())
  const out = {}
  for (const id of ids) if (all.has(id)) out[id] = all.get(id)
  return out
}

/** Read the ACTIVE page: outlines, paragraphs (nested ones such as list items, in page order) + table cells, text and HTML. */
async function readActive(ctx, maxChars = MAX_TEXT) {
  const { notebook, section, page, pageObj } = await activeIds(ctx)
  if (!page) throw new ExecError('NotFound', 'No page is open in OneNote')
  const { outlines, omitted, incomplete } = await readOutlines(ctx, pageObj)
  let text = outlines.map(o => o.paragraphs.map(p => p.text || '').join('\n')).join('\n\n')
  const cut = text.length > maxChars
  if (cut) text = text.slice(0, maxChars)
  const truncated = cut || incomplete // A42: a read that missed nested paragraphs is never presented as whole
  // Same section id list_sections shows (OneNote for the web can report one raw id for several sections).
  const own = resolveSection(await allSections(ctx), null, section)
  return { notebookId: notebook.id, notebookName: notebook.name, sectionId: own ? own.id : section.id, sectionName: section.name, pageId: page.id, title: page.title, webUrl: page.webUrl, outlines, text, truncated, omitted, ...(incomplete ? { incomplete } : {}) }
}

/** Position for a new outline: under the lowest existing content (OneNote reports no heights). */
async function belowContent(ctx, pageObj) {
  const contents = pageObj.contents
  contents.load('items/top,items/type')
  await ctx.sync()
  const tops = contents.items.map(c => c.top).filter(Number.isFinite)
  return tops.length ? Math.max(...tops) + 240 : 90
}

async function requireActive(ctx, pageId) {
  const page = ctx.application.getActivePageOrNull()
  page.load('id,title,webUrl')
  await ctx.sync()
  if (page.isNullObject) throw new ExecError('NotFound', 'No page is open in OneNote')
  if (pageId && page.id !== pageId) throw new ExecError('Conflict', 'NotActivePage: OneNote add-ins can only read or write the page that is open; open it first (onenote_open_page)', 'NotActivePage')
  return page
}

async function outlinesById(ctx, pageObj) {
  const contents = pageObj.contents
  contents.load('items/id,items/type')
  await ctx.sync()
  const list = contents.items.filter(c => c.type === 'Outline').map(c => { const o = c.outline; o.load('id'); const ps = o.paragraphs; ps.load('items/id,items/type'); return { c, o, ps } })
  await ctx.sync()
  const firsts = list.map(x => { const p = x.ps.items[0]; if (p && p.type === 'RichText') { const rt = p.richText; rt.load('text'); return rt } return null })
  await ctx.sync()
  return new Map(list.map((x, i) => [x.o.id, { ...x, firstText: firsts[i] ? firsts[i].text : '' }]))
}

// ---- tools -------------------------------------------------------------------------------------
const tools = {
  async get_context(ctx) {
    const { notebook, section, page } = await activeIds(ctx)
    const own = resolveSection(await allSections(ctx), null, section)
    return { notebook, section: own ? { id: own.id, name: own.name } : section, page }
  },

  async list_sections(ctx) {
    const secs = await allSections(ctx)
    return { sections: secs.map(s => ({ id: s.id, name: s.name, path: s.path })) }
  },

  async list_pages(ctx, { sectionId, all, maxPages }) {
    const limit = Math.min(MAX_PAGES, Math.max(1, Math.floor(maxPages || MAX_PAGES)))
    const { notebook, section } = await activeIds(ctx)
    const secs = await allSections(ctx)
    const one = all ? null : resolveSection(secs, sectionId, section)
    const targets = all ? secs : (one ? [one] : [])
    if (!all && !targets.length) throw new ExecError('NotFound', 'No such section in this notebook')
    const pages = []
    const skipped = []
    let truncated = false
    let stoppedForTime = false
    let visited = 0
    const started = Date.now()
    for (const s of targets) {
      // A27: a notebook with many sections took longer than Hermes waits (60 s); while the pane was busy it could
      // take no other command. Answer within the budget with what was read, marked incomplete.
      if (Date.now() - started > listBudgetMs) { truncated = true; stoppedForTime = true; break }
      visited++
      let items
      try { items = await pagesOf(ctx, s) } catch (err) { skipped.push({ id: s.id, path: s.path, reason: officeError(err).code || 'unreadable' }); continue }
      for (const p of items) {
        if (pages.length >= limit) { truncated = true; break }
        pages.push(pageInfo(p, s))
      }
      if (truncated) break
    }
    const out = { notebookId: notebook.id, pages, sectionsRead: (stoppedForTime ? visited : targets.length) - skipped.length, skipped, complete: !truncated && !skipped.length, truncated }
    if (stoppedForTime) out.stoppedEarly = `time: stopped after ${Math.round(listBudgetMs / 1000)} s having read ${visited} of ${targets.length} sections; list one section (sectionId) instead of all`
    return out
  },

  async get_page(ctx, { pageId, maxChars }) {
    if (pageId) await requireActive(ctx, pageId)
    return readActive(ctx, maxChars ? Math.min(MAX_TEXT, Math.max(1, Math.floor(maxChars))) : MAX_TEXT)
  },

  async open_page(ctx, { pageId }) {
    str(pageId, 'pageId', 300)
    const before = await activeIds(ctx)
    const hit = await findPage(ctx, pageId)
    if (!hit) throw new ExecError('NotFound', 'No page with that id in the open notebook')
    ctx.application.navigateToPage(hit.page)
    let navError = null
    try { await ctx.sync() } catch (err) { navError = err }
    // OneNote for the web opens a page in ANOTHER section but fails the navigating batch
    // (GeneralException): check what is actually open, in fresh batches, before reporting.
    // A27: a navigation that did not report an error can still land a moment later, so poll either way.
    let openId = null
    let probeError = null
    for (let i = 0; i < 6 && openId !== pageId; i++) {
      if (i) await new Promise(r => setTimeout(r, NAV_POLL_MS))
      try {
        openId = await oneRun(async c => { const p = c.application.getActivePageOrNull(); p.load('id'); await c.sync(); return p.isNullObject ? null : p.id })
        probeError = null
      } catch (err) { openId = null; probeError = err }
    }
    if (openId !== pageId) {
      if (navError) throw navError
      if (probeError) throw probeError // every check failed: execute() tells a lost OneNote context from a refusal
      throw new ExecError('Unavailable', 'OneNote did not open the page (another page is still open); nothing was changed')
    }
    return { ...pageInfo(hit.page, hit.section), previousPageId: before.page ? before.page.id : null, ...(navError ? { navigationReportedError: true } : {}) }
  },

  async create_page(ctx, { title, html: body, sectionId, allowDuplicateTitle, onDuplicateTitle }, note) {
    const requested = str(title, 'title', 255)
    let t = requested
    const content = body === undefined ? null : html(body)
    const { section } = await activeIds(ctx)
    const secs = await allSections(ctx)
    const target = resolveSection(secs, sectionId, section)
    if (!target) throw new ExecError('NotFound', 'No such section in this notebook')
    const existing = await pagesOf(ctx, target)
    const key = s => String(s || '').trim().toLowerCase()
    const taken = new Set(existing.map(p => key(p.title)))
    const dup = existing.find(p => key(p.title) === key(t))
    let titleAdjusted = null
    if (dup && !allowDuplicateTitle && onDuplicateTitle === 'next') {
      // A27: a re-run of the demo (same email, same day) meets the page an earlier run made. Pick the next free
      // title in this section ("<title> (2)", "(3)" ...); the existing page is never opened, read or changed.
      for (let i = 2; i <= DUP_TITLE_MAX && !titleAdjusted; i++) {
        const suffix = ` (${i})`
        const cand = requested.slice(0, 255 - suffix.length).trimEnd() + suffix
        if (!taken.has(key(cand))) titleAdjusted = { requested, existingPageId: dup.id, existingTitle: dup.title, n: i, cand }
      }
      if (!titleAdjusted) throw new ExecError('Conflict', `DuplicateTitle: "${dup.title}" and the numbered titles up to (${DUP_TITLE_MAX}) already exist in ${target.path}; nothing was created. Pass a different title.`, 'DuplicateTitle')
      t = titleAdjusted.cand
    } else if (dup && !allowDuplicateTitle) throw new ExecError('Conflict', `DuplicateTitle: "${dup.title}" already exists in ${target.path} (pageId ${dup.id}); open it instead, or set allowDuplicateTitle`, 'DuplicateTitle')
    const before = await activeIds(ctx)
    const page = target.obj.addPage(t)
    page.load('id,title,webUrl,clientUrl,pageLevel')
    await ctx.sync()
    note({ created: { pageId: page.id } }) // F8 D2: if the answer is lost from here on, op_status still names the page
    // Content can only be written to the active page: open the new page first.
    ctx.application.navigateToPage(page)
    await ctx.sync()
    let outlineId = null
    let partial = null
    if (content) {
      try {
        const o = page.addOutline(40, 90, content)
        o.load('id')
        await ctx.sync()
        outlineId = o.id
      } catch (err) {
        // The page exists now and cannot be deleted by an add-in: report it, never hide it.
        partial = `The page was created but its content could not be added (${officeError(err).code}).`
      }
    }
    const adjusted = titleAdjusted ? { titleAdjusted: { requested, used: t, existingPageId: titleAdjusted.existingPageId, existingTitle: titleAdjusted.existingTitle } } : {}
    return { ...pageInfo(page, target), outlineId, partial, previousPageId: before.page ? before.page.id : null, ...adjusted, created: { pageId: page.id, outlineIds: outlineId ? [outlineId] : [] } }
  },

  async append(ctx, { html: body, pageId, outlineId, left, top }, note) {
    const content = html(body)
    const page = await requireActive(ctx, pageId)
    if (outlineId) {
      const map = await outlinesById(ctx, page)
      const hit = map.get(outlineId)
      if (!hit) throw new ExecError('NotFound', 'No outline with that id on the open page')
      const beforeIds = new Set(hit.ps.items.map(p => p.id))
      hit.o.appendHtml(content)
      await ctx.sync()
      const ps = hit.o.paragraphs
      ps.load('items/id')
      await ctx.sync()
      const added = ps.items.map(p => p.id).filter(id => !beforeIds.has(id))
      return { pageId: page.id, outlineId, paragraphIds: added, created: { paragraphIds: added, fingerprints: await writtenFingerprints(ctx, page, added) } }
    }
    note({ before: { pageId: page.id, outlineIds: await outlineIdsOn(ctx, page) } })
    const y = Number.isFinite(top) ? top : await belowContent(ctx, page)
    const o = page.addOutline(Number.isFinite(left) ? left : 40, y, content)
    o.load('id')
    await ctx.sync()
    note({ created: { outlineIds: [o.id] } })
    return { pageId: page.id, outlineId: o.id, created: { outlineIds: [o.id], fingerprints: await writtenFingerprints(ctx, page, [o.id]) } }
  },

  async insert_table(ctx, { rows, title, pageId }, note) {
    if (!Array.isArray(rows) || !rows.length || rows.length > 100 || !rows.every(r => Array.isArray(r) && r.length === rows[0].length && r.length >= 1 && r.length <= 20)) {
      throw new ExecError('Validation', 'rows must be a rectangular 1-100 x 1-20 array of strings')
    }
    const values = rows.map(r => r.map(v => String(v ?? '').slice(0, 2000)))
    const page = await requireActive(ctx, pageId)
    note({ before: { pageId: page.id, outlineIds: await outlineIdsOn(ctx, page) } })
    const y = await belowContent(ctx, page)
    const heading = `<p><b>${escapeHtml(title ? String(title).slice(0, 200) : 'Table')}</b></p>`
    const o = page.addOutline(40, y, heading)
    o.load('id')
    await ctx.sync()
    note({ created: { outlineIds: [o.id] } })
    o.appendTable(values.length, values[0].length, values)
    await ctx.sync()
    return { pageId: page.id, outlineId: o.id, rows: values.length, columns: values[0].length, created: { outlineIds: [o.id], fingerprints: await writtenFingerprints(ctx, page, [o.id]) } }
  },

  // ---- office_insert_image (plugin-level tool): a Hermes-owned outline holding the picture ----
  async insert_image(ctx, { html: body, pageId, left, top, alt }) {
    const content = pictureHtml(body)
    const page = await requireActive(ctx, pageId)
    const y = Number.isFinite(top) ? top : await belowContent(ctx, page)
    const o = page.addOutline(Number.isFinite(left) ? left : 40, y, content)
    o.load('id')
    const ps = o.paragraphs
    ps.load('items/id,items/type')
    await ctx.sync()
    const pictures = ps.items.filter(p => p.type === 'Image')
    let altTextApplied = null
    if (alt !== undefined && pictures.length) {
      // OneNote.Image.description is the picture's alternative text (OneNoteApi 1.1).
      try { const d = altText(alt); for (const p of pictures) p.image.description = d; await ctx.sync(); altTextApplied = true } catch { altTextApplied = false }
    }
    return { pageId: page.id, outlineId: o.id, images: pictures.length, altTextApplied, created: { outlineIds: [o.id], fingerprints: await writtenFingerprints(ctx, page, [o.id]) } }
  },

  async image_status(ctx, { pageId, outlineId }) {
    const page = await requireActive(ctx, str(pageId, 'pageId', 300))
    const hit = (await outlinesById(ctx, page)).get(str(outlineId, 'outlineId', 300))
    return { exists: !!hit, images: hit ? hit.ps.items.filter(p => p.type === 'Image').length : 0 }
  },

  // ---- internal commands (plugin only; not model tools) ----
  // Ownership is proven twice: the plugin's ledger lists the id, and the content still has the
  // fingerprint recorded when Hermes wrote it (expect: id -> sha256). An id without an expected
  // fingerprint, or whose content changed (a person edited it), is kept and reported in notOwned.
  async delete_owned(ctx, { pageId, outlineIds, paragraphIds, expect }) {
    const page = await requireActive(ctx, str(pageId, 'pageId', 300))
    const map = await outlinesById(ctx, page)
    const fps = await fingerprintsOf(ctx, page)
    refuseIncomplete(fps, [...(outlineIds || []), ...(paragraphIds || [])])
    const want = expect && typeof expect === 'object' ? expect : {}
    const matches = id => typeof want[id] === 'string' && fps.get(id) === want[id]
    const deleted = []
    const missing = []
    const notOwned = []
    for (const id of outlineIds || []) {
      const hit = map.get(id)
      if (!hit) { missing.push(id); continue }
      if (!matches(id)) { notOwned.push(id); continue }
      hit.c.delete()
      deleted.push(id)
    }
    const wanted = new Set(paragraphIds || [])
    if (wanted.size) {
      for (const { ps } of map.values()) {
        for (const p of ps.items) {
          if (!wanted.has(p.id)) continue
          wanted.delete(p.id)
          if (!matches(p.id)) { notOwned.push(p.id); continue }
          p.delete()
          deleted.push(p.id)
        }
      }
      missing.push(...wanted)
    }
    await ctx.sync()
    return { pageId: page.id, deleted, missing, notOwned }
  },

  async replace_owned(ctx, { pageId, deleteOutlineIds, expect, html: body, marker, left, top }) {
    const content = html(body)
    const mark = str(marker || MARKER, 'marker', 200)
    if (!content.includes(escapeHtml(mark))) throw new ExecError('Validation', 'replacement HTML must carry the Hermes marker')
    const page = await requireActive(ctx, str(pageId, 'pageId', 300))
    const map = await outlinesById(ctx, page)
    const fps = await fingerprintsOf(ctx, page)
    refuseIncomplete(fps, deleteOutlineIds || [])
    const want = expect && typeof expect === 'object' ? expect : {}
    const deleted = []
    const notOwned = []
    for (const id of deleteOutlineIds || []) {
      const hit = map.get(id)
      if (!hit) continue
      // Ownership: the plugin's ledger said Hermes made it, it still starts with the marker, AND its
      // content is exactly what Hermes wrote (fingerprint). An edited block stays the person's.
      if (!hit.firstText.startsWith(mark) || typeof want[id] !== 'string' || fps.get(id) !== want[id]) { notOwned.push(id); continue }
      hit.c.delete()
      deleted.push(id)
    }
    await ctx.sync()
    const y = Number.isFinite(top) ? top : await belowContent(ctx, page)
    const o = page.addOutline(Number.isFinite(left) ? left : 40, y, content)
    o.load('id')
    await ctx.sync()
    return { pageId: page.id, outlineId: o.id, deleted, notOwned, created: { outlineIds: [o.id], fingerprints: await writtenFingerprints(ctx, page, [o.id]) } }
  }
}

export function escapeHtml(s) {
  return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;').replace(/'/g, '&#39;')
}

export const MUTATING = new Set(['create_page', 'append', 'insert_table', 'delete_owned', 'replace_owned', 'insert_image'])
export const SUPPORTED_TOOLS = Object.freeze([...Object.keys(tools), 'op_status'])

export async function execute(tool, args) {
  const a = args && typeof args === 'object' ? args : {}
  if (tool === 'op_status') {
    const e = journalGet(String(a.opId || ''))
    return { ok: true, data: e ? { opId: a.opId, ...e } : { opId: a.opId, state: 'unknown' } }
  }
  const fn = Object.prototype.hasOwnProperty.call(tools, tool) ? tools[tool] : null
  if (!fn) return { ok: false, error: { category: 'Unsupported', message: `Pane does not implement ${tool}` } }
  const mutating = MUTATING.has(tool)
  const opId = typeof a.opId === 'string' && a.opId ? a.opId.slice(0, 80) : null
  if (mutating) {
    if (!opId) return { ok: false, error: { category: 'Validation', message: 'mutations need an opId from the plugin' } }
    const prior = journalGet(opId)
    if (prior) {
      // Never run the same operation twice (a lost result must be reconciled, not replayed).
      if (prior.state === 'done') return { ...prior.result, duplicate: true }
      return { ok: false, error: { category: 'ExecutionUncertain', message: 'This operation already started in this pane; check the page before doing anything else.' } }
    }
  }
  try {
    // F8 D2: a write journals what it is about to touch and what it made (before / created) as it goes, so after a
    // lost answer op_status can still tell which block or page is Hermes' (nothing else reads these).
    const note = patch => { if (mutating) journalSet(opId, { ...journalGet(opId), ...patch }) }
    const out = await oneRun(async ctx => {
      if (a.expectNotebookId) await expectNotebook(ctx, a.expectNotebookId)
      if (mutating) journalSet(opId, { state: 'started', tool, at: Date.now() })
      return fn(ctx, a, note)
    })
    const result = { ok: true, data: out }
    if (mutating) journalSet(opId, { state: 'done', tool, at: Date.now(), result })
    return result
  } catch (err) {
    let result = { ok: false, error: officeError(err) }
    // A27: seen live, after a navigation to a page in another section every OneNote request from this pane failed
    // with GeneralException (reads included) while the pane kept polling. When even the smallest read fails the
    // same way, say what the user can do instead of passing on one more GeneralException.
    if (result.error.code === 'GeneralException' && !(await contextAnswers())) {
      result = { ok: false, error: { category: 'PaneReload', code: 'OfficeContextLost', message: contextLostMessage(mutating) } }
    }
    // A refused identity check or validation never started anything; anything later may be partial.
    if (mutating && journalGet(opId)) journalSet(opId, { ...progressOf(journalGet(opId)), state: 'failed', tool, at: Date.now(), result })
    return result
  }
}

function progressOf(e) {
  return { ...(e && e.before ? { before: e.before } : {}), ...(e && e.created ? { created: e.created } : {}) }
}

/** Can this pane still read the smallest thing (the open notebook's id)? Any failure counts as no. */
export async function contextAnswers() {
  try {
    await oneRun(async ctx => { const nb = ctx.application.getActiveNotebook(); nb.load('id'); await ctx.sync() })
    return true
  } catch {
    return false
  }
}

export function contextLostMessage(mutating) {
  return 'OneNote stopped answering this add-in: even reading which notebook is open fails (GeneralException), so ' +
    'every OneNote request from this pane will fail until it is reloaded. ' +
    (mutating ? 'This change may or may not have been made: after the reload, read the page before trying again. '
      : 'Nothing was changed. ') +
    'Ask the user to reload the OneNote browser tab (or close and reopen the Workbench pane); the pane reconnects ' +
    'to this chat by itself. Then try again once.'
}

/** Snapshot of the active page for visit-indexing (pane-pushed event; read-only). */
export async function captureActive() {
  try {
    const data = await oneRun(ctx => readActive(ctx))
    return data
  } catch {
    return null
  }
}

export async function activePageId() {
  try {
    return await oneRun(async ctx => {
      const nb = ctx.application.getActiveNotebook()
      nb.load('id')
      const p = ctx.application.getActivePageOrNull()
      p.load('id')
      await ctx.sync()
      return { notebookId: nb.id, pageId: p.isNullObject ? null : p.id }
    })
  } catch {
    return null
  }
}
