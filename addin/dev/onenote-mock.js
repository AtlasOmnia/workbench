// SPDX-License-Identifier: Apache-2.0
// Strict in-memory stand-in for the OneNote JS API subset (OneNoteApi 1.1) the executor uses.
// DEV/TEST ONLY. Not Microsoft code. Enforces what bites real add-ins:
//  - load()/sync(): reading an unloaded property throws PropertyNotLoaded;
//  - page CONTENT (Page.contents, addOutline) works only for the ACTIVE page; titles/ids/urls of
//    every page are readable without navigating;
//  - there is no Page.delete / Section.delete and RichText.text is read-only;
//  - navigateToPage changes the active page at sync time.
import { ClientResult, Loadable, LoadableCollection, MockContextBase, OfficeError, installOffice, makeRun, persistence } from './mock-core.js'

let uid = 0
const nid = p => `${p}-${++uid}`

// Stricter than mock-core: loading a property the real OneNote class does not have throws
// (e.g. OneNote.Table has no `values`), so an invented API member fails here, not in production.
function checkProps(obj, names) {
  for (const n of names) if (!(n in obj)) throw new OfficeError('PropertyNotFound', `The property '${n}' does not exist on this OneNote object`)
}
class SLoadable extends Loadable {
  _commit() { checkProps(this, [...this._pending]); super._commit() }
}
class SCollection extends LoadableCollection {
  _commit() { const names = [...this._pending].filter(n => n !== 'items' && n !== 'count'); super._commit(); for (const x of this._snap) checkProps(x, names) }
}

function escapeText(s) { return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;') }

export function htmlToText(html) {
  return String(html)
    .replace(/<br\s*\/?>/gi, '\n')
    .replace(/<[^>]+>/g, '')
    .replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/&#39;/g, "'").replace(/&nbsp;/g, ' ').replace(/&amp;/g, '&')
}

// Live OneNote for the web (F8 D1; the harness readback in scripts/livetest/live_driver.py): the items of a <ul>/<ol>
// are CHILD paragraphs (Paragraph.paragraphs) of the paragraph above them; a nested list's items are children of its item.
const listTags = () => /<(\/?)(ul|ol|li)\b[^>]*>/gi

/** The <li> items of one list's inner HTML: { html, children } per item, a nested list's items as children. */
function listItems(inner) {
  const items = []
  const re = listTags()
  let depth = 0
  let cur = null
  let textStart = 0
  let nestStart = -1
  let m
  while ((m = re.exec(inner))) {
    const close = m[1] === '/'
    const tag = m[2].toLowerCase()
    if (tag === 'li') {
      if (depth) continue
      if (!close) { cur = { html: '', children: [] }; items.push(cur); textStart = re.lastIndex; nestStart = -1 } else if (cur) { cur.html += inner.slice(textStart, m.index); cur = null }
      continue
    }
    if (!close) {
      if (depth++ === 0 && cur) { cur.html += inner.slice(textStart, m.index); nestStart = re.lastIndex }
    } else if (depth > 0 && --depth === 0 && cur && nestStart >= 0) {
      cur.children.push(...listItems(inner.slice(nestStart, m.index)))
      textStart = re.lastIndex
      nestStart = -1
    }
  }
  return items
}

function listParagraphs(items) {
  return items.map(it => ({ id: nid('p'), type: 'RichText', html: it.html, children: listParagraphs(it.children) }))
}

/** Split supported HTML into paragraphs the way OneNote does for addOutline/appendHtml (approximately). */
export function htmlToParagraphs(html) {
  const src = String(html)
  const out = []
  const re = listTags()
  let depth = 0
  let start = 0
  let from = 0
  let any = false
  const flat = part => { const got = flatParagraphs(part); if (got.length) any = true; out.push(...got) }
  let m
  while ((m = re.exec(src))) {
    if (m[2].toLowerCase() === 'li') continue
    if (m[1] !== '/') {
      if (depth++ === 0) { flat(src.slice(from, m.index)); start = re.lastIndex }
    } else if (depth > 0 && --depth === 0) {
      any = true
      const kids = listParagraphs(listItems(src.slice(start, m.index)))
      const above = out[out.length - 1]
      if (above && above.type === 'RichText') above.children.push(...kids)
      else out.push(...kids) // a list at the very top of the block has no paragraph above it
      from = re.lastIndex
    }
  }
  flat(src.slice(from))
  if (!any && src.trim()) out.push({ id: nid('p'), type: 'RichText', html: src, children: [] })
  return out
}

function flatParagraphs(html) {
  const out = []
  const re = /<table[^>]*>([\s\S]*?)<\/table>|<(p|h[1-6]|li|div)[^>]*>([\s\S]*?)<\/\2>|(<img\b[^>]*>)/gi
  const image = tag => ({ id: nid('p'), type: 'Image', html: tag, image: { description: (/\balt="([^"]*)"/i.exec(tag) || [])[1] || '' }, children: [] })
  let m
  let any = false
  while ((m = re.exec(String(html)))) {
    any = true
    if (m[4] !== undefined) {
      out.push(image(m[4]))
    } else if (m[3] !== undefined && /<img\b/i.test(m[3])) {
      // OneNote turns an <img> into its own Image paragraph; text around it stays RichText.
      for (const tag of m[3].match(/<img\b[^>]*>/gi)) out.push(image(tag))
      const rest = m[3].replace(/<img\b[^>]*>/gi, '')
      if (rest.replace(/<[^>]+>/g, '').trim()) out.push({ id: nid('p'), type: 'RichText', html: rest, children: [] })
    } else if (m[1] !== undefined) {
      const rows = [...m[1].matchAll(/<tr[^>]*>([\s\S]*?)<\/tr>/gi)].map(r => [...r[1].matchAll(/<t[dh][^>]*>([\s\S]*?)<\/t[dh]>/gi)].map(c => htmlToText(c[1])))
      out.push({ id: nid('p'), type: 'Table', rows, children: [] })
    } else {
      out.push({ id: nid('p'), type: 'RichText', html: m[3], children: [] })
    }
  }
  if (!any && String(html).trim()) out.push({ id: nid('p'), type: 'RichText', html: String(html), children: [] })
  return out
}

export class MockNotebook {
  /**
   * spec: { id?, name, sections: [{ name, pages: [{ title, html?: string[] }] }], groups?: [{ name, sections: [...] }] }
   */
  constructor(spec = { name: 'Notebook', sections: [{ name: 'Section 1', pages: [{ title: 'Untitled page' }] }] }) {
    this.id = spec.id || nid('nb')
    this.name = spec.name
    const mkSection = s => ({
      id: nid('sec'), name: s.name, locked: !!s.locked, apiId: s.apiId,
      pages: (s.pages || []).map(p => this._mkPage(p))
    })
    const mkGroup = g => ({ id: nid('grp'), name: g.name, sections: (g.sections || []).map(mkSection), groups: (g.groups || []).map(mkGroup) })
    this.sections = (spec.sections || []).map(mkSection)
    this.groups = (spec.groups || []).map(mkGroup)
    const first = this.allSections().find(s => s.pages.length && !s.locked)
    this.active = first ? { section: first.id, page: first.pages[0].id } : { section: null, page: null }
    this.syncs = 0
    this.navigations = []
  }
  _mkPage(p) {
    const id = nid('page')
    const page = { id, title: p.title, level: p.level || 0, outlines: [] }
    for (const h of p.html || []) page.outlines.push({ id: nid('ol'), left: 40, top: 90 + 120 * page.outlines.length, paragraphs: htmlToParagraphs(h) })
    return page
  }
  webUrl(id) { return `https://onenote.test/${this.id}/${id}` }
  allSections() {
    const out = []
    const walk = g => { out.push(...g.sections); g.groups.forEach(walk) }
    walk({ sections: this.sections, groups: this.groups })
    return out
  }
  sectionOf(pageId) { return this.allSections().find(s => s.pages.some(p => p.id === pageId)) }
  page(id) { for (const s of this.allSections()) { const p = s.pages.find(x => x.id === id); if (p) return p } return null }
  pageByTitle(title) { for (const s of this.allSections()) { const p = s.pages.find(x => x.title === title); if (p) return p } return null }
  activePage() { return this.active.page ? this.page(this.active.page) : null }
  /** Simulate the user clicking another page (no add-in involvement). */
  userOpen(pageId) { this.active = { section: this.sectionOf(pageId).id, page: pageId } }
  fingerprint() {
    const para = x => { const v = x.type === 'Table' ? x.rows : x.type === 'Image' ? [x.html, x.image.description] : x.html; return x.children && x.children.length ? [v, x.children.map(para)] : v }
    return JSON.stringify(this.allSections().map(s => [s.name, s.pages.map(p => [p.title, p.outlines.map(o => o.paragraphs.map(para))])]))
  }
}

function sectionPath(nb, sectionId) {
  const walk = (g, prefix) => {
    for (const s of g.sections) if (s.id === sectionId) return [...prefix, s.name]
    for (const sub of g.groups) { const r = walk(sub, [...prefix, sub.name]); if (r) return r }
    return null
  }
  return walk({ sections: nb.sections, groups: nb.groups }, [])
}

class NParagraph extends SLoadable {
  // siblings: the array this paragraph lives in (the outline's paragraphs, or its parent paragraph's children).
  constructor(ctx, page, outline, para, siblings) { super(ctx); this._page = page; this._o = outline; this._p = para; this._sib = siblings || outline.paragraphs; this._rt = null; this._tbl = null; this._kids = null }
  get id() { this._need('id'); return this._p.id }
  get type() { this._need('type'); return this._p.type }
  /** Paragraph.paragraphs: the paragraphs under this one (list items under the paragraph above the list). */
  get paragraphs() {
    if (!this._kids) {
      const p = this._p
      this._kids = new SCollection(this._ctx, () => { this._ctx._requireActive(this._page); if (!p.children) p.children = []; return p.children.map(x => new NParagraph(this._ctx, this._page, this._o, x, p.children)) })
    }
    return this._kids
  }
  get richText() {
    if (this._p.type !== 'RichText') throw new OfficeError('InvalidOperation', 'Paragraph is not RichText')
    if (!this._rt) {
      const p = this._p; const ctx = this._ctx
      this._rt = new (class extends SLoadable {
        get text() { this._need('text'); return htmlToText(p.html) }
        set text(_v) { throw new OfficeError('InvalidOperation', 'RichText.text is read-only') }
        getHtml() { return new ClientResult(ctx, () => p.html) }
      })(ctx)
    }
    return this._rt
  }
  get table() {
    if (this._p.type !== 'Table') throw new OfficeError('InvalidOperation', 'Paragraph is not a Table')
    if (!this._tbl) {
      // OneNote.Table has rowCount/columnCount/rows (no values): text lives in row -> cell -> paragraphs.
      const p = this._p; const ctx = this._ctx; const page = this._page; const outline = this._o
      const cell = (r, c) => new (class extends SLoadable {
        get id() { this._need('id'); return `${p.id}-r${r}c${c}` }
        get rowIndex() { this._need('rowIndex'); return r }
        get cellIndex() { this._need('cellIndex'); return c }
        get paragraphs() {
          if (!this._ps) this._ps = new SCollection(ctx, () => [new NParagraph(ctx, page, outline, { id: `${p.id}-r${r}c${c}-p0`, type: 'RichText', html: escapeText(p.rows[r][c]), children: (p.cellKids && p.cellKids[`${r},${c}`]) || [] })])
          return this._ps
        }
      })(ctx)
      const row = r => new (class extends SLoadable {
        get id() { this._need('id'); return `${p.id}-r${r}` }
        get cellCount() { this._need('cellCount'); return p.rows[r].length }
        get cells() { if (!this._cs) this._cs = new SCollection(ctx, () => p.rows[r].map((_, c) => cell(r, c))); return this._cs }
      })(ctx)
      this._tbl = new (class extends SLoadable {
        get id() { this._need('id'); return `${p.id}-t` }
        get rowCount() { this._need('rowCount'); return p.rows.length }
        get columnCount() { this._need('columnCount'); return p.rows[0] ? p.rows[0].length : 0 }
        get rows() { if (!this._rs) this._rs = new SCollection(ctx, () => p.rows.map((_, r) => row(r))); return this._rs }
      })(ctx)
    }
    return this._tbl
  }
  get image() {
    if (this._p.type !== 'Image') throw new OfficeError('InvalidOperation', 'Paragraph is not an Image')
    const p = this._p; const ctx = this._ctx
    return new (class extends SLoadable {
      get description() { this._need('description'); return p.image.description }
      set description(v) { ctx._queue(() => { p.image.description = String(v) }) }
      get width() { this._need('width'); return p.image.width ?? null }
      get height() { this._need('height'); return p.image.height ?? null }
    })(ctx)
  }
  delete() {
    this._ctx._queue(() => {
      this._ctx._requireActive(this._page)
      const i = this._sib.indexOf(this._p)
      if (i < 0) throw new OfficeError('ItemNotFound', 'Paragraph no longer exists')
      this._sib.splice(i, 1)
    })
  }
}

class NOutline extends SLoadable {
  constructor(ctx, page, outline) { super(ctx); this._page = page; this._o = outline; this._paras = null }
  get id() { this._need('id'); return this._o.id }
  get paragraphs() {
    if (!this._paras) {
      this._paras = new SCollection(this._ctx, () => { this._ctx._requireActive(this._page); return this._o.paragraphs.map(p => new NParagraph(this._ctx, this._page, this._o, p)) })
    }
    return this._paras
  }
  appendHtml(html) {
    this._ctx._queue(() => { this._ctx._requireActive(this._page); if (this._ctx._model.failAppend) throw new OfficeError('GeneralException', 'append failed'); this._o.paragraphs.push(...htmlToParagraphs(html)) })
  }
  appendTable(rowCount, columnCount, values) {
    const t = new (class extends SLoadable { get id() { this._need('id'); return this._id } })(this._ctx)
    t._id = nid('tbl')
    this._ctx._queue(() => {
      this._ctx._requireActive(this._page)
      if (!Number.isInteger(rowCount) || !Number.isInteger(columnCount) || rowCount < 1 || columnCount < 1) throw new OfficeError('InvalidArgument', 'bad table size')
      const rows = Array.from({ length: rowCount }, (_, r) => Array.from({ length: columnCount }, (_, c) => String((values && values[r] && values[r][c]) ?? '')))
      this._o.paragraphs.push({ id: nid('p'), type: 'Table', rows, children: [] })
    })
    return t
  }
}

class NPageContent extends SLoadable {
  constructor(ctx, page, outline) { super(ctx); this._page = page; this._o = outline; this._outline = null }
  get id() { this._need('id'); return `pc-${this._o.id}` }
  get type() { this._need('type'); return 'Outline' }
  get left() { this._need('left'); return this._o.left }
  get top() { this._need('top'); return this._o.top }
  get outline() { if (!this._outline) this._outline = new NOutline(this._ctx, this._page, this._o); return this._outline }
  delete() {
    this._ctx._queue(() => {
      this._ctx._requireActive(this._page)
      const i = this._page.outlines.indexOf(this._o)
      if (i < 0) throw new OfficeError('ItemNotFound', 'Content no longer exists')
      this._page.outlines.splice(i, 1)
    })
  }
}

class NPage extends SLoadable {
  constructor(ctx, page) { super(ctx); this._pg = page; this._contents = null }
  get _nb() { return this._ctx._model }
  get id() { this._need('id'); return this._pg.id }
  get title() { this._need('title'); return this._pg.title }
  set title(v) { this._ctx._queue(() => { this._pg.title = String(v) }) }
  get webUrl() { this._need('webUrl'); return this._nb.webUrl(this._pg.id) }
  get clientUrl() { this._need('clientUrl'); return `onenote:mock/${this._pg.id}` }
  get pageLevel() { this._need('pageLevel'); return this._pg.level }
  get contents() {
    if (!this._contents) {
      // Real OneNote for the web (observed live 2026-11-06, A8): once a Page proxy's `contents` has been loaded, later
      // loads of that SAME proxy's collection in the same context do not see outlines added since (addOutline); a
      // fresh proxy (application.getActivePage()) does. The mock freezes membership at the first load likewise.
      let members = null
      this._contents = new SCollection(this._ctx, () => {
        this._ctx._requireActive(this._pg)
        if (!members) members = new Set(this._pg.outlines)
        return this._pg.outlines.filter(o => members.has(o)).map(o => new NPageContent(this._ctx, this._pg, o))
      })
    }
    return this._contents
  }
  addOutline(left, top, html) {
    const o = { id: nid('ol'), left, top, paragraphs: [] }
    const out = new NOutline(this._ctx, this._pg, o)
    this._ctx._queue(() => {
      this._ctx._requireActive(this._pg)
      if (this._ctx._model.failAddOutline) throw new OfficeError('GeneralException', 'addOutline failed')
      o.paragraphs = htmlToParagraphs(html)
      this._pg.outlines.push(o)
    })
    return out
  }
}

class NSection extends SLoadable {
  constructor(ctx, s) { super(ctx); this._s = s; this._pages = null }
  // apiId: what real OneNote for the web reports as Section.id, which after a reload can be the SAME for
  // several sections of a notebook (observed live 2026-11-05); the mock keeps its own unique id inside.
  get id() { this._need('id'); return this._s.apiId || this._s.id }
  get name() { this._need('name'); return this._s.name }
  get webUrl() { this._need('webUrl'); return `https://onenote.test/${this._ctx._model.id}/${this._s.id}` }
  get pages() {
    if (!this._pages) {
      this._pages = new SCollection(this._ctx, () => {
        if (this._s.locked) throw new OfficeError('AccessDenied', 'Section is locked')
        return this._s.pages.map(p => new NPage(this._ctx, p))
      })
    }
    return this._pages
  }
  addPage(title) {
    const pg = { id: nid('page'), title: String(title), level: 0, outlines: [] }
    const out = new NPage(this._ctx, pg)
    this._ctx._queue(() => { if (this._s.locked) throw new OfficeError('AccessDenied', 'Section is locked'); this._s.pages.push(pg) })
    return out
  }
}

class NGroup extends SLoadable {
  constructor(ctx, g) { super(ctx); this._g = g }
  get id() { this._need('id'); return this._g.id }
  get name() { this._need('name'); return this._g.name }
  get sections() { return new SCollection(this._ctx, () => this._g.sections.map(s => new NSection(this._ctx, s))) }
  get sectionGroups() { return new SCollection(this._ctx, () => this._g.groups.map(g => new NGroup(this._ctx, g))) }
}

class NNotebook extends SLoadable {
  get id() { this._need('id'); return this._ctx._model.id }
  get name() { this._need('name'); return this._ctx._model.name }
  get sections() { return new SCollection(this._ctx, () => this._ctx._model.sections.map(s => new NSection(this._ctx, s))) }
  get sectionGroups() { return new SCollection(this._ctx, () => this._ctx._model.groups.map(g => new NGroup(this._ctx, g))) }
}

class MockContext extends MockContextBase {
  constructor(nb) {
    super()
    this._model = nb
    // A navigation that lands a little later (navLagRuns batches after the one that asked for it).
    if (nb.pendingNav && --nb.pendingNav.left <= 0) { nb.active = nb.pendingNav.active; nb.pendingNav = null }
    const ctx = this
    this.application = {
      getActiveNotebook: () => new NNotebook(ctx),
      getActiveSection: () => { const s = nb.allSections().find(x => x.id === nb.active.section); if (!s) ctx._fail(new OfficeError('ItemNotFound', 'No active section')); return new NSection(ctx, s || { id: '', name: '', pages: [] }) },
      getActivePage: () => { const p = nb.activePage(); if (!p) ctx._fail(new OfficeError('ItemNotFound', 'No active page')); return new NPage(ctx, p || { id: '', title: '', outlines: [] }) },
      getActivePageOrNull: () => {
        const p = nb.activePage()
        const page = new NPage(ctx, p || { id: '', title: '', outlines: [] })
        Object.defineProperty(page, 'isNullObject', { get: () => !p })
        return page
      },
      navigateToPage: page => {
        ctx._queue(() => {
          if (nb.failNavigate) throw new OfficeError('GeneralException', 'navigation failed')
          const target = page._pg
          if (!nb.page(target.id)) throw new OfficeError('ItemNotFound', 'Page not found')
          nb.navigations.push(target.id)
          const fromSection = nb.active.section
          const next = { section: nb.sectionOf(target.id).id, page: target.id }
          if (nb.navLagRuns) { nb.pendingNav = { active: next, left: nb.navLagRuns }; return }
          nb.active = next
          // Live OneNote for the web (A26 cold run): after opening a page in ANOTHER section, every later request
          // from the add-in failed with GeneralException (reads included) while the pane kept running.
          if (nb.crossSectionNavigationBreaksContext && fromSection !== nb.active.section) nb.contextLost = true
          // Live OneNote for the web (2026-11-05): opening a page in ANOTHER section does open it, but the
          // batch that navigated fails with GeneralException (the section reloads under the add-in).
          if (nb.crossSectionNavigationThrows && fromSection !== nb.active.section) throw new OfficeError('GeneralException', 'Something went wrong.')
        })
      }
    }
  }
  async sync() {
    if (this._model.contextLost) {
      this._model.syncs++
      this._ops = []
      this._dirty.clear()
      throw new OfficeError('GeneralException', 'Something went wrong. Check the OfficeExtension.Error.debugInfo for more information.')
    }
    return super.sync()
  }
  _requireActive(pg) {
    if (!pg || pg.id !== this._model.active.page) throw new OfficeError('InvalidOperation', 'Page content is only available for the active page')
  }
}

const store = persistence('hermes-mock-onenote',
  nb => ({ uid, nb }),
  x => { uid = Math.max(uid, x.uid || 0); return Object.assign(Object.create(MockNotebook.prototype), x.nb) })

const selectionHandlers = []
export function fireSelectionChanged() { for (const h of selectionHandlers) h({ type: 'documentSelectionChanged' }) }

export function installOneNoteMock(nb) {
  nb = nb || store.load() || new MockNotebook()
  globalThis.OneNote = { run: makeRun(MockContext, nb, () => store.save(nb)) }
  installOffice('OneNote', { url: '' })
  globalThis.Office.HostType.OneNote = 'OneNote'
  globalThis.Office.EventType = { DocumentSelectionChanged: 'documentSelectionChanged' }
  globalThis.Office.context.document.addHandlerAsync = (type, handler, cb) => { selectionHandlers.push(handler); cb && cb({ status: 'succeeded' }) }
  globalThis.__hermesMockNotebook = nb
  globalThis.__hermesMockFingerprint = () => nb.fingerprint()
  return nb
}

export { sectionPath }
