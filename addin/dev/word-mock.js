// SPDX-License-Identifier: Apache-2.0
// Strict in-memory stand-in for the Office.js Word subset the executor uses.
// DEV/TEST ONLY. The document is a list of paragraphs; ranges live inside one
// paragraph and are re-anchored when earlier text in the same paragraph changes.
// body.insertOoxml(..., 'Replace') reproduces Word's extra trailing paragraph so
// the executor's undo trim is exercised.
import { ClientResult, Loadable, LoadableCollection, MockContextBase, OfficeError, installOffice, makeRun, persistence } from './mock-core.js'

let uid = 0
let blk = 0

const imgCount = html => (String(html).match(/<img\b/gi) || []).length
const htmlParas = html => {
  const out = []
  for (const m of String(html).matchAll(/<(p|h[1-6]|li|div)[^>]*>([\s\S]*?)<\/\1>/gi)) out.push(m[2].replace(/<[^>]+>/g, ''))
  return out.length ? out : [String(html).replace(/<[^>]+>/g, '')]
}

/** Insert new paragraphs as one block (picture or HTML) at sync time; returns the block id.
 *  joinLast (body 'End' inserts): like Word for the web (seen live, A9), when a content control ends the document the
 *  inserted pictures land inside THAT control, not in the range the insert returns. */
function insertBlock(ctx, where, paras, joinLast = false) {
  const id = `blk-${++blk}`
  ctx._queue(() => {
    const d = ctx._model
    const at = where(d)
    const prev = joinLast ? d.paragraphs[d.paragraphs.length - 1] : null
    const into = prev && prev.blk && d.controls.some(c => c.blk === prev.blk) ? prev.blk : null
    d.paragraphs.splice(at, 0, ...paras.map(p => ({ text: '', style: 'Normal', isListItem: false, ...p, blk: into && p.pictures ? into : id })))
  })
  return id
}

class WControl extends Loadable {
  constructor(ctx, rec) { super(ctx); this._rec = rec }
  get id() { this._need('id'); return this._rec.id }
  get tag() { this._need('tag'); return this._rec.tag }
  set tag(v) { this._ctx._queue(() => { this._rec.tag = String(v) }) }
  set title(v) { this._ctx._queue(() => { this._rec.title = String(v) }) }
  set appearance(v) { this._ctx._queue(() => { this._rec.appearance = String(v) }) }
  get inlinePictures() {
    const ctx = this._ctx
    return new LoadableCollection(ctx, () => ctx._model.paragraphs.filter(p => p.blk === this._rec.blk).flatMap(p => Array.from({ length: p.pictures || 0 }, () => new Loadable(ctx))))
  }
  insertInlinePictureFromBase64(b64, loc) {
    if (loc !== 'Start') throw new OfficeError('InvalidArgument', 'the mock inserts a control picture at Start only')
    const pic = picOf(this._ctx, b64)
    this._ctx._queue(() => {
      const p = this._ctx._model.paragraphs.find(x => x.blk === this._rec.blk)
      if (!p) throw new OfficeError('ItemNotFound', 'content control')
      p.pictures = (p.pictures || 0) + 1; p.picture = pic; p.text = '\ufffc' + p.text
    })
    return new WPicture(this._ctx, this._rec.blk, pic)
  }
  delete(keepContent) {
    this._ctx._queue(() => {
      const d = this._ctx._model
      const i = d.controls.indexOf(this._rec)
      if (i < 0) throw new OfficeError('ItemNotFound', 'content control')
      d.controls.splice(i, 1)
      if (!keepContent) d.paragraphs = d.paragraphs.filter(p => p.blk !== this._rec.blk)
    })
  }
}

function makeControl(ctx, blkId) {
  const rec = { id: ++uid, tag: '', title: '', appearance: 'BoundingBox', blk: blkId }
  ctx._queue(() => ctx._model.controls.push(rec))
  return new WControl(ctx, rec)
}

class WPicture extends Loadable {
  constructor(ctx, blkId, pic) { super(ctx); this._blk = blkId; this._pic = pic }
  insertContentControl() { return makeControl(this._ctx, this._blk) }
}
for (const k of ['altTextDescription', 'altTextTitle', 'width', 'height', 'lockAspectRatio']) {
  Object.defineProperty(WPicture.prototype, k, {
    get() { this._need(k); return this._pic[k] },
    set(v) { this._ctx._queue(() => { this._pic[k] = v }) }
  })
}

function picOf(ctx, b64) {
  if (typeof b64 !== 'string' || !/^(iVBORw0KGgo|\/9j\/)/.test(b64)) ctx._fail(new OfficeError('InvalidArgument', 'not an image'))
  return { kind: typeof b64 === 'string' && b64.startsWith('/9j/') ? 'jpeg' : 'png' }
}

function insertPicture(ctx, b64, where, joinLast = false) {
  const pic = picOf(ctx, b64)
  const id = insertBlock(ctx, where, [{ text: '\ufffc', pictures: 1, picture: pic }], joinLast)
  return new WPicture(ctx, id, pic)
}

function insertHtml(ctx, html, where, joinLast = false) {
  const pics = imgCount(html)
  const id = insertBlock(ctx, where, htmlParas(html).map((t, i) => ({ text: t, pictures: i === 0 ? pics : 0 })), joinLast)
  return { insertContentControl: () => makeControl(ctx, id) }
}

const newPara = text => ({ text: String(text), style: 'Normal', isListItem: false })
const controlled = (d, p) => !!(p && p.blk && d.controls.some(c => c.blk === p.blk))

/** A paragraph bound to one paragraph object (resolved on first use), so it survives inserts before it. */
class WPara extends Loadable {
  constructor(ctx, resolve) { super(ctx); this._resolve = resolve; this._ref = null }
  get _para() {
    if (!this._ref) this._ref = this._resolve(this._ctx._model)
    if (!this._ref) throw new OfficeError('ItemNotFound', 'paragraph')
    return this._ref
  }
  get _index() { const i = this._ctx._model.paragraphs.indexOf(this._para); if (i < 0) throw new OfficeError('ItemNotFound', 'paragraph'); return i }
  get text() { this._need('text'); return this._para.text }
  get inlinePictures() { const ctx = this._ctx; return new LoadableCollection(ctx, () => Array.from({ length: this._para.pictures || 0 }, () => new Loadable(ctx))) }
  // Mock controls are paragraph-level blocks, so a paragraph never holds a control of its own.
  get contentControls() { return new LoadableCollection(this._ctx, () => []) }
  get parentContentControlOrNullObject() {
    const self = this
    return new (class extends Loadable {
      get isNullObject() { this._need('isNullObject'); return !controlled(self._ctx._model, self._para) }
      insertParagraph(text, loc) {
        const p = newPara(text)
        self._ctx._queue(() => {
          const d = self._ctx._model; const b = self._para.blk
          const idx = d.paragraphs.map((x, i) => (x.blk === b ? i : -1)).filter(i => i >= 0)
          d.paragraphs.splice(loc === 'Before' ? idx[0] : idx[idx.length - 1] + 1, 0, p)
        })
        return new WPara(self._ctx, () => p)
      }
    })(this._ctx)
  }
  insertParagraph(text, loc) {
    if (loc !== 'Before' && loc !== 'After') throw new OfficeError('InvalidArgument', `Bad location ${loc}`)
    const p = newPara(text)
    this._ctx._queue(() => { const i = this._index; this._ctx._model.paragraphs.splice(loc === 'Before' ? i : i + 1, 0, p) })
    return new WPara(this._ctx, () => p)
  }
  insertHtml(html, loc) {
    if (loc !== 'Replace') throw new OfficeError('InvalidArgument', 'the mock inserts paragraph HTML with Replace only')
    const id = `blk-${++blk}`
    const pics = imgCount(html)
    this._ctx._queue(() => {
      const i = this._index
      this._ctx._model.paragraphs.splice(i, 1, ...htmlParas(html).map((t, k) => ({ ...newPara(t), pictures: k === 0 ? pics : 0, blk: id })))
    })
    return { insertContentControl: () => makeControl(this._ctx, id) }
  }
  insertContentControl() {
    const id = `blk-${++blk}`
    this._ctx._queue(() => { this._para.blk = id })
    return makeControl(this._ctx, id)
  }
}

export class MockDocument {
  constructor(paragraphs = ['']) {
    this.paragraphs = paragraphs.map(t => ({ text: t, style: 'Normal', isListItem: false }))
    this.comments = []
    this.controls = [] // content controls { id, tag, title, appearance, blk }
    this.marks = [] // { p, start, end, props }
    this.selection = { p: 0, start: 0, end: 0 }
    this.title = ''
    this.syncs = 0
    this.live = new Set()
  }
  get text() { return this.paragraphs.map(p => p.text).join('\n') }
  setSelection(p, start, end) { this.selection = { p, start, end } }
  /** Replace [start,end) in paragraph p with text (may contain \n); returns the new span of the inserted text's last line. */
  splice(p, start, end, text) {
    const para = this.paragraphs[p]
    const before = para.text.slice(0, start); const after = para.text.slice(end)
    const lines = String(text).split('\n')
    const delta = lines[0].length - (end - start)
    if (lines.length === 1) {
      para.text = before + text + after
      for (const r of this.live) if (r._p === p && r._start >= end && r._start !== start) { r._start += delta; r._end += delta }
      return { p, start, end: start + text.length }
    }
    para.text = before + lines[0]
    const extra = lines.slice(1).map(t => ({ text: t, style: para.style, isListItem: false }))
    extra[extra.length - 1].text += after
    this.paragraphs.splice(p + 1, 0, ...extra)
    return { p: p + lines.length - 1, start: 0, end: lines[lines.length - 1].length }
  }
  serialize() { return `<pkg:package mock="1">${JSON.stringify({ paragraphs: this.paragraphs, comments: this.comments, marks: this.marks })}</pkg:package>` }
}

class WRange extends Loadable {
  constructor(ctx, p, start, end) { super(ctx); this._p = p; this._start = start; this._end = end; ctx._model.live.add(this) }
  get _doc() { return this._ctx._model }
  get text() { this._need('text'); return this._doc.paragraphs[this._p].text.slice(this._start, this._end) }
  get font() {
    const self = this
    return new Proxy({}, { set(_t, k, v) { self._ctx._queue(() => self._doc.marks.push({ p: self._p, start: self._start, end: self._end, props: { [k]: v } })); return true } })
  }
  insertText(text, loc) {
    const out = new WRange(this._ctx, this._p, this._start, this._start)
    this._ctx._queue(() => {
      const d = this._doc
      let at
      if (loc === 'Replace') at = d.splice(this._p, this._start, this._end, text)
      else if (loc === 'Start' || loc === 'Before') at = d.splice(this._p, this._start, this._start, text)
      else if (loc === 'End' || loc === 'After') at = d.splice(this._p, this._end, this._end, text)
      else throw new OfficeError('InvalidArgument', `Bad location ${loc}`)
      out._p = at.p; out._start = String(text).includes('\n') ? at.start : at.end - String(text).length; out._end = at.end
    })
    return out
  }
  insertOoxml(x, loc) {
    const text = [...String(x).matchAll(/<w:t[^>]*>([^<]*)<\/w:t>/g)].map(m => m[1]).join('')
    return this.insertText(text, loc === 'Replace' ? 'Replace' : loc)
  }
  select() { this._ctx._queue(() => { this._doc.selection = { p: this._p, start: this._start, end: this._end } }) }
  _where(loc) {
    if (loc === 'Replace') throw new OfficeError('InvalidArgument', 'the mock never replaces the selection with a block')
    return loc === 'Before' || loc === 'Start' ? () => this._p : () => this._p + 1
  }
  insertInlinePictureFromBase64(b64, loc) { return insertPicture(this._ctx, b64, this._where(loc)) }
  insertHtml(html, loc) { return insertHtml(this._ctx, html, this._where(loc)) }
  compareLocationWith(o) {
    return new ClientResult(this._ctx, () => {
      if (this._p !== o._p) return this._p < o._p ? 'Before' : 'After'
      if (this._end < o._start) return 'Before'
      if (this._end === o._start) return 'AdjacentBefore'
      if (this._start > o._end) return 'After'
      if (this._start === o._end) return 'AdjacentAfter'
      return 'Overlaps'
    })
  }
  insertComment(text) {
    const c = new (class extends Loadable { get id() { this._need('id'); return this._id } })(this._ctx)
    c._id = `comment-${++uid}`
    this._ctx._queue(() => this._doc.comments.push({ id: c._id, text, anchor: this._doc.paragraphs[this._p].text.slice(this._start, this._end) }))
    return c
  }
  get paragraphs() {
    const ctx = this._ctx
    const col = new LoadableCollection(ctx, () => [new WParagraph(ctx, this._p)])
    col.getLast = col.getFirst = () => new WPara(ctx, d => d.paragraphs[this._p])
    return col
  }
}

class WParagraph extends Loadable {
  constructor(ctx, i) { super(ctx); this._i = i }
  get _para() { return this._ctx._model.paragraphs[this._i] }
  get text() { this._need('text'); return this._para.text }
  get style() { this._need('style'); return this._para.style }
  set style(v) { this._ctx._queue(() => { this._para.style = v }) }
  set styleBuiltIn(v) { this._ctx._queue(() => { this._para.style = v.replace(/^Heading(\d)$/, 'Heading $1') }) }
  get isListItem() { this._need('isListItem'); return this._para.isListItem }
  delete() { this._ctx._queue(() => { const d = this._ctx._model; if (d.paragraphs.length <= 1) throw new OfficeError('InvalidOperation', 'last paragraph'); d.paragraphs.splice(this._i, 1) }) }
}

class MockContext extends MockContextBase {
  constructor(doc) {
    super(); this._model = doc
    const ctx = this
    const q = fn => ctx._queue(fn)
    const body = new (class extends Loadable {
      get text() { this._need('text'); return doc.text }
      get paragraphs() {
        const col = new LoadableCollection(ctx, () => doc.paragraphs.map((_, i) => new WParagraph(ctx, i)))
        col.getFirst = () => new WPara(ctx, d => d.paragraphs[0])
        col.getLast = () => new WPara(ctx, d => d.paragraphs[d.paragraphs.length - 1])
        return col
      }
      insertParagraph(text, loc) {
        const p = newPara(text)
        q(() => { if (loc === 'Start') doc.paragraphs.unshift(p); else if (loc === 'End') doc.paragraphs.push(p); else throw new OfficeError('InvalidArgument', `Bad location ${loc}`) })
        return new WPara(ctx, () => p)
      }
      get tables() { return new LoadableCollection(ctx, () => []) }
      search(text, { matchCase } = {}) {
        if (text.length > 255) ctx._fail(new OfficeError('SearchStringInvalidOrTooLong', 'The search string is too long'))
        return new LoadableCollection(ctx, () => {
          const out = []
          doc.paragraphs.forEach((p, i) => {
            const hay = matchCase ? p.text : p.text.toLowerCase(); const needle = matchCase ? text : text.toLowerCase()
            let k = hay.indexOf(needle)
            while (k >= 0 && needle) { out.push(new WRange(ctx, i, k, k + needle.length)); k = hay.indexOf(needle, k + needle.length) }
          })
          return out
        })
      }
      insertText(text, loc) {
        const last = doc.paragraphs.length - 1
        return loc === 'Start' ? new WRange(ctx, 0, 0, 0).insertText(text, 'Start') : new WRange(ctx, last, doc.paragraphs[last].text.length, doc.paragraphs[last].text.length).insertText(text, 'End')
      }
      getOoxml() { return new ClientResult(ctx, () => doc.serialize()) }
      insertOoxml(x, loc) {
        const m = /^<pkg:package mock="1">(.*)<\/pkg:package>$/s.exec(String(x))
        if (m && loc === 'Replace') {
          q(() => { const s = JSON.parse(m[1]); doc.paragraphs = [...s.paragraphs, { text: '', style: 'Normal', isListItem: false }]; doc.comments = s.comments; doc.marks = s.marks })
          return
        }
        const text = [...String(x).matchAll(/<w:t[^>]*>([^<]*)<\/w:t>/g)].map(mm => mm[1]).join('')
        const out = new WRange(ctx, 0, 0, 0) // like Word, the inserted range; re-anchored when the insert runs
        q(() => {
          if (loc === 'Start') doc.paragraphs.unshift({ text, style: 'Normal', isListItem: false }); else doc.paragraphs.push({ text, style: 'Normal', isListItem: false })
          out._p = loc === 'Start' ? 0 : doc.paragraphs.length - 1; out._start = 0; out._end = text.length
        })
        return out
      }
      insertInlinePictureFromBase64(b64, loc) { return insertPicture(ctx, b64, loc === 'Start' ? () => 0 : d => d.paragraphs.length, loc === 'End') }
      insertHtml(html, loc) { return insertHtml(ctx, html, loc === 'Start' ? () => 0 : d => d.paragraphs.length, loc === 'End') }
      getComments() {
        return new LoadableCollection(ctx, () => doc.comments.map(c => Object.assign(new (class extends Loadable { get id() { this._need('id'); return c.id } })(ctx), {
          delete: () => q(() => { doc.comments.splice(doc.comments.indexOf(c), 1) })
        })))
      }
    })(ctx)
    this.document = {
      body,
      contentControls: {
        getByTag: tag => new LoadableCollection(ctx, () => doc.controls.filter(c => c.tag === tag).map(rec => new WControl(ctx, rec)))
      },
      getSelection: () => new WRange(ctx, doc.selection.p, doc.selection.start, doc.selection.end),
      properties: new (class extends Loadable { get title() { this._need('title'); return doc.title } get author() { this._need('author'); return 'mock' } })(ctx)
    }
  }
}

const wordStore = persistence('hermes-mock-word',
  d => ({ uid, paragraphs: d.paragraphs, comments: d.comments, controls: d.controls, marks: d.marks, selection: d.selection, title: d.title }),
  x => { uid = Math.max(uid, x.uid || 0); const d = new MockDocument(['']); Object.assign(d, { paragraphs: x.paragraphs, comments: x.comments, controls: x.controls || [], marks: x.marks, selection: x.selection, title: x.title }); return d })

export function installWordMock(doc) {
  doc = doc || wordStore.load() || new MockDocument()
  globalThis.Word = { run: makeRun(MockContext, doc, () => wordStore.save(doc)) }
  installOffice('Word', { url: 'mock://Document1.docx' })
  globalThis.__hermesMockDocument = doc
  globalThis.__hermesMockFingerprint = () => JSON.stringify({ p: doc.paragraphs, c: doc.comments.map(c => [c.text, c.anchor]), m: doc.marks, cc: doc.controls })
  return doc
}
