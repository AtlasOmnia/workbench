// SPDX-License-Identifier: Apache-2.0
// Strict in-memory stand-in for the Office.js PowerPoint subset the executor uses.
// DEV/TEST ONLY. Slide export is a JSON copy encoded as base64 so delete-undo and
// duplicate round-trip through insertSlidesFromBase64 exactly like the real flow.
import { ClientResult, Loadable, LoadableCollection, MockContextBase, OfficeError, installOffice, makeRun, persistence } from './mock-core.js'

let uid = 255
const b64 = s => (typeof Buffer !== 'undefined' ? Buffer.from(s, 'utf8').toString('base64') : btoa(unescape(encodeURIComponent(s))))
const unb64 = s => (typeof Buffer !== 'undefined' ? Buffer.from(s, 'base64').toString('utf8') : decodeURIComponent(escape(atob(s))))

export class MockPresentation {
  constructor() {
    const layouts = [
      { id: 'L-title', name: 'Title Slide', placeholders: ['CenterTitle', 'Subtitle'] },
      { id: 'L-content', name: 'Title and Content', placeholders: ['Title', 'Body'] },
      { id: 'L-titleonly', name: 'Title Only', placeholders: ['Title'] },
      { id: 'L-blank', name: 'Blank', placeholders: [] }
    ]
    this.masters = [{ id: 'M-1', name: 'Office Theme', layouts }]
    this.slides = []
    this.selected = []
    this.selectedShapes = []
    this.syncs = 0
    this.addSlide({ layoutId: 'L-title' })
    this.selected = [this.slides[0].id]
  }
  layout(id) { for (const m of this.masters) { const l = m.layouts.find(x => x.id === id); if (l) return { m, l } } return null }
  newShape(fields) { return { id: String(++uid), name: fields.name || `Shape ${uid}`, left: 0, top: 0, width: 100, height: 50, text: '', font: { bold: false, italic: false, color: '#000000', size: 18, name: 'Calibri' }, ...fields } }
  addSlide({ layoutId, slideMasterId } = {}, at) {
    const hit = this.layout(layoutId || 'L-content')
    if (!hit || (slideMasterId && hit.m.id !== slideMasterId)) throw new OfficeError('InvalidArgument', 'Unknown layout')
    const slide = { id: `${++uid}#`, layoutId: hit.l.id, masterId: hit.m.id, shapes: hit.l.placeholders.map((t, i) => this.newShape({ type: 'Placeholder', placeholderType: t, name: t, top: 40 + i * 120, width: 800, height: 100 })) }
    if (at === undefined) this.slides.push(slide); else this.slides.splice(at, 0, slide)
    return slide
  }
  slide(id) { return this.slides.find(s => s.id === id) }
}

class Shape extends Loadable {
  constructor(ctx, slideId, id) { super(ctx); this._slideId = slideId; this._id = id }
  get _s() { const sl = this._ctx._model.slide(this._slideId); const s = sl && sl.shapes.find(x => x.id === this._id); if (!s) throw new OfficeError('ItemNotFound', 'shape'); return s }
  get id() { this._need('id'); return this._id }
  get name() { this._need('name'); return this._s.name }
  get type() { this._need('type'); return this._s.type }
  get left() { this._need('left'); return this._s.left }
  get top() { this._need('top'); return this._s.top }
  get width() { this._need('width'); return this._s.width }
  get height() { this._need('height'); return this._s.height }
  // PowerPointApi 1.10
  get altTextDescription() { this._need('altTextDescription'); return this._s.altTextDescription ?? '' }
  set altTextDescription(v) { this._ctx._queue(() => { this._s.altTextDescription = String(v) }) }
  get altTextTitle() { this._need('altTextTitle'); return this._s.altTextTitle ?? '' }
  set altTextTitle(v) { this._ctx._queue(() => { this._s.altTextTitle = String(v) }) }
  get placeholderFormat() { const self = this; return new (class extends Loadable { get type() { this._need('type'); return self._s.placeholderType } })(this._ctx) }
  get textFrame() {
    const ctx = this._ctx; const self = this
    const font = new (class extends Loadable {})(ctx)
    for (const k of ['bold', 'italic', 'color', 'size', 'name']) {
      Object.defineProperty(font, k, { get() { this._need(k); return self._s.font[k] }, set(v) { ctx._queue(() => { self._s.font[k] = v }) } })
    }
    const textRange = new (class extends Loadable {
      get text() { this._need('text'); if (self._s.type === 'Image') throw new OfficeError('InvalidArgument', 'no text frame'); return self._s.text }
      set text(v) { ctx._queue(() => { self._s.text = v }) }
    })(ctx)
    textRange.font = font
    return { textRange }
  }
  delete() { this._ctx._queue(() => { const sl = this._ctx._model.slide(this._slideId); sl.shapes.splice(sl.shapes.indexOf(this._s), 1) }) }
}

class Slide extends Loadable {
  constructor(ctx, id) { super(ctx); this._id = id }
  get _sl() { const s = this._ctx._model.slide(this._id); if (!s) throw new OfficeError('ItemNotFound', 'slide'); return s }
  get id() { this._need('id'); return this._id }
  get shapes() {
    const ctx = this._ctx; const self = this
    // Live (A9 probe): loading a slide's shapes in the same PowerPoint.run that added the slide fails the batch with
    // GeneralException in PowerPoint for the web; a fresh PowerPoint.run reads them fine.
    if (ctx._addedSlides && ctx._addedSlides.has(self._id)) ctx._fail(new OfficeError('GeneralException', 'GeneralException'))
    const col = new LoadableCollection(ctx, () => self._sl.shapes.map(s => new Shape(ctx, self._id, s.id)))
    col.getItem = id => { const p = ctx._model.slide(self._id); if (!p || !p.shapes.some(s => s.id === id)) ctx._fail(new OfficeError('ItemNotFound', `shape ${id}`)); return new Shape(ctx, self._id, id) }
    const add = fields => { const shape = ctx._model.newShape(fields); ctx._queue(() => self._sl.shapes.push(shape)); return new Shape(ctx, self._id, shape.id) }
    col.addTextBox = (text, o = {}) => add({ type: 'TextBox', text, ...o })
    col.addGeometricShape = (shapeType, o = {}) => { if (!/^[A-Za-z]+$/.test(shapeType)) ctx._fail(new OfficeError('InvalidArgument', 'shapeType')); return add({ type: 'GeometricShape', geometry: shapeType, ...o }) }
    return col
  }
  get layout() { const self = this; return new (class extends Loadable { get id() { this._need('id'); return self._sl.layoutId } })(this._ctx) }
  delete() { this._ctx._queue(() => { const m = this._ctx._model; m.slides.splice(m.slides.indexOf(this._sl), 1); m.selected = m.selected.filter(x => x !== this._id) }) }
  moveTo(i) { this._ctx._queue(() => { const m = this._ctx._model; const s = this._sl; m.slides.splice(m.slides.indexOf(s), 1); m.slides.splice(i, 0, s) }) }
  exportAsBase64() { return new ClientResult(this._ctx, () => b64(JSON.stringify(this._sl))) }
  applyLayout(layout) {
    this._ctx._queue(() => {
      const hit = this._ctx._model.layout(layout._layoutId)
      const sl = this._sl
      sl.layoutId = hit.l.id; sl.masterId = hit.m.id
      const keep = sl.shapes.filter(s => s.type !== 'Placeholder' || s.text)
      const have = new Set(keep.filter(s => s.type === 'Placeholder').map(s => s.placeholderType))
      sl.shapes = [...keep, ...hit.l.placeholders.filter(t => !have.has(t)).map(t => this._ctx._model.newShape({ type: 'Placeholder', placeholderType: t, name: t }))]
    })
  }
}

class MockContext extends MockContextBase {
  // A picture from setSelectedDataAsync reads back with id "0" until PowerPoint for the web assigns its real id a
  // moment later (seen live, A8/A9); the mock assigns it two syncs after the insert, then selects the picture.
  async sync() {
    await super.sync()
    const ps = this._model.pendingSlide
    if (ps && this._model.syncs >= ps.at) {
      this._model.selected = this._model.selected.map(x => (x === ps.slide.id ? ps.realId : x))
      ps.slide.id = ps.realId
      this._model.pendingSlide = null
    }
    const pend = this._model.pending
    if (pend && this._model.syncs >= pend.at) {
      pend.shape.id = pend.realId
      this._model.selectedShapes = [[pend.slideId, pend.realId]]
      this._model.pending = null
    }
  }

  constructor(pres) {
    super(); this._model = pres
    const ctx = this
    const q = fn => ctx._queue(fn)
    const slides = new LoadableCollection(ctx, () => pres.slides.map(s => new Slide(ctx, s.id)))
    slides.getItem = id => { if (!pres.slide(id)) ctx._fail(new OfficeError('ItemNotFound', `slide ${id}`)); return new Slide(ctx, id) }
    slides.add = (opts = {}) => q(() => {
      const s = pres.addSlide(opts)
      ;(ctx._addedSlides ||= new Set()).add(s.id)
      // Live (A11 probe): a new slide reads back with a provisional id for a moment, then gets its real one; the
      // provisional id is then unknown to getItem. The mock flips it three syncs later.
      pres.pendingSlide = { slide: s, realId: s.id, at: pres.syncs + (pres.provisionalSyncs ?? 3) }  // A19: settable, live it lasts up to ~2 s
      s.id = `${4123571000 + uid}#123571113`
    })
    const layoutObj = (l, isNull) => Object.assign(new (class extends Loadable { get id() { this._need('id'); return l.id } get name() { this._need('name'); return l.name } })(ctx), { isNullObject: isNull, _layoutId: l.id })
    const masters = new LoadableCollection(ctx, () => pres.masters.map(m => {
      const o = new (class extends Loadable { get id() { this._need('id'); return m.id } get name() { this._need('name'); return m.name } })(ctx)
      const layouts = new LoadableCollection(ctx, () => m.layouts.map(l => layoutObj(l, false)))
      layouts.getItemOrNullObject = id => { const l = m.layouts.find(x => x.id === id); return layoutObj(l || { id, name: '' }, !l) }
      o.layouts = layouts
      return o
    }))
    this.presentation = {
      slides,
      slideMasters: masters,
      getSelectedSlides: () => new LoadableCollection(ctx, () => pres.selected.filter(id => pres.slide(id)).map(id => new Slide(ctx, id))),
      getSelectedShapes: () => new LoadableCollection(ctx, () => pres.selectedShapes.map(([sid, id]) => new Shape(ctx, sid, id))),
      getSelectedTextRangeOrNullObject: () => Object.assign(new (class extends Loadable { get text() { this._need('text'); return '' } })(ctx), { isNullObject: true }),
      // Live: re-selecting the slide a shape is selected on keeps that shape selected; another slide clears it.
      setSelectedSlides: ids => q(() => { if (ids.join() !== pres.selected.join()) pres.selectedShapes = []; pres.selected = [...ids] }),
      pageSetup: new (class extends Loadable { get slideWidth() { this._need('slideWidth'); return 960 } get slideHeight() { this._need('slideHeight'); return 540 } })(ctx),
      insertSlidesFromBase64: (data, { targetSlideId } = {}) => q(() => {
        const src = JSON.parse(unb64(data))
        const copy = { ...src, id: `${++uid}#`, shapes: src.shapes.map(s => ({ ...s, id: String(++uid), font: { ...s.font } })) }
        const at = targetSlideId ? pres.slides.findIndex(s => s.id === targetSlideId) + 1 : 0
        pres.slides.splice(at, 0, copy)
      })
    }
  }
}

const pptStore = persistence('hermes-mock-powerpoint',
  p => ({ uid, masters: p.masters, slides: p.slides, selected: p.selected, selectedShapes: p.selectedShapes }),
  x => { const p = new MockPresentation(); uid = Math.max(uid, x.uid || 0); Object.assign(p, { masters: x.masters, slides: x.slides, selected: x.selected, selectedShapes: x.selectedShapes }); return p })

export function installPowerPointMock(pres) {
  pres = pres || pptStore.load() || new MockPresentation()
  globalThis.__hermesPptTiming = { settleMs: 0, settleQuietMs: 0 }  // the executor's settle/retry waits, instant against the mock
  globalThis.PowerPoint = { run: makeRun(MockContext, pres, () => pptStore.save(pres)) }
  installOffice('PowerPoint', {
    url: 'mock://Presentation1.pptx',
    document: {
      setSelectedDataAsync(data, opts, cb) {
        const sl = pres.slide(pres.selected[0])
        if (!sl || opts.coercionType !== 'image') { cb({ status: 'failed', error: { code: 5001, message: 'no slide' } }); return }
        // Live (A9): with a shape selected and no imageLeft/imageTop, PowerPoint for the web tries to write into that
        // selection and refuses.
        if (opts.imageLeft === undefined && opts.imageTop === undefined && pres.selectedShapes.length) {
          setTimeout(() => cb({ status: 'failed', error: { code: 2001, message: 'Cannot write to the current selection.' } }), 0)
          return
        }
        const shape = pres.newShape({ type: 'Image', left: opts.imageLeft || 0, top: opts.imageTop || 0, width: opts.imageWidth || 100, height: opts.imageHeight || 100 })
        pres.pending = { shape, realId: shape.id, slideId: sl.id, at: pres.syncs + 2 }
        shape.id = '0'
        sl.shapes.push(shape)
        setTimeout(() => cb({ status: 'succeeded' }), 0)
      }
    }
  })
  globalThis.__hermesMockPresentation = pres
  globalThis.__hermesMockFingerprint = () => JSON.stringify(pres.slides.map(sl => ({ l: sl.layoutId, sh: sl.shapes.map(x => [x.type, x.text, x.font, x.placeholderType, x.left, x.top, x.width, x.height, x.altTextDescription]) })))
  return pres
}
