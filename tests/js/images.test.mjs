// SPDX-License-Identifier: Apache-2.0
// node --test tests/js/ : office_insert_image pane commands (insert_image + image_status) per host,
// against the strict Office.js mocks. Undo must remove exactly the inserted picture and nothing else.
import assert from 'node:assert/strict'
import { test } from 'node:test'
import { installMock, MockWorkbook, modelFingerprint } from '../../addin/dev/office-mock.js'
import { installWordMock, MockDocument } from '../../addin/dev/word-mock.js'
import { installPowerPointMock, MockPresentation } from '../../addin/dev/powerpoint-mock.js'
import { installOneNoteMock, MockNotebook } from '../../addin/dev/onenote-mock.js'
import * as excel from '../../addin/excel/executor.js'
import * as word from '../../addin/word/executor.js'
import * as ppt from '../../addin/powerpoint/executor.js'
import * as one from '../../addin/onenote/executor.js'
import { imageBase64 } from '../../addin/shared/images.js'

const PNG = 'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR4nGNgYGD4DwABBAEAX+XDSwAAAABJRU5ErkJggg=='
const JPEG = '/9j/4AAQSkZJRgABAQAAAQABAAD/2wBDAP//////////////////////////////////////////////////////////////////////////////////////2wBDAf//////////////////////////////////////////////////////////////////////////////////////wAARCAABAAEDASIAAhEBAxEB/8QAFAABAAAAAAAAAAAAAAAAAAAACP/EABQQAQAAAAAAAAAAAAAAAAAAAAD/2gAMAwEAAhADEAA/AKp//9k='

async function ok(exec, tool, args) { const r = await exec(tool, args); assert.equal(r.ok, true, `${tool}: ${JSON.stringify(r)}`); return r.data }
function supported(requirements) { globalThis.Office.context.requirements = { isSetSupported: (set, v) => requirements(set, v) } }

test('shared base64 guard: PNG/JPEG only, data-URI prefix tolerated, size capped', () => {
  assert.equal(imageBase64(PNG).kind, 'png')
  assert.equal(imageBase64(`data:image/jpeg;base64,${JPEG}`).kind, 'jpeg')
  for (const bad of ['', 'R0lGODlhAQABAAAAACw=', 'PHN2Zz4=', 'not base64!', 42]) assert.throws(() => imageBase64(bad), e => e.category === 'Validation')
  assert.throws(() => imageBase64('iVBORw0KGgo' + 'A'.repeat(6 * 1024 * 1024)), e => e.category === 'Capacity')
})

// ---- Excel ----------------------------------------------------------------------------------------
test('Excel: picture anchored at a cell with alt text, read back by id, undo removes only it', async () => {
  excel.executor.clearUndo()
  const wb = installMock(new MockWorkbook())
  wb.sheets[0].shapes.push({ id: 'user-shape', name: 'Mine', type: 'GeometricShape', left: 1, top: 1, width: 5, height: 5, altTextDescription: '', altTextTitle: '', lockAspectRatio: false })
  const before = modelFingerprint(wb)
  const d = await ok(excel.execute, 'insert_image', { base64: PNG, alt: 'Q4 chart — source: https://img.example/a.png', title: 'Q4 chart', sheet: 'Sheet1', cell: 'C3', width: 120 })
  assert.equal(d.undo.available, true)
  assert.equal(d.sheet, 'Sheet1')
  assert.equal(d.cell, 'C3')
  assert.equal(d.altTextApplied, true)
  const pic = wb.sheets[0].shapes.find(s => s.id === d.shapeId)
  assert.deepEqual([pic.type, pic.left, pic.top, pic.width, pic.altTextDescription, pic.altTextTitle], ['Image', 96, 30, 120, 'Q4 chart — source: https://img.example/a.png', 'Q4 chart'])
  assert.deepEqual(await ok(excel.execute, 'image_status', { sheet: 'Sheet1', shapeId: d.shapeId }), { exists: true, type: 'Image', images: 1 })
  assert.equal((await ok(excel.execute, 'undo_last')).undone, 'insert_image')
  assert.equal(modelFingerprint(wb), before)
  assert.equal((await ok(excel.execute, 'image_status', { sheet: 'Sheet1', shapeId: d.shapeId })).exists, false)
})

test('Excel: active sheet by default; ExcelApi 1.9 required; bad bytes refused before Office', async () => {
  excel.executor.clearUndo()
  const wb = installMock(new MockWorkbook())
  const d = await ok(excel.execute, 'insert_image', { base64: JPEG, alt: 'a', title: 'a' })
  assert.equal(d.sheet, 'Sheet1')
  assert.equal(wb.sheets[0].shapes.length, 1)
  let r = await excel.execute('insert_image', { base64: 'R0lGODlhAQABAAAAACw=', alt: 'a', title: 'a' })
  assert.equal(r.error.category, 'Validation')
  supported((set, v) => !(set === 'ExcelApi' && v === '1.9'))
  r = await excel.execute('insert_image', { base64: PNG, alt: 'a', title: 'a' })
  assert.equal(r.error.category, 'Unsupported')
  assert.equal(wb.sheets[0].shapes.length, 1)
})

// ---- Word -----------------------------------------------------------------------------------------
const wfp = doc => JSON.stringify({ p: doc.paragraphs, c: doc.controls })

test('Word: picture at the end in a tagged content control; status by tag; undo deletes exactly it', async () => {
  word.executor.clearUndo()
  const doc = installWordMock(new MockDocument(['Title', 'Body text.']))
  // The document does not end in an empty paragraph: one is added (Word never deletes the final paragraph mark, so a
  // picture there could not be removed cleanly) and the picture goes in its own paragraph just before it.
  const before = JSON.stringify({ p: [...doc.paragraphs, { text: '', style: 'Normal', isListItem: false }], c: doc.controls })
  const d = await ok(word.execute, 'insert_image', { base64: PNG, alt: 'A fox — Photo: J. Doe', title: 'A fox', location: 'end', tag: 'hermes-image-1', width: 200 })
  assert.equal(d.tag, 'hermes-image-1')
  assert.equal(d.altTextApplied, true)
  assert.deepEqual(doc.paragraphs.map(p => p.text), ['Title', 'Body text.', '\ufffc', ''])
  assert.deepEqual(doc.paragraphs[2].picture, { kind: 'png', altTextDescription: 'A fox — Photo: J. Doe', altTextTitle: 'A fox', lockAspectRatio: true, width: 200 })
  assert.deepEqual(doc.controls.map(c => [c.tag, c.title, c.appearance]), [['hermes-image-1', 'Hermes image', 'BoundingBox']])
  assert.deepEqual(await ok(word.execute, 'image_status', { tag: 'hermes-image-1' }), { exists: true, controls: 1, pictures: 1 })
  assert.equal((await ok(word.execute, 'undo_last')).undone, 'insert_image')
  assert.equal(wfp(doc), before)
  assert.deepEqual(await ok(word.execute, 'image_status', { tag: 'hermes-image-1' }), { exists: false, controls: 0, pictures: 0 })
})

test('Word: HTML at the cursor keeps the user text; start location; guards', async () => {
  word.executor.clearUndo()
  const doc = installWordMock(new MockDocument(['One', 'Two']))
  doc.setSelection(0, 1, 1)
  const before = wfp(doc)
  const d = await ok(word.execute, 'insert_image', { html: '<p>Card <img src="data:image/png;base64,' + PNG + '" alt="x"></p>', alt: 'card', location: 'selection', tag: 'hermes-image-2' })
  assert.equal(d.kind, 'html')
  assert.deepEqual(doc.paragraphs.map(p => p.text), ['One', 'Card ', 'Two'])
  assert.equal((await ok(word.execute, 'image_status', { tag: 'hermes-image-2' })).pictures, 1)
  await ok(word.execute, 'undo_last')
  assert.equal(wfp(doc), before)
  await ok(word.execute, 'insert_image', { base64: PNG, alt: 'a', title: 'a', location: 'start', tag: 'hermes-image-3' })
  assert.equal(doc.paragraphs[0].pictures, 1)
  for (const bad of [{ html: '<script>x</script>', tag: 't' }, { html: '<p onclick="x">y</p>', tag: 't' }, { base64: PNG, tag: '' },
    { base64: PNG, location: 'middle', tag: 't' }]) {
    const r = await word.execute('insert_image', { alt: 'a', title: 'a', ...bad })
    assert.equal(r.error.category, 'Validation', JSON.stringify(bad))
  }
})

test('Word: HTML after an earlier picture at the end lands in its OWN control; undo removes exactly it (A9 live defect)', async () => {
  word.executor.clearUndo()
  const doc = installWordMock(new MockDocument(['HOL live test document.']))
  await ok(word.execute, 'insert_image', { base64: PNG, alt: 'a', title: 'a', location: 'end', tag: 'hermes-image-p' })
  const afterFirst = wfp(doc)
  await ok(word.execute, 'insert_image', { html: '<p><b>Card</b></p><p><img src="data:image/png;base64,' + PNG + '" alt="x"></p>', alt: 'card', location: 'end', tag: 'hermes-image-h' })
  assert.deepEqual(await ok(word.execute, 'image_status', { tag: 'hermes-image-h' }), { exists: true, controls: 1, pictures: 1 })
  assert.deepEqual(await ok(word.execute, 'image_status', { tag: 'hermes-image-p' }), { exists: true, controls: 1, pictures: 1 })
  assert.equal((await ok(word.execute, 'undo_last')).undone, 'insert_image')
  assert.equal(wfp(doc), afterFirst)
  assert.deepEqual(await ok(word.execute, 'image_status', { tag: 'hermes-image-p' }), { exists: true, controls: 1, pictures: 1 })
})

test('Word: a picture after an earlier picture gets its own control too; a trailing empty paragraph is reused, undo is exact', async () => {
  word.executor.clearUndo()
  const doc = installWordMock(new MockDocument(['HOL live test document.', '']))
  const before = wfp(doc)
  await ok(word.execute, 'insert_image', { base64: PNG, alt: 'a', title: 'a', location: 'end', tag: 'hermes-image-1a' })
  await ok(word.execute, 'insert_image', { base64: JPEG, alt: 'b', title: 'b', location: 'end', tag: 'hermes-image-1b' })
  for (const tag of ['hermes-image-1a', 'hermes-image-1b']) assert.deepEqual(await ok(word.execute, 'image_status', { tag }), { exists: true, controls: 1, pictures: 1 })
  assert.equal(doc.paragraphs[doc.paragraphs.length - 1].text, '') // the document still ends in its own empty paragraph
  await ok(word.execute, 'undo_last')
  await ok(word.execute, 'undo_last')
  assert.equal(wfp(doc), before)
})

test('Word: undo refuses when the tagged control is gone (never deletes something else)', async () => {
  word.executor.clearUndo()
  const doc = installWordMock(new MockDocument(['One']))
  await ok(word.execute, 'insert_image', { base64: PNG, alt: 'a', title: 'a', tag: 'hermes-image-4' })
  doc.controls.length = 0 // the user removed the content control by hand
  const r = await word.execute('undo_last')
  assert.equal(r.ok, false)
  assert.equal(r.error.category, 'NotFound')
})

// ---- PowerPoint -----------------------------------------------------------------------------------
function deck() {
  ppt.executor.clearUndo()
  const p = installPowerPointMock(new MockPresentation())
  p.addSlide({ layoutId: 'L-content' })
  return p
}
const pfp = p => JSON.stringify(p.slides.map(s => s.shapes.map(x => [x.id, x.type, x.altTextDescription])))

test('PowerPoint: picture on slide 2 with alt text (PowerPointApi 1.10); status; undo removes only it', async () => {
  const p = deck()
  const before = pfp(p)
  const d = await ok(ppt.execute, 'insert_image', { base64: PNG, alt: 'Logo — source: https://img.example/l.png', title: 'Logo', slideIndex: 1, left: 40, top: 50, width: 100, height: 80 })
  assert.equal(d.slideIndex, 1)
  assert.equal(d.altTextApplied, true)
  const shape = p.slides[1].shapes.find(s => s.id === d.shapeId)
  assert.deepEqual([shape.type, shape.left, shape.top, shape.width, shape.height, shape.altTextDescription, shape.altTextTitle], ['Image', 40, 50, 100, 80, 'Logo — source: https://img.example/l.png', 'Logo'])
  assert.deepEqual(await ok(ppt.execute, 'image_status', { slideId: d.slideId, shapeId: d.shapeId }), { exists: true, type: 'Image' })
  assert.equal((await ok(ppt.execute, 'undo_last')).undone, 'insert_image')
  assert.equal(pfp(p), before)
  assert.equal((await ok(ppt.execute, 'image_status', { slideId: d.slideId, shapeId: d.shapeId })).exists, false)
})

test('PowerPoint: without PowerPointApi 1.10 the picture lands but altTextApplied is false', async () => {
  const p = deck()
  supported((set, v) => !(set === 'PowerPointApi' && v === '1.10'))
  const d = await ok(ppt.execute, 'insert_image', { base64: PNG, alt: 'x', title: 'x', slideIndex: 0 })
  assert.equal(d.altTextApplied, false)
  assert.equal(p.slides[0].shapes.filter(s => s.type === 'Image').length, 1)
  assert.equal(d.undo.available, true)
})

test('PowerPoint: two pictures in a row, no position, the first left selected: both land with real ids; undo removes exactly the second (A9)', async () => {
  const p = deck()
  const before = pfp(p)
  const a = await ok(ppt.execute, 'insert_image', { base64: PNG, alt: 'one', title: 'one', slideIndex: 0 })
  assert.notEqual(a.shapeId, '0')
  assert.deepEqual(await ok(ppt.execute, 'image_status', { slideId: a.slideId, shapeId: a.shapeId }), { exists: true, type: 'Image' })
  const afterFirst = pfp(p)
  const b = await ok(ppt.execute, 'insert_image', { base64: PNG, alt: 'two', title: 'two', slideIndex: 0 })
  assert.notEqual(b.shapeId, a.shapeId)
  assert.deepEqual(await ok(ppt.execute, 'image_status', { slideId: b.slideId, shapeId: b.shapeId }), { exists: true, type: 'Image' })
  assert.equal(p.slides[0].shapes.filter(s => s.type === 'Image').length, 2)
  assert.equal((await ok(ppt.execute, 'undo_last')).undone, 'insert_image')
  assert.equal(pfp(p), afterFirst)
  assert.equal((await ok(ppt.execute, 'undo_last')).undone, 'insert_image')
  assert.equal(pfp(p), before)
})

test('PowerPoint: undo of a picture the user already deleted says NotFound and removes nothing', async () => {
  const p = deck()
  const d = await ok(ppt.execute, 'insert_image', { base64: PNG, alt: 'x', title: 'x', slideIndex: 1 })
  const sl = p.slides[1]
  sl.shapes.splice(sl.shapes.findIndex(s => s.id === d.shapeId), 1)
  const left = pfp(p)
  const r = await ppt.execute('undo_last')
  assert.equal(r.ok, false)
  assert.equal(r.error.category, 'NotFound')
  assert.equal(pfp(p), left)
})

test('PowerPoint: legacy add_image still works unchanged', async () => {
  const p = deck()
  const d = await ok(ppt.execute, 'add_image', { base64Image: PNG, slideIndex: 1 })
  assert.equal(d.format, 'png')
  assert.equal(p.slides[1].shapes.filter(s => s.type === 'Image').length, 1)
})

// ---- OneNote --------------------------------------------------------------------------------------
function notebook() {
  one.clearJournal()
  const nb = new MockNotebook({ id: 'nb-1', name: 'NB', sections: [{ name: 'S', pages: [{ title: 'Meeting', html: ['<p>User notes</p>'] }] }] })
  installOneNoteMock(nb)
  return nb
}

test('OneNote: image outline on the open page, alt text set, status, delete_owned removes exactly it', async () => {
  const nb = notebook()
  const page = nb.activePage()
  const before = nb.fingerprint()
  const html = `<p><img src="data:image/png;base64,${PNG}" alt="Whiteboard"></p><p><i>Photo by Ana — https://img.example/w.png</i></p>`
  const d = await ok(one.execute, 'insert_image', { html, alt: 'Whiteboard — Photo by Ana', pageId: page.id, opId: 'op-img-1', expectNotebookId: 'nb-1' })
  assert.equal(d.images, 1)
  assert.equal(d.altTextApplied, true)
  assert.deepEqual(d.created.outlineIds, [d.outlineId])
  assert.deepEqual(Object.keys(d.created.fingerprints), [d.outlineId])
  const outline = page.outlines.find(o => o.id === d.outlineId)
  assert.deepEqual(outline.paragraphs.map(x => x.type), ['Image', 'RichText'])
  assert.equal(outline.paragraphs[0].image.description, 'Whiteboard — Photo by Ana')
  assert.deepEqual(await ok(one.execute, 'image_status', { pageId: page.id, outlineId: d.outlineId, expectNotebookId: 'nb-1' }), { exists: true, images: 1 })
  await ok(one.execute, 'delete_owned', { pageId: page.id, outlineIds: [d.outlineId], expect: d.created.fingerprints, opId: 'op-undo-1', expectNotebookId: 'nb-1' })
  assert.equal(nb.fingerprint(), before)
  assert.deepEqual(await ok(one.execute, 'image_status', { pageId: page.id, outlineId: d.outlineId }), { exists: false, images: 0 })
})

test('OneNote: needs an opId, the open page and a matching notebook; refuses active markup', async () => {
  const nb = notebook()
  const page = nb.activePage()
  const img = `<p><img src="data:image/png;base64,${PNG}"></p>`
  let r = await one.execute('insert_image', { html: img, pageId: page.id })
  assert.equal(r.error.category, 'Validation')
  r = await one.execute('insert_image', { html: img, pageId: 'other', opId: 'op-a' })
  assert.equal(r.error.code, 'NotActivePage')
  r = await one.execute('insert_image', { html: img, opId: 'op-b', expectNotebookId: 'nb-other' })
  assert.equal(r.error.code, 'IdentityMismatch')
  for (const bad of ['<p><img src="x" onerror="alert(1)"></p>', '<script>x</script>', '<p><a href="javascript:x">y</a></p>', '<iframe src="https://x"></iframe>']) {
    r = await one.execute('insert_image', { html: bad, opId: `op-${bad.length}` })
    assert.equal(r.error.category, 'Validation', bad)
  }
  // a data: image is allowed (the generic append guard is for note text, not for pictures)
  assert.equal((await one.execute('insert_image', { html: img, opId: 'op-c' })).ok, true)
  assert.equal(page.outlines.length, 2)
})

// ---- A6 #4/#5: the pane's second check on plugin-sanitized HTML (fail closed) ----------------------
test('shared pictureHtml: no stylesheet, no URL-bearing style attribute, at least one picture', async () => {
  const { pictureHtml } = await import('../../addin/shared/images.js')
  const img = `<img src="data:image/png;base64,${PNG}" alt="x">`
  // the plugin's normalized output passes
  assert.equal(pictureHtml(`<p><span style="color:red;font-weight:bold">a</span>${img}</p>`).includes('color:red'), true)
  const refused = [
    `<style>@font-face{font-family:x;src:https://attacker.invalid/font.woff2}</style><p>ok</p>${img}`,
    `<style>p{background-image:image-set("https://attacker.invalid/p.png" 1x)}</style><p>ok</p>${img}`,
    `<STYLE>p{color:red}</STYLE>${img}`,
    `<link rel="stylesheet" href="https://attacker.invalid/s.css">${img}`,
    `<p style="background:url(https://attacker.invalid/a)">t</p>${img}`,
    `<p style='background-image:image-set("https://attacker.invalid/p.png" 1x)'>t</p>${img}`,
    `<p style=background:u\\72l(https://attacker.invalid/e)>t</p>${img}`,
    `<p style="background:&#117;rl(https://attacker.invalid/ent)">t</p>${img}`,
    `<p style="src:https://attacker.invalid/f">t</p>${img}`,
    `<p style="width:expression(alert(1))">t</p>${img}`,
    `<p STYLE="color:red;/**/background:x">t</p>${img}`,
    // no picture at all (A6 #5): the picture tool never inserts text-only HTML
    '<p>plain text</p>',
    '<p>an &lt;img&gt; mentioned in text</p>'
  ]
  for (const bad of refused) assert.throws(() => pictureHtml(bad), e => e.category === 'Validation', bad)
})

test('Word and OneNote insert_image refuse text-only HTML before touching the document (A6 #5)', async () => {
  word.executor.clearUndo()
  const doc = installWordMock(new MockDocument(['One']))
  const before = wfp(doc)
  const r = await word.execute('insert_image', { html: '<p>plain text</p>', alt: 'a', tag: 'hermes-image-t' })
  assert.equal(r.ok, false)
  assert.equal(r.error.category, 'Validation')
  assert.equal(wfp(doc), before)
  const nb = notebook()
  const nbBefore = nb.fingerprint()
  const o = await one.execute('insert_image', { html: '<p>plain text</p>', opId: 'op-text-only', expectNotebookId: 'nb-1' })
  assert.equal(o.ok, false)
  assert.equal(o.error.category, 'Validation')
  assert.equal(nb.fingerprint(), nbBefore)
})

test('shared pictureHtml: alt text that reads like a style attribute is not mistaken for one', async () => {
  const { pictureHtml } = await import('../../addin/shared/images.js')
  const h = `<p><img src="data:image/png;base64,${PNG}" alt="code style=url(clean) &amp; more"></p>`
  assert.equal(pictureHtml(h), h)
})
