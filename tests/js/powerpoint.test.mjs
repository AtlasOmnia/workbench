// SPDX-License-Identifier: Apache-2.0
// node --test tests/js/ : PowerPoint executor against the strict Office.js mock.
import assert from 'node:assert/strict'
import { test } from 'node:test'
import { installPowerPointMock, MockPresentation } from '../../addin/dev/powerpoint-mock.js'
import { execute, executor, SUPPORTED_TOOLS, MUTATING } from '../../addin/powerpoint/executor.js'
// office_insert_image pane commands (plugin-level tool; undo cases live in images.test.mjs)
const IMAGE_COMMANDS = ['insert_image', 'image_status']

const CANONICAL = ['get_presentation_overview', 'read_selected_slides', 'read_all_slides', 'create_slide', 'create_text_slide', 'add_text_box',
  'update_slide_text', 'add_speaker_notes', 'add_shape', 'read_slide_content', 'delete_slide', 'duplicate_slide', 'reorder_slide',
  'find_and_replace_text', 'list_slide_layouts', 'set_slide_layout', 'format_text', 'add_image']
const PNG = 'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR4nGNgYGD4DwABBAEAX+XDSwAAAABJRU5ErkJggg=='

function fresh() {
  executor.clearUndo()
  const p = installPowerPointMock(new MockPresentation())
  p.slides[0].shapes[0].text = 'Hermes deck'
  p.slides[0].shapes[1].text = 'Live from the chat'
  p.addSlide({ layoutId: 'L-content' })
  p.slides[1].shapes[0].text = 'Agenda'
  p.slides[1].shapes[1].text = 'One\nTwo'
  return p
}
const fp = p => JSON.stringify(p.slides.map(s => ({ l: s.layoutId, sh: s.shapes.map(x => [x.type, x.text, x.font, x.placeholderType]) })))
async function ok(tool, args = {}) { const r = await execute(tool, args); assert.equal(r.ok, true, `${tool}: ${JSON.stringify(r)}`); return r.data }

test('pane advertises all 18 PowerPoint tools plus undo_last', () => {
  assert.deepEqual([...SUPPORTED_TOOLS].sort(), [...CANONICAL, ...IMAGE_COMMANDS, 'undo_last'].sort())
})

test('read tools: overview, all slides, selected slides, slide content, layouts', async () => {
  fresh()
  const o = await ok('get_presentation_overview')
  assert.equal(o.slideCount, 2)
  assert.equal(o.slides[0].title, 'Hermes deck')
  assert.deepEqual(o.selectedSlideIndexes, [0])
  assert.deepEqual((await ok('read_all_slides')).slides[1].text, ['Agenda', 'One\nTwo'])
  assert.equal((await ok('read_selected_slides')).slides[0].index, 0)
  const c = await ok('read_slide_content', { slideIndex: 1 })
  assert.equal(c.shapes.length, 2)
  const l = await ok('list_slide_layouts')
  assert.equal(l.slideMasters[0].layouts.length, 4)
  const t = await ok('read_all_slides', { maxLength: 40 })
  assert.equal(t.truncated, true)
})

const shapeOf = (p, i, k) => p.slides[i].shapes[k].id
const CASES = [
  ['create_slide', () => ({ layoutId: 'L-blank' })],
  ['create_text_slide', () => ({ title: 'Plan', bullets: ['Ship', 'Measure'] })],
  ['create_text_slide', () => ({ title: 'Blank', body: 'Body', layoutId: 'L-blank' })],
  ['add_text_box', () => ({ text: 'Note', slideIndex: 1, left: 10, top: 10, width: 200, height: 40 })],
  ['update_slide_text', p => ({ text: 'New agenda', slideIndex: 1, shapeId: shapeOf(p, 1, 0) })],
  ['add_shape', () => ({ shapeType: 'Rectangle', slideIndex: 0 })],
  ['delete_slide', () => ({ slideIndex: 0 })],
  ['duplicate_slide', () => ({ slideIndex: 1 })],
  ['reorder_slide', () => ({ fromIndex: 1, toIndex: 0 })],
  ['find_and_replace_text', () => ({ findText: 'o', replaceText: '0', slideIndex: 1 })],
  ['set_slide_layout', () => ({ slideIndex: 1, layoutId: 'L-titleonly' })],
  ['format_text', p => ({ slideIndex: 0, shapeId: shapeOf(p, 0, 0), font: { bold: true, color: '#FF0000', size: 40 } })],
  ['add_image', () => ({ base64Image: PNG, slideIndex: 1, left: 20, top: 20, width: 50, height: 50 })]
]
for (const [tool, argsOf] of CASES) {
  test(`${tool} applies and undo restores the presentation`, async () => {
    const p = fresh()
    const before = fp(p)
    const d = await ok(tool, argsOf(p))
    assert.equal(d.undo.available, true, JSON.stringify(d.undo))
    assert.notEqual(fp(p), before)
    assert.equal((await ok('undo_last')).undone, tool)
    assert.equal(fp(p), before, `${tool} undo mismatch`)
  })
}

test('every mutating PowerPoint tool has an undo case or is documented Unsupported', () => {
  const covered = new Set(CASES.map(c => c[0]))
  for (const t of MUTATING) assert.ok(covered.has(t) || t === 'add_speaker_notes' || IMAGE_COMMANDS.includes(t), t)
})

test('guards: last slide, bad indexes, speaker notes, bad image, missing shape', async () => {
  const p = fresh()
  await ok('delete_slide', { slideIndex: 1 })
  let r = await execute('delete_slide', { slideIndex: 0 })
  assert.equal(r.error.category, 'Conflict')
  r = await execute('reorder_slide', { fromIndex: 0, toIndex: 5 })
  assert.equal(r.error.category, 'Validation')
  r = await execute('add_speaker_notes', { text: 'x' })
  assert.equal(r.error.category, 'Unsupported')
  r = await execute('add_image', { base64Image: 'bm90IGFuIGltYWdl', slideIndex: 0 })
  assert.equal(r.error.category, 'Validation')
  r = await execute('format_text', { slideIndex: 0, shapeId: 'nope', font: { bold: true } })
  assert.equal(r.error.code, 'ItemNotFound')
  assert.equal(p.slides.length, 1)
})

// A19 (probed live): a new slide's provisional id can outlast the settle reads; the fill then meets it at its getItem
// (GeneralException live, ItemNotFound here). The id is read again and the fill retried; the receipt and the undo
// carry the slide's real id.
test('create_text_slide fills the new slide even when its provisional id outlasts the settle reads', async () => {
  const p = fresh()
  p.provisionalSyncs = 6
  const before = p.slides.length
  const d = await ok('create_text_slide', { title: 'Late id', bullets: ['One', 'Two'] })
  assert.equal(p.slides.length, before + 1)
  const s = p.slides[d.slideIndex]
  assert.equal(d.slideId, s.id)
  assert.ok(!/#123571113$/.test(d.slideId), `receipt carries the provisional id ${d.slideId}`)
  assert.ok(s.shapes.some(x => x.text === 'Late id'), JSON.stringify(s.shapes.map(x => x.text)))
  assert.ok(s.shapes.some(x => String(x.text).includes('One')))
  await ok('undo_last')
  assert.equal(p.slides.length, before)
})

