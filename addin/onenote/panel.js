// SPDX-License-Identifier: Apache-2.0
// Live Links / Backlinks panel for the OneNote pane. Read-only: it shows what the paired Hermes
// chat's local index knows about the page that is open (sent back in reply to the pane's own
// page-visit event) and never writes to the notebook. Clicking an entry opens that page through
// the executor's open_page, the same path Hermes uses; the visit watcher then refreshes the panel.
//
// Titles and labels are the user's note text: they are DATA and are only ever set with
// textContent (never innerHTML), so a title like <img onerror=...> renders as those characters.

const MAX_LABEL = 120
const MAX_ID = 300

function text(v, n = MAX_LABEL) {
  const s = String(v == null ? '' : v).replace(/[\u0000-\u001f\u007f]/g, ' ')
  return s.length > n ? `${s.slice(0, n - 1)}…` : s
}

function entries(list) {
  return (Array.isArray(list) ? list : [])
    .filter(x => x && typeof x.pageId === 'string' && x.pageId && x.pageId.length <= MAX_ID)
    .map(x => ({ pageId: x.pageId, label: text(x.label || x.title), title: text(x.title, 200) }))
}

/** Normalize the plugin's panel payload into plain strings (exported for the tests). */
export function panelModel(panel) {
  if (!panel || typeof panel !== 'object' || typeof panel.pageId !== 'string') return null
  const cov = panel.coverage && typeof panel.coverage === 'object' ? panel.coverage : {}
  const n = Number.isFinite(cov.indexed) ? cov.indexed : 0
  const m = Number.isFinite(cov.discovered) ? cov.discovered : 0
  const more = panel.more && typeof panel.more === 'object' ? panel.more : {}
  return {
    pageId: panel.pageId,
    title: text(panel.title, 200),
    links: entries(panel.links),
    backlinks: entries(panel.backlinks),
    unresolved: (Array.isArray(panel.unresolved) ? panel.unresolved : []).map(t => text(t)),
    ambiguous: (Array.isArray(panel.ambiguous) ? panel.ambiguous : []).map(t => text(t)),
    moreLinks: Number.isFinite(more.links) ? more.links : 0,
    moreBacklinks: Number.isFinite(more.backlinks) ? more.backlinks : 0,
    coverage: `indexed ${n} of ${m} pages`,
    hint: text(panel.hint || 'Ask Hermes to run brain_sync to index more pages.', 300)
  }
}

function el(doc, tag, txt, cls) {
  const e = doc.createElement(tag)
  if (txt != null) e.textContent = txt
  if (cls) e.className = cls
  return e
}

function list(doc, heading, items, more, onOpen) {
  const box = el(doc, 'div', null, 'links-group')
  box.append(el(doc, 'h2', heading))
  if (!items.length) {
    box.append(el(doc, 'p', 'none', 'fine'))
    return box
  }
  const ul = el(doc, 'ul')
  for (const it of items) {
    const li = el(doc, 'li')
    const b = el(doc, 'button', it.label, 'link')
    b.type = 'button'
    b.title = `Open “${it.title}” in OneNote`
    b.dataset.pageId = it.pageId
    b.addEventListener('click', () => onOpen(it.pageId))
    li.append(b)
    ul.append(li)
  }
  box.append(ul)
  if (more > 0) box.append(el(doc, 'p', `and ${more} more`, 'fine'))
  return box
}

/**
 * Render the panel into `root` (replacing what was there). Returns the model, or null when the
 * payload is unusable (the panel is then hidden). `onOpen(pageId)` is called on a click.
 */
export function renderLinksPanel(root, panel, { onOpen = () => {}, doc = globalThis.document } = {}) {
  const m = panelModel(panel)
  if (!root) return m
  if (!m) {
    root.hidden = true
    root.replaceChildren()
    return null
  }
  const parts = [
    list(doc, 'Links', m.links, m.moreLinks, onOpen),
    list(doc, 'Backlinks', m.backlinks, m.moreBacklinks, onOpen)
  ]
  if (m.unresolved.length) parts.push(el(doc, 'p', `Not created yet: ${m.unresolved.join(', ')}`, 'fine'))
  if (m.ambiguous.length) parts.push(el(doc, 'p', `Ambiguous (several pages share this title): ${m.ambiguous.join(', ')}`, 'fine'))
  parts.push(el(doc, 'p', `${m.coverage}. ${m.hint}`, 'fine coverage'))
  root.replaceChildren(...parts)
  root.dataset.pageId = m.pageId
  root.hidden = false
  return m
}

/**
 * Wiring used by the pane: `show(res)` takes the bridge's reply to a page-visit event; a click opens
 * the page with the executor's read-only open_page (the user's own navigation, no opId, no write),
 * then asks the visit watcher to pick up the new page, whose reply refreshes the panel.
 */
export function createLinksPanel({ root, execute, refresh = () => {}, doc = globalThis.document }) {
  let status = null
  async function open(pageId) {
    const r = await execute('open_page', { pageId })
    if (!r || !r.ok) {
      if (!status) { status = el(doc, 'p', null, 'fine err'); root.append(status) }
      status.textContent = 'Could not open that page. It may have been moved or deleted; ask Hermes to run brain_sync.'
      return false
    }
    await refresh()
    return true
  }
  function show(res) {
    status = null
    if (!res || !res.accepted || !res.panel) return null
    return renderLinksPanel(root, res.panel, { onOpen: open, doc })
  }
  return { show, open }
}
