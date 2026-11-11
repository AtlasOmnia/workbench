// SPDX-License-Identifier: Apache-2.0
// OneNote task pane entry: shared pairing/bridge UI + the OneNote executor + visit-indexing.
//
// Visit-indexing: while paired, when the user opens a different page (DocumentSelectionChanged
// if OneNote fires it, else a ~2 s poll of the active page id), the pane reads THAT page (the
// only page an add-in may read) and posts it to the paired Hermes chat, which indexes it in a
// local SQLite file. The pane never navigates on its own and never writes for this.
import { startPane } from '../shared/taskpane.js'
import { SUPPORTED_TOOLS, activePageId, execute } from './executor.js'
import { POLL_MS, createVisitWatcher } from './visits.js'
import { createLinksPanel } from './panel.js'

let watcher = null

startPane({
  app: 'onenote',
  hostType: 'OneNote',
  label: 'OneNote',
  execute,
  SUPPORTED_TOOLS,
  describeDocument: async () => {
    const ids = await activePageId()
    return ids ? { id: ids.notebookId, name: 'OneNote notebook' } : { name: 'OneNote notebook' }
  },
  onPaired: (client, { $ }) => {
    if (watcher) return
    // Live Links / Backlinks for the open page, from the reply to each visit event (read-only).
    const links = createLinksPanel({ root: $('links-panel'), execute, refresh: () => watcher.tick().catch(() => null) })
    watcher = createVisitWatcher({
      client,
      onIndexed: (res, page) => {
        links.show(res)
        const el = $('brain')
        if (!el || !res || !res.accepted) return
        el.hidden = false
        // A snapshot from the last page the user opened, not a live counter (Hermes' own reads and crawls
        // also index pages); say so instead of showing a number that silently goes stale.
        const title = page && page.title ? `“${String(page.title).slice(0, 80)}”` : 'this page'
        el.textContent = res.coverage
          ? `Indexed ${title} on this computer. At that moment Hermes had indexed ${res.coverage.indexed} of ${res.coverage.discovered} known pages.`
          : `Indexed ${title} on this computer.`
      }
    })
    const kick = () => { watcher.tick().catch(() => {}) }
    try {
      const O = globalThis.Office
      if (O && O.context && O.context.document && typeof O.context.document.addHandlerAsync === 'function' && O.EventType) {
        O.context.document.addHandlerAsync(O.EventType.DocumentSelectionChanged, kick, () => {})
      }
    } catch { /* event not available in this host: the poll below covers it */ }
    setInterval(kick, POLL_MS)
    kick()
  },
  // Dev/test only (?mock=1). dev/ is excluded from deploy builds, so this import cannot load in production.
  installMock: async () => {
    const m = await import('../dev/onenote-mock.js')
    if (globalThis.OneNote) return
    let nb
    if (!sessionStorage.getItem('hermes-mock-onenote') && new URLSearchParams(location.search).get('seed') === 'demo') {
      nb = new m.MockNotebook(await (await fetch('../dev/demo-notebook.json')).json())
    }
    m.installOneNoteMock(nb)
  }
})
