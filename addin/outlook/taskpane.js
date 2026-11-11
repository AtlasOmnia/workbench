// SPDX-License-Identifier: Apache-2.0
// Outlook task pane entry: shared pairing/bridge UI + the Outlook (Mailbox) executor.
import { startPane } from '../shared/taskpane.js'
import { SUPPORTED_TOOLS, execute, executor, itemMode } from './executor.js'

startPane({
  app: 'outlook',
  hostType: 'Outlook',
  label: 'Outlook',
  execute,
  SUPPORTED_TOOLS,
  // No item details at pairing: only that this is Outlook mail or calendar and whether the item is read or written.
  describeDocument: async () => {
    const it = globalThis.Office?.context?.mailbox?.item
    const calendar = String(it?.itemType || '').toLowerCase() === 'appointment'
    return { name: calendar ? 'Outlook calendar' : 'Outlook mail', mode: it ? itemMode(it) : 'none' }
  },
  // The pane is pinnable and follows the selected message: undo history belongs to the message it was made on,
  // so a different item clears it (the executor also refuses an undo whose item is no longer open).
  onPaired: () => {
    if (globalThis.__hermesOutlookItemHandler) return
    globalThis.__hermesOutlookItemHandler = true
    const mb = globalThis.Office?.context?.mailbox
    const ev = globalThis.Office?.EventType?.ItemChanged
    if (mb && ev && typeof mb.addHandlerAsync === 'function') {
      try { mb.addHandlerAsync(ev, () => executor.clearUndo()) } catch { /* older Outlook: undo still checks the item */ }
    }
  },
  // Dev/test only (?mock=1). dev/ is excluded from deploy builds, so this import cannot load in production.
  installMock: async () => { const m = await import('../dev/outlook-mock.js'); if (!globalThis.Office) m.installOutlookMock() }
})
