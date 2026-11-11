// SPDX-License-Identifier: Apache-2.0
// Excel task pane entry: shared pairing/bridge UI + the Excel executor.
import { startPane } from '../shared/taskpane.js'
import { SUPPORTED_TOOLS, execute } from './executor.js'

startPane({
  app: 'excel',
  hostType: 'Excel',
  label: 'Excel',
  execute,
  SUPPORTED_TOOLS,
  // Dev/test only (?mock=1). dev/ is excluded from deploy builds, so this import cannot load in production.
  installMock: async () => { const m = await import('../dev/office-mock.js'); if (!globalThis.Excel) m.installMock() }
})
