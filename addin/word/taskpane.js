// SPDX-License-Identifier: Apache-2.0
// Word task pane entry: shared pairing/bridge UI + the Word executor.
import { startPane } from '../shared/taskpane.js'
import { SUPPORTED_TOOLS, execute } from './executor.js'

startPane({
  app: 'word',
  hostType: 'Word',
  label: 'Word',
  execute,
  SUPPORTED_TOOLS,
  // Dev/test only (?mock=1). dev/ is excluded from deploy builds, so this import cannot load in production.
  installMock: async () => { const m = await import('../dev/word-mock.js'); if (!globalThis.Word) m.installWordMock() }
})
