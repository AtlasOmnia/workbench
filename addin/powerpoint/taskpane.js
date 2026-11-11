// SPDX-License-Identifier: Apache-2.0
// PowerPoint task pane entry: shared pairing/bridge UI + the PowerPoint executor.
import { startPane } from '../shared/taskpane.js'
import { SUPPORTED_TOOLS, execute } from './executor.js'

startPane({
  app: 'powerpoint',
  hostType: 'PowerPoint',
  label: 'PowerPoint',
  execute,
  SUPPORTED_TOOLS,
  // Dev/test only (?mock=1). dev/ is excluded from deploy builds, so this import cannot load in production.
  installMock: async () => { const m = await import('../dev/powerpoint-mock.js'); if (!globalThis.PowerPoint) m.installPowerPointMock() }
})
