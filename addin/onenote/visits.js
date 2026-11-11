// SPDX-License-Identifier: Apache-2.0
// Visit watcher for the OneNote pane: when the user opens another page, read THAT page (the only
// one an add-in may read) and post it to the paired Hermes chat for local indexing. Read-only.
import { activePageId, captureActive } from './executor.js'

export const POLL_MS = 2000

/** Visit watcher; exported for the tests (tick() runs one check). */
export function createVisitWatcher({ client, capture = captureActive, active = activePageId, onIndexed = () => {} }) {
  let last = null
  let busy = false
  async function tick() {
    if (busy || !client.token) return null
    busy = true
    try {
      const now = await active()
      if (!now || !now.pageId || now.pageId === last) return null
      const page = await capture()
      if (!page || page.pageId !== now.pageId) return null // the user moved on mid-read; next tick
      last = now.pageId
      const res = await client.postEvent({ kind: 'page-visit', document: { id: page.notebookId }, page })
      onIndexed(res, page)
      return res
    } finally {
      busy = false
    }
  }
  return { tick, reset: () => { last = null } }
}
