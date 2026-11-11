// SPDX-License-Identifier: Apache-2.0
// Pure helpers for the live-test harness (no Electron import, so node --test can load them).
//
// openTabScript: the host-page script that adds one <webview> tab and resolves with its guest
// webContents id. <webview>.getWebContentsId() THROWS until the guest's dom-ready ("The WebView
// must be attached to the DOM and the dom-ready event emitted before this method can be called"),
// and did-attach comes before dom-ready. So the id is tried on did-attach and again on dom-ready,
// a throw only means "not yet", and the deadline stays armed until the promise has settled: the
// script always resolves or rejects, it never hangs.
//
// harnessHash: a short sha256 over the harness files, so run_macbook.sh can tell a window that
// still runs older code (compared to the same files it is about to sync) and restart it.
const crypto = require('node:crypto')
const fs = require('node:fs')
const path = require('node:path')

const HARNESS_FILES = ['main.js', 'tabscript.js']

// bare: what the harness logs or returns for a URL. Never a query string or fragment (Office URLs carry document
// and drive ids), never a OneDrive / SharePoint personal path (it carries the account's drive id or user name: A8
// found them in `tab list` URLs), and no other 16-hex-digit drive id in a path segment.
const PERSONAL_HOSTS = /(^|\.)(onedrive\.live\.com|docs\.live\.net|1drv\.ms|sharepoint\.com)$/i
// Outlook on the web: everything after /mail/ can carry a message, folder or conversation id.
const OUTLOOK_HOSTS = /^outlook\.(live\.com|office\.com|office365\.com|cloud\.microsoft)$/i
function bare(u) {
  let x
  try { x = new URL(u) } catch { return String(u || '').slice(0, 80) }
  if (PERSONAL_HOSTS.test(x.hostname)) return x.origin + '/<path>'
  if (OUTLOOK_HOSTS.test(x.hostname)) {
    const first = x.pathname.split('/')[1] || ''
    const rest = x.pathname.slice(first.length + 2)
    return x.origin + '/' + (first ? first + '/' : '') + (rest ? '<path>' : '')
  }
  const p = x.pathname.replace(/\/personal\/[^/]+/gi, '/personal/<cid>').replace(/\/[0-9a-f]{16}(?=\/|$)/gi, '/<cid>')
  return x.origin + p
}

// Partitions (cookie jars). Every Office tab shares one persistent jar, so the one-time sign-in covers them all.
// Outlook signs a DIFFERENT throwaway account in (one with an Outlook.com mailbox), so outlook* tabs get a jar of
// their own: signing that account in can never replace the Excel/Word/PowerPoint/OneNote account session.
// Overrides: an explicit `partition` on tab open, or HOL_TAB_PARTITIONS="name=persist:hol-x,..." (exact tab
// names); both must be a persist:hol-* name, so no in-memory jar and no jar of anything else on the machine.
const DEFAULT_PARTITION = 'persist:hol-livetest'
const TAB_PARTITIONS = Object.freeze({ outlook: 'persist:hol-outlook' })
const PARTITION_RE = /^persist:hol-[a-z0-9][a-z0-9-]{0,40}$/
function checkPartition(p) {
  if (typeof p !== 'string' || !PARTITION_RE.test(p)) throw new Error(`bad partition ${JSON.stringify(String(p)).slice(0, 60)} (want persist:hol-<name>)`)
  return p
}
function partitionFor(name, explicit, overrides) {
  if (explicit != null) return checkPartition(explicit)
  const n = String(name || '')
  for (const pair of String(overrides || '').split(',').map(x => x.trim()).filter(Boolean)) {
    const i = pair.indexOf('=')
    if (i > 0 && pair.slice(0, i).trim() === n) return checkPartition(pair.slice(i + 1).trim())
  }
  if (/^outlook([-_].*)?$/.test(n)) return TAB_PARTITIONS.outlook
  return DEFAULT_PARTITION
}

function harnessHash(dir) {
  const h = crypto.createHash('sha256')
  for (const f of HARNESS_FILES) h.update(fs.readFileSync(path.join(dir, f)))
  return h.digest('hex').slice(0, 12)
}

function openTabScript({ el, name, partition, url, timeoutMs = 60000 }) {
  const j = JSON.stringify
  return `new Promise((resolve, reject) => {
    let done = false
    const settle = (fn, v) => { if (done) return; done = true; clearTimeout(t); fn(v) }
    const t = setTimeout(() => settle(reject, new Error('webview guest id not available within ${Math.round(timeoutMs / 1000)} s (no dom-ready)')), ${Number(timeoutMs)})
    try {
      const w = document.createElement('webview'); w.id = ${j(el)}
      w.setAttribute('partition', ${j(partition)})
      w.setAttribute('webpreferences', 'contextIsolation=yes,nodeIntegration=no,sandbox=yes')
      w.setAttribute('src', ${j(url || 'about:blank')})
      const tryId = () => { let id; try { id = w.getWebContentsId() } catch { return } if (typeof id === 'number' && id > 0) settle(resolve, id) }
      w.addEventListener('did-attach', tryId)
      w.addEventListener('dom-ready', tryId)
      document.getElementById('pane').appendChild(w)
      const s = document.createElement('span'); s.dataset.tab = ${j(name)}; s.textContent = ${j(name)}
      document.getElementById('tabs').appendChild(s)
    } catch (e) { settle(reject, e) }
  })`
}

module.exports = { DEFAULT_PARTITION, HARNESS_FILES, TAB_PARTITIONS, bare, harnessHash, openTabScript, partitionFor }
