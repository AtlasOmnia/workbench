// SPDX-License-Identifier: Apache-2.0
// Shared task-pane controller for the Excel, Word and PowerPoint bridge panes:
// load Office.js (or the dev mock), resume or pair with the Hermes chat, hand
// commands to the host executor, and show plain-language state and errors.
import { BINDING_IDLE_MS, BridgeClient } from './bridge-client.js'
import { BUILD } from './build-info.js'

const OFFICE_JS = 'https://appsforoffice.microsoft.com/lib/1/hosted/office.js'

function loadScript(src) {
  return new Promise((resolve, reject) => {
    const s = document.createElement('script')
    s.src = src
    s.onload = resolve
    s.onerror = () => reject(new Error(`Could not load Office.js from ${src}. Check the network, then reload this pane.`))
    document.head.appendChild(s)
  })
}

/** Plain-language text for bridge/Office error categories shown in the activity list: what happened + what to do. */
export const ERROR_TEXT = {
  Validation: 'The document could not accept that request. Hermes can read the document and try a corrected request.',
  Office: 'Office refused the change. Nothing was changed by this step; check the document, then ask Hermes again.',
  Partial: 'Only part of this step went through. Hermes was told exactly what changed and what did not; check the document before asking again.',
  Unsupported: 'This Office version cannot do that from an add-in. Do it by hand, or ask Hermes for another way.',
  Conflict: 'That would clash with something already in the document (for example a name in use). Ask Hermes to pick another.',
  NotFound: 'Hermes could not find that sheet, range, slide or page. It may have been renamed or deleted.',
  NothingToUndo: 'Nothing left to undo in this pane (undo history is cleared when the pane reloads).',
  Capacity: 'That is too large for one step. Ask Hermes to split it up.',
  ExecutionUncertain: 'Unclear whether that change applied. Look at the document before asking for anything else; Hermes will not retry it.',
  Unavailable: 'Office was busy or not ready. Wait a moment, then ask Hermes again.',
  Protocol: 'The pane and Hermes disagreed about a message. Reload this pane; it reconnects to the same chat.',
  PaneReload: 'OneNote stopped answering this add-in. Reload this browser tab (or close and reopen the Workbench pane); it reconnects to the same chat.'
}

export function describeError(err) {
  if (!err) return ''
  const base = ERROR_TEXT[err.category] || 'The command failed. Ask Hermes what happened.'
  const code = err.code && err.code !== 'OfficeError' ? ` (${err.code})` : ''
  return `${base}${code}`
}

/** "write_range" -> "Write range": a readable action name for the activity list (the tool id stays in the title). */
export function actionName(tool) {
  const t = String(tool || '').replace(/_/g, ' ').trim()
  return t ? t[0].toUpperCase() + t.slice(1) : 'Command'
}

/** One activity-list line for a finished command. */
export function activityText(s) {
  if (s.ok) return `${actionName(s.last)} ✓`
  return `${actionName(s.last)} ✗ — ${describeError(s.error)}`
}

/**
 * Pane footer: the add-in build this pane actually loaded, for stale-cache diagnosis. An unbuilt checkout says
 * so; a build from a tree with uncommitted changes is labelled, never shown as a clean release.
 */
export function buildLabel(info = BUILD) {
  if (!info || !info.build) return 'Workbench add-in · unbuilt source (dev)'
  const parts = [`Workbench add-in ${info.version || '?'}`, `build ${info.build}`]
  if (info.commit) parts.push(`source ${info.commit}${info.dirty ? ' + uncommitted changes' : ''}`)
  else parts.push('source unknown')
  return parts.join(' · ')
}

/** Document setting Office reads to reopen this add-in's task pane whenever the document opens. */
export const AUTO_SHOW_SETTING = 'Office.AutoShowTaskpaneWithDocument'

/**
 * F3: after the user pairs (their consent for this document), ask Office to reopen the Hermes pane
 * with the document. Needs document settings (Excel, Word, PowerPoint); OneNote has none -> false.
 */
export async function enableAutoShow(office = globalThis.Office) {
  const settings = office && office.context && office.context.document && office.context.document.settings
  if (!settings || typeof settings.set !== 'function' || typeof settings.saveAsync !== 'function') return false
  try {
    if (settings.get(AUTO_SHOW_SETTING) === true) return true
    settings.set(AUTO_SHOW_SETTING, true)
    const ok = await new Promise(resolve => settings.saveAsync(r => resolve(!!r && r.status === ((office.AsyncResultStatus || {}).Succeeded || 'succeeded'))))
    return ok
  } catch { return false }
}

export function startPane({ app, hostType, label, execute, SUPPORTED_TOOLS, installMock, describeDocument, onPaired }) {
  const params = new URLSearchParams(location.search)
  const $ = id => document.getElementById(id)
  let autoShowTried = false
  const MESSAGES = {
    starting: 'Starting…',
    ready: `Not connected yet. In your Hermes chat, ask "connect to ${label}", then type the 8-character code it shows.`,
    resuming: 'Reconnecting to your Hermes chat…',
    pairing: 'Looking for Hermes on this computer…',
    'no-bridge': 'Hermes is not answering on this computer. Check that Hermes is running with the hermes-office-desktop plugin enabled. If the browser asked to allow access to local devices, choose Allow, then press Pair again.',
    rejected: 'That code was not accepted. Codes work once and expire after 10 minutes; ask Hermes for a new one.',
    'rate-limited': 'Too many codes were tried. Wait a minute, then ask Hermes for a new code.',
    paired: 'Connected to your Hermes chat. Changes are listed below as Hermes makes them.',
    working: 'Hermes is working…',
    reconnecting: 'Connection to Hermes lost; retrying by itself. Your document is not affected.',
    'waiting-bridge': 'Hermes paused this pane\'s connection while it was idle. It reconnects by itself the next time Hermes uses Office in your chat, or ask Hermes for a new code.',
    unpaired: 'Not connected. Ask Hermes for a new code to reconnect.',
    'wrong-host': `This pane only works inside ${label}. Open it from ${label}.`
  }

  function setState(s) {
    const st = $('status')
    st.dataset.state = s.state
    let text = s.state === 'working' ? `Hermes is working: ${actionName(s.tool).toLowerCase()}…` : (MESSAGES[s.state] || s.state)
    if (s.state === 'rejected' && s.category === 'RateLimited') text = MESSAGES['rate-limited']
    if (s.state === 'paired' && s.resumed) text = 'Reconnected to your Hermes chat after the pane reloaded.'
    if (s.state === 'unpaired' && s.reason === 'expired') text = 'The previous pairing has expired or Hermes restarted. Ask Hermes for a new code.'
    if (s.state === 'unpaired' && s.reason === 'revoked') text = 'Hermes disconnected this pane (another pane or chat took over, or the chat ended). Ask Hermes for a new code to reconnect.'
    st.textContent = text
    const paired = ['paired', 'working', 'reconnecting'].includes(s.state)
    if (paired) $('code').value = '' // single-use; never leave it on screen
    $('pair-form').hidden = paired || s.state === 'resuming'
    $('disconnect').hidden = !paired
    if (s.last) {
      const li = document.createElement('li')
      li.textContent = activityText(s)
      li.title = String(s.last || '')
      if (!s.ok) li.className = 'err'
      $('activity').prepend(li)
      while ($('activity').children.length > 8) $('activity').lastChild.remove()
    }
    globalThis.__hermesPaneState = s.state
    if (s.state === 'paired' && !autoShowTried) { autoShowTried = true; enableAutoShow().then(v => { globalThis.__hermesAutoShow = v }) }
    if (s.state === 'paired' && onPaired && globalThis.__hermesBridgeClient) onPaired(globalThis.__hermesBridgeClient, { $ })
  }

  async function loadOffice() {
    if (params.get('mock') === '1') {
      // Dev/test only; dev/ is never deployed, so this import fails in production.
      await installMock()
    } else {
      await loadScript(OFFICE_JS)
    }
    return globalThis.Office.onReady()
  }

  async function main() {
    const build = $('build')
    if (build) build.textContent = buildLabel()
    globalThis.__hermesPaneBuild = BUILD
    setState({ state: 'starting' })
    let info
    try {
      info = await loadOffice()
    } catch (err) {
      $('status').textContent = String(err.message || err)
      $('status').dataset.state = 'error'
      return
    }
    if (info.host !== hostType) return setState({ state: 'wrong-host' })
    const client = new BridgeClient({
      app,
      execute,
      onState: setState,
      describe: async () => ({
        supportedTools: SUPPORTED_TOOLS,
        host: { host: hostType, platform: String(info.platform || ''), version: String(globalThis.Office.context?.diagnostics?.version || '') },
        document: describeDocument ? await describeDocument() : { name: String(globalThis.Office.context?.document?.url || '').split(/[\\/]/).pop().split('?')[0].slice(0, 200) }
      })
    })
    globalThis.__hermesBridgeClient = client
    $('code').addEventListener('input', e => {
      const raw = e.target.value.toUpperCase().replace(/[^0-9A-Z]/g, '').slice(0, 8)
      e.target.value = raw.length > 4 ? `${raw.slice(0, 4)}-${raw.slice(4)}` : raw
    })
    $('pair-form').addEventListener('submit', e => {
      e.preventDefault()
      // Called from a user gesture: lets the browser show its local-access permission prompt.
      client.pair($('code').value)
    })
    $('disconnect').addEventListener('click', () => client.disconnect())
    setState({ state: 'ready' })
    // A14: a bridge that gave its idle port back comes back on its next Office tool call, maybe on another port.
    const resumed = await client.resume({ waitMs: BINDING_IDLE_MS })
    if (!resumed.ok && !['Unavailable', 'Superseded'].includes(resumed.category) && globalThis.__hermesPaneState !== 'unpaired') setState({ state: 'ready' })
  }

  main()
}
