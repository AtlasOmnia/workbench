// SPDX-License-Identifier: Apache-2.0
// Electron harness that mirrors the Hermes Desktop preview pane (Electron 44.9.3,
// <webview> guest, contextIsolation+sandbox, no allowpopups, no permission handler
// on the preview partition) and loads a simulated "Office for the web" page that
// iframes the real bridge add-in task pane. Network topology is simulated with
// Chromium switches so the Office and add-in origins count as PUBLIC while the
// Hermes bridge is LOOPBACK, which is what triggers Local Network Access.
//
// Spike/test tool only. It never touches the real preview partition or app data.
const { app, BrowserWindow, session } = require('electron')
const fs = require('node:fs')
const path = require('node:path')
const readline = require('node:readline')

const VARIANT = process.env.HARNESS_VARIANT || 'allow-default'
const CODE = process.env.HARNESS_CODE || ''
const OUT = process.env.HARNESS_OUT
const OFFICE_PORT = process.env.HARNESS_OFFICE_PORT || '8443'
const ADDIN_PORT = process.env.HARNESS_ADDIN_PORT || '8444'
const PARTITION = 'spike-preview' // in-memory; NOT persist:hermes-preview
const APP = process.env.HARNESS_APP || 'excel'

app.setPath('userData', path.join(OUT, 'electron-userdata'))
// HARNESS_PAGE_IP=127.0.0.1 (default) relabels the page servers PUBLIC via ip-address-space-overrides;
// a real LAN address (e.g. 10.0.0.1) makes them genuinely LOCAL (private) with no override at all.
const PAGE_IP = process.env.HARNESS_PAGE_IP || '127.0.0.1'
app.commandLine.appendSwitch('host-resolver-rules', `MAP office.test ${PAGE_IP}, MAP addin.test ${PAGE_IP}`)
if (process.env.HARNESS_FEATURES) app.commandLine.appendSwitch('enable-features', process.env.HARNESS_FEATURES)
if (PAGE_IP === '127.0.0.1') app.commandLine.appendSwitch('ip-address-space-overrides', `127.0.0.1:${OFFICE_PORT}=public,127.0.0.1:${ADDIN_PORT}=public`)

const log = obj => process.stdout.write(JSON.stringify({ t: Date.now(), ...obj }) + '\n')

app.whenReady().then(async () => {
  const ses = session.fromPartition(PARTITION)
  ses.setCertificateVerifyProc((req, cb) => cb(/\.test$/.test(req.hostname) ? 0 : -3))
  if (VARIANT === 'allow-deny') {
    ses.setPermissionRequestHandler((_wc, permission, cb, details) => { log({ ev: 'permission-request', permission, origin: details.requestingUrl && new URL(details.requestingUrl).origin, decision: 'deny' }); cb(false) })
    ses.setPermissionCheckHandler((_wc, permission) => { log({ ev: 'permission-check', permission, decision: 'deny' }); return false })
  } else if (VARIANT === 'allow-log') {
    ses.setPermissionRequestHandler((_wc, permission, cb, details) => { log({ ev: 'permission-request', permission, origin: details.requestingUrl && new URL(details.requestingUrl).origin, decision: 'grant' }); cb(true) })
  }
  // *-default variants: no handler at all, exactly like the Desktop preview partition today.

  const win = new BrowserWindow({ width: 1280, height: 860, show: false, webPreferences: { webviewTag: true, contextIsolation: true, sandbox: true } })
  win.webContents.setWindowOpenHandler(() => ({ action: 'deny' }))
  const iframeAllow = VARIANT.startsWith('noallow') ? 'none' : 'allow'
  const seed = APP === 'onenote' ? '&seed=demo' : ''
  const src = `https://office.test:${OFFICE_PORT}/office.html?iframe=${iframeAllow}&addin=${encodeURIComponent(`https://addin.test:${ADDIN_PORT}/office/live/${APP}/taskpane.html?mock=1${seed}`)}`
  const host = `<!doctype html><html><body style="margin:0"><webview id="w" partition="${PARTITION}" src="${src}" webpreferences="contextIsolation=yes,nodeIntegration=no,sandbox=yes" style="width:100vw;height:100vh;display:flex"></webview></body></html>`
  fs.writeFileSync(path.join(OUT, 'host.html'), host)

  app.on('web-contents-created', (_e, wc) => {
    if (wc.getType() !== 'webview') return
    wc.setWindowOpenHandler(() => ({ action: 'deny' }))
    wc.on('console-message', (details) => {
      const d = details && typeof details === 'object' ? details : {}
      log({ ev: 'console', level: d.level, message: String(d.message || '').slice(0, 400), frame: d.frame && d.frame.url ? new URL(d.frame.url).origin : undefined })
    })
    wc.on('did-finish-load', () => drive(wc))
  })
  await win.loadFile(path.join(OUT, 'host.html'))
  log({ ev: 'host-loaded', variant: VARIANT })

  const rl = readline.createInterface({ input: process.stdin })
  rl.on('line', async line => {
    const msg = JSON.parse(line)
    if (msg.cmd === 'screenshot') {
      const img = await win.webContents.capturePage()
      fs.writeFileSync(msg.path, img.toPNG())
      log({ ev: 'screenshot', path: msg.path })
    } else if (msg.cmd === 'eval') {
      log({ ev: 'eval', id: msg.id, value: await paneEval(msg.js) })
    } else if (msg.cmd === 'reload') {
      await paneEval('(setTimeout(() => location.reload(), 30), true)')
      await new Promise(r => setTimeout(r, 800))
      let st = null
      for (let i = 0; i < 120; i++) {
        st = await paneEval('globalThis.__hermesPaneState || null')
        if (st === 'paired') break
        await new Promise(r => setTimeout(r, 250))
      }
      log({ ev: 'reloaded', state: st })
    } else if (msg.cmd === 'workbook') {
      log({ ev: 'workbook', data: await paneEval(`(() => { const wb = globalThis.__hermesMockWorkbook; return wb ? { sheets: wb.sheets.map(s => ({ name: s.name, used: s.usedBox(), cells: [...s.cells].filter(([,c]) => c.v !== null || c.f !== null).map(([k, c]) => [k, c.f || c.v]) })), active: wb.active, tables: wb.tables, charts: wb.charts.map(c => ({ name: c.name, type: c.type, title: c.title.text })) } : null })()`) })
    } else if (msg.cmd === 'quit') {
      app.quit()
    }
  })
})

let guest = null
function paneFrame() {
  if (!guest) return null
  return guest.mainFrame.framesInSubtree.find(f => { try { const u = new URL(f.url); return u.hostname === 'addin.test' && u.pathname.endsWith('/taskpane.html') } catch { return false } }) || null
}
async function paneEval(js) {
  const f = paneFrame()
  if (!f) return { error: 'no pane frame' }
  try { return await f.executeJavaScript(js, true) } catch (e) { return { error: String(e.message || e) } }
}

let driven = false
async function drive(wc) {
  if (driven) return
  driven = true
  guest = wc
  const outer = await wc.mainFrame.executeJavaScript(`({ crossOriginIsolated, allow: document.querySelector('iframe')?.getAttribute('allow') || null })`)
  log({ ev: 'office-page', ...outer })
  // Wait for the task pane to reach "ready".
  for (let i = 0; i < 60; i++) {
    const s = await paneEval('globalThis.__hermesPaneState || null')
    if (s === 'ready') break
    await new Promise(r => setTimeout(r, 250))
  }
  const policy = await paneEval(`({ state: globalThis.__hermesPaneState, origin: location.origin, isSecureContext, lnaPolicy: document.featurePolicy ? { lna: document.featurePolicy.allowsFeature('local-network-access'), loopback: document.featurePolicy.allowsFeature('loopback-network') } : 'n/a' })`)
  log({ ev: 'pane-ready', ...policy })
  // Simulate the human typing the code from the Hermes chat and pressing Pair.
  await paneEval(`(() => { const i = document.getElementById('code'); i.value = ${JSON.stringify(CODE)}; i.dispatchEvent(new Event('input')); document.getElementById('pair-form').requestSubmit(); return true })()`)
  log({ ev: 'code-entered' })
  let last = null
  for (let i = 0; i < 80; i++) {
    const s = await paneEval(`({ state: globalThis.__hermesPaneState, status: document.getElementById('status').textContent, net: globalThis.__hermesBridgeClient && globalThis.__hermesBridgeClient.lastNetworkError || null })`)
    if (JSON.stringify(s) !== JSON.stringify(last)) { log({ ev: 'pane-state', ...s }); last = s }
    if (s.state === 'paired' || s.state === 'no-bridge' || s.state === 'rejected') break
    await new Promise(r => setTimeout(r, 250))
  }
  log({ ev: 'pairing-settled', state: last && last.state })
}
