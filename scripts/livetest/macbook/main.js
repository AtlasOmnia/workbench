// SPDX-License-Identifier: Apache-2.0
// Live-test harness for the client Mac (client only). Mirrors the Hermes Desktop preview pane:
// Electron 44.9.3, a <webview> guest with contextIsolation + sandbox, no allowpopups and no
// permission handler on its partition. Unlike the spike harness it loads REAL Office for the
// web, in a PERSISTENT partition of its own (userData under ~/hermes-office-livetest), so a
// one-time human sign-in survives between runs. It never opens the real Hermes Desktop or
// Chrome profiles and never reads cookies or tokens.
//
// Tabs (demo loop): several <webview> guests in this one window, one per Office document, like the
// Desktop preview pane's tabs. One is visible; hidden ones stay attached (visibility:hidden), so the
// loop can check what a hidden pane does. Commands target the current tab, or `tab: <name>` explicitly.
//
// Control: a loopback-only JSON endpoint (127.0.0.1:HOL_CONTROL_PORT, default 9334) that the
// Studio-side driver reaches through an ssh -L forward. Commands are mechanical (navigate,
// eval in a frame, trusted typing/clicks, CDP file chooser, screenshots); the driver decides.
const { app, BrowserWindow, session, webContents } = require('electron')
const fs = require('node:fs')
const http = require('node:http')
const path = require('node:path')
const { DEFAULT_PARTITION, bare, harnessHash, openTabScript, partitionFor } = require('./tabscript.js')

const HOL_VERSION = '2026-11-06.4'
// The harness files this window was started from; run_macbook.sh restarts a window whose hash
// differs from the files it syncs (a stale window from an older main.js keeps running otherwise).
const HOL_HASH = harnessHash(__dirname)
const ROOT = process.env.HOL_ROOT || path.join(require('node:os').homedir(), 'hermes-office-livetest')
const PORT = Number(process.env.HOL_CONTROL_PORT || 9334)
// Loopback-only DevTools endpoint: Electron's in-process debugger does not expose the file-chooser
// commands, so the harness talks CDP to itself here (never reachable from another machine).
const CDP_PORT = Number(process.env.HOL_CDP_PORT || 9335)
// Per-tab partitions: tabscript.js partitionFor (outlook* tabs get persist:hol-outlook, the rest share one jar).
const TAB_PARTITION_OVERRIDES = process.env.HOL_TAB_PARTITIONS || ''
const START_URL = process.env.HOL_START_URL || 'https://excel.cloud.microsoft/'
const LOG = path.join(ROOT, 'logs', 'electron.log')

fs.mkdirSync(path.join(ROOT, 'logs'), { recursive: true })
app.setPath('userData', path.join(ROOT, 'userdata'))
app.setPath('sessionData', path.join(ROOT, 'userdata'))
app.commandLine.appendSwitch('remote-debugging-port', String(CDP_PORT))
app.commandLine.appendSwitch('remote-debugging-address', '127.0.0.1')
app.commandLine.appendSwitch('disable-renderer-backgrounding')
app.commandLine.appendSwitch('disable-background-timer-throttling')
app.commandLine.appendSwitch('disable-backgrounding-occluded-windows')

const log = obj => fs.appendFileSync(LOG, JSON.stringify({ t: new Date().toISOString(), ...obj }) + '\n')
// URLs are logged and returned through bare() (tabscript.js): no query, fragment or OneDrive personal path.

// A stray rejection (a CDP socket closing mid-navigation) must not take the harness window down.
process.on('uncaughtException', e => log({ ev: 'uncaught', error: String((e && e.message) || e).slice(0, 300) }))
process.on('unhandledRejection', e => log({ ev: 'unhandled-rejection', error: String((e && e.message) || e).slice(0, 300) }))
app.on('render-process-gone', (_e, _wc, d) => log({ ev: 'render-process-gone', reason: d.reason }))
app.on('child-process-gone', (_e, d) => log({ ev: 'child-process-gone', type: d.type, reason: d.reason }))
app.on('window-all-closed', () => { log({ ev: 'window-all-closed' }); app.quit() })
let win = null
let guest = null  // the current tab's guest webContents (the target of commands without `tab`)
let current = null
const tabs = new Map()  // name -> { wc, el, partition }
const partitions = new Set([DEFAULT_PARTITION])  // every jar a tab of this window has used
const fileChooser = { result: null }

function hostHtml() {
  return `<!doctype html><html><head><meta charset="utf-8"><title>HOL live test</title><style>
  html,body{margin:0;height:100%;font:13px -apple-system,system-ui,sans-serif;background:#1e1f24;color:#e8e8ea}
  #wrap{display:flex;height:100vh}
  #chat{width:330px;flex:none;display:flex;flex-direction:column;border-right:1px solid #333}
  #chat h1{font-size:13px;margin:0;padding:10px 12px;border-bottom:1px solid #333;color:#9ecbff}
  #lines{flex:1;overflow:auto;padding:8px 12px}
  #lines div{margin:0 0 8px;white-space:pre-wrap;word-break:break-word}
  #lines .user{color:#ffd479} #lines .tool{color:#8fd18f;font-family:ui-monospace,monospace;font-size:11px}
  #lines .note{color:#bbb;font-style:italic}
  #pane{flex:1;position:relative;display:flex}
  #pane webview{position:absolute;top:0;left:0;width:100%;height:100%;visibility:hidden}
  #pane webview.on{visibility:visible}
  #tabs{display:flex;flex-wrap:wrap;gap:4px;padding:6px 12px;border-bottom:1px solid #333}
  #tabs span{padding:2px 8px;border-radius:4px;background:#2a2b31;color:#999;font-size:11px}
  #tabs span.on{background:#3b5b8a;color:#fff}
  #mask{position:absolute;top:0;right:0;width:360px;height:56px;background:#222;color:#888;display:none;
        align-items:center;justify-content:center;font-size:11px;z-index:9;pointer-events:none}
  #gate{position:absolute;left:0;right:0;bottom:0;padding:10px 14px;background:#7a4b00;color:#fff;z-index:8;display:none}
  </style></head><body><div id="wrap"><div id="chat"><h1>Hermes chat (live-test harness)</h1><div id="tabs"></div><div id="lines"></div></div>
  <div id="pane"><div id="gate"></div><div id="mask">account area hidden</div></div></div></body></html>`
}

app.whenReady().then(async () => {
  win = new BrowserWindow({ width: 1500, height: 940, title: 'HOL live test', webPreferences: { webviewTag: true, contextIsolation: true, sandbox: true, backgroundThrottling: false } })
  win.webContents.setWindowOpenHandler(() => ({ action: 'deny' }))
  app.on('web-contents-created', (_e, wc) => {
    if (wc.getType() !== 'webview') return
    wc.setWindowOpenHandler(d => { log({ ev: 'popup-denied', url: bare(d.url) }); return { action: 'deny' } })
    wc.on('did-navigate', (_e2, url) => log({ ev: 'navigate', url: bare(url) }))
  })
  const file = path.join(ROOT, 'logs', 'host.html')
  fs.writeFileSync(file, hostHtml())
  await win.loadFile(file)
  log({ ev: 'host-loaded', partition: DEFAULT_PARTITION, electron: process.versions.electron })
  http.createServer(onRequest).listen(PORT, '127.0.0.1', () => log({ ev: 'control-listening', port: PORT }))
  try { await openTab('main', START_URL) } catch { /* logged as tab-open-failed */ }
})

// ---- CDP over the loopback DevTools endpoint: the file chooser -------------------------
// armFiles: attach to every page/iframe/webview target, intercept the next file chooser any of
// them opens and fill it with the given local files (DOM.setFileInputFiles). No OS dialog.
async function cdpArm(paths, timeoutMs) {
  const ver = await (await fetch(`http://127.0.0.1:${CDP_PORT}/json/version`)).json()
  const ws = new WebSocket(ver.webSocketDebuggerUrl)
  await new Promise((res, rej) => { ws.onopen = res; ws.onerror = () => rej(new Error('cdp ws error')) })
  let id = 0
  const pending = new Map()
  const send = (method, params = {}, sessionId) => new Promise((res, rej) => {
    const msg = { id: ++id, method, params }
    if (sessionId) msg.sessionId = sessionId
    pending.set(msg.id, { res, rej })
    ws.send(JSON.stringify(msg))
  })
  const sessions = new Map()
  let finish
  const done = new Promise(r => { finish = r })
  const attach = async info => {
    if (!['page', 'iframe', 'webview'].includes(info.type) || [...sessions.values()].includes(info.targetId)) return
    try {
      const { sessionId } = await send('Target.attachToTarget', { targetId: info.targetId, flatten: true })
      sessions.set(sessionId, info.targetId)
      await send('Page.enable', {}, sessionId).catch(() => {})
      await send('Page.setInterceptFileChooserDialog', { enabled: true }, sessionId)
    } catch (e) { log({ ev: 'cdp-attach-failed', type: info.type, error: String(e.message || e).slice(0, 120) }) }
  }
  ws.onmessage = async m => {
    const d = JSON.parse(m.data)
    if (d.id && pending.has(d.id)) { const p = pending.get(d.id); pending.delete(d.id); d.error ? p.rej(new Error(d.error.message)) : p.res(d.result); return }
    if (d.method === 'Target.targetCreated') attach(d.params.targetInfo)
    if (d.method === 'Page.fileChooserOpened') {
      log({ ev: 'file-chooser-opened', mode: d.params.mode })
      try {
        await send('DOM.setFileInputFiles', { files: paths, backendNodeId: d.params.backendNodeId }, d.sessionId)
        finish({ ok: true, files: paths.map(p => path.basename(p)) })
      } catch (e) { finish({ ok: false, error: String(e.message || e) }) }
    }
  }
  await send('Target.setDiscoverTargets', { discover: true })
  const { targetInfos } = await send('Target.getTargets')
  for (const t of targetInfos) await attach(t)
  log({ ev: 'file-chooser-armed', sessions: sessions.size })
  setTimeout(() => finish({ ok: false, error: 'no file chooser opened' }), timeoutMs || 30000)
  const result = done.then(async r => {
    for (const sid of sessions.keys()) await send('Page.setInterceptFileChooserDialog', { enabled: false }, sid).catch(() => {})
    ws.close()
    return r
  })
  return { armed: sessions.size, result }
}

// ---- tabs --------------------------------------------------------------------------------
const elId = name => 'w-' + String(name).replace(/[^A-Za-z0-9_-]/g, '_')
const opening = new Map()  // name -> in-flight openTab promise (a second open of the same tab joins it)
function openTab(name, url, partition) {
  if (tabs.has(name)) return Promise.resolve(tabs.get(name))
  if (!opening.has(name)) opening.set(name, openTabOnce(name, url, partition).finally(() => opening.delete(name)))
  return opening.get(name)
}
async function openTabOnce(name, url, explicitPartition) {
  const el = elId(name)
  const part = partitionFor(name, explicitPartition, TAB_PARTITION_OVERRIDES)
  let id
  try {
    // tabscript.js explains why the id is read on dom-ready and why this always settles.
    id = await win.webContents.executeJavaScript(openTabScript({ el, name, partition: part, url }))
  } catch (e) {
    // A tab that never produced a guest leaves nothing behind, so a later open starts clean.
    await win.webContents.executeJavaScript(`(() => { const w = document.getElementById(${JSON.stringify(el)}); if (w) w.remove();
      for (const s of document.querySelectorAll('#tabs span')) if (s.dataset.tab === ${JSON.stringify(name)}) s.remove(); return true })()`).catch(() => {})
    log({ ev: 'tab-open-failed', tab: name, error: String((e && e.message) || e).slice(0, 200) })
    throw e
  }
  const wc = webContents.fromId(id)
  if (!wc) throw new Error(`tab ${name}: no guest webContents`)
  const t = { wc, el, partition: part }
  tabs.set(name, t)
  partitions.add(part)
  log({ ev: 'tab-opened', tab: name, partition: part, url: bare(url) })
  if (!current) await showTab(name)
  return t
}
async function showTab(name) {
  const t = tabs.get(name)
  if (!t) throw new Error(`no tab ${name}`)
  await win.webContents.executeJavaScript(`(() => {
    for (const w of document.querySelectorAll('#pane webview')) w.classList.toggle('on', w.id === ${JSON.stringify(t.el)})
    for (const s of document.querySelectorAll('#tabs span')) s.classList.toggle('on', s.dataset.tab === ${JSON.stringify(name)})
    return true })()`)
  guest = t.wc
  current = name
  log({ ev: 'tab-shown', tab: name })
}
async function closeTab(name) {
  const t = tabs.get(name)
  if (!t) return
  tabs.delete(name)
  cdpSessions.delete(t.wc.id)
  await win.webContents.executeJavaScript(`(() => { const w = document.getElementById(${JSON.stringify(t.el)}); if (w) w.remove();
    for (const s of document.querySelectorAll('#tabs span')) if (s.dataset.tab === ${JSON.stringify(name)}) s.remove(); return true })()`)
  if (current === name) { current = null; guest = null; const next = tabs.keys().next(); if (!next.done) await showTab(next.value) }
  log({ ev: 'tab-closed', tab: name })
}
function tabInfo() {
  return { current, tabs: [...tabs.entries()].map(([name, t]) => ({ name, visible: name === current, partition: t.partition, url: t.wc.isDestroyed() ? null : bare(t.wc.getURL()) })) }
}
// The guest a command targets: `tab` names one (visible or not); without it, the current (visible) tab.
function target(msg) {
  if (msg && msg.tab) {
    const t = tabs.get(msg.tab)
    if (!t) throw new Error(`no tab ${msg.tab}`)
    return t.wc
  }
  if (!guest) throw new Error('no guest yet')
  return guest
}

// ---- frames ------------------------------------------------------------------------------
function frames(wc) {
  if (!wc || wc.isDestroyed()) return []
  return wc.mainFrame.framesInSubtree
}
function findFrame(wc, match) {
  if (!match || match === 'top') return wc && wc.mainFrame
  return frames(wc).find(f => { try { return f.url.includes(match) } catch { return false } }) || null
}
function withTimeout(promise, ms, what) {
  let t
  return Promise.race([promise, new Promise((_, rej) => { t = setTimeout(() => rej(new Error(what)), ms) })]).finally(() => clearTimeout(t))
}
async function evalIn(wc, match, js, gesture = true) {
  const f = findFrame(wc, match)
  if (!f) return { __error: `no frame matching ${match}` }
  try { return await withTimeout(f.executeJavaScript(js, gesture), 20000, 'frame script timed out (frame detached?)') } catch (e) { return { __error: String(e.message || e).slice(0, 500) } }
}

// ---- commands ----------------------------------------------------------------------------
const sleep = ms => new Promise(r => setTimeout(r, ms))

// Trusted input goes through CDP Input.* on the <webview> guest target (sendInputEvent did not
// reach Office's handlers); mouse events hit-test into OOPIFs, key events go to the focused frame.
// With several tabs, the guest's DevTools target is found by a nonce set in its main world.
const cdpSessions = new Map()  // webContents id -> { ws, send }
async function openCdp(url) {
  const ws = new WebSocket(url)
  await new Promise((res, rej) => { ws.onopen = res; ws.onerror = () => rej(new Error('cdp ws error')) })
  let id = 0
  const pending = new Map()
  ws.onmessage = m => { const d = JSON.parse(m.data); const p = pending.get(d.id); if (p) { pending.delete(d.id); d.error ? p.rej(new Error(d.error.message)) : p.res(d.result) } }
  const send = (method, params = {}) => withTimeout(new Promise((res, rej) => { const i = ++id; pending.set(i, { res, rej }); ws.send(JSON.stringify({ id: i, method, params })) }), 10000, 'cdp input timed out')
  return { ws, send }
}
async function guestSession(wc) {
  const hit = cdpSessions.get(wc.id)
  if (hit && hit.ws.readyState === 1) return hit
  const nonce = 'hol-' + Math.random().toString(36).slice(2)
  await withTimeout(wc.executeJavaScript(`(window.__holNonce = ${JSON.stringify(nonce)}, true)`, false), 10000, 'tab nonce timed out')
  const list = await (await fetch(`http://127.0.0.1:${CDP_PORT}/json/list`)).json()
  for (const t of list.filter(x => x.type === 'webview' && x.webSocketDebuggerUrl)) {
    let s
    try { s = await openCdp(t.webSocketDebuggerUrl) } catch { continue }
    let v = null
    try { v = ((await s.send('Runtime.evaluate', { expression: 'window.__holNonce', returnByValue: true })).result || {}).value } catch {}
    if (v === nonce) {
      const send = (m, p) => s.send(m, p).catch(e => { cdpSessions.delete(wc.id); try { s.ws.close() } catch {} throw e })
      const sess = { ws: s.ws, send }
      cdpSessions.set(wc.id, sess)
      return sess
    }
    try { s.ws.close() } catch {}
  }
  throw new Error('no webview target for this tab')
}

async function trustedClick(wc, x, y) {
  const { send } = await guestSession(wc)
  const at = { x: Math.round(x), y: Math.round(y) }
  await send('Input.dispatchMouseEvent', { type: 'mouseMoved', ...at })
  await sleep(60)
  await send('Input.dispatchMouseEvent', { type: 'mousePressed', ...at, button: 'left', clickCount: 1 })
  await sleep(80)
  await send('Input.dispatchMouseEvent', { type: 'mouseReleased', ...at, button: 'left', clickCount: 1 })
}

const KEYS = { Enter: { key: 'Enter', code: 'Enter', windowsVirtualKeyCode: 13, text: '\r' }, Escape: { key: 'Escape', code: 'Escape', windowsVirtualKeyCode: 27 },
  Tab: { key: 'Tab', code: 'Tab', windowsVirtualKeyCode: 9 } }

// Offset of a frame's viewport inside the guest: walk up, finding each iframe element by name
// (WebFrameMain.name), else the only iframe, else the one whose src origin matches.
async function frameOffset(f) {
  let x = 0, y = 0
  while (f.parent) {
    const parent = f.parent
    const r = await withTimeout(parent.executeJavaScript(`(() => { const all = [...document.querySelectorAll('iframe')];
      const byName = all.filter(i => i.name && i.name === ${JSON.stringify(f.name || '')});
      let el = byName[0] || (all.length === 1 ? all[0] : null);
      if (!el) { const o = ${JSON.stringify((() => { try { return new URL(f.url).origin } catch { return '' } })())}; el = all.find(i => { try { return new URL(i.src, location.href).origin === o } catch { return false } }) }
      // The content box starts after the border AND the padding (Outlook's add-in dialog iframe has padding-top 17px).
      if (!el) return null; const b = el.getBoundingClientRect(), cs = getComputedStyle(el);
      return { x: b.x + el.clientLeft + (parseFloat(cs.paddingLeft) || 0), y: b.y + el.clientTop + (parseFloat(cs.paddingTop) || 0) } })()`, false), 10000, 'parent frame timed out')
    if (!r) throw new Error('cannot locate iframe element for a frame')
    x += r.x; y += r.y
    f = parent
  }
  return { x, y }
}
const COMMANDS = {
  async version() { return { version: HOL_VERSION, hash: HOL_HASH, electron: process.versions.electron, tabs: tabs.size } },
  // The add-in site sends no Cache-Control, so Chromium caches its scripts heuristically: every run
  // starts from the bytes actually served (sign-in cookies are untouched).
  async clearCache() { for (const p of partitions) await session.fromPartition(p).clearCache(); return { ok: true, partitions: [...partitions] } },
  // Tabs: open (creates it if missing, never navigates an existing one), show (make it visible and the
  // default target), close, list. The Office tabs share one partition, so the one-time sign-in covers them;
  // outlook* tabs (or an explicit persist:hol-* `partition`) use their own jar (tabscript.js partitionFor).
  async tab({ op, name, url, partition }) {
    if (op === 'open') { await openTab(name, url, partition); return tabInfo() }
    if (op === 'show') { await showTab(name); return tabInfo() }
    if (op === 'close') { await closeTab(name); return tabInfo() }
    if (op === 'list') return tabInfo()
    throw new Error(`tab: unknown op ${op}`)
  },
  async state(msg) {
    const wc = msg.tab ? target(msg) : guest
    return { url: wc ? bare(wc.getURL()) : null, title: wc ? wc.getTitle() : null, loading: wc ? wc.isLoading() : null,
      frames: frames(wc).map(f => ({ url: bare(f.url), name: f.name })), tab: msg.tab || current, tabs: tabInfo().tabs }
  },
  async navigate(msg) { const wc = target(msg); await wc.loadURL(msg.url).catch(e => ({ error: String(e.message || e) })); return { url: bare(wc.getURL()) } },
  // Disposable test documents: full URLs stay in ~/hermes-office-livetest/state.json on this machine only.
  async docs(msg) {
    const { op, host } = msg
    const wc = target(msg)
    const file = path.join(ROOT, 'state.json')
    const st = fs.existsSync(file) ? JSON.parse(fs.readFileSync(file, 'utf8')) : {}
    if (op === 'get') return { known: !!st[host] }
    if (op === 'remember') { st[host] = wc.getURL(); fs.writeFileSync(file, JSON.stringify(st, null, 2), { mode: 0o600 }); return { ok: true } }
    if (op === 'forget') { delete st[host]; fs.writeFileSync(file, JSON.stringify(st, null, 2), { mode: 0o600 }); return { ok: true } }
    if (op === 'open') { if (!st[host]) throw new Error(`no stored document for ${host}`); await wc.loadURL(st[host]).catch(() => {}); return { url: bare(wc.getURL()) } }
    throw new Error(`docs: unknown op ${op}`)
  },
  async reloadFrame(msg) { return { value: await evalIn(target(msg), msg.frame, '(setTimeout(() => location.reload(), 30), true)') } },
  async eval(msg) { return { value: await evalIn(target(msg), msg.frame, msg.js, msg.gesture !== false) } },
  async host({ js }) { return { value: await win.webContents.executeJavaScript(js, true) } },
  async gate({ text }) {
    await win.webContents.executeJavaScript(`(() => { const g = document.getElementById('gate'); g.textContent = ${JSON.stringify(text || '')}; g.style.display = ${JSON.stringify(text ? 'block' : 'none')}; return true })()`)
    return { ok: true }
  },
  async chat({ cls, text, clear }) {
    await win.webContents.executeJavaScript(`(() => { const l = document.getElementById('lines'); if (${!!clear}) l.textContent = '';
      if (${JSON.stringify(text || '')}) { const d = document.createElement('div'); d.className = ${JSON.stringify(cls || 'agent')}; d.textContent = ${JSON.stringify(String(text || '').slice(0, 1500))}; l.appendChild(d); l.scrollTop = l.scrollHeight } return true })()`)
    return { ok: true }
  },
  async focus() { win.show(); win.focus(); guest && guest.focus(); return { ok: true } },
  async type(msg) {
    const { send } = await guestSession(target(msg))
    for (const ch of String(msg.text)) { await send('Input.insertText', { text: ch }); await sleep(30) }
    return { ok: true }
  },
  async key(msg) {
    const { send } = await guestSession(target(msg))
    const k = KEYS[msg.keyCode] || { key: msg.keyCode, text: msg.keyCode }
    await send('Input.dispatchKeyEvent', { type: 'keyDown', ...k })
    await send('Input.dispatchKeyEvent', { type: 'keyUp', ...k, text: undefined })
    return { ok: true }
  },
  async click(msg) { await trustedClick(target(msg), msg.x, msg.y); return { ok: true } },
  // Arm the CDP file chooser with local files; the next chooser any frame opens gets them.
  async armFiles({ paths, timeoutMs }) {
    for (const p of paths) if (!fs.existsSync(p)) return { ok: false, error: `missing ${path.basename(p)}` }
    const { armed, result } = await cdpArm(paths, timeoutMs)
    fileChooser.result = result
    return { ok: true, sessions: armed }
  },
  async awaitFiles() {
    if (!fileChooser.result) return { ok: false, error: 'not armed' }
    const r = await fileChooser.result
    fileChooser.result = null
    return r
  },
  // Trusted click on an element found by `js` (returns {x, y} in that frame's viewport).
  async clickIn(msg) {
    const wc = target(msg)
    const f = findFrame(wc, msg.frame)
    if (!f) throw new Error(`no frame matching ${msg.frame}`)
    const pt = await withTimeout(f.executeJavaScript(msg.js, true), 20000, 'frame script timed out')
    if (!pt || typeof pt.x !== 'number') return { ok: false, found: false }
    const off = await frameOffset(f)
    await trustedClick(wc, pt.x + off.x, pt.y + off.y)
    return { ok: true, x: Math.round(pt.x + off.x), y: Math.round(pt.y + off.y), label: pt.label || null }
  },
  async screenshot({ file, mask }) {
    if (mask !== false) await win.webContents.executeJavaScript(`document.getElementById('mask').style.display='flex'`)
    await sleep(150)
    const img = await win.webContents.capturePage()
    if (mask !== false) await win.webContents.executeJavaScript(`document.getElementById('mask').style.display='none'`)
    const out = path.join(ROOT, 'shots', file)
    fs.mkdirSync(path.dirname(out), { recursive: true })
    fs.writeFileSync(out, img.toPNG())
    return { ok: true, file: out }
  },
  async quit() { setTimeout(() => app.quit(), 50); return { ok: true } }
}

function onRequest(req, res) {
  if (req.socket.remoteAddress !== '127.0.0.1' && req.socket.remoteAddress !== '::ffff:127.0.0.1') { res.writeHead(403); return res.end() }
  let body = ''
  req.on('data', c => { body += c; if (body.length > 1e6) req.destroy() })
  req.on('end', async () => {
    let out
    try {
      const msg = JSON.parse(body || '{}')
      const fn = COMMANDS[msg.cmd]
      if (!fn) throw new Error(`unknown cmd ${msg.cmd}`)
      if (!guest && !msg.tab && !['quit', 'state', 'version', 'clearCache', 'tab', 'gate', 'chat', 'host', 'screenshot'].includes(msg.cmd)) throw new Error('no guest yet')
      out = await fn(msg)
    } catch (e) { out = { __error: String(e.message || e) } }
    const data = JSON.stringify(out === undefined ? null : out)
    res.writeHead(200, { 'Content-Type': 'application/json' })
    res.end(data)
  })
}
