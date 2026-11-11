// SPDX-License-Identifier: Apache-2.0
// Minimal Chrome DevTools Protocol client for live testing a WebView2 task pane (Node >= 22, no deps).
//   node cdp.mjs targets <port>
//   node cdp.mjs eval <port> <url-substring> <file.js | -e "expr">   (awaits promises, prints JSON)
import { readFileSync } from 'node:fs'

const [cmd, port, match, src, expr] = process.argv.slice(2)
const list = async () => (await fetch(`http://127.0.0.1:${port}/json/list`)).json()

async function evaluate(wsUrl, expression) {
  const ws = new WebSocket(wsUrl)
  await new Promise((res, rej) => { ws.onopen = res; ws.onerror = () => rej(new Error('ws error')) })
  let timer
  const out = await new Promise((res, rej) => {
    ws.onmessage = m => { const d = JSON.parse(m.data); if (d.id === 1) res(d) }
    ws.send(JSON.stringify({ id: 1, method: 'Runtime.evaluate', params: { expression, awaitPromise: true, returnByValue: true, timeout: 120000 } }))
    timer = setTimeout(() => rej(new Error('cdp timeout')), 130000)
  }).finally(() => clearTimeout(timer))
  ws.close()
  if (out.result && out.result.exceptionDetails) throw new Error(JSON.stringify(out.result.exceptionDetails).slice(0, 800))
  return out.result && out.result.result ? out.result.result.value : out
}

try {
  if (cmd === 'targets') {
    console.log(JSON.stringify((await list()).map(t => ({ type: t.type, url: t.url, title: t.title }))))
  } else if (cmd === 'eval') {
    const t = (await list()).find(x => x.url.includes(match) && x.webSocketDebuggerUrl)
    if (!t) { console.log(JSON.stringify({ cdpError: `no target matching ${match}` })); process.exit(2) }
    const code = src === '-e' ? expr : readFileSync(src, 'utf8')
    console.log(JSON.stringify(await evaluate(t.webSocketDebuggerUrl, code)))
  } else {
    console.error('usage: cdp.mjs targets <port> | eval <port> <match> <file|-e expr>'); process.exit(64)
  }
  process.exit(0)
} catch (e) {
  console.log(JSON.stringify({ cdpError: String(e.message || e) }))
  process.exit(1)
}
