// SPDX-License-Identifier: Apache-2.0
// node --test tests/js/ : pane BridgeClient reconnect / resume / one-pane logic against a fake bridge.
import assert from 'node:assert/strict'
import { test } from 'node:test'
import { BridgeClient } from '../../addin/shared/bridge-client.js'

class Store { constructor() { this.m = new Map() } getItem(k) { return this.m.has(k) ? this.m.get(k) : null } setItem(k, v) { this.m.set(k, String(v)) } removeItem(k) { this.m.delete(k) } }

function fakeBridge() {
  const b = { tokens: new Set(), queue: [], results: [], pairs: [], down: false, calls: [] }
  b.fetch = async (url, init = {}) => {
    const u = new URL(url)
    b.calls.push(u.pathname)
    if (b.down) throw new TypeError('Failed to fetch')
    const auth = ((init.headers || {}).Authorization || '').replace('Bearer ', '')
    const body = init.body ? JSON.parse(init.body) : null
    const json = (status, v) => ({ status, json: async () => v })
    if (u.port !== '47615') throw new TypeError('Failed to fetch')
    switch (u.pathname) {
      case '/v1/hello': return json(200, { ok: true, service: 'hermes-office-desktop', protocol: 1, apps: ['excel', 'word', 'powerpoint'] })
      case '/v1/pair': {
        b.pairs.push(body)
        if (body.code !== 'GOOD-CODE') return json(401, { ok: false, error: { category: 'Authentication' } })
        if (body.previousToken) b.tokens.delete(body.previousToken)
        const t = `tok-${b.pairs.length}`; b.tokens.add(t)
        return json(200, { ok: true, paneToken: t, bindingId: 'bind' })
      }
      case '/v1/resume': return b.tokens.has(auth) ? json(200, { ok: true, bindingId: 'bind' }) : json(401, { ok: false, error: { category: 'Authentication' } })
      case '/v1/next': {
        if (!b.tokens.has(auth)) return json(401, { ok: false })
        await new Promise(r => setTimeout(r, 5))
        return json(200, { ok: true, command: b.queue.shift() || null })
      }
      case '/v1/result': b.results.push(body); return json(200, { ok: true })
      case '/v1/bye': b.tokens.delete(auth); return json(200, { ok: true })
    }
    return json(404, {})
  }
  return b
}

const until = async (fn, ms = 2000) => { const end = Date.now() + ms; while (Date.now() < end) { if (fn()) return true; await new Promise(r => setTimeout(r, 5)) } return false }
function client(b, store, states) {
  return new BridgeClient({ app: 'excel', fetchImpl: b.fetch, storage: store, onState: s => states.push(s), sleep: () => new Promise(r => setTimeout(r, 1)),
    execute: async (tool) => ({ ok: true, data: { tool } }), describe: () => ({ supportedTools: ['read_range'] }) })
}

test('pair, run a command, then a reloaded pane resumes the same binding without a new code', async () => {
  const b = fakeBridge(); const store = new Store(); const states = []
  const c1 = client(b, store, states)
  assert.equal((await c1.pair('GOOD-CODE')).ok, true)
  b.queue.push({ requestId: 'r1', tool: 'read_range', arguments: {} })
  assert.ok(await until(() => b.results.length === 1))
  c1.stopped = true // the page goes away (reload) without saying bye
  const states2 = []
  const c2 = client(b, store, states2)
  const r = await c2.resume()
  assert.equal(r.ok, true)
  assert.equal(c2.token, 'tok-1')
  assert.ok(states2.some(s => s.state === 'paired' && s.resumed))
  b.queue.push({ requestId: 'r2', tool: 'read_range', arguments: {} })
  assert.ok(await until(() => b.results.length === 2))
  c2.stopped = true
})

test('resume with a token the bridge no longer knows clears it and asks for a code', async () => {
  const b = fakeBridge(); const store = new Store(); const states = []
  store.setItem('hermes-office-desktop:excel', JSON.stringify({ base: 'https://localhost:47615', token: 'old' }))
  const c = client(b, store, states)
  const r = await c.resume()
  assert.equal(r.ok, false)
  assert.equal(store.getItem('hermes-office-desktop:excel'), null)
  assert.ok(states.some(s => s.state === 'unpaired' && s.reason === 'expired'))
})

test('resume while Hermes is down keeps the saved pairing for later', async () => {
  const b = fakeBridge(); const store = new Store(); const states = []
  store.setItem('hermes-office-desktop:excel', JSON.stringify({ base: 'https://localhost:47615', token: 'tok-x' }))
  b.down = true
  const r = await client(b, store, states).resume()
  assert.equal(r.category, 'Unavailable')
  assert.notEqual(store.getItem('hermes-office-desktop:excel'), null)
})

test('network loss shows reconnecting, then recovers on the same token', async () => {
  const b = fakeBridge(); const store = new Store(); const states = []
  const c = client(b, store, states)
  await c.pair('GOOD-CODE')
  b.down = true
  assert.ok(await until(() => states.some(s => s.state === 'reconnecting')))
  b.down = false
  assert.ok(await until(() => states.some(s => s.state === 'paired' && s.reconnected)))
  b.queue.push({ requestId: 'r1', tool: 'read_range', arguments: {} })
  assert.ok(await until(() => b.results.length === 1))
  c.stopped = true
})

test('revoked token (other pane/chat took over) stops the loop and clears storage', async () => {
  const b = fakeBridge(); const store = new Store(); const states = []
  const c = client(b, store, states)
  await c.pair('GOOD-CODE')
  b.tokens.clear()
  assert.ok(await until(() => states.some(s => s.state === 'unpaired' && s.reason === 'revoked')))
  assert.equal(store.getItem('hermes-office-desktop:excel'), null)
  assert.equal(c.token, null)
})

test('re-pairing sends the previous token so the old chat binding is dropped', async () => {
  const b = fakeBridge(); const store = new Store(); const states = []
  const c = client(b, store, states)
  await c.pair('GOOD-CODE')
  await c.pair('GOOD-CODE')
  assert.equal(b.pairs[1].previousToken, 'tok-1')
  assert.equal(b.tokens.has('tok-1'), false)
  assert.equal(c.token, 'tok-2')
  c.stopped = true
})

test('bad code and missing bridge give distinct states; disconnect says bye and forgets', async () => {
  const b = fakeBridge(); const store = new Store(); const states = []
  const c = client(b, store, states)
  assert.equal((await c.pair('BAD')).category, 'Authentication')
  assert.ok(states.some(s => s.state === 'rejected'))
  b.down = true
  assert.equal((await c.pair('GOOD-CODE')).category, 'Unavailable')
  assert.ok(states.some(s => s.state === 'no-bridge'))
  b.down = false
  await c.pair('GOOD-CODE')
  await c.disconnect()
  assert.equal(store.getItem('hermes-office-desktop:excel'), null)
  assert.ok(b.calls.includes('/v1/bye'))
})

test('every posted result carries the pane-measured elapsedMs', async () => {
  const b = fakeBridge(); const store = new Store(); const states = []
  b.queue.push({ requestId: 'r1', tool: 'read_range', arguments: {} })
  const c = new BridgeClient({ app: 'excel', fetchImpl: b.fetch, storage: store, onState: s => states.push(s), sleep: () => new Promise(r => setTimeout(r, 1)),
    execute: async () => { await new Promise(r => setTimeout(r, 20)); return { ok: true, data: {} } }, describe: () => ({}) })
  assert.equal((await c.pair('GOOD-CODE')).ok, true)
  assert.ok(await until(() => b.results.length === 1))
  const r = b.results[0].result
  assert.equal(r.ok, true)
  assert.ok(Number.isInteger(r.elapsedMs) && r.elapsedMs >= 15, `elapsedMs=${r.elapsedMs}`)
  await c.disconnect()
})

// --- A14: a profile's bridge gives its idle port back and takes a (maybe different) port again -----------
// ports: port -> { instance, tokens:Set, queue:[], results:[], auths:[] } or absent (nothing listening).
function portWorld() {
  const w = { ports: new Map(), bridges: new Map(), calls: [] }
  w.bridge = (instance) => { const b = { instance, tokens: new Set(), queue: [], results: [], auths: [] }; w.bridges.set(instance, b); return b }
  w.fetch = async (url, init = {}) => {
    const u = new URL(url)
    const b = w.ports.get(Number(u.port))
    w.calls.push(`${u.port}${u.pathname}`)
    if (!b) throw new TypeError('Failed to fetch')
    if (b === 'released') return { status: 503, json: async () => { throw new Error('no body') } }
    const auth = ((init.headers || {}).Authorization || '').replace('Bearer ', '')
    if (auth) b.auths.push(auth)
    const body = init.body ? JSON.parse(init.body) : null
    const json = (status, v) => ({ status, json: async () => v })
    switch (u.pathname) {
      case '/v1/hello': return json(200, { ok: true, service: 'hermes-office-desktop', protocol: 1, apps: ['excel'], instance: b.instance })
      case '/v1/pair': { if (body.code !== `CODE-${b.instance}`) return json(401, { ok: false, error: { category: 'Authentication' } }); const t = `tok-${b.instance}`; b.tokens.add(t); return json(200, { ok: true, paneToken: t, bindingId: 'bind' }) }
      case '/v1/resume': return b.tokens.has(auth) ? json(200, { ok: true, bindingId: 'bind' }) : json(401, { ok: false, error: { category: 'Authentication' } })
      case '/v1/next': { if (!b.tokens.has(auth)) return json(401, { ok: false }); await new Promise(r => setTimeout(r, 3)); return json(200, { ok: true, command: b.queue.shift() || null }) }
      case '/v1/result': b.results.push(body); return json(200, { ok: true })
    }
    return json(404, {})
  }
  return w
}
function wclient(w, store, states, extra = {}) {
  return new BridgeClient({ app: 'excel', fetchImpl: w.fetch, storage: store, onState: s => states.push(s), sleep: () => new Promise(r => setTimeout(r, 2)),
    execute: async (tool) => ({ ok: true, data: { tool } }), describe: () => ({}), ...extra })
}

test('released port: the pane finds its bridge by instance on the port it takes again, never sending its token to another', async () => {
  const w = portWorld(); const store = new Store(); const states = []
  const a = w.bridge('A'); const other = w.bridge('B')
  w.ports.set(47615, a)
  const c = wclient(w, store, states)
  assert.equal((await c.pair('CODE-A')).ok, true)
  assert.equal(JSON.parse(store.getItem('hermes-office-desktop:excel')).instance, 'A')
  a.queue.push({ requestId: 'r1', tool: 'read_range', arguments: {} })
  assert.ok(await until(() => a.results.length === 1))
  // A gives its idle port back; B takes it; nothing answers as A for a while, then A takes 47617.
  w.ports.set(47615, 'released')
  assert.ok(await until(() => states.some(s => s.state === 'reconnecting')))
  w.ports.set(47615, other)
  await new Promise(r => setTimeout(r, 30))
  w.ports.set(47617, a)
  a.queue.push({ requestId: 'r2', tool: 'read_range', arguments: {} })
  assert.ok(await until(() => a.results.length === 2))
  assert.equal(c.base, 'https://localhost:47617')
  assert.equal(JSON.parse(store.getItem('hermes-office-desktop:excel')).base, 'https://localhost:47617')
  assert.deepEqual(other.auths, []) // B never saw A's pane token
  assert.ok(states.some(s => s.state === 'paired' && s.reconnected))
  assert.ok(!states.some(s => s.state === 'unpaired'))
  c.stopped = true
})

test('a 401 from a bridge that is not ours is a moved port, not a revocation', async () => {
  const w = portWorld(); const store = new Store(); const states = []
  const a = w.bridge('A'); const other = w.bridge('B')
  w.ports.set(47615, a)
  const c = wclient(w, store, states)
  await c.pair('CODE-A')
  w.ports.set(47615, other) // swapped between two polls
  w.ports.set(47616, a)
  assert.ok(await until(() => c.base === 'https://localhost:47616'))
  a.queue.push({ requestId: 'r1', tool: 'read_range', arguments: {} })
  assert.ok(await until(() => a.results.length === 1))
  assert.equal(c.base, 'https://localhost:47616')
  assert.ok(!states.some(s => s.state === 'unpaired'))
  c.stopped = true
})

test('our own bridge rejecting the token still revokes', async () => {
  const w = portWorld(); const store = new Store(); const states = []
  const a = w.bridge('A'); w.ports.set(47615, a)
  const c = wclient(w, store, states)
  await c.pair('CODE-A')
  a.tokens.clear()
  assert.ok(await until(() => states.some(s => s.state === 'unpaired' && s.reason === 'revoked')))
  assert.equal(store.getItem('hermes-office-desktop:excel'), null)
})

test('a reloaded pane waits for its released bridge, then resumes on the new port', async () => {
  const w = portWorld(); const store = new Store(); const states = []
  const a = w.bridge('A'); a.tokens.add('tok-A'); const other = w.bridge('B')
  w.ports.set(47615, other) // the saved port now belongs to another profile
  store.setItem('hermes-office-desktop:excel', JSON.stringify({ base: 'https://localhost:47615', token: 'tok-A', instance: 'A', seenAt: Date.now() }))
  const c = wclient(w, store, states)
  const pending = c.resume({ waitMs: 60000, intervalMs: 5 })
  assert.ok(await until(() => states.some(s => s.state === 'waiting-bridge')))
  w.ports.set(47618, a) // A's next Office tool call took a port again
  const r = await pending
  assert.equal(r.ok, true)
  assert.equal(c.base, 'https://localhost:47618')
  assert.ok(states.some(s => s.state === 'paired' && s.resumed && s.moved))
  assert.deepEqual(other.auths, [])
  c.stopped = true
})

test('a saved pairing older than the bridge keeps pairings is dropped instead of waited for', async () => {
  const w = portWorld(); const store = new Store(); const states = []
  store.setItem('hermes-office-desktop:excel', JSON.stringify({ base: 'https://localhost:47615', token: 'tok-A', instance: 'A', seenAt: Date.now() - 16 * 60 * 1000 }))
  const r = await wclient(w, store, states).resume({ waitMs: 60000, intervalMs: 5 })
  assert.equal(r.ok, false)
  assert.equal(store.getItem('hermes-office-desktop:excel'), null)
  assert.ok(states.some(s => s.state === 'unpaired' && s.reason === 'expired'))
})

test('a new pairing while the pane waits for its old bridge stops the wait', async () => {
  const w = portWorld(); const store = new Store(); const states = []
  const b2 = w.bridge('C'); w.ports.set(47619, b2)
  store.setItem('hermes-office-desktop:excel', JSON.stringify({ base: 'https://localhost:47615', token: 'tok-A', instance: 'A', seenAt: Date.now() }))
  const c = wclient(w, store, states)
  const pending = c.resume({ waitMs: 60000, intervalMs: 5 })
  assert.ok(await until(() => states.some(s => s.state === 'waiting-bridge')))
  assert.equal((await c.pair('CODE-C')).ok, true)
  assert.equal((await pending).category, 'Superseded')
  assert.equal(c.base, 'https://localhost:47619')
  c.stopped = true
})

// --- A19: pairing never hangs on a bridge that is going away (seen live: excel run 3 of hol-a19-896ff28-all-3 sat at
// 'pairing' "Looking for Hermes on this computer…" while another profile's idle bridge gave its port back) ----------
test('a bridge that goes away between hello and pair does not leave the pane stuck: the code still reaches the next bridge', async () => {
  const w = portWorld(); const store = new Store(); const states = []
  const gone = w.bridge('B'); const ours = w.bridge('A')
  w.ports.set(47615, gone); w.ports.set(47616, ours)
  const realFetch = w.fetch
  w.fetch = async (url, init) => { // 47615 answers hello, then its listener is closed before the pair request
    if (url.endsWith(':47615/v1/pair')) throw new TypeError('Failed to fetch')
    return realFetch(url, init)
  }
  const c = wclient(w, store, states)
  const r = await c.pair('CODE-A')
  assert.equal(r.ok, true)
  assert.equal(c.base, 'https://localhost:47616')
  assert.equal(states.at(-1).state, 'paired')
  c.stopped = true
})

test('a pair request that fails on every bridge ends in no-bridge, never an endless pairing', async () => {
  const w = portWorld(); const store = new Store(); const states = []
  w.ports.set(47615, w.bridge('B'))
  const realFetch = w.fetch
  w.fetch = async (url, init) => { if (url.endsWith('/v1/pair')) throw new TypeError('Failed to fetch'); return realFetch(url, init) }
  const c = wclient(w, store, states)
  const r = await c.pair('CODE-B')
  assert.equal(r.ok, false)
  assert.equal(r.category, 'Unavailable')
  assert.equal(states.at(-1).state, 'no-bridge')
})

test('a hello probe or pair request that never answers is bounded by a timeout', async () => {
  const w = portWorld(); const store = new Store(); const states = []
  const ours = w.bridge('A'); w.ports.set(47615, w.bridge('B')); w.ports.set(47616, ours)
  const realFetch = w.fetch
  const hang = (init) => new Promise((resolve, reject) => {
    const sig = init && init.signal
    if (sig) sig.addEventListener('abort', () => reject(Object.assign(new Error('The operation was aborted'), { name: 'AbortError' })))
  })
  w.fetch = async (url, init) => (url.includes(':47615/') ? hang(init) : realFetch(url, init)) // accepts, never answers
  const c = wclient(w, store, states, { probeTimeoutMs: 30, requestTimeoutMs: 60 })
  const t0 = Date.now()
  const r = await c.pair('CODE-A')
  assert.equal(r.ok, true)
  assert.equal(c.base, 'https://localhost:47616')
  assert.ok(Date.now() - t0 < 1000, `pair took ${Date.now() - t0} ms`)
  c.stopped = true
  // the pair request itself hanging: bounded as well
  const w2 = portWorld(); const states2 = []
  w2.ports.set(47615, w2.bridge('B'))
  const real2 = w2.fetch
  w2.fetch = async (url, init) => (url.endsWith('/v1/pair') ? hang(init) : real2(url, init))
  const c2 = wclient(w2, new Store(), states2, { probeTimeoutMs: 30, requestTimeoutMs: 60 })
  const r2 = await c2.pair('CODE-B')
  assert.equal(r2.ok, false)
  assert.equal(states2.at(-1).state, 'no-bridge')
})

test('the pane looks for Hermes over https://localhost on the five ports and drops a saved plain-http base', async () => {
  const seen = []
  const store = new Store()
  store.setItem('hermes-office-desktop:excel', JSON.stringify({ base: 'http://127.0.0.1:47615', token: 'old', instance: 'A' }))
  const fetchImpl = async (url) => { seen.push(url); throw new TypeError('Failed to fetch') }
  const c = new BridgeClient({ app: 'excel', fetchImpl, storage: store, execute: async () => ({ ok: true }) })
  assert.deepEqual(await c.discover(), [])
  assert.deepEqual(seen, [47615, 47616, 47617, 47618, 47619].map(p => `https://localhost:${p}/v1/hello`))
  assert.equal(c._saved(), null) // a pairing saved by an older pane over plain http is never reused
  store.setItem('hermes-office-desktop:excel', JSON.stringify({ base: 'https://localhost:47616', token: 'tok' }))
  assert.equal(c._saved().base, 'https://localhost:47616')
})
