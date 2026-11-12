// SPDX-License-Identifier: Apache-2.0
// Pane side of the Hermes Office Desktop bridge: discover the local Hermes
// bridge on 127.0.0.1, pair with a one-time code the user reads from their
// Hermes chat, then long-poll for commands and post results.
//
// Network: only https://localhost:<47615-47619>, the plugin's loopback bridge (HTTPS with the Hermes profile's
// local CA, which the user trusts once). By default this pane is served by that same bridge, so its own port is
// same-origin; another of the five ports, or a pane from a self-hosted site, is a CORS request the bridge checks
// by Origin. From a self-hosted (public) page, Chromium Local Network Access treats this as a public->loopback
// request; `targetAddressSpace: 'loopback'` lets the browser show (or, in an embedder, auto-resolve) the permission.
//
// Reconnect: the pane token is kept in sessionStorage (this tab only, cleared
// when the tab closes) so reloading the pane resumes the same chat binding
// without a new code. A token the bridge no longer knows (Hermes restarted,
// pairing expired, chat disconnected) is discarded and the pane asks for a code.
//
// Moving port (A14): a Hermes profile's bridge gives its port back after 10 idle
// minutes and takes the first free one of 47615-47619 again on its next Office
// tool call, possibly a different port (another profile may hold the old one).
// /v1/hello carries the bridge's random `instance` id; the pane remembers it and
// finds ITS bridge by that id on whichever port it is now, without ever sending
// its token to another profile's bridge.

export const DEFAULT_PORTS = [47615, 47616, 47617, 47618, 47619] // must match bridge_server.py
const PROTOCOL = 1
// bridge_server.BINDING_IDLE_SECONDS: a pairing unseen this long is gone on the Hermes side as well.
export const BINDING_IDLE_MS = 15 * 60 * 1000
const SEEN_SAVE_MS = 30 * 1000
// A19: a hello probe or a pairing request that never answers (a port whose listener is being closed) must not hold
// the pane at "Looking for Hermes" forever; it counts as that port not answering.
export const PROBE_TIMEOUT_MS = 4000
export const REQUEST_TIMEOUT_MS = 15000

function sessionStore() {
  try { return globalThis.sessionStorage || null } catch { return null }
}

export class BridgeClient {
  constructor({ app, ports = DEFAULT_PORTS, fetchImpl = globalThis.fetch.bind(globalThis), onState = () => {}, execute, describe, storage = sessionStore(), sleep, now, probeTimeoutMs = PROBE_TIMEOUT_MS, requestTimeoutMs = REQUEST_TIMEOUT_MS }) {
    this.app = app
    this.ports = ports
    this.fetch = fetchImpl
    this.onState = onState
    this.execute = execute
    this.describe = describe || (() => ({}))
    this.storage = storage
    this.sleep = sleep || (ms => new Promise(r => setTimeout(r, ms)))
    this.now = now || (() => Date.now())
    this.probeTimeoutMs = probeTimeoutMs
    this.requestTimeoutMs = requestTimeoutMs
    this.key = `hermes-office-desktop:${app}`
    this.base = null
    this.token = null
    this.instance = null // the paired bridge's /v1/hello instance id (null: an older bridge without one)
    this.instances = new Map() // base -> instance, from the last discover()
    this.seenAt = 0
    this.savedAt = 0
    this.stopped = false
    this.loopRunning = false
  }

  _opts(init = {}) {
    // targetAddressSpace is the LNA opt-in; unknown fetch init keys are ignored by older engines.
    return { mode: 'cors', credentials: 'omit', cache: 'no-store', targetAddressSpace: 'loopback', ...init }
  }

  async _json(url, init, timeoutMs = 0) {
    // timeoutMs > 0: abort a request that has not answered by then (it rejects like a network error).
    const ctl = timeoutMs > 0 && typeof AbortController === 'function' ? new AbortController() : null
    const timer = ctl ? setTimeout(() => ctl.abort(), timeoutMs) : null
    try {
      const res = await this.fetch(url, this._opts(ctl ? { ...init, signal: ctl.signal } : init))
      let body = null
      try { body = await res.json() } catch { body = null }
      return { status: res.status, body }
    } finally {
      if (timer) clearTimeout(timer)
    }
  }

  _save() {
    if (!this.storage) return
    try {
      if (this.token && this.base) {
        this.storage.setItem(this.key, JSON.stringify({ base: this.base, token: this.token, ...(this.instance ? { instance: this.instance } : {}), seenAt: this.seenAt || this.now() }))
        this.savedAt = this.now()
      } else this.storage.removeItem(this.key)
    } catch { /* storage may be disabled; reload then needs a new code */ }
  }

  _saved() {
    if (!this.storage) return null
    try {
      const v = JSON.parse(this.storage.getItem(this.key) || 'null')
      if (v && typeof v.token === 'string' && typeof v.base === 'string' && /^https:\/\/localhost:\d+$/.test(v.base)) return v
    } catch { /* ignore */ }
    return null
  }

  /** Probe every candidate port; returns bases answering as this service (instances: base -> instance id). */
  async discover() {
    const found = []
    this.instances = new Map()
    for (const port of this.ports) {
      const base = `https://localhost:${port}`
      try {
        const { status, body } = await this._json(`${base}/v1/hello`, { method: 'GET' }, this.probeTimeoutMs)
        if (status === 200 && body && body.service === 'hermes-office-desktop' && body.protocol === PROTOCOL &&
            (!Array.isArray(body.apps) || body.apps.includes(this.app))) {
          found.push(base)
          if (typeof body.instance === 'string') this.instances.set(base, body.instance)
        }
      } catch (err) {
        this.lastNetworkError = String((err && err.message) || err)
      }
    }
    return found
  }

  /** Where the bridge with this instance id listens now (hello probes only: no token is sent). */
  async _findBridge(instance) {
    const bases = await this.discover()
    return bases.find(base => this.instances.get(base) === instance) || null
  }

  /**
   * After a pane reload: rebind to the same chat with the tab-scoped token, if it is still valid.
   * waitMs > 0: when the paired bridge is not listening right now (it gave its idle port back), keep looking
   * every intervalMs until it is back, the pairing would have expired anyway, or the user pairs anew.
   */
  async resume({ waitMs = 0, intervalMs = 5000 } = {}) {
    const saved = this._saved()
    if (!saved) return { ok: false, category: 'NoSavedPairing' }
    this.onState({ state: 'resuming' })
    const deadline = this.now() + Math.max(0, waitMs)
    let waiting = false
    while (true) {
      if (this.token && this.token !== saved.token) return { ok: false, category: 'Superseded' } // paired anew meanwhile
      let base = saved.base
      if (saved.instance) {
        try {
          base = await this._findBridge(saved.instance)
        } catch (err) {
          base = null
          this.lastNetworkError = String((err && err.message) || err)
        }
        if (!base) {
          if (saved.seenAt && this.now() - saved.seenAt > BINDING_IDLE_MS) return this._expired()
          if (this.now() < deadline) {
            if (!waiting) { waiting = true; this.onState({ state: 'waiting-bridge' }) }
            await this.sleep(intervalMs)
            continue
          }
          this.onState({ state: 'no-bridge', detail: this.lastNetworkError || '' })
          return { ok: false, category: 'Unavailable' }
        }
      }
      try {
        const describe = await this.describe()
        const { status, body } = await this._json(`${base}/v1/resume`, {
          method: 'POST', headers: { 'Content-Type': 'application/json', Authorization: `Bearer ${saved.token}` },
          body: JSON.stringify({ app: this.app, ...describe })
        })
        if (status === 200 && body && body.ok) {
          this.base = base
          this.token = saved.token
          this.instance = saved.instance || null
          this.seenAt = this.now()
          this.stopped = false
          this._save()
          this.onState({ state: 'paired', bindingId: body.bindingId, resumed: true, ...(base !== saved.base ? { moved: true } : {}) })
          this._loop()
          return { ok: true, bindingId: body.bindingId }
        }
        return this._expired(body)
      } catch (err) {
        // Hermes not reachable right now: keep the saved pairing so a later reload can still resume.
        this.lastNetworkError = String((err && err.message) || err)
        if (saved.instance && this.now() < deadline) { await this.sleep(intervalMs); continue }
        this.onState({ state: 'no-bridge', detail: this.lastNetworkError })
        return { ok: false, category: 'Unavailable' }
      }
    }
  }

  _expired(body) {
    this.storage && this.storage.removeItem(this.key)
    this.onState({ state: 'unpaired', reason: 'expired' })
    return { ok: false, category: (body && body.error && body.error.category) || 'Authentication' }
  }

  async pair(code) {
    this.onState({ state: 'pairing' })
    const bases = await this.discover()
    if (!bases.length) {
      this.onState({ state: 'no-bridge', detail: this.lastNetworkError || '' })
      return { ok: false, category: 'Unavailable' }
    }
    const describe = await this.describe()
    const previousToken = this.token || (this._saved() || {}).token || undefined
    let lastCategory = 'Authentication'
    let answered = 0
    for (const base of bases) {
      let status, body
      try {
        ({ status, body } = await this._json(`${base}/v1/pair`, {
          method: 'POST', headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ code, app: this.app, ...describe, ...(previousToken ? { previousToken } : {}) })
        }, this.requestTimeoutMs))
      } catch (err) {
        // A19: this bridge went away after its hello (its idle port given back) or never answered: try the next.
        this.lastNetworkError = String((err && err.message) || err)
        continue
      }
      answered++
      if (status === 200 && body && body.ok && typeof body.paneToken === 'string') {
        this.base = base
        this.token = body.paneToken
        this.instance = this.instances.get(base) || null
        this.seenAt = this.now()
        this.stopped = false
        this._save()
        this.onState({ state: 'paired', bindingId: body.bindingId })
        this._loop()
        return { ok: true, bindingId: body.bindingId }
      }
      lastCategory = (body && body.error && body.error.category) || lastCategory
      if (lastCategory === 'RateLimited') break
    }
    if (!answered) {
      this.onState({ state: 'no-bridge', detail: this.lastNetworkError || '' })
      return { ok: false, category: 'Unavailable' }
    }
    this.onState({ state: 'rejected', category: lastCategory })
    return { ok: false, category: lastCategory }
  }

  async _loop() {
    if (this.loopRunning) return
    this.loopRunning = true
    let backoff = 500
    let recovering = false
    let verify = false // after a loss: hello must show OUR instance at this.base before the token goes there again
    const lost = async () => {
      // The bridge stopped answering here (Hermes gone, or this profile gave its idle port back). A bridge with
      // an instance id is looked for by hello on every port before the token is sent anywhere again.
      this.onState({ state: 'reconnecting' })
      recovering = true
      if (this.instance && this.seenAt && this.now() - this.seenAt > BINDING_IDLE_MS) {
        this.token = null; this._save(); this.onState({ state: 'unpaired', reason: 'expired' }); return false
      }
      if (this.instance) {
        let found = null
        try { found = await this._findBridge(this.instance) } catch { found = null }
        if (found && found !== this.base) { this.base = found; this._save(); verify = false; return true } // moved: poll there now
        verify = !found // still ours at this.base (hello said so), or nowhere yet: check again before the token goes out
      }
      await this.sleep(backoff)
      backoff = Math.min(backoff * 2, 8000)
      return true
    }
    try {
      while (!this.stopped && this.token) {
        let res
        const token = this.token
        if (verify && this.instance && !(await this._stillOurs())) {
          if (!(await lost())) break
          continue
        }
        verify = false
        try {
          res = await this._json(`${this.base}/v1/next?wait=20`, { method: 'GET', headers: { Authorization: `Bearer ${token}` } })
        } catch {
          if (!(await lost())) break
          continue
        }
        if (this.token !== token) continue // re-paired meanwhile
        if (res.status === 503 || (res.status === 401 && this.instance && !(await this._stillOurs()))) {
          // 503: a listener this bridge has released; 401 from a bridge that is not ours: our port moved.
          if (!(await lost())) break
          continue
        }
        if (recovering) this.onState({ state: 'paired', reconnected: true })
        recovering = false
        backoff = 500
        if (res.status === 401) { this.token = null; this._save(); this.onState({ state: 'unpaired', reason: 'revoked' }); break }
        this.seenAt = this.now()
        if (this.now() - this.savedAt > SEEN_SAVE_MS) this._save()
        const cmd = res.body && res.body.command
        if (!cmd) continue
        this.onState({ state: 'working', tool: cmd.tool })
        const started = Date.now()
        const executed = await this.execute(cmd.tool, cmd.arguments)
        // Pane-measured execution time (Office.js batches included); Hermes adds its own roundTripMs.
        const result = { ...(executed && typeof executed === 'object' ? executed : { ok: false }), elapsedMs: Date.now() - started }
        try {
          await this._json(`${this.base}/v1/result`, {
            method: 'POST', headers: { 'Content-Type': 'application/json', Authorization: `Bearer ${token}` },
            body: JSON.stringify({ requestId: cmd.requestId, result })
          })
        } catch { /* Hermes reports ExecutionUncertain on its side; never re-run here. */ }
        this.onState({ state: 'paired', last: cmd.tool, ok: !!result.ok, error: result.ok ? null : result.error, ms: Date.now() - started })
      }
    } finally {
      this.loopRunning = false
    }
  }

  /** Is the bridge at this.base still the one we paired with? (hello only) */
  async _stillOurs() {
    try {
      const { status, body } = await this._json(`${this.base}/v1/hello`, { method: 'GET' }, this.probeTimeoutMs)
      return status === 200 && !!body && body.instance === this.instance
    } catch {
      return false
    }
  }

  /** Tell the paired Hermes chat what the pane observed (e.g. a OneNote page visit). Never carries a command. */
  async postEvent(event) {
    if (!this.token || !this.base) return { ok: false, category: 'NotPaired' }
    try {
      const { status, body } = await this._json(`${this.base}/v1/event`, {
        method: 'POST', headers: { 'Content-Type': 'application/json', Authorization: `Bearer ${this.token}` },
        body: JSON.stringify(event)
      })
      return status === 200 && body ? body : { ok: false, status }
    } catch (err) {
      return { ok: false, category: 'Unavailable', detail: String((err && err.message) || err) }
    }
  }

  async disconnect() {
    this.stopped = true
    const token = this.token
    this.token = null
    this._save()
    if (token && this.base) {
      try {
        await this._json(`${this.base}/v1/bye`, { method: 'POST', headers: { 'Content-Type': 'application/json', Authorization: `Bearer ${token}` }, body: '{}' })
      } catch { /* best effort */ }
    }
    this.onState({ state: 'unpaired', reason: 'user' })
  }
}
