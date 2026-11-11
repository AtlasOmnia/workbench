// SPDX-License-Identifier: Apache-2.0
// Shared core for the strict in-memory Office.js stand-ins (DEV/TEST ONLY).
// Enforces Office.js's load()/sync() rule: reading a property that was not
// loaded and synced throws PropertyNotLoaded, and writes/actions are queued
// until sync(). Not Microsoft code; no fidelity is claimed beyond what the
// executors use.

export class OfficeError extends Error {
  constructor(code, message) { super(message); this.code = code; this.name = 'RichApi.Error' }
}

export class Loadable {
  constructor(ctx) { this._ctx = ctx; this._pending = new Set(); this._loaded = new Set() }
  load(props) {
    const list = typeof props === 'string' ? props.split(',') : Array.isArray(props) ? props : Object.keys(props || {})
    for (const p of list) this._pending.add(p.trim().replace(/^items\//, '').split('/')[0])
    this._ctx._dirty.add(this)
    return this
  }
  _commit() { for (const p of this._pending) this._loaded.add(p); this._pending.clear() }
  _need(p) { if (!this._loaded.has(p)) throw new OfficeError('PropertyNotLoaded', `The property '${p}' is not available. Call load() and sync().`) }
}

/** Collection whose `items` become readable after any load() + sync(). */
export class LoadableCollection extends Loadable {
  constructor(ctx, list) { super(ctx); this._list = list; this._snap = null }
  _commit() {
    super._commit()
    // Like Office.js, items are a snapshot taken at sync time.
    this._snap = this._list().map(x => { x._loaded = new Set([...x._loaded, ...this._loaded]); return x })
  }
  get items() {
    if (!this._snap) throw new OfficeError('PropertyNotLoaded', "The property 'items' is not available. Call load() and sync().")
    return this._snap
  }
  get count() { this._need('count'); return this._snap.length }
}

export class ClientResult {
  constructor(ctx, fn) { this._ctx = ctx; this._fn = fn; this._ready = false; ctx._queue(() => { this._v = fn(); this._ready = true }) }
  get value() { if (!this._ready) throw new OfficeError('PropertyNotLoaded', 'ClientResult value is not available until sync()'); return this._v }
}

export class MockContextBase {
  constructor() { this._ops = []; this._dirty = new Set(); this._failOnSync = null; this.trackedObjects = { add() {}, remove() {} } }
  _queue(fn) { this._ops.push(fn) }
  _fail(err) { if (!this._failOnSync) this._failOnSync = err }
  async sync() {
    this._model.syncs++
    if (this._failOnSync) { const e = this._failOnSync; this._failOnSync = null; this._ops = []; throw e }
    const ops = this._ops; this._ops = []
    try {
      for (const op of ops) op()
    } catch (e) { this._dirty.clear(); throw e }
    // Like Office.js, a failed batch is discarded; the context stays usable for the next sync.
    try { for (const o of this._dirty) o._commit() } finally { this._dirty.clear() }
  }
}

export function makeRun(Ctx, model, afterRun = () => {}) {
  return async fn => {
    const ctx = new Ctx(typeof model === 'function' ? model() : model)
    try { const out = await fn(ctx); await ctx.sync(); return out } finally { afterRun() }
  }
}

/** In a browser pane, keep the mock document across pane reloads (like a real open document). */
export function persistence(key, serialize, restore) {
  const store = (() => { try { return globalThis.sessionStorage || null } catch { return null } })()
  return {
    load() { if (!store) return null; const raw = store.getItem(key); return raw ? restore(JSON.parse(raw)) : null },
    save(model) { if (store) store.setItem(key, JSON.stringify(serialize(model))) }
  }
}

export function installOffice(host, { url, platform = 'OfficeOnline', document = {} } = {}) {
  globalThis.Office = {
    HostType: { Excel: 'Excel', Word: 'Word', PowerPoint: 'PowerPoint' },
    PlatformType: { OfficeOnline: 'OfficeOnline', PC: 'PC' },
    CoercionType: { Image: 'image', Text: 'text' },
    onReady: cb => Promise.resolve({ host, platform }).then(info => { cb && cb(info); return info }),
    context: {
      document: { url, ...document },
      diagnostics: { version: 'mock', platform },
      requirements: { isSetSupported: () => true }
    }
  }
}
