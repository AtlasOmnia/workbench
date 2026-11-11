// SPDX-License-Identifier: Apache-2.0
// Host-neutral executor runtime shared by the Excel, Word and PowerPoint panes.
//
// Each host module supplies a table of canonical tools. A read tool returns plain
// JSON. A mutating tool returns { data, undo } where `undo` is either an async
// function (ctx) => void that restores the captured prior state, or null plus an
// `undoUnavailable` reason. The runtime keeps a bounded in-memory undo stack per
// pane (lost when the pane reloads; Office.js has no host undo API) and exposes
// it as the `undo_last` command. Every result is a bridge envelope:
//   { ok: true, data } | { ok: false, error: { category, code?, message } }

export const UNDO_LIMIT = 20

export class ExecError extends Error {
  constructor(category, message, code) {
    super(message)
    this.category = category
    if (code) this.code = code
  }
}

export function str(v, name, max = 255) {
  if (typeof v !== 'string' || !v || v.length > max) throw new ExecError('Validation', `${name} must be a non-empty string`)
  return v
}

export function int(v, name, { min = 0, max = Number.MAX_SAFE_INTEGER } = {}) {
  if (!Number.isInteger(v) || v < min || v > max) throw new ExecError('Validation', `${name} must be an integer between ${min} and ${max}`)
  return v
}

/** Requirement-set guard: real Office exposes Office.context.requirements; the mocks report everything supported. */
export function requireSet(set, version, what) {
  const req = globalThis.Office && globalThis.Office.context && globalThis.Office.context.requirements
  if (req && typeof req.isSetSupported === 'function' && !req.isSetSupported(set, version)) {
    throw new ExecError('Unsupported', `${what} needs ${set} ${version}, which this Office version does not provide`)
  }
}

export function officeError(err) {
  if (err instanceof ExecError) {
    const e = { category: err.category, message: err.message }
    if (err.code) e.code = err.code
    return e
  }
  // Office.js errors carry a code (ItemNotFound, InvalidArgument, ...); keep the code, bound the text.
  const code = err && typeof err.code === 'string' ? err.code : 'OfficeError'
  const message = String((err && err.message) || 'Office.js failed').slice(0, 300)
  return { category: 'Office', code, message }
}

/**
 * Best-effort change highlight. Runs only after a mutating tool's own sync succeeded, so the write has
 * already landed: a failed select never fails the write; it comes back as a note in the result.
 * `fn(ctx)` queues the selection and returns a short description of the target.
 */
export async function highlightChange(ctx, fn) {
  try {
    const target = await fn(ctx)
    await ctx.sync()
    return { selected: true, target: String(target || '').slice(0, 200) }
  } catch (err) {
    const e = officeError(err)
    return { selected: false, note: `The change applied, but it could not be highlighted (${e.code || e.category}: ${e.message})`.slice(0, 300) }
  }
}

/**
 * Build execute(tool, args) for one host.
 * @param {object} opts
 * @param {object} opts.tools      name -> async (ctx, args) => data | { data, undo, undoUnavailable }
 * @param {Set<string>} opts.mutating names that change the document
 * @param {(fn) => Promise<any>} opts.run host batch runner (Excel.run / Word.run / PowerPoint.run)
 */
export function createExecutor({ tools, mutating, run }) {
  const stack = []
  const supported = Object.freeze([...Object.keys(tools), 'undo_last'])

  async function execute(tool, args) {
    if (tool === 'undo_last') return undoLast()
    const fn = Object.prototype.hasOwnProperty.call(tools, tool) ? tools[tool] : null
    if (!fn) return { ok: false, error: { category: 'Unsupported', message: `Pane does not implement ${tool}` } }
    try {
      const out = await run(async ctx => {
        const o = await fn(ctx, args || {})
        // Change highlight (writes only): after the write has synced, select what changed in the same batch.
        if (mutating.has(tool) && o && typeof o.highlight === 'function') o.highlightResult = await highlightChange(ctx, o.highlight)
        return o
      })
      if (!mutating.has(tool)) return { ok: true, data: out }
      const data = out && Object.prototype.hasOwnProperty.call(out, 'data') ? out.data : out
      let undo
      if (out && typeof out.undo === 'function') {
        stack.push({ tool, label: out.undoLabel || tool, fn: out.undo, partial: out.undoPartial || null })
        while (stack.length > UNDO_LIMIT) stack.shift()
        undo = { available: true, depth: stack.length }
        if (out.undoPartial) undo.partial = out.undoPartial
      } else {
        undo = { available: false, reason: (out && out.undoUnavailable) || 'Office.js cannot reverse this change' }
      }
      const result = { ...(data && typeof data === 'object' && !Array.isArray(data) ? data : { result: data }), undo }
      if (out && out.highlightResult) result.highlight = out.highlightResult
      return { ok: true, data: result }
    } catch (err) {
      return { ok: false, error: officeError(err) }
    }
  }

  async function undoLast() {
    const entry = stack.pop()
    if (!entry) return { ok: false, error: { category: 'NothingToUndo', message: 'There is no Hermes change to undo in this pane (undo history is cleared when the pane reloads).' } }
    try {
      await run(ctx => entry.fn(ctx))
      const data = { undone: entry.tool, label: entry.label, remaining: stack.length }
      if (entry.partial) data.partial = entry.partial
      return { ok: true, data }
    } catch (err) {
      return { ok: false, error: { ...officeError(err), message: `Undo of ${entry.tool} failed: ${officeError(err).message}` } }
    }
  }

  return { execute, SUPPORTED_TOOLS: supported, undoDepth: () => stack.length, clearUndo: () => { stack.length = 0 } }
}
