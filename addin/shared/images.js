// SPDX-License-Identifier: Apache-2.0
// Picture helpers shared by the Excel, Word, PowerPoint and OneNote panes (office_insert_image).
// The plugin already checked the bytes (PNG/JPEG magic, <= 4 MiB) before sending them; the pane
// re-checks only what Office.js would otherwise choke on, and never fetches anything itself.
import { ExecError, str } from './runtime.js'

// 4 MiB of binary is ~5.6 MB of base64; a little headroom, no more.
export const MAX_IMAGE_BASE64 = 6 * 1024 * 1024
export const MAX_IMAGE_HTML = 8 * 1024 * 1024

export function base64Kind(b64) {
  const head = String(b64).slice(0, 16)
  if (head.startsWith('iVBORw0KGgo')) return 'png'
  if (head.startsWith('/9j/')) return 'jpeg'
  return null
}

/** Bare base64 of a PNG/JPEG (a data:image/png|jpeg;base64, prefix is stripped) and its kind. */
export function imageBase64(v) {
  if (typeof v !== 'string' || !v) throw new ExecError('Validation', 'base64 must be a non-empty string')
  const b64 = v.replace(/^data:image\/(png|jpeg);base64,/, '')
  if (b64.length > MAX_IMAGE_BASE64) throw new ExecError('Capacity', 'image is larger than 4 MiB')
  if (!/^[A-Za-z0-9+/]+={0,2}$/.test(b64)) throw new ExecError('Validation', 'base64 is not valid base64')
  const kind = base64Kind(b64)
  if (!kind) throw new ExecError('Validation', 'only PNG or JPEG images can be inserted')
  return { b64, kind }
}

export function altText(v, name = 'alt') {
  return str(typeof v === 'string' ? v.slice(0, 2000) : v, name, 2000)
}

/**
 * HTML the plugin already sanitized. Re-checked here as markup (inside tags only, so note text that
 * happens to read "one = two" is not refused): no scripts, frames, embedded objects, forms, event
 * handlers, javascript: URLs, stylesheets (<style>/<link>) or style attributes that could name a
 * resource, and at least one picture (office_insert_image never inserts text-only HTML).
 * The plugin only ever emits style="prop:value;..." from a small allowlist, so anything else fails closed.
 */
const BAD_STYLE = /url|image|expression|@|import|font-face|src\s*:|javascript|behavior|binding|cross-fade|element|var\s*\(|attr\s*\(|\\|<|>|\/\*|\*\/|["'`(){}]|&|--/i
// One attribute at a time, so an alt text that merely reads "style=..." is not mistaken for a style attribute.
const ATTR = /\s([^\s"'>\/=]+)(?:\s*=\s*(?:"([^"]*)"|'([^']*)'|([^\s"'=<>`]+)))?/g
export function pictureHtml(v) {
  if (typeof v !== 'string' || !v.trim() || v.length > MAX_IMAGE_HTML) throw new ExecError('Validation', 'html must be non-empty and at most 8 MB')
  if (/<\s*\/?\s*(script|iframe|object|embed|form|frame|frameset|applet|base|meta|link|style|svg|math)\b/i.test(v) ||
      /<[^>]*\son[a-z]+\s*=/i.test(v) || /<[^>]*javascript:/i.test(v)) {
    throw new ExecError('Validation', 'html contains markup an add-in must not insert')
  }
  for (const tag of v.match(/<[a-z][^>]*>/gi) || []) {
    for (const m of tag.slice(tag.search(/[\s>]/)).matchAll(ATTR)) {
      if (m[1].toLowerCase() !== 'style') continue
      const css = m[2] ?? m[3] ?? m[4] ?? ''
      if (m[2] === undefined || BAD_STYLE.test(css)) throw new ExecError('Validation', 'html carries a style the add-in must not insert')
    }
  }
  if (!/<img\s[^>]*src\s*=\s*"data:image\/(png|jpeg);base64,/i.test(v)) {
    throw new ExecError('Validation', 'html has no picture: office_insert_image inserts at least one PNG/JPEG image')
  }
  return v
}
