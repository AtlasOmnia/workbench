<!-- SPDX-License-Identifier: Apache-2.0 -->
# PowerPoint: native, readable slide design

## Design before decoration

Resolve audience, purpose, exact facts, intended viewing size and existing brand
constraints. Read `powerpoint_get_presentation_overview` and
`powerpoint_read_slide_content` for target slides before editing. Use returned
zero-based slide indexes and real shape IDs, not guessed slide numbers. Respect
truncation. Keep surrounding slides and existing content unless replacement is
explicitly requested.

Give each slide one message and one dominant title. Establish hierarchy with
scale, spacing, alignment and contrast before adding shapes. Use consistent
margins and a simple grid; align text boxes rather than scattering them. Keep a
single focal system: a clear claim with evidence, a diagram with labels, or a
short comparison. Do not give every element equal weight. Remove or split excess
content before shrinking fonts. Keep editable text as text, not an image.

Use one readable font family (at most two), consistent roles and a restrained
palette: neutral background, high-contrast text, one accent. A starting range is
28–36 pt titles and 18–24 pt body; verify actual viewing size and template rather
than treating numbers as acceptance. Never use color alone to communicate meaning.
Keep labels and sources readable, avoid full-slide dense paragraphs and busy
backgrounds, and leave breathing room around the title and evidence.

## Use the supported surface

- Prefer `powerpoint_create_text_slide` for a title plus body/bullets: it provides
  safe non-overlapping default layout rather than many guessed text boxes.
  Inspect the resulting slide before adding anything else.
- `powerpoint_list_slide_layouts` returns actual layouts/master IDs. Use those
  with `powerpoint_create_slide` or `powerpoint_set_slide_layout` only when the
  host supports them; do not invent a layout ID or promise arbitrary theme editing.
- For necessary custom composition, `powerpoint_add_text_box` accepts text,
  slideIndex, left, top, width and height in points. Check available slide geometry
  in the host before choosing coordinates; keep boxes within bounds and reserve
  non-overlapping title/body zones. The tools do not expose a general existing
  shape move/resize operation: do not invent one, or delete/recreate content to
  hide that limitation. Ask for user adjustment or agree the default layout.
- Read `powerpoint_read_slide_content` for actual shapeId values, then use
  `powerpoint_format_text` with explicit slideIndex and shapeId. Formatting is
  nested under `font`: bold, italic, color, size and name, not Word's fontSize.
  Formatting requires the host's PowerPointApi 1.4; respect Unsupported errors.
  Preserve unrequested properties. `powerpoint_update_slide_text` with explicit
  shapeId/slideIndex changes text, not layout.
- Never call `powerpoint_add_speaker_notes`: speaker notes are unsupported in
  this snapshot. Offer a separate Word speaker script if the user wants it.
  Do not promise animations, SmartArt or full theme control through these tools.

Make one write, then independent readback of slide content/order before the next
mutation. New slides are selected when the receipt's highlight succeeds; still
prefer an explicit verified index for subsequent edits. A highlight-only failure
is not permission to duplicate a slide. On uncertain writes, inspect the actual
slide inventory/content; never retry creation blindly. Confirm deletion or broad
replacement and consult undo availability before promising recovery.

## Verification is more than a receipt

Readback checks exact wording, shape IDs, positions/dimensions where returned,
slide count/order and preserved neighbors. It does not prove visual hierarchy,
font rendering, contrast, clipping, overlap or line wrapping. Inspect the actual
slide in the host at presentation size and thumbnail size; use an authorized
screenshot/visual review if available, otherwise ask the user to confirm these
visual checks. Fix only the observed issue with approval, then read back again.
Do not claim a rendered review from a text-only result.

This is static tool-contract guidance, not a new native-host acceptance run.
Report schema/mock checks, native-host readback and visual acceptance separately.
A clean bundle or build is not evidence that a specific deck looks correct.

## Distinguish registration from source implementations

The registered canonical `powerpoint_add_image` and `powerpoint_duplicate_slide`
descriptions still say unsupported/diagnostic-only. Their executor source does
contain image insertion via Office Common API and duplication via export/reinsert;
source paths and strict mocks do not resolve the stale model-facing contract or
qualify every host. Do not advertise those diagnostic-only entries as portable
operations. Speaker notes really throw Unsupported in this snapshot.

Separately, registered shared `office_insert_image` supports app powerpoint:
exactly one permitted local PNG/JPEG path or guarded public HTTPS PNG/JPEG URL,
explicit slideIndex and point geometry, alt text/credit, one approval and bounded
readback status. HTML returns NeedsRaster for PowerPoint: render separately with
permission then supply a PNG path. This dedicated path must not be over-banned
because the canonical add_image description is stale. Read status/uncertainty,
then verify slide content and visuals; do not rasterize an entire editable deck.
