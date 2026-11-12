<!-- SPDX-License-Identifier: Apache-2.0 -->
# Reference

How Workbench fits together, every tool family, its limits, the bridge's security and the full disclosures.

## How it fits together

```
Hermes Desktop
├── chat (agent)  ── excel_* / word_* / powerpoint_* ──► plugin (this repo, Python)
│                                                         └─ loopback HTTPS bridge, https://localhost:47615-47619
└── preview pane (Electron <webview>)                            ▲  long-poll / result (same origin by default)
    └── Excel | Word | PowerPoint for the web (user signs in)    │
        └── "Workbench" task pane     served by the bridge: /office/live/<app>/taskpane.html
              └─ Office.js  →  the live document
```

* **The add-in comes from the plugin.** By default the bridge itself serves the add-in: the pinned tree's `addin/`
  files on `https://localhost:<port>/office/live/` (the port the bridge bound: 47615, or the next free one), only
  names listed in `addin/SHA256SUMS` whose sha256 still matches, never `addin/dev`. The pane and the bridge are then
  one origin. The bridge's certificate comes from a local certificate authority created for this Hermes profile
  (`plugin-data/hermes-office-desktop/tls`); Office loads the pane only after you trust it once yourself, with the
  command `office_connect` (`data.trust`), `/workbench trust` or `hermes workbench trust` prints. The plugin never
  changes certificate trust.
* **One add-in base URL.** The plugin setting `addin_base` (default `https://localhost:47615/office/live/`, whose
  port follows the bound port) is the one source of truth: the manifests' task-pane, icon and support URLs and
  every URL in `firstTimeSetup` come from it, and the bridge's CORS check compares the request's `Origin` with its
  origin (scheme, host and port) only. A self-hosted site is optional (see "Self-hosting the add-in"); the image
  fetch User-Agent names its support page, or the public repository for the loopback default. The older setting `allowed_origin` still works: an origin alone means
  `<origin>/office/live/`. If both are set they must agree (else `Configuration` error); `addin_base` set to the
  default value does not override an `allowed_origin`.
* **Manifests, no download.** `office_connect` writes the five manifests for that base (`scripts/build_addin.py`'s
  `render_manifests`, the same text the site build emits) to `~/Documents/Omnia Works/add-in manifests/` (plugin
  setting `manifest_dir`), returns the app's file in `data.manifestPath`, and its `firstTimeSetup` says: Home >
  Add-ins > More Add-ins > My Add-ins > Manage My Add-ins > Upload My Add-in > Browse, choose `manifest-<app>.xml` in
  Documents > Omnia Works > add-in manifests, then Upload (OneNote: Insert > Office Add-ins > My Add-ins > Manage My
  Add-ins > Upload My Add-in). This is the path that worked in every live run. Writes are atomic and idempotent; a file whose hash is not the one the plugin last
  wrote there (recorded in `.omnia-works-manifests.json`) is the user's and is never overwritten: the plugin writes
  `manifest-<app> (Omnia Works).xml` instead. If the folder cannot be written, `manifestPath` is absent and
  `firstTimeSetup` says to set `manifest_dir` and run `/workbench manifest` (a self-hosted site: it names
  `<addin_base>manifest-<app>.xml`). `/workbench manifest` in a chat, or `hermes workbench manifest` in a terminal
  (which finds this profile's running bridge over HTTPS and never starts one), writes the five manifests for the
  bridge's current port.
* `office_connect {"app": ...}` mints a single-use 8-character code bound to **the chat
  session that asked**. The user types it into that app's task pane. That is the only way
  a pane gets a bearer token; another chat (or any other local page) cannot drive it.
  **One chat, one pane** per app: re-pairing a chat replaces its old pane, and a pane that
  re-pairs with another chat drops its old binding.
* Tools reuse the native bridge's canonical schemas byte for byte: Excel 27 of 28
  (`schemas/excel.json` sha256 `0bce18cb…956f`), Word 12 (`0c3fc553…2bbc`), PowerPoint 18
  (`f1c0e19a…1728`), plus `office_connect`, `office_status`, `office_disconnect`,
  `office_undo`, `office_insert_image` (62 tools); OneNote adds 8 host tools and 8 brain tools; Outlook adds 4
  mail tools, 3 calendar tools and `brain_from_email` (86 registered). One of them does nothing:
  `powerpoint_add_speaker_notes` is kept for schema parity but always answers `Unsupported` (Office.js for the web
  has no speaker-notes API). A registered tool is not a live-tested capability: see
  [live-evidence.md](live-evidence.md) for what the live runs exercised.
* **Excel plans:** `excel_execute_plan` is deliberately not registered. Its native guards
  (trusted turn/task/tool-call metadata, one inspect + one execute per turn, operation
  fingerprints, replay reports) depend on a companion-side plan authority and host turn
  metadata that a browser task pane cannot provide; porting them without that authority
  would give a weaker guard under the same name. The 27 primitives cover the same edits.
* **What a receipt guarantees.** Every edit asks for approval first. Every result says what happened: a
  mutation the pane acknowledged (most primitives), or for composite and picture tools a receipt with its own
  readback (`verified` = written and read back as expected; `applied` = written, not read back; `partial`,
  `uncertain`, `failed`). A `verified` receipt proves the checks it names, not every property of the document.
  There is no universal post-write readback and no transaction across several tools.
* **Undo:** every mutating tool reports `data.undo.available` (with `undoUnavailable` and the reason when it is
  false); `office_undo` reverts the newest Hermes change (LIFO, 20 deep, per pane). The pane snapshots what it will
  change (Excel: formulas, number formats and cell formatting of the range; Word: body OOXML; PowerPoint: created
  ids, prior text/font/layout/position, a base64 export before a slide delete). Not undoable, or only partly:
  `excel_delete_sheet`; password protect/unprotect (the password is never stored); an Excel edit over more than
  10,000 cells (no snapshot); a Word edit when the body is too large to snapshot; PowerPoint text replaced through
  a bare selection; a slide delete when the slide export (PowerPointApi 1.8) is unavailable (with it, the slide
  comes back with a new id); a OneNote page Hermes created (add-ins cannot delete pages: undo opens it for you to
  delete). Undo history is lost on pane reload (OneNote's change ledger survives it). Outlook drafts and event
  forms are autosaved by Outlook itself; the plugin never calls a save API.
* **Read limits:** `word_read_document_text` returns at most the first 20,000 characters of the document body
  (not headers, footnotes, comments or tracked changes), `outlook_get_item` and `calendar_get_event` at most
  20,000 characters of the body, `excel_read_range` at most 5,000 cells per call, and the PowerPoint read tools
  return readable shape text within an output budget (no notes, no pictures). A cut-off read says `truncated`;
  the skill tells the model to say a summary covers only the part it read.
* **OneNote (web only, OneNoteApi 1.1):** 8 host tools (`onenote_get_context`,
  `_list_sections`, `_list_pages`, `_get_page`, `_create_page`, `_append`, `_insert_table`,
  `_open_page`) and 8 brain tools (`brain_sync`, `brain_ask`, `brain_links`,
  `brain_link_both_ways`, `brain_decision_page`, `brain_make_deliverables`, `brain_why`, `brain_resurface`).
  OneNote add-ins can read the content of **only the open page**; titles and urls of every
  page are cheap. So the index grows by *visit*: while paired, each page you open is read
  by the pane and indexed in `<Hermes home>/plugin-data/hermes-office-desktop/brain.db`
  (SQLite, FTS5 when available, per profile, per notebook). When one gateway serves several profiles,
  each profile's index is its own file; a call or pane event that cannot be tied to the owning profile
  is refused rather than written elsewhere. `brain_sync` adds the whole
  notebook's titles without navigating; `brain_sync {"crawl": true}` (approval) opens pages
  one by one and returns to where you were. Every answer says "indexed N of M pages".
* **Links without rewriting your text:** `[[Title]]` / `[[Title|label]]` are resolved against
  page titles (exact, then case/space-insensitive; duplicates are reported as ambiguous;
  `Section/Title` disambiguates). Hermes keeps one "Links & Backlinks" block per page that
  it owns and regenerates (delete its own block, add the new one). Unresolved titles come
  back as an offer to create the page.
  Hermes only replaces blocks this profile wrote; blocks from another profile or person are left alone.
  The ledger stores a fingerprint of each block's text when Hermes writes it; a block you edited
  (even with its "Links & Backlinks" heading intact) is never deleted by a refresh or by
  `office_undo`: it is kept, reported as `partial` / `notOwned` with the next step, and treated as
  yours from then on (the pane re-checks the fingerprint right before any delete).
* **Two-way links, like Obsidian (within the add-in's limits).** OneNote add-ins can read and
  write only the open page, so two-way linking has two parts:
  * **Live Links / Backlinks panel (read-only).** While paired, the pane shows the open page's
    links and backlinks from the local index every time you open a page (it comes back in the
    reply to the pane's own page-visit message over the existing 127.0.0.1 bridge: no new port,
    no new secret). Each entry opens that page when clicked. It says "indexed N of M pages" and
    reminds you to run `brain_sync` for coverage: backlinks come only from pages Hermes has
    indexed. The panel never writes to the notebook; note text in it is rendered as text.
  * **`brain_link_both_ways`** (approval) takes the `[[links]]` the index recorded for one page
    (by `pageId` or title) and, for each linked page (at most 10, **all named in the one
    approval prompt**), visibly opens it, refreshes only this profile's "Links & Backlinks"
    block there (now showing the backlink), reads it back, and finally reopens the page you were
    on. A page that already carries someone else's Links & Backlinks block is left as is and
    reported (`skipped`). Each page gets a receipt (`verified`, `applied`, `partial`,
    `uncertain`, `failed`, `skipped`, `not run`) and `status` is the weakest; nothing is
    retried, and an uncertain write or a page change by you stops the run. Every updated page
    has its own ledger row, so `office_undo {"app": "onenote"}` reverts them newest first, one
    call per page. If the link list changed after the approval was shown, the call is refused
    before anything is opened.
* **OneNote undo = change ledger:** every OneNote write gets an operation id and a ledger row
  before it runs; the pane re-checks the paired notebook right before writing. `office_undo
  {"app": "onenote"}` removes Hermes-added blocks (and survives pane reloads). Add-ins cannot
  delete pages: undoing a created page marks it "manual delete" and opens it for you. A write
  whose result was lost is reconciled through the pane's operation journal, never re-run blindly.
* **Deliverables:** `brain_make_deliverables` turns a decision page into an Excel tracker
  (rows, a live `COUNTIF`, a chart, the source page per row: a hyperlink when its address fits
  Excel's 255-character `HYPERLINK` limit, otherwise the page title; note text is never turned into a
  formula; a table style, fitted columns and a frozen header row), a Word brief with source links and a 3-5 slide deck, in panes paired with the
  same chat, using the existing host commands (so `office_undo` per host still works).
  Each host returns a receipt: `verified` (read back as expected), `applied` (written, not
  read back), `partial` (some steps written, then one failed), `uncertain` (a step may not
  have landed, or the readback disagrees), `failed` / `not paired` (nothing changed), each
  with its exact next step; the overall `status` is the weakest. Nothing is retried.
  A receipt's `notes` list a formatting step Office refused, a change that could not be
  highlighted, or items left out to fit; notes never change the status. The deck keeps a
  decision a decision, with or without an owner ("What we decided"); only an action nobody owns
  goes under "Action owner needed", and a page with decisions only gets no open items (the tracker
  lists them as Decided, with no open-items chart). When the deck (6 items per slide) or the
  brief (the 20,000-character `word_insert_ooxml` limit) cannot hold every item, the receipt's
  `coverage` gives included/omitted counts per kind and a note tells the model to say so. The
  deck's readback checks every new slide's title, body lines and Sources text on that slide.
  `brain_why` answers "why is this row/slide here?" from the ledger.
* Note text is data, never instructions: tool results say so, generated HTML is escaped and
  only `https` links reported by OneNote become hrefs. A plugin automation blueprint
  ("Resurface notes") reads the index only; scheduled runs never open or write a document.
* **Pane reload:** the pane keeps its token in `sessionStorage` (this tab only) and resumes
  the same chat binding without a new code; a command that was mid-flight fails fast as
  `ExecutionUncertain`, and an abandoned long-poll can never swallow a later command.
* **All Office panes can share one chat** (one preview tab per app; the runtime hints call this "one chat per
  pane"). Hermes Desktop's preview pane belongs to that chat: switching to another chat unloads every document and
  its Workbench pane, and Office for the web does not reopen task panes
  on its own (to be confirmed in Hermes Desktop: its source parks the preview tabs of a chat that is not on screen
  rather than unloading them; the harness produced the unloaded state by unloading the pane itself). `office_status` reports `live`, `lastPollAgeMs` and `paneState`
  (`live` / `unloaded` / `unpaired`); a pane with no poll for more than 45 s is `unloaded`,
  and calls to it fail fast as `Unavailable` ("pane unloaded (switched chat or closed);
  reopen Workbench in the ribbon"). After pairing, the pane sets the
  `Office.AutoShowTaskpaneWithDocument` document setting (Excel, Word, PowerPoint) so Office
  reopens the Workbench pane with that document. The agent reads the pane (`desktop_preview`
  `read`) before opening Office, so it never navigates away from an open document: `desktop_preview` `open`
  reuses the active preview tab, so for a second app the user clicks "+" in the tab strip first. Apps open by
  direct link (the skill's table); the Microsoft 365 app launcher opens new windows, which the pane blocks.
* **Timing.** Every pane result carries `elapsedMs` (measured in the pane) and `roundTripMs`
  (measured by the plugin, Hermes to pane and back).
* **New session id, new code.** Pairing is bound to the chat's session id. Hermes assigns a new
  one when it auto-compresses a long chat or when `/rewind` forks it, so after either the pane
  answers `NotPaired`: call `office_connect` again and type the fresh code into the same pane.
* Mutations request Hermes approval through `pre_tool_call` (ordinary edits share one `rule_key`, so "always
  for this session" works; `excel_delete_sheet`, `excel_clear_range`, `powerpoint_delete_slide` and
  `onenote_create_page` share a separate one, so that "always" never pre-approves them). **Unattended runs** are blocked outright, for reads too, even
  when an operator configured `approvals.cron_mode`/`unattended_mode` to approve. A run is
  unattended when `ctx.current_cron_execution()` reports a cron fire (including a job hosted
  inside the gateway), when the `HERMES_CRON_SESSION` marker is set, or when the session
  platform is `webhook`, `msgraph_webhook`, `api_server` or `cron`. Both markers are read per
  turn through `gateway.session_context.get_session_env` (the process environment when that
  module is not importable). Nothing in the plugin waits for a person; every error carries a `hint`.

## Outlook on the web (fifth app)

* **Tools:** `outlook_get_item` (read-only: the ONE open message's subject, from, to/cc, date, body as text and
  attachment names/sizes; never attachment contents, never other mailbox items), `outlook_insert` (in a draft you
  are writing: text at the cursor, or the subject; approval; `office_undo {"app": "outlook"}` restores the prior
  body or subject and only on that same draft), `outlook_draft_reply` and `outlook_draft_new` (approval: Outlook
  opens a pre-filled reply to the sender, or a new message with a subject and body and no recipients, through
  `displayReplyForm` / `displayNewMessageForm`; **limitation, seen live in A10b:** Outlook on the web opens these forms
  as a popout window, and the Hermes Desktop preview pane, a `<webview>` without `allowpopups`, blocks popouts, so the
  request succeeds but no draft appears there; the receipt says `requested`, never `opened`, and its next step is to
  click Reply yourself and let Hermes `outlook_insert` the text into that draft), and `brain_from_email` (indexes the open email into this
  profile's local brain index as `email:<key>`; with decisions it also creates a OneNote decision page from it
  through `brain_decision_page`, quotes checked against the email; when that section already has a page with the same
  title, from an earlier run, the new page gets the next free title, `<title> (2)` up to `(20)`, the old page is never
  opened or changed, and the receipt's `titleAdjusted` says so). If OneNote for the web stops answering the pane (every
  request fails with `GeneralException`, seen after a navigation to another section), the pane answers `PaneReload`:
  reload the OneNote tab. For a change (create, append, table, links block, picture, undo) that answer means the
  change may or may not have been made: the ledger keeps it `uncertain`, the receipt says to reload the tab and read
  the page, a multi-page run (`brain_link_both_ways`) stops there, and an identical call is held back until a
  successful read shows whether it landed: `onenote_get_page` of that page (for a new page, `onenote_list_pages` with
  `all`, complete). If it landed the call is refused (`AlreadyApplied`, nothing sent); if not, it runs once. Retrying,
  a failed read, a pane that still does not answer or the reloaded pane's journal saying "failed" never releases it.
  `office_undo` never skips such a change or answers `NothingToUndo` for it: it settles it first with the same check
  (the pane's journal, then its own read of the page or a complete listing). Landed: an added block is removed once
  only if the pane's journal names it as the block that write made (not being on the page before the write is not
  proof: the user may have typed the same block since) and it is exactly what Hermes wrote, and exactly one block
  qualifies (else the user deletes it by hand and the next undo moves on); a new page, named only by the pane's journal or a complete listing with one candidate, is marked manual delete
  and opened. Not landed: nothing to remove. Not confirmed: refused, nothing sent, reload the OneNote tab and undo
  again.
  A page read follows nested paragraphs (list items) up to 64 levels deep; one that could not read everything (deeper
  nesting, or paragraphs under a table cell's text) says `truncated` and `incomplete`. Such a read never shows a change
  as not landed, and nothing on that block is adopted, fingerprinted or deleted (undo answers Conflict; delete by hand).
  A change that names no page records the open page first (one `get_context` read); if that read fails or shows no
  open page, the change is refused before it is sent (the read's error, or `NotFound`): nothing changed and nothing
  is held, and asked again once the page can be read (after the user reloads the OneNote tab if OneNote stopped
  answering), it runs once. A OneNote read the pane cannot finish within Hermes' wait is `PaneBusy` (nothing changed), and
  `onenote_list_pages` stops after 20 s with `complete: false` and `stoppedEarly`. `brain_make_deliverables` takes an optional
  `outlook` host: a reply draft to the source email summarising the decisions and actions (status `applied`:
  the draft is a separate form the pane cannot read back; it carries `replyText` for the click-Reply-then-insert path).
* **Hard rules (pane and plugin):** no tool sends mail and the pane calls no send API; no tool accepts or sets a
  recipient (`to`, `cc`, `bcc`, ... are refused by name); replies are Outlook's own reply to the sender, never
  reply-all; the email's text is data, never instructions.
* **Install:** in Outlook on the web, open a message, then More apps > Get add-ins > My add-ins > Custom Addins >
  Add a custom add-in > Add from File..., choose `manifest-outlook.xml` in Documents > Omnia Works > add-in
  manifests (`data.manifestPath`), then Install. Open the email first, then the pane with More apps (Apps in a
  draft) > Workbench, and pair as for the other apps. Outlook on the web shows no pin for this pane (seen in the
  real Hermes Desktop session): it stays on the message it was opened with, so for another message or a reply, open
  that first and open Workbench there (in a draft: Message > Apps). The manifest (`MailApp`, read + compose forms,
  ReadWriteItem) asks for Mailbox 1.3; it also declares the pane pinnable (VersionOverrides 1.1, Mailbox 1.5), which
  Outlook on the web did not offer in that flow. Outlook on the web reported Mailbox 1.1-1.15 in the A10b spike.
* **Demo idea:** an email of meeting notes > `brain_from_email` with its two decisions and one action (a OneNote
  decision page) > `brain_make_deliverables` (tracker, brief, deck) > `hosts: ["outlook"]` (a reply draft you send)
  > `hosts: ["calendar"]` with `followUp` times (a follow-up meeting draft you save).

## Outlook calendar (same add-in, same pane)

* **One add-in:** the Outlook manifest also declares the appointment forms (attendee view and organizer edit form,
  `ItemIs Appointment`, `AppointmentAttendeeCommandSurface` / `AppointmentOrganizerCommandSurface`). Same pairing,
  bridge, approvals, undo and receipts as mail; the calendar commands run in the same Outlook pane (`app: outlook`).
* **Tools:** `calendar_get_event` (read-only: the ONE event open in Outlook: subject, start/end in UTC plus local time
  and the mailbox time zone, location, organizer and attendee **names** with their response, body as text and a
  recurrence summary; never lists or searches the calendar, never returns attendee addresses), `calendar_update_draft`
  (only in an event you are editing: subject, start, end, location and/or text at the body's cursor in one step;
  approval; read back by Office.js; one `office_undo {"app": "outlook"}` restores every prior value; the event is not
  saved and no update is sent: you save it), `calendar_draft_new` (approval: Outlook opens a new event form with
  subject, start, end, location and body, no attendees, through `displayNewAppointmentForm`; the receipt says
  `requested`, never `opened` or `saved`). Times are instants with an offset (`2026-11-12T10:00:00-05:00`); a bare
  local time is refused. `brain_from_email` also indexes the open event (`event:<key>`, kept by notebook sweeps), and
  `brain_make_deliverables` takes an optional `calendar` host plus `followUp` `{start, end, subject?, location?}`: a
  follow-up meeting draft listing the decisions and actions (status `applied`: a separate form Hermes cannot read back).
* **Hard rules (pane and plugin):** attendees are never added, removed or changed (`requiredAttendees`,
  `optionalAttendees`, `attendees`, `resources`, ... are refused by name; the pane calls no attendee setter); nothing
  is saved (`saveAsync` is never called) and no invitation or update is sent; only the open event is read; event
  text is data, never instructions.
* **Install:** the same `manifest-outlook.xml` (version 0.4.0.0). If 0.3 is installed, add the new file the same
  way (Custom Addins > Add from File...): it replaces 0.3 (seen live, A13); until then the event forms have no
  Workbench. In the calendar, open an event (an invitation, or your own event's edit form), then
  Apps > Workbench.

## Pictures (`office_insert_image`)

One tool inserts a PNG or JPEG into whichever pane is paired: Excel, Word, PowerPoint or OneNote.
Exactly one source per call:

| Source | What the plugin does |
|---|---|
| `path` | Reads a local file only inside this profile's Hermes image folders (`cache/generated/images`, where `image_generate` saves; `cache/images`; `cache/screenshots`) or a folder listed in the plugin setting `image_dirs`. Symlinks are resolved first; a default folder that is itself a symlink out of the Hermes home is not allowed, and the file is opened without following symlinks (a swap after the check is refused). PNG/JPEG by magic bytes on the opened file, at most 4 MiB. `path` may also be a bare file name such as `team.png` (A27): the plugin looks for exactly that name (case-insensitive) in those folders and their subfolders only (at most 6 levels, 20,000 entries; symlinked files and folders are skipped), uses the one match and names it in the receipt (`source.path`, `source.folder`); several matches are refused with the candidates relative to their image folder, and no match names the folders searched. |
| `url` | Fetches an `https` URL itself (see Disclosures for the guards) and records the URL and the optional `credit` in the alt text (OneNote also adds a caption line). |
| `html` | Word and OneNote insert it natively after an allowlist sanitizer (no `<style>` or stylesheets; `style=""` keeps only simple formatting properties, never anything that can name a resource; at least one picture must survive, otherwise it is refused, and the readback must show a picture); Excel and PowerPoint return `NeedsRaster` (render it with the browser tool, screenshot to PNG, call again with `path`). No headless browser is bundled. |

| App | Office.js | API set |
|---|---|---|
| Excel | `worksheet.shapes.addImage(base64)`; anchor from `Range.left/top`; `Shape.altTextDescription` | ExcelApi 1.9 (Range position 1.3) |
| Word | `insertInlinePictureFromBase64` / `insertHtml`, wrapped in a tagged content control (`insertContentControl`, `contentControls.getByTag`) | WordApi 1.2 (picture), 1.1 (HTML, content controls) |
| PowerPoint | the existing `add_image` path: `Office.context.document.setSelectedDataAsync(..., {coercionType: image})`; alt text via `Shape.altTextDescription` | ImageCoercion 1.1 + PowerPointApi 1.5 (slide selection); alt text PowerPointApi 1.10 |
| OneNote | `page.addOutline(left, top, '<img src="data:...">')`; `Image.description` for alt text | OneNoteApi 1.1 |

Each insert asks for one approval ("PowerPoint · Q4 deck · insert image on slide 2 (undoable)"),
then reads the picture back by its id (shape id, content-control tag or outline id). The receipt
says `verified`, `applied`, `partial` (alt text could not be set) or `uncertain`, with the next
step; nothing is retried. `office_undo` for that app removes exactly the inserted picture
(OneNote: through the change ledger). `powerpoint_add_image` is kept unchanged.

## Limits: five bridge ports, shared by every profile

The pane can only look for Hermes on five fixed loopback ports (https://localhost:47615-47619, bound on 127.0.0.1). Under a
multiplexed gateway each Hermes profile that uses Workbench runs its own bridge (its own pairings,
tokens and OneNote index; nothing is shared between profiles), and each bridge needs one of those ports.

* **At most five profiles at the same time.** A port is taken on the first `office_connect` (or, for a chat
  that is still paired, on its next Office tool call or `office_status`).
* **Idle profiles give their port back.** After 10 minutes with nothing to serve (no pane polling, no
  unused pairing code, no command or Office tool call in progress, no request being handled) a profile's
  bridge closes its port. Pairings are kept in memory: if that chat's pane is still open (or reopened
  within the 15-minute pairing lifetime), the next Office tool call takes the first free port again,
  which may be a different one, and the pane finds its bridge there by itself (it looks for its bridge's
  random `instance` id in `/v1/hello` on all five ports and never sends its token to another profile's
  bridge). While a pane is open and polling, its port is never given back.
* **The bridge serves the pane, so the pane loads only while a bridge listens on its port.** Reopening Workbench
  after an idle release shows Office's add-in error until the bridge takes a port again: ask Hermes (`office_status`
  or "connect to <app>"), then reopen Workbench. A pane that stayed open waits and reconnects by itself, also on a
  different port: every Workbench bridge accepts panes served by its five sibling ports.
* **When all five are taken**, `office_connect` fails with `PortsBusy` and `office_status` reports
  `bridgeBusy`, both saying: "All five Workbench bridge ports (127.0.0.1:47615-47619) are in use: five
  Hermes profiles are using Workbench right now (or another program holds a port). Close a Workbench
  pane that another profile's chat paired, or wait up to 10 minutes: a profile with no open pane gives its
  port back after 10 minutes. Then call office_connect again."
* `office_status` never takes a port for a chat that has no pairing (it only probes whether one is free,
  `bridgePortsFree`), and reports `idleReleaseInSeconds` while a listening bridge is idle.
* Why not one shared bridge per process that routes by profile? Each profile loads its own copy of the
  plugin, possibly a different version, with its own `addin_base` and its own OneNote index; one
  copy serving another profile's pairings would put every profile's pane traffic, tokens and page-visit
  writes through whichever copy started first, undo the per-profile isolation (c7ce69b), and stop every
  profile's bridge whenever that copy is unloaded. Giving idle ports back keeps each profile's bridge
  separate and is enough for the realistic case: several profiles tried the plugin, a few use it at once.

## Bridge security

The pairing code protects against web pages, not against other programs on the same computer.

* Binds `127.0.0.1` only, HTTPS only, fixed ports 47615-47619 (the pane probes exactly these); on Windows the port
  is bound exclusively. An idle bridge gives its port back (see "Limits"), and a request on a connection to a
  released listener gets 503.
* `Host` must be exactly `localhost:<bound port>` (else 421: DNS-rebinding guard, no numeric-host fallback).
* Every `/v1` route requires an allowed `Origin`. Default: one of the five Workbench loopback origins
  `https://localhost:47615` to `https://localhost:47619`: the pane's own port is same-origin, the others are sibling
  Workbench bridges serving the same hash-checked files (a pane needs them after an idle release moved its bridge,
  or when several profiles share one uploaded manifest). A same-origin `GET` carries no `Origin`; it is accepted
  only in this default and only when `Sec-Fetch-Site` is absent or `same-origin`. Any other origin gets 403. With a
  self-hosted `addin_base`: exactly that site's origin, as before.
* Add-in files (`/office/live/...`): `GET`/`HEAD` only (else 405); only names listed in `addin/SHA256SUMS` whose
  sha256 matches at that request (a Windows checkout with CRLF line endings is sent in its LF form, and only when
  that form is the pinned content: the bytes sent always match the list); no symlinks, nothing outside `addin/`, never `addin/dev`; encoded separators
  (`%2f`), `..`, backslashes, NUL and empty segments get 404; a foreign `Origin` gets 403. Fixed MIME types,
  `X-Content-Type-Options: nosniff`, the panes' own Content-Security-Policy, `Referrer-Policy: no-referrer`. No
  `frame-ancestors`: Office for the web frames the pane from several Microsoft hosts.
* TLS: a per-profile local CA (RSA 2048, name-constrained to `localhost` and `127.0.0.1`, 825 days) and a
  `localhost` certificate (renewed under the same CA), owner-only files in `plugin-data/hermes-office-desktop/tls`.
  The plugin never adds it to an OS trust store; you run the printed command yourself.
* Codes: 30-symbol alphabet, 8 chars, 10-minute TTL, single use, app-specific; 5 wrong
  tries burn live codes; 20 failures/minute trips a global limit.
* Pane tokens: 256-bit, kept in the pane tab's `sessionStorage` (cleared when the tab
  closes), revoked by re-pairing, `office_disconnect`, pane close, or 15 minutes unseen.
* Bounded bodies (8 MiB), bounded waits, no redirects, no automatic retries. A command the
  pane never picked up reports `Timeout` (not executed); one it picked up but never
  answered reports `ExecutionUncertain`, naming what to check (the workbook, document,
  presentation, page, draft or event).
* The plugin reads no credentials and needs none. For the unattended check it reads only
  `ctx.current_cron_execution()` and two non-secret Hermes session markers
  (`HERMES_CRON_SESSION`, `HERMES_SESSION_PLATFORM`).

## Requirements

Windows: one real Hermes Desktop session passed the 5-prompt story on 2026-11-10 with a single profile.
Windows: two Hermes profiles at once are not live-tested (port 47615 is no longer bound twice).
Requires the Hermes Desktop app (macOS 13+ or Windows); the browser Webapp (`webapp.enabled`, v0.23.0) cannot host Office for the web in its preview pane; Desktop must run on the same machine as the Hermes gateway (or with the bridge ports forwarded).

* On Windows, keep the app's tab in front while Hermes writes to it; background tabs can make writes slow and uncertain.
* The Webapp's preview pane is a sandboxed cross-origin `<iframe>`: Microsoft refuses to be framed there,
  sign-in and storage are blocked, and the frame cannot reach `127.0.0.1`. `desktop_preview` `read`
  reports `surface: 'iframe'` in that case; the agent then tells you to switch to the Desktop app instead
  of minting a pairing code.
* The bridge listens on the gateway host's loopback (`127.0.0.1:47615-47619`). A Desktop window chatting
  with a remote or SSH gateway needs those ports forwarded to the machine that shows the pane, and that machine must
  trust the profile's local CA (`ca.pem` from the gateway's `plugin-data/hermes-office-desktop/tls`).
* A Microsoft account and Office for the web (OneNote, Excel, Word or PowerPoint); for Outlook, an account
  with an Outlook.com (or Microsoft 365) mailbox.

## Self-hosting the add-in

Optional; the default needs no site (the plugin serves its own pinned add-in). To host the pane yourself:
`python3 scripts/build_addin.py --base https://<host>/<any path>/ --out <dir>`, serve `<dir>`
over HTTPS at that URL, and set `plugins.entries.hermes-office-desktop.settings.addin_base: https://<host>/<any path>/`
in the profile's `config.yaml` (or the older `allowed_origin: https://<host>` when the site is at
`https://<host>/office/live/`). The bridge then accepts exactly that origin and serves no add-in files; the pane still
reaches the bridge on `https://localhost`, so the local CA trust step applies too. `.github/workflows/pages.yml`
builds such a site on GitHub Pages by hand only (`workflow_dispatch`); it is never the default. These steps follow the tested path; a full walk-through from a clean profile by
someone new to the project is still pending.

## Disclosures

* Opens a loopback HTTPS listener on `127.0.0.1:47615-47619` (in the process that runs Hermes
  tools) the first time `office_connect` runs; it closes after 10 idle minutes and opens again on the
  next Office tool call of a still-paired chat (or the next `office_connect`).
* **The task pane comes from the catalog pin.** By default the bridge serves the add-in page from the plugin's own
  `addin/` files, only those `addin/SHA256SUMS` lists with a matching sha256 (`scripts/build_addin.py --bundle`
  regenerates that list, the icons and `shared/build-info.js`). The Workbench pane footer shows the build it loaded
  (`Workbench add-in 0.4.0 · build <id> · source unknown` for the bundle, which cannot name its own commit; a
  self-hosted site build adds `source <commit>`): `<id>` is the first 12 hex of the sha256 over the `SHA256SUMS`
  lines of every other file, so a pane from an older browser cache shows a different id. It is a label for
  diagnosis, not an attestation. A self-hosted `addin_base` is outside the pin by definition.
* **A local certificate authority.** The first bridge start creates a CA and a `localhost` certificate for this
  profile in `plugin-data/hermes-office-desktop/tls` (owner-only). Trusting it is up to you: the plugin prints the
  command and never runs it. Remove that trust when you stop using Workbench.
* Needs Office for the web and a Microsoft account (you sign in yourself in the pane). Edits
  ask for Hermes approval. No telemetry.
* **Outbound HTTPS, for one purpose only:** the plugin fetches an image URL the agent passes to
  `office_insert_image` (`url`, or an `https` `<img>` inside `html`), and nothing else. Guards:
  `https` on port 443 only, no credentials in the URL; the host is resolved and every address
  must be public (loopback, private, link-local, CGNAT `100.64/10`, multicast, reserved and
  unspecified addresses are refused, IPv4-mapped/6to4/Teredo IPv6 included); the connection goes
  to the address that was checked; at most 3 redirects, each re-checked from scratch;
  `Content-Type` must be `image/png` or `image/jpeg` and match the file's magic bytes; the body
  is streamed into a hard 4 MiB cap; one 15 s deadline covers the whole fetch. No cookies, no
  authentication, no retries. The image bytes go only to the paired pane on `127.0.0.1`.
* Local image files are read only from this profile's Hermes image folders and the folders you
  list in `image_dirs`.
* The add-in page loads Microsoft's Office.js from `appsforoffice.microsoft.com`; it sends
  document data only to `127.0.0.1`.
* Document content returned by tools enters the Hermes conversation like any tool result.
* OneNote: while a OneNote pane is paired, the page you open is sent to this Hermes on
  `127.0.0.1` and stored, unencrypted, in this profile's `plugin-data/hermes-office-desktop/`
  `brain.db` (page text, titles, urls, links, tags, a ledger of Hermes' writes). Delete that
  file to forget everything. Nothing is uploaded anywhere by the plugin; text you ask about is
  sent to your configured Hermes model like any other tool result.
* **Outlook:** when you ask (any `outlook_*` tool or `brain_from_email`), the open message's subject, sender,
  To/Cc names and addresses, date, body text and attachment names/sizes are sent to this Hermes on `127.0.0.1`
  and enter the conversation like any tool result (so they reach your configured model); `brain_from_email`
  also stores them in `brain.db`. Attachment contents and other mailbox items are never read. Hermes never
  sends mail and never adds, removes or changes recipients: drafts open in Outlook for you to review, send or
  discard.
* **Outlook calendar:** when you ask (any `calendar_*` tool, or `brain_from_email` on an open event), the open event's
  subject, times and time zone, location, organizer and attendee names with their responses (no addresses), body
  text and recurrence summary are sent to this Hermes on `127.0.0.1` and enter the conversation;
  `brain_from_email` also stores them in `brain.db`. The calendar is never listed or searched. Hermes never adds,
  removes or changes attendees, never saves events and never sends invitations or updates.
* OneNote two-way links: the pane's Links / Backlinks panel is read-only (index data in the
  reply to its own visit message; it never writes). `brain_link_both_ways` navigates visibly
  through up to 10 linked pages and writes on each, so it always asks for approval and the
  prompt names every page.
* Developer-only: `scripts/livetest/chat.py` (used by `live_driver.py` and `warm.py`) imports
  private Hermes modules (`hermes_cli.gateway_runtime`, `hermes_cli.gateway_client`) to drive a
  live chat. `hermes plugins validate` warns about them (catalog rule 9). Hermes never loads
  these scripts; the plugin's runtime code uses only public surfaces.
* Without Microsoft Graph the plugin cannot see notebooks other than the open one, cannot
  sync in the background, cannot delete or move pages, cannot edit your text in place and
  does not read OneNote desktop or mobile.

## Layout

| Path | What |
| --- | --- |
| `__init__.py`, `bridge_server.py`, `bridge_tls.py`, `validation.py`, `plugin.yaml` | Hermes plugin (stdlib, plus `cryptography` for the local CA) |
| `onenote.py`, `brain.py` | OneNote host + brain tools; the SQLite index, link resolution, ledger, provenance |
| `outlook.py` | Outlook mail and calendar tools, `brain_from_email` (emails and events), the reply-draft and follow-up meeting deliverables |
| `images.py` | `office_insert_image` sources: allowed local folders, the guarded HTTPS image fetch, the HTML sanitizer |
| `skills/office-desktop/`, `skills/daily-note/`, `skills/meeting-to-decisions/` | Bundled skills: pairing/editing/undo; the daily-note template; meeting to decisions to deliverables |
| `addin/` | Bridge add-in, served by the bridge: `SHA256SUMS` (exactly the files it serves), icons (`assets/`), manifest templates (TaskPaneApp for 4 hosts, MailApp for Outlook), shared pane UI/bridge client/runtime, one executor per host, Office.js mocks and the fictional demo notebook (`dev/`, never deployed) |
| `scripts/build_addin.py` | `--bundle` refreshes `addin/SHA256SUMS`, icons and `shared/build-info.js`; `--base` builds a site to self-host (5 manifests, `SHA256SUMS`) |
| `scripts/demo/demo_notebook.py`, `docs/demo.md` | Demo storyline and the seed content for a real OneNote notebook |
| `scripts/serve_addin.py`, `scripts/live/` | Local live testing only: loopback HTTPS server, coverage plans, runner, CDP driver for desktop Office |
| `scripts/livetest/` | Automated live test in Office for the web: client Mac Electron harness (`macbook/`), Studio-side chat + UI driver, `run_macbook.sh` |
| `spike/electron-pane-harness/` | Electron 44 pane harness (all four hosts) and the Chrome control |
| `tests/` | `pytest` (plugin/bridge) and `node --test` (executors, bridge client) |
