<!-- SPDX-License-Identifier: Apache-2.0 -->
# Demo: "Your notes remember the decision. Hermes turns it into the work."

Target length about 85 seconds (under the 90-second limit). One loop, shown working:
OneNote is the memory; Excel, Word and PowerPoint are the deliverables.

Everything in the notebook is fictional (Lumen Bikes, a city e-bike subscription; people
Priya, Sam, Lee). No personal or work data. The content lives in
`addin/dev/demo-notebook.json`; `python3 scripts/demo/demo_notebook.py --list` prints it and
`--prompt` prints one chat message that has Hermes create the pages (each with approval).

## Before recording (off camera, disclosed in a caption)

1. OneNote for the web, Excel, Word and PowerPoint for the web are open in the Hermes Desktop
   preview pane, signed in to a test Microsoft account, each with the "Workbench" add-in
   uploaded (Upload My Add-in) and paired to the same chat with its one-time code (the user
   types the code; Hermes never does).
2. Notebook "Lumen Bikes (demo)" with sections Projects, Meetings, Daily, seeded by the script.
   `brain_sync`, then open each page once (or approve `brain_sync {"crawl": true}`). The
   coverage caption is copied from the tool result on camera ("indexed N of M pages"), never
   typed in advance: pages created during the take change both numbers.
3. Blank workbook, blank document, blank presentation. Approvals: choose "always for this
   session" on the first edit only if the recording says so; otherwise show each approval.

## Storyline (about 85 s)

| Time | On screen | What proves it is real |
| --- | --- | --- |
| 0-7 s | OneNote beside the chat. Title card: "Workbench — works with Excel, Word, PowerPoint, OneNote and Outlook (mail and calendar) on the web", then "Your notes remember the decision. Hermes turns it into the work." | Pane badge "Connected to your Hermes chat". |
| 7-20 s | "Start today's daily note and link it to [[Q4 Launch]]." Approval card, then `Daily 2026-11-05` appears in the pane with its "Links & Backlinks" block. | Real approval prompt; the page is created in the notebook. |
| 20-30 s | Hermes opens Q4 Launch: its backlinks now list the daily note, the launch meeting and the pricing review. | `brain_links` block regenerated, user text untouched. |
| 30-42 s | "What did we decide about pricing?" Two-sentence answer citing *Launch meeting 2026-11-03* and *Pricing review 2026-10-20*; click one, the page opens in the pane. Caption: the reply's first line, copied verbatim from `brain_ask`'s `caption` ("indexed N of M pages, local index"). | `brain_ask` excerpts + `onenote_open_page`. |
| 42-55 s | On the meeting page: "Record the decisions and actions, then refresh this page's links." Decision page appears: two decisions, three actions; "Check supplier lead times" shows owner **Unknown (not in the notes)**. Back on the meeting page, its new "Links & Backlinks" block lists the decision page. | `brain_decision_page` quotes checked against the source; unknowns stay unknown. The `brain_links` block on the meeting page is the **last OneNote change** and the one the final beat undoes. |
| 55-72 s | "Turn this into the tracker, brief and deck." Excel: a styled tracker table (header row, fitted columns, frozen header), `Open items` = 3 (live COUNTIF), chart by owner, selected as it lands. Word: one-page brief with source links. PowerPoint: 4-5 slides ending with Sources. Each write selects what changed (range, chart, new slide) so the eye lands on it. | Per-host receipts with read-backs: `verified` / `applied` / `partial` / `uncertain` / `failed` / `not paired`, plus an overall status that is the weakest of them. Sequential host switches, no implied transaction. |
| 72-78 s | In Excel, point at the "Draft the pilot plan" row and ask by its text: "Why is 'Draft the pilot plan' in the tracker?" Answer quotes the source line, names the decision, and opens the meeting page in the pane; "Check supplier lead times" stays owner Unknown. | `brain_why` from the ledger (looked up by the item text you typed, not by the selection); no match or several matches is said out loud, never filled from another item. |
| 78-85 s | "Undo the last OneNote change." The meeting page's Hermes "Links & Backlinks" block disappears; the user's own notes on that page are untouched. End card: "Workbench · works with Excel, Word, PowerPoint, OneNote and Outlook (mail and calendar) on the web · local index · approvals · undo · no Graph, no sign-in to Hermes". | `office_undo {"app": "onenote"}` removes the block written in the 42-55 s beat. |

Cut order if over time: the provenance beat (show as a still), then the daily note (start at
the pricing question). Never cut approval, the coverage caption, or undo. If the links
refresh in the 42-55 s beat is cut, end the undo beat on Excel instead ("Undo the last Excel
change", `office_undo {"app": "excel"}` removes the tracker chart): without that block, the
last OneNote change is a created page, which Hermes cannot undo.

## Undo: what it can and cannot do

- `office_undo` reverts the most recent Hermes change in that app, one step per call:
  OneNote blocks Hermes wrote (Links & Backlinks, appended blocks, tables), Excel/Word/PowerPoint
  edits made through the pane.
- **Pages Hermes created (the daily note, the decision page) are not undone.** OneNote add-ins
  cannot delete pages; `office_undo` on a created page only opens it and marks it "manual
  delete" in the ledger. Delete those pages yourself after the take (right-click the page >
  Delete page).

## What the demo must not claim

- Not "your whole knowledge base": say "indexed N of M pages" exactly as the tool returns it.
- Not background sync, other notebooks, OneNote desktop/mobile, or page deletion by Hermes.
- If a host step fails on camera, keep the failure (the receipt says so) or re-record; do not
  edit a failure into a success. Label any cut or speed-up.

## Demo idea: email to decision page to deliverables to reply draft (Outlook on the web)

Not part of the timed loop above. With Outlook on the web (Workbench open on the email; there is no pin, so the
pane stays on that email) and OneNote paired to the same chat:

1. Open the meeting-notes email. "Turn this email into a decision page." (`brain_from_email` with the email's
   decisions and actions, quotes copied from the email: the OneNote decision page opens, linked to "Email: ...").
2. "Make the tracker, brief and deck from it." (`brain_make_deliverables`.)
3. "Draft a reply with what we decided." (`brain_make_deliverables` with `hosts: ["outlook"]`: Outlook opens a
   reply to the sender listing the decisions and actions.) The user reads it and sends it; Hermes never sends
   mail and never changes recipients.

## Demo beat: decision to follow-up meeting draft (Outlook calendar)

Not part of the timed loop above. After the decision page exists (from OneNote notes or from an email) and the
Workbench pane is open in Outlook on a message or event being read:

1. "Set up a follow-up meeting next Thursday 10:00-10:30 in the Studio." (`brain_make_deliverables` with
   `hosts: ["calendar"]` and `followUp` `{start: "...T10:00:00-05:00", end: "...T10:30:00-05:00", location: "Studio"}`:
   Outlook is asked to open a new event form titled "Follow-up: <decision page>" whose body lists the decisions and
   actions.) The receipt is `applied` / `requested`, never `saved`.
2. The user adds attendees, if any, and saves it. Hermes never adds attendees, never saves and never sends invitations.
3. Optional: in an event's edit form, "Move it to 10:30 and set the location to Room 2."
   (`calendar_update_draft`, read back; "Undo that." restores both with one `office_undo`.)

## Prompts (copy/paste)

1. `Start today's daily note from the daily-note template and link it to [[Q4 Launch]].`
2. `Open Q4 Launch and refresh its links.`
3. `What did we decide about pricing?`
4. `Open the launch meeting and record its decisions and actions, then refresh the meeting page's links.`
5. `Turn that decision page into the tracker, the brief and the deck.`
6. `Why is "Draft the pilot plan" in the tracker?`
7. `Undo the last OneNote change.`
