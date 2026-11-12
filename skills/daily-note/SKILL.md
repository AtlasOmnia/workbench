---
name: daily-note
description: Use when the user asks to start today's note or a daily page in OneNote. Creates one dated page from this template, linked to the pages they mention, through onenote_create_page and brain_links.
version: 0.4.0
license: Apache-2.0
metadata:
  hermes:
    tags: [onenote, daily-note, template, second-brain]
---

<!-- SPDX-License-Identifier: Apache-2.0 -->
# Daily note (OneNote template)

Needs a OneNote pane paired with this chat (`office_connect` with `app: "onenote"`).

1. `onenote_list_pages` for the open section. If a page titled `Daily YYYY-MM-DD` (today, in the
   user's timezone) exists, open it with `onenote_open_page` instead of creating a duplicate.
2. Otherwise call `onenote_create_page` (asks for approval) with title `Daily YYYY-MM-DD` and this
   HTML, filling only what you know. Leave a line empty rather than inventing content:

   ```html
   <h2>Focus</h2><ul><li></li></ul>
   <h2>Meetings</h2><ul><li></li></ul>
   <h2>Notes</h2><p></p>
   <h2>Links</h2><p>[[Yesterday's daily title]] · [[Project pages the user named]]</p>
   ```

   `[[Title]]` and `[[Title|label]]` are Hermes link syntax; the user's text is never rewritten.
3. Call `brain_links` on the new page. It writes the one Hermes-owned "Links & Backlinks" block.
   For any title it reports as not created yet, ask the user before calling `onenote_create_page`.
4. To show the backlink, open the linked page (`onenote_open_page`) and call `brain_links` there.

Undo: `office_undo` with `app: "onenote"` removes Hermes blocks. Pages cannot be deleted by an
add-in; undo marks a created page for the user to delete and opens it.
