<!-- SPDX-License-Identifier: Apache-2.0 -->
# Demo recording script (85 seconds)

For the person recording the demo. The story is one email turning into finished work: a OneNote decision page, a
tracker with a chart, a brief, a deck with a picture, a reply and a follow-up meeting request. Each result can be
traced back to the email. The harness runs this story two ways before anyone records it: `story` adds a routing
note to each prompt, and `story-bare` types these prompts exactly as you will (`scripts/livetest/run_macbook.sh
story-bare`; results are in [live-evidence.md](live-evidence.md)).

## Set-up (off camera)

* Hermes Desktop with ONE new chat. All five panes are paired to that chat, one preview tab per app: Outlook on the
  web with the meeting-notes email open in reading view, OneNote (an empty section), Excel (a blank workbook), Word
  (a short document) and PowerPoint (one title slide).
* Open each app by its direct link, and let Hermes open it: ask "open Outlook" (then OneNote, Excel, Word,
  PowerPoint) in the chat. Before each app after the first, click **+** in the preview pane's tab strip so the new
  app lands in a blank tab and the panes already paired stay open. Hermes opens `https://outlook.live.com/mail/0/`,
  `https://www.onenote.com/notebooks`, `https://excel.cloud.microsoft/`, `https://word.cloud.microsoft/` and
  `https://powerpoint.cloud.microsoft/` (work or school Outlook: `https://outlook.office.com/mail/`). For a new
  file, click Blank workbook, Blank document or Blank presentation on the app's start page. Do not use the
  Microsoft 365 app launcher: it opens new windows, which the pane blocks.
* Set every app up off camera, then pair all five panes again right before the take: a code lasts 10 minutes, and
  after Hermes restarts every pane must be paired again. Only you type the codes.
* Outlook: open the meeting-notes email first, then More apps > Workbench (the pane has no pin; it stays on that
  email). If the email is not in the Inbox, search All folders (it may be archived).
* OneNote: the first response confirms the tab opened and asks which notebook you want (unless you named it),
  then guides the observed My Files route; it does not create one by default. Confirm My Files is visible
  before recommending it. You click My Files, open the notebook in its separate window, copy the final
  opened-notebook browser address, and paste it into the Workbench OneNote tab's address bar. Not the file-list
  address, an `onenote:` app link, or a sharing link. Pair only after the notebook loads inside Workbench.
  Only My Files opening a separate window is user-confirmed, not end-to-end open/copy/paste/pair success.
  If My Files is absent, inspect/report the current page; if paste fails, capture the actual error and use only
  a verified fallback. Never automatically fall back to sharing/permission changes. Already open notebook:
  skip onboarding. For the first-response check, see [harness.md](harness.md#onenote-first-response-acceptance-manual-fresh-chat).
* Proposed OneNote recording checklist: after concurrent smoke work finishes and in a safe chat, verify the
  first response, then let the user perform My Files selection, external notebook opening, final-address copy,
  and Workbench paste; do not compete for the mouse. Record each step separately, then confirm the notebook
  loaded here before the user enters a pairing code. Stop on an absent control or paste error; record the
  actual state, not a successful take. Do not present this checklist as completed recording evidence.
* Excel: no cell may be in editing mode (press Esc), or Excel refuses the tracker.
* Click the Hermes chat box before typing each prompt: an Office app, the Outlook reply editor most of all, can keep
  the keyboard, and your words would land in the document.
* You click the app tabs. Hermes cannot bring a tab to the front and works only on the tab in front; if it asks
  you to click a tab, do that and continue.
* Record from a profile with self-learning off. If it is on, say on screen that the profile is warm (it may have
  learned the demo).
* Do not switch to another chat during set-up or recording: that unloads every pane. If it happens, come back and
  reopen Workbench from each app's ribbon to pair again.
* The email: subject "HOL live test: pilot sync". Body: "Decision: we ship the beta on Friday. Decision: we keep the
  current pilot scope. Action: Sam writes the FAQ by 12 November."
* The team picture is `team.png` in the profile's Hermes images folder: `cache/images/demo/team.png` under the
  profile's Hermes home (for example `~/.hermes/profiles/omnia-demo/cache/images/demo/team.png`). Prompt 2 names it
  by file name, so Hermes needs no memory of which picture is meant; the harness puts its own stand-in PNG at the
  same path in its test profile and types the same words. No image is generated on camera.
  With the bare prompt Hermes picks where the picture goes; the routed harness runs place it explicitly (left 600,
  top 300, width 288, height 162) because the default position covers the slide title. If it covers the title on
  camera, drag it aside or record again.
* Approvals: click Allow once for each prompt (Hermes Desktop labels it "Run once"). Do not choose "always".
* The account's address must not be visible. Crop or blur the message list, the address fields and the account
  menu. These are personal areas, the same ones the harness blurs in its screenshots.

## Prompts

Type these exactly, with nothing added. The `story-bare` harness mode types exactly these words and judges only the
outcome (what the documents and the index hold afterwards, and the receipts). The `story` mode adds a bracketed note
to each one, and that note is tool routing: it names the tool to call and its arguments (decision page id, section,
slide number, picture path, times). A presenter does not type it, so `story` results show the tools work, and
`story-bare` results show the bare prompts work.

1. `Read this email and turn it into a decision page in my notes, quoting the email for each decision and action.`
2. `Now make the tracker, the brief and the deck from that page, put team.png from my Hermes images folder on the title slide, draft my reply, and request a follow-up meeting next Thursday 10:00-10:30 in the Studio.`
3. `Why is the FAQ row in the tracker?`
4. `I clicked Reply. Fill in the reply text.`
5. `Undo that.`

Prompt 1 is one step because that is what Hermes does: asked to add an email to the notes, it can index the email
and make the decision page from it in one approved call. The script had two prompts here through `c94cd8d`; in
both bare runs on that commit the page already existed after the first prompt, so the second had nothing to do.

## Timeline

Times are on-screen seconds after editing. A live turn takes longer than its slot here. Measured on `d1f7aae` (chat
turn, the bare prompts, cold run 1 of 2, both passed every beat): prompt 1 13 s, prompt 2 63 s, prompt 3 6 s,
prompt 4 8 s, prompt 5 6 s ([live-evidence.md](live-evidence.md)). Each **CUT** marks where to cut that wait out.

| Time | On screen | Say / do | Expected result |
|---|---|---|---|
| 0-3 s | The meeting-notes email in Outlook | Caption: "Meeting notes end up in email. Then someone has to retype them into five apps." | The problem is on screen within 3 seconds |
| 3-22 s | Outlook, with the chat beside it, then OneNote | Type prompt 1, click Allow, then switch to OneNote | The email is indexed as "Email: HOL live test: pilot sync" and a new decision page opens. It shows "Source: Email: HOL live test: pilot sync", two decisions (beta on Friday, the current pilot scope) and one action (Sam, FAQ, due 12 November). Each item has a quote from the email, and none says "not found in the source page" |
| | | **CUT** while the turn runs | |
| 22-50 s | PowerPoint, then Excel, then Word | Type prompt 2, click Allow for each step. Show the deck first, then switch tabs | PowerPoint: four new slides (the decision page title, "What we decided" with both decisions, "Who does what" with Sam's FAQ action, "Sources" naming the email and the decision record) and the picture on the title slide. A decision with no owner stays under "What we decided"; an "Action owner needed" slide appears only for an action nobody owns. Excel: a "Decision tracker" sheet with the FAQ row (owner Sam, status Open, source the email), "Open items" = 1 and an "Open items by owner" chart. Word: the brief at the end of the document, with the decision record linked |
| 50-52 s | The chat | Point at the receipts | Excel, Word and PowerPoint say `verified`. The reply and the follow-up say `applied`: they were requested and not yet sent or saved. Outlook on the web opens those as popouts, which the pane blocks, so nothing appears yet. That is the honest result, not an error |
| 52-62 s | Excel, FAQ row selected | Type prompt 3 | The answer quotes the email ("Action: Sam writes the FAQ by 12 November.") and names the decision page and the email "HOL live test: pilot sync" |
| | | **CUT** | |
| 62-75 s | Outlook | Click Reply on the email, open Workbench from the draft's Message tab > Apps (not Home > More apps), click the chat box, then type prompt 4, click Allow. Hermes refuses while the pane is still on the email; that is correct | The reply draft fills with "Thanks for the notes. Here is what we decided:", the decisions, the action and the decision page title. It is addressed only to the sender. Do not click Send |
| 75-82 s | Outlook | Type prompt 5, click Allow | The inserted text disappears from the draft. The draft is still open and unsent |
| 82-85 s | End card | Caption: "One email, five apps. Every change came back with a receipt and can be traced to its source; the reply was undone in one step. Nothing was sent." | |

After recording, discard the reply draft and any event form Outlook opened. Hermes never sends or saves them.

## Where to cut, and what not to do

* Cut only the waiting while a turn runs (the two **CUT** marks, plus each approval click). Label any cut or
  speed-up on screen, for example with "sped up".
* Do not cut a failure into a success. If a receipt says `uncertain`, `partial` or `failed`, keep it on screen or
  record again.
* Say "requested", not "created", for the reply and the follow-up meeting: Hermes asked Outlook to open them.
* Do not show the account's address, the message list or other mail. Do not click Send or Save.
