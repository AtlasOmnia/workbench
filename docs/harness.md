<!-- SPDX-License-Identifier: Apache-2.0 -->
# Test harness

Unit tests, desktop-Office live tests on Windows, and the automated live test in Office for the web.
Results: [live-evidence.md](live-evidence.md).

`scripts/livetest/` and `spike/electron-pane-harness/` are developer-only. Their local paths, private tailnet and host defaults describe the author's test setup, not a supported install requirement; override them for your own environment.

## Unit tests

CI (`.github/workflows/tests.yml`, ubuntu-latest and macos-latest) runs exactly this, with Hermes Agent
v0.23.0 (`2bd378ed`) checked out beside the plugin so the tests that import Hermes run instead of skipping:

```sh
python -m pip install pytest==9.1.1 ruamel.yaml==0.18.17 PyYAML==6.0.3        # Python 3.11, Node 24
PYTHONPATH=<hermes-agent checkout> python -m pytest -q -c tests/pytest.ini tests -rs
node --test tests/js/*.mjs
```

Locally, with `uv`:

```sh
PYTHONPATH=<hermes-agent checkout> uv run --no-project --python 3.12 --with pytest --with ruamel.yaml --with pyyaml python -m pytest -q -c tests/pytest.ini tests
node --test tests/js/*.test.mjs
uv run --no-project --python 3.12 --with pyyaml python spike/electron-pane-harness/run.py   # needs Electron 44 at HARNESS_ELECTRON
```

## OneNote first-response acceptance (manual, fresh chat)

`tests/test_onenote_first_response_contract.py` checks the static instruction contract,
not observed fresh-model behavior. Passing it does not prove skill loading, installed-copy
parity, or the model's first response. Existing Office tool/story receipts do not prove this gate.

After the authorized demo-copy sync, verify the installed `skills/office-desktop/SKILL.md`
against the candidate source on each target. Use the real Hermes Desktop app, not the browser
webapp or the Office harness's prepared notebook. Wait for any concurrent smoke test to finish;
do not replace another app's active document. Switching chats unloads preview documents, so use
a safe new chat with a blank active preview tab and no notebook link in its context.

1. Send only `open OneNote tab`. Record the prompt, loaded skill/tool calls, tab-open result,
   observed landing controls, and the FIRST assistant response after opening, not a later troubleshooting turn.
2. Pass only if that response says the **tab** opened (not that a notebook is ready), asks which
   notebook the user wants, says up front that clicking a notebook in the list opens a window the pane
   blocks, and, when My Files is observed, immediately guides: click My Files for a separate window, open
   the notebook there, copy its final browser address, paste it into the chat so Hermes opens it in the
   Workbench OneNote tab, then connect only after it loads here. No default notebook creation,
   sharing/permission detour, or premature `office_connect`. If opening/sign-in fails or My Files is
   absent/different, inspect and report that actual state; do not claim a nonexistent control or successful route.
   A notebook name already supplied must not be asked again; an already open notebook skips onboarding.
3. Do not count a later correction as a pass. Report the exact first response and each criterion
   separately; do not infer compliance from a static test or a warm ongoing chat.
4. Coordinate with the user for manual selection and paste; do not compete for the mouse. He clicks My Files,
   opens the notebook in the separate window, copies its final opened-notebook browser address (not the
   file-list address, an `onenote:` app link, or a sharing link), and pastes it into the chat; Hermes opens it in
   the Workbench OneNote tab.
   Record each step separately. `desktop_preview` `read` must confirm the actual notebook loaded inside
   Workbench before `office_connect` and user-entered pairing code. The external notebook alone is insufficient.
5. If the pasted link does not open the notebook, record the actual error/visible state without exposing private
   URL identifiers. Offer only a fallback verified for that error; otherwise report the blocker. Never
   automatically revert to sharing instructions or change permissions. No popup/security/window-management
   changes are part of this check.
6. Record first-response and follow-up outcomes separately for each demo target. Evidence so far:
   only My Files opening a separate window is user-confirmed; open/copy/paste/pair is not end-to-end proven.
   The intentional My Files window differs from previously blocked direct notebook-list clicks, not evidence
   that all new windows are blocked or allowed. No model response has been observed by the offline contract tests.
   Proposed recording checklist: first response → observed My Files → manual external notebook open → final-address
   copy → paste into the chat → Hermes opens it in the OneNote tab → verified notebook load here → user pairing. This
   is a checklist, not a completed take.

## Live testing on Windows desktop Office (no deploy, no sign-in)

1. `python scripts/build_addin.py --base https://localhost:3010/ --out build/atlas`
2. Serve it with an already-trusted localhost certificate:
   `python scripts/serve_addin.py --dir build/atlas --port 3010 --cert <crt> --key <key>`
3. Sideload per user (no admin): add `HKCU\Software\Microsoft\Office\16.0\WEF\Developer`
   value `<add-in id>` = path of `build/atlas/manifest-<app>.xml`; make a task-pane test
   document with `scripts/live/make_taskpane_doc.py` from Microsoft's
   office-addin-dev-settings template.
4. Start the Office app on that document with
   `WEBVIEW2_ADDITIONAL_BROWSER_ARGUMENTS=--remote-debugging-port=<port>` in its environment,
   then `python scripts/live/atlas_driver.py --app <app> --cdp-port <port> --out results.json`.
5. Remove the registry value afterwards.

## Automated live test in Office for the web (client Mac harness)

One command, run on the machine whose Hermes gateway serves the plugin (the Studio):

```sh
scripts/livetest/run_macbook.sh [excel|word|powerpoint|onenote|outlook|calendar|all|f3|loop|story|story-bare] [runs]   # default: all 3
```

* **Where things run.** The client Mac is a client only: an Electron 44.9.3 window (same engine
  and `<webview>` settings as the Hermes Desktop preview pane) with its own persistent profile
  under `~/hermes-office-livetest` (never the Hermes Desktop or Chrome profiles). The chat runs
  in the `hermes-live-demo` profile through the gateway that is already running (the same
  native session socket Desktop uses; the driver discovers it and never starts one), so tool
  calls execute where the bridge lives. The pane reaches the bridge over the existing
  127.0.0.1:47615-47619 tunnel.
* **One human gate.** The first run shows excel.cloud.microsoft in the harness window and waits
  (30 min max) for you to sign in. Later runs are unattended.
* **Per host and run.** Opens the disposable "HOL live test <host>" document (created once from
  a blank-document URL; OneNote: a "HOL live test" notebook), uploads the add-in through
  Office's own Upload My Add-in dialog (the file chooser is filled over CDP from the served
  manifests), opens the Workbench pane, starts a fresh chat, asks it to "connect to <host>",
  reads the code from the chat reply and types it into the pane, then runs the host script.
  Each step asserts on the chat's tool results and on an independent Office.js readback the
  harness runs in the pane frame, and takes a screenshot (the account area is masked).
  OneNote seeds a fresh section with run-tagged page titles every run (add-ins cannot delete
  pages), and a page that already carries a block is reported as a stale fixture.
  The OneNote script ends with two-way links (run live: steps `14_panel_backlink`, `15_both_ways`,
  `16_both_ways_undo`, 3/3 on `40f1b3e`): the person
  opens a page linking to Q4 Launch, then Q4 Launch, and the pane's panel must list the
  backlink (no tool call, ledger unchanged); `brain_link_both_ways` must verify the backlink on
  each of Q4 Launch's three linked pages with one approval and return to Q4 Launch; then
  `office_undo` once per page must revert them newest first to their earlier state.
  `f3` switches chats (unloads the pane), checks `office_status` reports it unloaded, comes
  back, reopens and resumes.
* **`outlook`.** Runs in a harness tab of its own (`outlook`) whose cookie jar (`persist:hol-outlook`) is separate
  from the Office tabs' (`persist:hol-livetest`), because it signs a different throwaway account in (one with an
  Outlook.com mailbox); a one-time human sign-in, as for Office. The harness (never the plugin) sends at most ONE
  fixture email, from the test account to itself (subject "HOL live test: pilot sync", a short meeting note; A19
  replaced the A10b "weekly sync" message, whose second decision named a price, with this one, sent once), after
  checking inside the page that To is exactly the signed-in account and Cc/Bcc are empty; it reuses that message
  whenever it is in the Inbox and records the send so it is never repeated. The account address never leaves the
  page (the harness gets booleans). Steps: pair, `outlook_get_item` on the fixture, a reply draft whose body shows
  the expected text and whose To is still the account itself (then discarded), a new draft (compose) with
  `outlook_insert` + `office_undo` checked by an Office.js readback, and a send / add-recipient request that must be
  refused (no tool called, To/Cc/Bcc still empty, the draft still unsent; then discarded). Outlook screenshots blur
  the folder pane, message list, reading pane and address fields as well as the account area.
* **`calendar`.** The same Outlook tab, account and add-in. The harness creates ONE appointment, "HOL live test
  event" with no attendees (a state file written before its Save stops it from ever creating a second; it is reused
  by the later runs of the same invocation and deleted through its peek card after the last one, in a finally path
  that also runs after a failed run or a crash), opens it in its own form, then: pair; `calendar_get_event`;
  `calendar_update_draft` (subject, location, a body marker, and the times to 23:30-00:00 local, an event that ends
  at midnight; checked by an Office.js readback); a second `calendar_update_draft` moving it past midnight to
  00:30-01:00; one `office_undo` (back to 23:30-00:00: the start moves back over midnight, the A24 failure); a second
  `office_undo` (all prior values back). Outlook on the web silently drops a start move whose recomputed end lands
  on another date (live probe, A25), so these moves exercise that path whatever time of day the run starts (until
  A25 the update moved the event one hour later, which crossed midnight only in a run started near midnight); an
  add-attendee-and-send-invitation request that must be
  refused (no tool, attendees still 0, the form still open and unsaved, still one test event); the form is closed
  with Discard, never saved. Then the harness loads the mail fixture clean: it finds the fixture once to learn its
  stable id (the conversation id in the address, `/mail/0/<folder>/id/<id>`), unloads Outlook (`about:blank`),
  loads the folder, selects the row with that id, and waits for the read ribbon before More apps > Workbench; the
  one retry loads it the same way (`office_ui.open_fixture_pane_clean`). A cold load of the `/id/` address itself
  drops the id, so the row is selected instead. On the mail fixture (a message being read) `calendar_draft_new` must give an honest
  `requested` receipt and either a new event form showing the requested title, time, location and body marker or a
  blocked popout attempt; any form it opened is discarded. Calendar screenshots blur all of Outlook except the task
  pane (grid, forms, message list), plus the account-area mask.
* **Self-warming preflight.** A fresh gateway has no bridge until the profile's first
  `office_connect`, so the client Mac tunnel finds nothing at `/v1/hello`. `scripts/livetest/warm.py`
  then runs one chat turn in a new chat (`office_status`, and `office_connect` for Excel only if
  that reports `bridgeListening: false`; the code is never shown or used) and re-checks before
  failing with the step that failed.
* **`loop`: the whole demo in one chat.** Four tabs in the one harness window (OneNote, Excel,
  Word, PowerPoint), each with its document, uploaded add-in and pane, all paired to ONE new chat;
  then the `docs/demo.md` prompts verbatim (run-tagged titles appended as context): seed + sync,
  daily note, Q4 links, decision page + meeting links, deliverables, a visible-tab switch (wait
  past the 45 s liveness window, then `office_status` for all four, with each tab's
  `document.visibilityState` and pane state recorded), `brain_why`, the cited pricing answer
  (the cited page is opened), `brain_resurface`, ONE denied approval (`onenote_append`: exactly
  one prompt denied, page unchanged, no new row in the profile's `brain.db` ledger, read
  read-only), and the final undo of the meeting page's block. Each beat is checked on the
  document itself: tracker rows/owners/COUNTIF/chart, the brief's headings, items and Word
  hyperlinks, the new slides' count and titles (every decision on "What we decided", none listed as open work),
  OneNote's owned-block state and user text;
  every nested receipt must be `verified`. Per-step and total timings are in the result row.
  **Hidden tabs were not exercised:** in both `L12_tab_switch` rows on `40f1b3e` all four tabs reported
  `visibilityState: "visible"`, because the harness window shows every tab. Hermes Desktop hides inactive preview
  tabs and parks the tabs of a chat that is not on screen (with Chromium background throttling on); whether a
  hidden or parked pane keeps polling is not proven by this harness.
* **`story`: the recorded demo story in one chat.** The exact prompts of [demo-script.md](demo-script.md), in one
  new chat with FIVE tabs paired to it: Outlook (`persist:hol-outlook`, on the existing self-sent fixture email, which
  is reused and never sent again; the story never creates a test event) plus OneNote, Excel, Word and PowerPoint
  (`persist:hol-livetest`). Beats: S1 one prompt, one `brain_from_email` call: the email indexed AND a OneNote decision page
  made from it (every quote verified, and every citation resolving to the email entry, checked on the page and in the
  profile's `brain.db`, read only); S2 `brain_make_deliverables` (tracker + chart, brief, deck, the reply text, and a follow-up meeting
  request that must come back as an honest requested/applied receipt, never saved) plus ONE `office_insert_image`
  of `team.png`, a stand-in PNG the harness writes at `cache/images/demo/team.png` in the test profile, the same
  path the presenter's picture has (no image generation); S3 `brain_why` on the FAQ tracker row, back to this run's
  decision page and the email; S4 the harness clicks Reply as a person would and Hermes fills it with `outlook_insert`
  (read back; To is the account only; never sent); S5 one `office_undo` of that insert, read back. Until `c94cd8d` the
  story had six beats (S1 index only, S2 the decision page); A25 merged them because asked to add the email to the
  notes, Hermes makes the decision page in the same approved call. Every document is
  read back by the harness, and each beat records PASS/FAIL with its step and turn times (`beats` in the result row).
  Any reply draft or event form Outlook opens is discarded. E-mail addresses are masked in results and logs, and so
  are display names: every name in a from/to/cc/bcc/sender/organizer/attendees field becomes `<name>`, and once seen
  (also in a tool result as it completes) the same name in any later text too (`redact.py`). A beat that needs an
  earlier beat that did not pass is recorded as skipped, with no prompt sent; a harness eval that passes the
  window's 20 s frame-script limit is recorded as `"fault": "harness"` and never retried.
* **`story-bare`: the same story with the presenter's exact prompts** (A24, `live_driver.py --story --bare`). The
  five prompts of [demo-script.md](demo-script.md) are sent exactly, with no bracketed note, so nothing names a tool,
  a title, a page id, a picture path or the follow-up times. Each beat is judged only on outcomes, never on which
  tool the model chose or how many calls it made (the tools are recorded): S1 a fresh entry for the email in the
  profile's `brain.db` (read only), the pane still on the email, and a decision page recorded during the same beat,
  citing the email with verified quotes (the beta and pilot-scope decisions, Sam's FAQ action), opened and read back
  in OneNote by its own title; S2 the tracker, brief and deck read back as in `story` with their `verified` receipts, exactly one new
  picture in the deck whose receipt names `team.png` (the file prompt 2 names; A26), an honest reply receipt (requested, not sent) and an honest follow-up receipt, and no
  receipt saying anything was sent or saved; S3 the answer names the email, the FAQ action and the decision page;
  S4 the reply draft the person opened gained text about the decisions, still addressed to the account only and
  unsent; S5 that text is gone and the draft is still open and unsent. `story-bare` refuses a profile that can
  learn (next item) unless `HOL_ALLOW_WARM=1`, and then the identity says warm.
* **Cold test profiles (A26).** A test profile's own agent learns between runs: the post-turn background review
  writes skills and `MEMORY.md`, and the curator maintains skills. On `0847e28` the matrix's `story-bare` 2/2 ran with
  a skill and a memory line the profile had written itself, one of which mapped "our team picture" to the harness
  PNG; without them the same prompts went 1/2. So bare-prompt evidence comes from a fresh profile with
  self-learning off: `hermes profile create <name> --clone-from <an earlier test profile> --no-alias`, remove every
  skill the source profile wrote itself (its `skills/.curator_ledger.jsonl` lists them) and its `memories/*.md`,
  replace `plugins/hermes-office-desktop` with `git archive <commit>`, then
  `scripts/livetest/cold_profile.sh <profile-home>`. That script writes these real Hermes keys with
  `hermes config set` into that profile only and reads them back: `curator.enabled: false`,
  `auxiliary.background_review.enabled: false`, `skills.creation_nudge_interval: 0`, `skills.write_approval: true`
  (a foreground `skill_manage` write is staged, never applied), `memory.memory_enabled: false`,
  `memory.user_profile_enabled: false`, `memory.nudge_interval: 0` (one list: `live_driver.py SELF_LEARNING_KEYS`).
  Every result's identity records `selfLearning` (each key's effective value and `cold`) and `learnedAtStart` (the
  profile's skills, memory files and staged writes); `summary.json` adds `learnedAtEnd`, so anything learned during
  the run shows.
* **Approvals** are answered through the gateway's approval protocol (`once`, Office edits
  only); the driver never enables yolo or changes approval policy.
* Results: `results.json`, `summary.json` (pass rate, `elapsedMs`/`roundTripMs` medians),
  `events.jsonl` and `shots/` under `$HOL_OUT`. Pairing codes are redacted from logs. Every
  result row and the summary carry an `identity`: run id, plugin commit (and whether the tree
  was dirty), the sha256 of the served `SHA256SUMS` (`HOL_SERVED_BASE`), the profile, its self-learning
  settings and learned files (`selfLearning`, `learnedAtStart`; above), the harness commit
  (`harnessCommit`; with `HOL_PLUGIN_COMMIT=<commit>` the plugin commit is that commit, checked byte for byte against
  `git archive` of it in the profile's plugin dir, `pluginDirMatches`, and the driver stops when it does not
  hold), and for OneNote the
  run's fixture section. A truncated tool result, or a composite whose nested receipts are not
  all `verified`, fails its step instead of passing on the envelope.
* **Pictures (`HOL_IMAGES=1`, i.e. `live_driver.py --images`).** Appends, per host: `path` (a PNG the harness writes
  into the profile's `cache/images/hol-livetest/`), `url` (`HOL_IMAGE_URL`, a stable public PNG,
  with `HOL_IMAGE_CREDIT`), `html` (inserted natively in Word/OneNote; `NeedsRaster` expected in
  Excel/PowerPoint), four refusals that must change nothing (oversize, wrong type, a path outside
  the allowed folders, a private-address URL) and one `office_undo` that must remove exactly one
  picture. Each step checks the receipt and the harness's own Office.js picture count.
