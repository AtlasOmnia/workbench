<!-- SPDX-License-Identifier: Apache-2.0 -->
# Live evidence

What was run live, on which commit, with which result. The README keeps only the one-line summary and the
medians; this page keeps the per-run detail. How the harness works: [harness.md](harness.md).

## 0.4.0: the add-in served by the plugin (not live-tested)

0.4.0 stops defaulting to a GitHub Pages site: the bridge serves the pinned `addin/` files itself over HTTPS on
`https://localhost:<port>/office/live/` with a per-profile local CA. Tested in unit and integration tests only
(`tests/test_addin_served.py`, plus a throwaway bridge process that served every `SHA256SUMS`-listed file over HTTPS,
verified against its generated CA, and refused the rest).

**One live run, Excel on macOS, 2026-11-12, commit `fe3d97a`** (real Hermes Desktop, not the harness; the plugin
copied from a `git archive` of that commit into an existing test profile, `addin_base` unset). A person ran the
printed trust command, uploaded the generated `manifest-excel.xml` and typed the pairing code; Excel for the web
(OfficeOnline 16.0.20709.42309) loaded the pane from the bridge on `https://localhost`. `office_status`: paired and
live. Workbook overview 398 ms and a range read 310 ms round trip; a one-cell write to A1 (undo available);
`office_undo`; a read back showed A1 blank. That is one run of one app. **Not live-tested with the plugin-served
add-in:** Word, PowerPoint, OneNote, Outlook, Windows, and a clean-profile `hermes plugins install` (the smoke test
below). The runs below used earlier builds and a private add-in site.

## Current: F8 `402f00e` (A45/A46)

Shipped runtime: `402f00e` (F8). A45 prepared the cold `office-live-c402f00e` profile and
verified the 33-file private add-in build, but **no live run started**: the client Mac was offline.
A46 reran with harness `a7e68d2` (harness/test-only OneNote empty-section setup fix; runtime unchanged).
The add-in served build-info `402f00e1d9e1`, clean, with served SHA256SUMS matching disk.

| A46 run directory | Result |
| --- | --- |
| `hol-a46-a7e68d2-all-1` (`HOL_IMAGES=1`) | Excel, Word, PowerPoint, OneNote **1/1 each**; image path/URL/HTML, refusals and undo passed on all four |
| `hol-a46-a7e68d2-story-bare-2` (cold) | **2/2 PASS**, first attempt |
| `hol-a46-a7e68d2-loop-1` | **1/1 PASS**, first attempt |

The loop exercised the harness's new Add page path in an empty OneNote section; one PowerPoint
upload retry in the bare story recovered. No gateway restart, mail send or event save was performed.
Standalone Outlook/mail and calendar results and timing rows remain **F5 `8fa1f63`**;
the F8 story includes its reply write/undo, not a new standalone calendar matrix.
F7 below is historical evidence, not the shipped runtime. Source: A45-report.md and A46-report.md
(private overnight evidence reports); no F8 timing medians were supplied there.

Earlier live runs used a private add-in site. The plugin-served add-in (the default since 0.4.0) passed one live Excel run on macOS (2026-11-12); other apps and Windows are not live-tested with it yet.

### Post-publish smoke test (owner; not run yet)

1. Publish `https://github.com/AtlasOmnia/workbench` (no Pages deployment: the plugin serves its own add-in).
   Fill the catalog's publish-time `sha` with the intended pinned 40-hex commit; submit only after this smoke passes.
2. Create a clean Hermes profile, install the published plugin, and leave `addin_base` and `allowed_origin` unset.
3. The person trusts that profile's local certificate authority with the command `/workbench trust` prints (the
   plugin never does it).
4. In Hermes Desktop open a disposable Excel workbook, connect Excel, upload its generated manifest (every URL
   `https://localhost:<port>/office/live/...`), and have the person type the pairing code into Workbench.
5. Approve one cell write, independently read it back, then call `office_undo` once and confirm the prior cells return.
6. Record the commit, the bridge port, receipts and outcome here. A private-origin pass does not satisfy this check.

## F7 `d1f7aae` (A31)

Plugin `d1f7aae` on a cold disposable profile `office-live-cd1f7aae` (cloned from the F6 cold profile with its
learning history moved aside; plugin dir byte-identical to `git archive d1f7aae`; self-learning off with
`scripts/livetest/cold_profile.sh`). Every result row: `pluginCommit` and `harnessCommit` `d1f7aae`,
`pluginDirMatches` true (133 files), `pluginDirty` false, `selfLearning.cold` true, served `SHA256SUMS` sha256
`739d699c…6b34fb2` (33 files; `addin/` is unchanged since F5, so the served add-in is the F5 build). Nothing was
learned during the matrix (`learnedAtEnd`: only the bundled skill, no memory file, no staged write). The shared
OneNote test notebook was not cleaned first. Every attempt listed, no reruns used (run directories
`hol-a31-d1f7aae-<mode>`):

| Mode | Result | What happened |
| --- | --- | --- |
| `story-bare 2` (cold) | **2/2 PASS** | the presenter's exact prompts, judged on outcomes only; run 1 made "Pilot sync, 6 November: decisions", run 2 met an earlier run's title and got the next free one, "Pilot sync decisions (6 November) (3)", and Hermes said the earlier page was left as it was; `team.png` passed by file name and found by the plugin both times |
| `HOL_IMAGES=1 all 1` | **5/5 PASS** | Excel, Word, PowerPoint, OneNote 1/1 each with pictures on, plus F3 |
| `loop 1` | **1/1 PASS** | the whole notes loop in one chat with four Office tabs (L00-L17 all passed) |

Not run on F7, because its diff changes no code they use: `calendar` and `outlook` (1/1 each on F5, below).

Cold bare-prompt result on `d1f7aae`: **2 of 2 runs**, first attempt (on `5df2609`: 2 of 2; on `8fa1f63`: 2 of 2; on
`0847e28`: 1 of 4).

Chat turns per beat (`story-bare` run 1 / run 2): S1 13.1 s / 13.1 s, S2 62.8 s / 59.3 s, S3 5.7 s / 7.0 s, S4 7.8 s /
8.2 s, S5 6.1 s / 6.3 s.

What changed from `5df2609` and why (from the A30 review; tests RED on `a7d387a`, GREEN on `d1f7aae`): on F6 the first
blocked retry of an uncertain OneNote change marked it "checked", so the next identical call wrote again without the
page ever being read. Now only a successful read releases it: `onenote_get_page` of the page the change targeted (the
open page is recorded before a change that names none), or, for a new page, a complete `onenote_list_pages` with
`all`. The ledger records what that read showed (`reconciled`, `reconciled_at`, `reconciled_by`). If the change is
there, the identical call is refused (`AlreadyApplied`, nothing sent); if it is not, it runs once. A blocked retry, a
failed read, a pane that still does not answer, or the pane's journal saying "failed" never releases it. The same
holds for `brain_link_both_ways` (its own read of each page is the check), the decision page, a picture and undo (an
undo after a lost delete reads the page first and sends nothing if Hermes' block is gone). Every OneNote change that
names no page now starts with one `get_context` read. This path is covered by tests; no `PaneReload`,
`ExecutionUncertain` or `AlreadyApplied` came up in the F7 runs.

Limit (A32, pinned by a test in A33): if that preliminary `get_context` read itself fails, the change is still sent
but with no page recorded, and no later read can settle it, because the ledger row names no page to match. If the
change's result is then lost, it stays held indefinitely: every identical call is refused (`ExecutionUncertain`) and
the identical change is not re-sent, even after the user reloads the tab and Hermes reads the page successfully. What
the user sees is that the change stays held; they reload the OneNote tab and check the page themselves (and add the
content by hand, or ask for a differently worded change, if it is missing). The outcome is a held write, never a
duplicate.

## `ff9771e` (A35): rerun after the skill change, incomplete

`ff9771e` changes only docs, the skill (`skills/office-desktop/SKILL.md`) and tests; the runtime is still `d1f7aae`.
Because the skill text steers the model, the cold matrix was run again on a fresh cold profile `office-live-cff9771e`
(plugin dir byte-identical to `git archive ff9771e`; every result row: `pluginCommit` and `harnessCommit` `ff9771e`,
`pluginDirMatches` true, `pluginDirty` false, `selfLearning.cold` true; served `SHA256SUMS` unchanged, `739d699c…`;
nothing learned during the runs). Every attempt (run directories `hol-a35-ff9771e-<mode>`):

| Mode | Result | What happened |
| --- | --- | --- |
| `story-bare 2` (cold) | **not run: 0/2, harness gate** | both runs stopped while preparing the Outlook tab: its sign-in had expired and the harness waited 30 minutes for a person to sign in (no one was there). No prompt was sent, so no beat ran. Not rerun |
| `HOL_IMAGES=1 all 1` | **5/5 PASS** | Excel, Word, PowerPoint, OneNote 1/1 each with pictures on, plus F3; first attempt |
| `loop 1` | 0/1, then **1/1 PASS** on the rerun | first attempt: every step passed except `L11_deliverables`, where the Word pane timed out during `brain_make_deliverables`, so its receipt was `uncertain` (Excel `verified`; Hermes reported PowerPoint `verified`); Hermes read the document instead of retrying and reported the brief there once, complete; the harness requires every receipt `verified`. The rerun passed L00-L17 |

So the cold bare-prompt result on `ff9771e` is **not established**; the current cold bare-prompt evidence stays
`d1f7aae` (above). Medians from the `ff9771e` `all 1` run (in pane / round trip): Excel 165/245 ms, Word 1496/1518 ms,
PowerPoint 512/593 ms, OneNote 32/119 ms.

## Real Hermes Desktop session, Studio, F7

The first end-to-end session in the Hermes Desktop app itself (not the harness), on the Studio Mac, plugin `d1f7aae`
(F7), one chat with all five panes paired (Excel, Word, PowerPoint, OneNote, Outlook mail). A person typed the bare
prompts from [demo-script.md](demo-script.md). **Warm, not cold evidence:** the profile had self-learning on and
learned during the session; the cold evidence stays the F7 harness runs above.

| Prompt | Result | What needed more |
| --- | --- | --- |
| 1 email to decision page | **PASS** | none; the fixture email was in Archive, and Hermes searched All folders and opened the pane on it |
| 2 tracker, brief, deck, reply, follow-up | **PASS after two follow-up turns** | first turn: Word `verified`; Excel `failed` with nothing changed (a cell was in editing mode); PowerPoint `uncertain` (its pane stopped answering mid-call), not retried; reply draft and follow-up `applied`. Second turn: the deck (4 slides) and `team.png` `verified`; Excel, slow (57 s round trip), read back first: nothing written, no blind retry. Third turn (Excel only): tracker, formula, owner summary and chart `verified` |
| 3 why is the FAQ row there | **PASS** | none; quoted the email, decision page linked |
| 4 fill the reply | **PASS on the second turn** | first turn Hermes refused, correctly: the pane was still on the read email. After the pane was opened from the draft's Message > Apps, the reply text was inserted and read back |
| 5 undo | **PASS** | none; the draft was back to the quoted original, read back |

Bare story: **5 of 5 prompts passed in one chat.** Nothing was sent or saved (no mail, no event).

Outside the plugin: Hermes (the gateway) was stopped mid-run by an external signal and restarted by the system,
not by the test. Pairings are kept in memory by design, so every pane had to be paired again; Hermes issued the codes
and refused to type them itself (the person types them). The PowerPoint pane stayed on "Hermes is working…" after the
restart and had to be closed and opened again.

What this changed in the guidance (README, [demo-script.md](demo-script.md), the skill): Outlook has no pin, so open
the email first and open Omnia Works inside the reply draft (Message > Apps); clicking a notebook in OneNote's list
opens a blocked new window (same-tab route by link: not yet verified live); Hermes cannot bring a tab to the front, so
the person clicks the app's tab; click the chat box before typing; Excel refuses changes in cell-editing mode; the
Add-ins button needs a few seconds after the file opens; a code lasts 10 minutes. Cosmetic: `team.png` at its
default position overlapped the title slide's subtitle, and the reply's bullet list became inline dashes.

## F6 `5df2609` (A29)

Plugin `5df2609` on a cold disposable profile `office-live-c5df2609` (cloned from the F5 cold profile with its
learning history moved aside; plugin dir byte-identical to `git archive 5df2609`; self-learning off with
`scripts/livetest/cold_profile.sh`). Every result row: `pluginCommit` and `harnessCommit` `5df2609`,
`pluginDirMatches` true (133 files), `pluginDirty` false, `selfLearning.cold` true, served `SHA256SUMS` sha256
`739d699c…6b34fb2` (33 files; `addin/` is unchanged since F5, so the served add-in is the F5 build). Nothing was
learned during the matrix (`learnedAtEnd`: only the bundled skill, no memory file, no staged write). The shared
OneNote test notebook was not cleaned first. Every attempt listed, no reruns used (run directories
`hol-a29-5df2609-<mode>`):

| Mode | Result | What happened |
| --- | --- | --- |
| `story-bare 2` (cold) | **2/2 PASS** | the presenter's exact prompts, judged on outcomes only; both runs met an earlier run's decision page title and got the next free one (run 1 "Pilot sync decisions (6 November) (2)", run 2 "Pilot sync, 6 November (3)"), and Hermes said the earlier page was left as it was; `team.png` passed by file name and found by the plugin at `cache/images/demo/team.png` |
| `loop 1` | **1/1 PASS** | the whole notes loop in one chat with four Office tabs: sync, daily note, links, decision page, deliverables, tab switch, why, cited ask, resurface, a denied approval, undo |
| `HOL_IMAGES=1 all 1` | **5/5 PASS** | Excel, Word, PowerPoint, OneNote 1/1 each with pictures on, plus F3 |

Not run on F6, because its diff changes no code they use: `calendar` and `outlook` (1/1 each on F5, below).

Cold bare-prompt result on `5df2609`: **2 of 2 runs**, first attempt (on `8fa1f63`: 2 of 2; on `0847e28`: 1 of 4).

Chat turns per beat (`story-bare` run 1 / run 2): S1 10.6 s / 12.0 s, S2 52.5 s / 50.9 s, S3 5.4 s / 9.7 s, S4 8.5 s /
7.5 s, S5 6.8 s / 6.0 s.

What changed from `8fa1f63` and why (from the A28 review; tests RED on `8fa1f63`, GREEN on `5df2609`): when a OneNote
change was sent and the pane then answered `PaneReload` (its message: "this change may or may not have been made"),
the plugin recorded the change as failed, `brain_link_both_ways` reported "Nothing was changed on this page", and an
identical later call could run a second time. Now that answer keeps the change `uncertain` in the ledger; the receipt
says to reload the OneNote tab and read the page; `brain_link_both_ways` stops there; the decision page, a picture and
undo report it as uncertain; and an identical call is held back until the page has been checked, also when the
reloaded pane's own journal lists the change as failed, and without using up that check while the pane is still not
answering. A `PaneReload` or `PaneBusy` on a read still means nothing changed. This path is covered by tests with the
pane's exact answer; it was not met live on `5df2609`. Correction (A30): on `5df2609` the hold was released by a second blocked
retry without any read of the page; F7 above fixes that.

## F5 `8fa1f63` (A27)

Plugin `8fa1f63` on a cold disposable profile `office-live-c8fa1f63` (cloned from the A26 cold profile with its
learning history moved aside; plugin dir byte-identical to `git archive 8fa1f63`; self-learning off with
`scripts/livetest/cold_profile.sh`). Every result row: `pluginCommit` and `harnessCommit` `8fa1f63`,
`pluginDirMatches` true (133 files), `pluginDirty` false, `selfLearning.cold` true, served `SHA256SUMS` sha256
`739d699c…6b34fb2` (33 files over https). Nothing was learned during the matrix (`learnedAtEnd` in every row: only the
bundled skill, no memory file, no staged write). The shared OneNote test notebook was not cleaned first. One matrix,
every attempt listed, no reruns (run directories `hol-a27-8fa1f63-<mode>`):

| Mode | Result | What happened |
| --- | --- | --- |
| `story-bare 2` (cold) | **2/2 PASS** | the presenter's exact prompts, judged on outcomes only; both runs met an earlier run's decision page title and got the next free one (run 1 "Pilot sync decisions (6 Nov) (2)", run 2 "Pilot sync, 6 November (2)"); Hermes told the user the earlier page was left unchanged; `team.png` passed by file name, found by the plugin at `cache/images/demo/team.png` (receipt `source.foundByName`, `source.path`), verified on the title slide |
| `story 1` | **1/1 PASS** | all five beats, routing note on each prompt |
| `HOL_IMAGES=1 all 1` | **5/5 PASS** | Excel, Word, PowerPoint, OneNote 1/1 each with pictures on, plus F3 |
| `calendar 1` | **1/1 PASS** | get, update, move past midnight, undo back over midnight, undo all prior values, refused attendee, new-event request; the one test event deleted at the end |
| `outlook 1` | **1/1 PASS** | pair, read, reply draft, the reply fallback, compose insert + undo, a refused send |
| `f3 1` | **PASS** | |

Cold bare-prompt result on `8fa1f63`: **2 of 2 runs**, first attempt (on `0847e28`: 1 of 4, below).

Chat turns per beat (`story-bare` run 1 / run 2; `story` run 1):

| Beat | `story-bare` (cold) | `story` |
| --- | --- | --- |
| S1 the email indexed and its decision page | 11.7 s / 13.2 s | 12.2 s |
| S2 tracker + chart, brief, deck, one picture, reply text, follow-up request | 52.8 s / 55.2 s | 43.8 s |
| S3 why is the FAQ row there | 6.1 s / 7.9 s | 7.2 s |
| S4 person clicks Reply, Hermes fills it | 5.9 s / 7.7 s | 6.2 s |
| S5 one undo of that insert | 3.3 s / 6.0 s | 3.3 s |

What changed from `0847e28` and why (each from a failure of the cold bare runs below; tests RED on `d4caa11`, GREEN
on `8fa1f63`):

* **A picture named by file name.** `office_insert_image` `path` may be a bare file name (`team.png`). The plugin
  looks for exactly that name, case-insensitively, in the allowed image folders and their subfolders only (bounded;
  symlinked files and folders are never followed), uses the one match and names it in the receipt; several matches
  are refused with the candidates relative to their folder; none names the folders searched. On `0847e28` one cold
  run answered "I couldn't find the file" without looking.
* **A decision page title an earlier run already used.** `brain_decision_page` and `brain_from_email` give the new
  page the next free title in that section, `<title> (2)` up to `(20)`, never open or change the existing page, and
  say so in the receipt (`titleAdjusted`). On `0847e28` both cold first runs stopped at `Conflict: DuplicateTitle`.
* **A OneNote pane that stopped answering.** From the A26 runs' events and tool results: (1) `onenote_list_pages`
  with `all` walked every section of the shared test notebook (one section per earlier run) for more than Hermes'
  60 s, was reported as "the edit may or may not have applied. Check the workbook", and kept the pane from picking up
  the next command; (2) after an `onenote_open_page` to a page in another section, every OneNote request from the pane
  failed with `GeneralException` while the pane kept polling (the add-in's OneNote context was unusable until the pane
  reloads). Now `list_pages` answers within 20 s (`complete: false`, `stoppedEarly`), a OneNote read that outlasts the
  wait is `PaneBusy` (nothing changed), `open_page` keeps checking for a navigation that lands late, and when even the
  smallest read fails the pane answers `PaneReload` ("reload the OneNote browser tab") instead of `GeneralException`.
  The `PaneReload` path is covered by tests (`tests/js/onenote.test.mjs`); it was not met live on `8fa1f63`.

Known limits are unchanged ([below](#known-limits-seen-on-40f1b3e)).

## Bare prompts, cold: `0847e28` (A26)

This is the bare-prompt evidence for `0847e28`. The `story-bare 2/2` in the matrix below ran **warm**: by then the
test profile's own agent had written itself a skill and a `MEMORY.md` line (one mapped "our team picture" to the
harness PNG). Run without them (A25, `hol-a25-0847e28-story-bare-2-cold`), the same prompts went 1/2: in run 1 the
deck got no picture, because "our team picture" names no file.

Since A26, prompt 2 names the file ("put team.png from my Hermes images folder on the title slide"; the presenter's
picture is `cache/images/demo/team.png` in the profile, and the harness writes its stand-in at the same path), and
the bare story runs on a cold profile: a fresh `office-live-c0847e28` (cloned, every skill the earlier profile wrote
itself and its learning history moved aside; plugin dir byte-identical to `git archive 0847e28`) with self-learning
off (`scripts/livetest/cold_profile.sh`; [harness.md](harness.md), "Cold test profiles"). Every result row:
`pluginCommit` `0847e28` (`pluginDirMatches` true, 132 files), `harnessCommit` `df3f7f3`, `selfLearning.cold` true,
served `2d05f2cd…960eedf`. Nothing was learned during either attempt (`learnedAtEnd`: only the bundled skill, no
memory file, no staged write). Two attempts, the second the one allowed rerun:

| Attempt | Run directory | Result | What happened |
| --- | --- | --- | --- |
| 1 | `hol-a26-0847e28-story-bare-2-cold` | **1/2** | run 1: S1 failed (below), so S2-S5 were skipped. Run 2: **PASS**, all five beats; `team.png` inserted on the title slide (receipt source `team.png`), chat turns S1 18.4 s, S2 77.3 s, S3 6.1 s, S4 7.9 s, S5 5.3 s |
| 2 | `hol-a26-0847e28-story-bare-2-cold-2` | **0/2** | run 1: S1 failed the same way. Run 2: S1 passed (11.8 s); S2 made and verified the tracker, brief and deck and requested the reply and follow-up honestly, but inserted no picture: the model answered that it "couldn't find the file" without calling any tool to look for it |

* **S1 in both first runs:** the model titled its decision page as earlier test runs had ("Pilot sync, 6 November",
  "Pilot sync decisions (6 Nov)"), the plugin refused the duplicate (`Conflict: DuplicateTitle`, the earlier page in
  section `HOL r3 1107-080246` of the shared test notebook), and the model's follow-up OneNote calls then failed (a
  60 s `onenote_list_pages` timeout and a command the pane never picked up in attempt 1; "OneNote did not open the
  page" and then `GeneralException` on every OneNote call in attempt 2). The model said each time that no page was
  made. Not a picture problem, and not judged further here.
* **The picture:** of the two cold runs that reached S2, one inserted `team.png` and one did not look for it. The
  skill now tells Hermes to look for a picture named by file name in its image folders first and never to invent or
  generate one ([SKILL.md](../skills/office-desktop/SKILL.md), Pictures); that text is not in `0847e28`'s plugin
  dir, so these runs did not have it. Cold bare-prompt result on `0847e28`: **1 of 4 runs**.

## F4 `0847e28` (A25)

Plugin `0847e28` (profile `office-live-n0847e28` == `git archive 0847e28`; served `SHA256SUMS` sha256
`2d05f2cd…960eedf`, 33 files over https; every result row's `identity.pluginCommit` is `0847e28`, `pluginDirty`
false). One matrix, every attempt, no reruns (`run_macbook.sh <mode> <n>`, run directories `hol-a25-0847e28-<mode>`):

| Mode | Result | What happened |
| --- | --- | --- |
| `calendar 3` | **3/3 PASS** | every step in every run: get, update to 23:30-00:00, move to 00:30-01:00 the next day, undo (back over midnight), undo (all prior values), refused attendee, new-event request; the one test event deleted at the end |
| `story 2` | **2/2 PASS** | all five beats, routing note on each prompt |
| `story-bare 2` | **2/2 PASS (warm)** | all five beats, the presenter's exact prompts then, judged on outcomes only; the profile had learned a skill and a memory line by then, so this is not the bare-prompt evidence (see [cold](#bare-prompts-cold-0847e28-a26) above) |
| `outlook 1` | **1/1 PASS** | pair, `outlook_get_item`, `outlook_draft_reply`, the reply fallback, compose insert + undo, a refused send |
| `HOL_IMAGES=1 all 1` | **5/5 PASS** | Excel, Word, PowerPoint, OneNote 1/1 each with pictures on, plus F3 |
| `f3 1` | **PASS** | |

What changed from `c94cd8d` and why:

* **Calendar undo over midnight (c94cd8d: calendar 0/3).** Outlook on the web acknowledges a `start.setAsync` but
  keeps the old start when the move would make it recompute an end on another date (seen live in a probe of 17
  scenarios, replayed in `tests/js/fixtures/owa-appointment-times-a25.json`). The c94cd8d runs started at 23:10, so
  the fixture was 23:30-00:00; the undo's start was dropped and the end then failed with `InvalidEndTime`, after the
  subject had already been restored (a partial undo that said "nothing was changed"). The pane now reads the times
  back after each step and gives the end room before a start Outlook would drop; a step that still does not apply
  is an error naming what the form shows, and an undo or rollback that cannot restore every field says which ones
  it did and did not restore (category `Partial`). Earlier candidates passed because their fixtures (13:30, 21:00,
  04:00, 09:30, 01:00 local) never crossed midnight; the calendar run now crosses it on purpose.
* **One prompt for the email and its decision page (c94cd8d: story-bare 0/2).** Asked to add the email to the notes,
  Hermes made the decision page in the same approved call, so the old second prompt had nothing to do and the bare
  beat failed. The script now asks once ("Read this email and turn it into a decision page in my notes, quoting the
  email for each decision and action.") and the story has five beats.

Chat turns per beat (run 1 / run 2):

| Beat | `story` | `story-bare` (warm) |
| --- | --- | --- |
| S1 the email indexed and its decision page | 12.7 s / 10.6 s | 19.0 s / 20.0 s |
| S2 tracker + chart, brief, deck, one picture, reply text, follow-up request | 45.4 s / 45.1 s | 86.1 s / 61.8 s |
| S3 why is the FAQ row there | 6.6 s / 7.5 s | 10.9 s / 6.6 s |
| S4 person clicks Reply, Hermes fills it | 6.3 s / 6.0 s | 6.8 s / 8.0 s |
| S5 one undo of that insert | 4.4 s / 2.7 s | 2.8 s / 2.8 s |

* In the warm `story-bare` S1 the model called `brain_from_email` twice in both runs (index, then decisions) and titled the
  page itself; each page had both decisions and Sam's FAQ action, every quote verified and cited to the email, read
  back on the page. Neither the call count nor the title is judged.
* Medians on `0847e28`, Hermes to pane and back: Excel 190 / 241 ms, Word 1224 / 1290 ms, PowerPoint 423 / 457 ms,
  OneNote 24 / 47 ms, Outlook (mail) 97 / 148.5 ms, Outlook calendar 170 / 324 ms (one run each, calendar three).

### Results on `c94cd8d` (F3) for modes the change does not touch

Profile `office-live-nc94cd8d`, served `739c878a…c08b9`; every row's `identity.pluginCommit` is `c94cd8d`. The
runtime change to `0847e28` is the Outlook calendar code plus one new error category's pane text and hint, so these
results still describe `0847e28` for the modes below; they were not rerun on it.

| Mode | Run directory | Result |
| --- | --- | --- |
| `HOL_IMAGES=1 all 3` | `hol-a24-c94cd8d-all-3` | **13/13 PASS** (Excel, Word, PowerPoint, OneNote 3/3 each, plus F3) |
| `loop 2` | `hol-a24-c94cd8d-loop-2` | **2/2 PASS** (the whole demo loop in one chat, four Office tabs) |
| `outlook 3` | `hol-a24-c94cd8d-outlook-3` | **3/3 PASS** |

The other `c94cd8d` modes were rerun on `0847e28` above: `calendar 3` 0/3 and `story-bare 2` 0/2 there (the two
findings), `story 2` 2/2 (the old six-beat script) and `f3 1` 1/1.

Known limits are unchanged from `40f1b3e` ([below](#known-limits-seen-on-40f1b3e)): a new calendar event is
requested, not opened in the pane; OneNote for the web can fail to load its editor; the model sometimes needs a
second ask; hidden tabs were not exercised.

## Release candidate `40f1b3e`

Plugin `40f1b3e` (profile `office-live-n40f1b3e` == `git archive 40f1b3e`; served `SHA256SUMS` sha256
`44d4a4f6…eabd6b7`, 33 files over https; every result row's `identity.pluginCommit` is `40f1b3e`, `pluginDirty`
false). Harness and docs commits after `40f1b3e` change no runtime file. Every attempt on this commit, in order
(run directories under the harness output root, `run_macbook.sh <mode> <n>`):

| Mode | Attempt (run directory) | Result | What happened |
| --- | --- | --- | --- |
| `HOL_IMAGES=1 all 3` | `hol-a19-40f1b3e-all-3` | **13/13 PASS** | Excel, Word, PowerPoint, OneNote 3/3 each, pictures on, plus F3 |
| `loop 2` | `hol-a19-40f1b3e-loop-2` | **2/2 PASS** | the whole demo loop in one chat, four Office tabs |
| `outlook 3` | `hol-a19-40f1b3e-outlook-3` | 2/3 | run 3 lost its chat at `insert` ("Gateway disconnected"): a planned gateway restart at 12:46, not a defect; `undo` and `refuse_send` then had no gateway |
| `story 2`, `calendar 3`, `f3 1` | `hol-a19-40f1b3e-story-2` (story only) | not run | same restart: the story got "gateway ... not ready (starting)" before its first turn; calendar and F3 stopped at the preflight. Nothing was tested |
| `outlook 3`, `story 2`, `calendar 3`, `f3 1` | rerun "gwrestart" (no directories) | not run | all four stopped at the preflight while the gateway was still starting |
| `outlook 3` | `hol-a19-40f1b3e-outlook-3-gw2` | **3/3 PASS** | pair, `outlook_get_item`, `outlook_draft_reply`, the reply fallback, compose insert + undo, a refused send |
| `story 2` | `hol-a19-40f1b3e-story-2-gw2` | 1/2 | run 1 FAIL at S2 (harness eval timeout, below; S3-S6 followed from it); run 2 **PASS**, all six beats |
| `calendar 3` | `hol-a19-40f1b3e-calendar-3-gw2` | **3/3 PASS** | every step in every run; the one test event deleted by the finally path |
| `f3 1` | `hol-a19-40f1b3e-f3-1-gw2` | **PASS** | |
| `story 2` | `hol-a19-40f1b3e-story-2-gw3` (one rerun) | 0/2 | both runs FAIL at S2 with the same harness eval timeout |
| `story 2` | `hol-a20-40f1b3e-h96d3510-story-2` (fixed harness `96d3510`) | 1/2 | run 1 FAIL at `P0_prepare_onenote`: OneNote for the web never showed its ribbon in 183 s, so no beat ran; run 2 **PASS**, all six beats. The rows record `identity.pluginCommit` `96d3510` (the checkout's HEAD); the profile's plugin was `40f1b3e` and `96d3510` changes no runtime file |

Story on `40f1b3e`, all attempts: 2 of 6 passed all six beats; 3 failed on the harness's page-open eval (fixed), 1
never started (OneNote editor). On the fixed harness: 1 of 2. Every story run above adds a bracketed note to each
prompt that names the tool and its arguments (see [demo-script.md](demo-script.md#prompts)); the bare prompts were
first run as `story-bare` on the next candidate.

* **Story S2, the harness timeout (3 of 4 story runs).** In each failed run the plugin's part of S2 succeeded:
  `outlook_get_item`, `onenote_list_sections` and `brain_from_email` were ok, the page was made in the run's own
  section, its quotes verified, and the reply said it was open. What failed was the harness's next step, the stand-in
  for the person clicking that page: an eval in the OneNote task pane that walked every section of the shared test
  notebook. It hit the harness window's fixed 20 s frame-script limit (`main.js` `evalIn`, message "frame script timed
  out (frame detached?)"; nothing showed a detached frame). The test notebook gains one section per run, and S2's time
  after the chat turn grew from 16-17 s (`5d840f0`) to 23-24 s (`77b5652`, gw2 run 2) and then past the limit. S3 then
  never sent its prompt (it needs S2's decision page id), and S4-S6 ran without the deliverables they ask about.
  Harness fix after `40f1b3e` (no runtime change): the story's page-open walks only the run's own section and reports
  how many sections it read and how long it took; a frame-script timeout is its own error, recorded as
  `"fault": "harness"`, and never retried; beats that need a beat that did not pass are recorded as
  `"skipped": "not run: needs ..."` with no prompt sent. Rerun live on the fixed harness:
  `hol-a20-40f1b3e-h96d3510-story-2` (table above).
* **S5 in those failed runs.** S5 called no Office tool: its note asked for "text exactly the replyText from the
  brain_make_deliverables outlook result", and that result did not exist because S3 never ran. The model declining
  (once after `clarify`) was correct, not a cascade of the eval timeout.
* **`brain_why` cites any session's artifact in the profile.** In `story-2-gw3` S4 found the tracker that the passing
  gw2 run 2 had made (`brain.db` persists across runs in a profile) and the harness rejected it as the wrong decision
  page. That is the product's behaviour: "why is this here?" answers from the profile's whole index.
* What each story beat checks, and the passing run's timings: see the `fe68c8d` section below (same checks; the
  fixture email is now "HOL live test: pilot sync"). gw2 run 2 on `40f1b3e`: chat turns S1 7.8 s, S2 12.7 s, S3
  56.5 s, S4 7.8 s, S5 6.2 s, S6 5.0 s (96.0 s); beats with the harness readbacks 165.3 s; whole run 550.1 s.
* Privacy: names. These runs' `results.json` / `events.jsonl` masked addresses but recorded the Outlook sender's
  display name (a tool result's `from.name`, the agent's reply text, a reply draft's quoted `From:` header). Those
  files were scrubbed after the runs. The harness now masks every display name in a from/to/cc/bcc/sender/
  organizer/attendees field as `<name>`, and the same name anywhere in later text (`scripts/livetest/redact.py`).
  Screenshots were already blurred over the reading pane, message list and address fields.

### Known limits seen on `40f1b3e`

* OneNote for the web can fail to load its editor (no ribbon after minutes; seen once in the story preparations on
  this commit). A person reloads the tab.
* Natural-language routing: "connect to <app>", the loop's own prompts and the refusals were bare prompts; the story
  prompts were not (they carry a routing note). All runs used one model (`claude-opus-5-5`).
* Hidden tabs: never exercised (both `L12_tab_switch` rows recorded all four tabs visible).
* Outlook calendar, new event: `calendar_draft_new` returns an honest `requested` receipt. Outlook's one popout
  attempt is blocked by the pane (as Hermes Desktop's would be), so no new event form opens in the pane; in all three
  runs the receipt said so and no form appeared. The person opens the new event in Outlook. The harness closes any
  event form left open, without saving it.
* The model sometimes needs a second ask. On this commit the model once called `clarify` before an insert (story gw3
  run 1, S5; the harness answers "Proceed as described in my message"). Earlier, on `77b5652`, it once called Hermes
  core's `manage_connections` for "connect to Outlook" and once added an unasked `brain_sync`; `3076fb1` tightened the
  tool descriptions. A person may have to repeat the request.

## Plugin `d8b3ec5` (F2, smoke)

`d8b3ec5` (configured add-in origin + local manifests): one run each, profile `office-live-nd8b3ec5`, served site
unchanged (`44d4a4f6…eabd6b7`): `HOL_IMAGES=1 all 1` PASS (Excel, Word, PowerPoint, OneNote, F3), `outlook 1` PASS,
`calendar 1` PASS (test event deleted), `story 1` PASS. The first attempt stopped at the preflight because the running
gateway took the new profile before its config enabled the plugin (a set-up race, fixed by the gateway's
`reload-plugins` for that profile, no restart).

## Plugin `d715acd`

* **Live verification on `d715acd` (profile `office-live-nd715acd`, served `SHA256SUMS` sha256
  `ea6e9363…acd57e75`, 33 files).** With pictures on, Excel, Word, PowerPoint and OneNote each passed 3/3
  (`hol-maint-d715acd-all-3`); the whole demo loop passed 2/2 (`hol-maint-d715acd-loop-2`); Outlook mail 1/1
  (`hol-maint-d715acd-outlook-1`); F3 passed (`hol-maint-d715acd-f3-1`). Outlook calendar: next section.

## Demo story on `fe68c8d`

* **The recorded demo story ([demo-script.md](demo-script.md)) passed 2/2 end to end on `fe68c8d`** (profile
  `office-live-nfe68c8d`, a fresh disposable profile made from `git archive fe68c8d`; served `SHA256SUMS` sha256
  `b192dff4…6b0b19`, 33 files; run `a18-story-2`, `run_macbook.sh story 2`). Each run used one new chat with five tabs
  paired to it: Outlook on the existing self-sent fixture email (reused, nothing sent) plus OneNote, Excel, Word and
  PowerPoint. Every beat was checked on the tool receipts and on the harness's own readback of each document.

| Beat | Run 1 / run 2 | Chat turn | Beat incl. harness readback |
| --- | --- | --- | --- |
| S1 `brain_from_email` indexes the email | PASS / PASS | 7.9 s / 7.8 s | 8.0 s / 7.9 s |
| S2 OneNote decision page from it | PASS / PASS | 11.8 s / 14.2 s | 28.8 s / 27.4 s |
| S3 tracker + chart, brief, deck, one picture, reply text, follow-up request | PASS / PASS | 48.8 s / 54.8 s | 61.6 s / 73.6 s |
| S4 `brain_why` on the FAQ tracker row | PASS / PASS | 10.0 s / 10.3 s | 10.5 s / 10.3 s |
| S5 person clicks Reply, Hermes fills it (`outlook_insert`) | PASS / PASS | 10.0 s / 9.9 s | 39.4 s / 39.3 s |
| S6 one `office_undo` of that insert | PASS / PASS | 5.2 s / 6.3 s | 7.5 s / 9.1 s |
| All six beats | 2/2 | 93.8 s / 103.4 s | 155.8 s / 167.7 s |

  Whole runs, including opening five documents, uploading the add-ins and pairing five panes: 539.6 s / 533.4 s.
  S5's beat time includes the harness's own steps: re-opening the email, clicking Reply, and opening the pane in the
  draft, where it resumed its pairing.
* What each beat checked:
  * S1: the email indexed as `Email: HOL live test: weekly sync`; the pane in read mode on the fixture.
  * S2: two decisions and Sam's FAQ action, with no quote unverified. The page itself shows
    `Source: Email: HOL live test: weekly sync`, and every brain.db row (read only) cites the email entry.
  * S3: Excel, Word and PowerPoint receipts `verified`. Read back: the tracker's FAQ row (owner Sam, Open, source
    the email), COUNTIF = 1, the "Open items by owner" chart; the brief with every item and the decision-record link;
    5 new slides ending with Sources; exactly one new picture on the title slide (a local PNG the harness writes, no
    image generation). The reply came back `applied` with `replyText`, not sent. The follow-up meeting came back as
    the calendar host `applied`, `requested: appointment`, not saved or sent. Outlook made 2 popout attempts, both
    blocked by the pane; no inline draft or event form appeared.
  * S4: FAQ row -> this run's decision page -> verified quote -> the email entry, and the answer names the email.
  * S5: the reply text in the draft the person opened, read back in the pane and on the page, with To only the
    account's own address and Cc/Bcc empty.
  * S6: the inserted text was gone and the draft was still open and unsent.
  After each run the harness discarded the reply draft. Nothing was sent or saved, and no test event was created.
* Before the fix: `story 2` on `51887bf` (profile `office-live-nd715acd`, served `ea6e9363`) passed S1-S4 in both
  runs. S5 and S6 failed in both runs on a harness defect (a readback key, and Hermes core's `tool_describe` counted
  as an Office tool), fixed in `67f8009`. A plugin copy defect was also fixed (RED/GREEN, `fe68c8d`): for an email
  source, the deck's Sources slide said "OneNote: Email: ...". The 2/2 above is on that fix. The Sources slide text is
  covered by the unit test; the live run checked the slide titles, not that slide's bullet text.

## Outlook calendar on `d715acd`

* **Outlook calendar on `d715acd`: every calendar tool step passed in every run that reached it; whole runs are
  not yet 3/3 because of one harness step.** Same profile and served bytes as above, harness commits `d40d201` ->
  `bc95aae` (plugin unchanged). Runs: `hol-maint-d715acd-calendar-3` 2/3, then `hol-a15-d40d201-calendar-3` 1/3,
  `hol-a15-ce0446d-calendar-3` 0/3, `hol-a15-925f43c-calendar-3` 1/3. Across them pair, `calendar_get_event`,
  `calendar_update_draft`, one `office_undo` and the refused attendee/invitation request passed 10/10 runs that
  reached them; `calendar_draft_new` (honest `requested` receipt, blocked popout) passed 4/4 runs in which the
  harness got the Omnia Works pane open on the mail fixture. Every failed run failed in the harness, not in a tool:
  two runs waited on a pane in the A14 `waiting-bridge` state (fixed, `ce0446d`), and the rest could not open the
  pane from More apps on the mail fixture after the event form, even after one retry that re-opens the message
  (open; a freshly opened message shows it at once). The harness now deletes its one test event in a finally
  path (after the last run, a failed run, or a run that raises); `hol-a15-925f43c-calendar-3` and
  `hol-a15-d40d201-calendar-3` deleted theirs that way.

## Medians

Live-test medians, quoted from the harness `summary.json` files (pane-measured `elapsedMs` /
plugin-measured `roundTripMs`, Hermes to pane and back; model and approval time excluded):

| Host | Runs passed | `elapsedMs` median | `roundTripMs` median | Plugin commit measured | Run |
| --- | --- | --- | --- | --- | --- |
| Excel | 3/3 | 187 ms | 230 ms | `40f1b3e` | `hol-a19-40f1b3e-all-3` |
| Word | 3/3 | 1372 ms | 1393 ms | `40f1b3e` | `hol-a19-40f1b3e-all-3` |
| PowerPoint | 3/3 | 448.5 ms | 503 ms | `40f1b3e` | `hol-a19-40f1b3e-all-3` |
| OneNote | 3/3 | 35 ms | 81 ms | `40f1b3e` | `hol-a19-40f1b3e-all-3` |
| Outlook (mail) | 3/3 | 95 ms | 167 ms | `40f1b3e` | `hol-a19-40f1b3e-outlook-3-gw2` |
| Outlook calendar | 3/3 | 147 ms | 257 ms | `40f1b3e` | `hol-a19-40f1b3e-calendar-3-gw2` |
| Demo story | 1/2 | 98 ms | 181 ms | `40f1b3e` | `hol-a19-40f1b3e-story-2-gw2` |

Medians include the picture steps. The `d715acd` figures (Excel 185/271, Word 792/815, PowerPoint 681.5/815.5,
OneNote 26/49, Outlook 131/174, calendar 128.5/159.5 ms) and earlier runs' (e.g. `b9f74ce`, `1992088`, `c6d20e6`) are
other runs and are not these numbers.

## Earlier runs

* **`b9f74ce`; with pictures on, every host passes 3/3.** Measured by the automated
  harness ([harness.md](harness.md)): a real chat in a Hermes gateway, real OneNote/Excel/Word/PowerPoint for the web in an
  Electron 44.9.3 window with the same `<webview>` settings as the Hermes Desktop preview pane (that
  window is not the Hermes Desktop app itself). On `b9f74ce` (served `SHA256SUMS` sha256
  `e4308a98…7f15e7cc`, the same add-in bytes as `c6d20e6`; plugin runtime unchanged since `c6d20e6`,
  `b9f74ce` changes only the harness readbacks): with pictures on (`HOL_IMAGES=1`,
  `hol-nightly-b9f74ce-images-all-3`) Excel, Word, PowerPoint and OneNote each passed 3/3, and F3
  passed. In Word runs 2 and 3 Word for the web showed its own "Sorry, your session expired" banner
  right after the harness's document reset; the harness reloaded the document once (no write) and the
  pane resumed its pairing, so those two runs started on Word's last saved document (earlier runs'
  text and pictures still in it) rather than the reset one; their checks compare before and after
  each step and passed. Last measured on `c6d20e6` and not rerun since (same plugin runtime): the
  whole demo loop in one chat with four Office tabs passed 2/2 (`hol-nightly-c6d20e6-loop-2`),
  including the decision page and its backlink, `brain_make_deliverables` into Excel, Word and
  PowerPoint verified by readback, the denied change and the final undo; F3 passed
  (`hol-nightly-c6d20e6-f3-1`); two profiles served by one gateway kept separate `brain.db` files,
  and profile B could neither replace nor undo profile A's Links & Backlinks block
  (`hol-nightly-c6d20e6-isolation`). The release catalog pin (`catalog/hermes-office-desktop.yaml`,
  `sha: TODO`) stays unset.

* **Outlook on the web (A10b, `1992088`): 3/3 passed** (`hol-a10b-1992088-outlook-3`, profile
  `office-live-n1992088`, served `SHA256SUMS` sha256 `927713b9…203eaa3`): pair, `outlook_get_item` on the one
  harness-sent fixture (read mode, its Sent Items copy: the self-addressed message never reached this new account's
  Inbox), `outlook_draft_reply` (receipt `requested`; Outlook's popout was attempted and blocked by the pane, as
  Hermes Desktop's would be), the reply fallback (the person clicks Reply, `outlook_insert`, the reply shows the text
  and its To is only the account's own address), compose `outlook_insert` + `office_undo` read back by Office.js, and
  a send / add-recipient request refused (no tool called, To/Cc/Bcc empty, draft unsent). Excel smoke on the same
  commit and served bytes: 1/1 (`hol-a10b-1992088-excel-1`). The medians above are the current `d715acd` figures.

## Windows: offline readiness

Offline checks on a Windows 11 PC (host label withheld) with Hermes Agent v0.23.0 (`2bd378ed`), plugin `git archive ff9771e`, no
Microsoft sign-in, no model call (A35). At the time of these offline checks `plugin.yaml` was
`platforms: [macos]` and nothing had run in Hermes Desktop or Office on Windows; the real Hermes Desktop session
that followed is recorded in the next section.

* Unit tests with the test PC's Hermes Python 3.11.15 (pytest 9.1.1 in a separate folder): **533 passed, 13 failed**. Node
  v22.22.1: **197/197**. `hermes plugins validate`: passed (the same two warnings as on macOS). All 13 failures are
  Windows-only (on macOS the same tree passes 546/546):
  * 11 are test or harness assumptions, not product faults: 3 tests read UTF-8 files without `encoding="utf-8"`
    (Windows defaults to cp1252); one uses `os.mkfifo` (no FIFOs on Windows); two expect `/` in the native
    `needsRaster.saveTo` paths; three check the macOS-only live harness (`bash -n` of a Windows path, and
    `live_driver.py` naming files with `\` where the archive uses `/`); one sets a 1 ns mtime that NTFS stores as 0
    (100 ns resolution); one sets `HOME`, which Windows `~` expansion ignores (it uses `USERPROFILE`).
    Since fixed in the tests and in `live_driver.py`, with LF line endings pinned in `.gitattributes` for the
    hashed schema copies and the shell scripts: the current unit suites pass on Windows (unit tests only).
  * 2 (`tests/test_port_release.py`) show a **real Windows defect**, confirmed directly: two bridges in one process both
    bound `127.0.0.1:47615`. Python's `http.server` sets `SO_REUSEADDR`, and on Windows that lets a second socket
    bind a port that is already listening, so a second Hermes profile would share 47615 instead of taking
    47616-47619, and a pane could reach the wrong profile's bridge (pairing then fails). Since fixed in code: on
    Windows the bridge binds with `SO_EXCLUSIVEADDRUSE` and no `SO_REUSEADDR`; a unit test checks it (not run live).
* Bridge: the plugin's own `BridgeServer`, started in a throwaway Python process, listened on `127.0.0.1:47615` only;
  `/v1/hello` with the allowed origin answered 200, another origin 403, a wrong pairing code 401, and a minted code
  paired (200, pane token issued). Nothing listened on 47615-47619 before or after.
* Plugin load: in a new profile with the plugin copied into its plugins folder and enabled, `hermes plugins list`
  shows it enabled, and the running Hermes gateway's reload-plugins for that profile loaded it (the `pre_tool_call`
  hook and the same 86 tools as on macOS). `platforms: [macos]` does not stop that: in Hermes 0.23.0 the field only
  makes a catalog install (`hermes plugins install`) refuse an unlisted OS. A catalog install on Windows would be
  refused until the field lists `windows`.
* Add-in site: from the test PC the five manifests, the five task-pane pages, their scripts, the shared stylesheet, the
  support page and the three icons all answered HTTP 200 with the same bytes as the served site.

## Real Hermes Desktop session, Windows, 2026-11-10

One session in the real Hermes Desktop app on a Windows 11 PC, 11:54 to 12:24 local time, one Hermes profile,
plugin runtime `402f00e` (F8) with the Workbench display name (Outlook and OneNote still showed the cached old
add-in name), throwaway Office files. This was the real app, not the automated harness. The session record names
a person for the two recoveries below; it does not say who entered the prompts. The five-prompt story:

| Prompt | Result |
|---|---|
| 1. Decision page in OneNote | Pass: page written and read back, each decision and action present exactly once |
| 2. Excel tracker, Word brief, PowerPoint deck | Pass after one recovery: the PowerPoint pane stopped answering in a background tab; the person reopened that pane, then the deck was verified (4 new slides, one image inserted with one approval) |
| 3. Trace an FAQ row to its source | Pass |
| 4. Insert text into an Outlook reply draft | Pass after the known Outlook step: the first attempt was refused safely because the pane was still on the read email; the person reopened Workbench from the draft's Message > Apps; insert verified, recipients unchanged, nothing sent |
| 5. Undo the insert | Pass: `office_undo` removed it; readback shows only the quoted original; nothing sent |

Not run in that session: two profiles at once, a catalog install on Windows, and any run against the public
add-in site. Seen and not fixed: work in a background tab was slow or stopped on Windows (it hit the PowerPoint
pane); the cause was not measured.

An earlier Windows session on 2026-11-09 (runtime `ff9771e`) ran prompt 1 only: a page read missed list items, the
agent wrote the decisions twice and undo did nothing. F8 fixed both; prompt 1 above is the retest of the
duplicate. Undo of an uncertain OneNote write was not retested on Windows.

## Not yet run live

* A14 idle port release (a sixth profile while five hold ports; a pane left open across a release): unit-tested
  only. See "Five bridge ports" in [reference.md](reference.md#limits-five-bridge-ports-shared-by-every-profile).
* Windows with two Hermes profiles running Workbench at once: not run (the port 47615 double bind is fixed; unit-tested only).
* A catalog install on Windows: not run (the Windows session used a manually copied plugin folder).
* The harness's Electron 44.9.3 window uses the same `<webview>` settings as the Hermes Desktop preview pane, but it
  is not the Hermes Desktop app itself.
