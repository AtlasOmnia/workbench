# SPDX-License-Identifier: Apache-2.0
"""A24: the copy says what the receipts guarantee, names the exceptions, and matches the result directories."""
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent


def _read(rel):
    return (ROOT / rel).read_text(encoding="utf-8")


def test_no_universal_readback_or_undo_promise():
    for rel in ("README.md", "docs/demo-script.md", "docs/reference.md", "skills/office-desktop/SKILL.md",
                "catalog/hermes-office-desktop.yaml", "addin/support.html"):
        text = " ".join(_read(rel).split())
        assert "read back from the document, and can be undone" not in text, rel
        assert "Every change was read back" not in text, rel
        assert "can be undone. Nothing was sent" not in text, rel


def test_undo_exceptions_and_receipt_guarantee_are_named():
    readme = " ".join(_read("README.md").split())
    for needle in ("too large to snapshot", "10,000 cells", "bare selection", "slide delete", "OneNote page Hermes created",
                   "lost when the pane reloads", "autosaved by Outlook itself", "`verified` (read back from the document)"):
        assert needle in readme, needle
    ref = " ".join(_read("docs/reference.md").split())
    assert "There is no universal post-write readback" in ref and "autosaved by Outlook itself" in ref


def test_speaker_notes_are_listed_as_unsupported_wherever_capabilities_are_listed():
    for rel in ("README.md", "docs/reference.md", "skills/office-desktop/SKILL.md", "catalog/hermes-office-desktop.yaml"):
        text = " ".join(_read(rel).split())
        assert "powerpoint_add_speaker_notes" in text and "Unsupported" in text, rel


def test_read_limits_and_truncation_are_disclosed():
    readme = " ".join(_read("README.md").split())
    assert "first 20,000 characters of the document body" in readme and "20,000 characters of a message or event body" in readme
    skill = " ".join(_read("skills/office-desktop/SKILL.md").split())
    assert "`truncated: true`, tell the user that your summary covers only the first part" in skill


def test_limits_cover_onenote_ribbon_long_chat_idle_and_brain_why():
    limits = _read("README.md").split("## Limits", 1)[1].split("\n## ", 1)[0]
    flat = " ".join(limits.split())
    assert "no ribbon after a minute?** Reload the tab" in flat
    assert "`NotPaired`: ask Hermes to connect again" in flat
    assert "more than 10 minutes?** Leave the pane open" in flat
    assert "earlier chat in the same profile" in flat


def test_first_connect_is_honest_and_leads_with_the_working_upload_path():
    readme = _read("README.md")
    assert "## First connect (about 2 minutes per app, once)" in readme and "60 seconds" not in readme
    sec = " ".join(readme.split("## First connect", 1)[1].split("\n## ", 1)[0].split())
    office = sec.index("**Excel, Word, PowerPoint:** Home > Add-ins > More Add-ins > Manage My Add-ins")
    assert "More Settings" not in sec and office > 0
    assert "**OneNote:** Insert > Office Add-ins > My Add-ins > Manage My Add-ins > Upload My Add-in" in sec


def test_harness_and_demo_script_say_what_was_and_was_not_exercised():
    harness = " ".join(_read("docs/harness.md").split())
    assert "Hidden tabs stay attached" not in harness and "Hidden tabs were not exercised" in harness
    assert "not yet run live" not in harness and "story-bare" in harness
    demo = " ".join(_read("docs/demo-script.md").split())
    assert "You do not need the note" not in demo and "that note is tool routing" in demo
    assert '"Still open"' not in demo
    ev = " ".join(_read("docs/live-evidence.md").split())
    assert "Not yet rerun live" not in ev and "2 of 6 passed all six beats" in ev
    assert "not yet rerun live" not in " ".join(_read("README.md").split())


def test_chat_switch_wording_is_marked_unconfirmed_in_reference_only():
    assert "to be confirmed in Hermes Desktop" in _read("docs/reference.md")
    for rel in ("README.md", "skills/office-desktop/SKILL.md"):
        assert "to be confirmed in Hermes Desktop" not in _read(rel), rel


def test_the_bare_prompt_evidence_is_the_cold_run_and_the_warm_2_of_2_is_labelled_warm():
    """A26: the matrix's story-bare 2/2 on 0847e28 ran on a profile that had learned the demo by itself; the cold
    result is the bare-prompt evidence, and no doc presents the warm one as it."""
    ev = " ".join(_read("docs/live-evidence.md").split())
    assert "## Bare prompts, cold: `0847e28` (A26)" in ev and "This is the bare-prompt evidence for `0847e28`" in ev
    assert "| `story-bare 2` | **2/2 PASS (warm)** |" in ev and "| `story-bare 2` | **2/2 PASS** |" not in ev
    assert "Cold bare-prompt result on `0847e28`: **1 of 4 runs**" in ev
    readme = " ".join(_read("README.md").split())
    assert "2/2 with the bare prompts" not in readme
    assert "on a cold test profile" in readme and "(warm) and is not the evidence" in readme


def test_f8_is_the_current_evidence_and_older_commits_are_cited_only_labelled():
    """A31 (was A29's F6 test): the README and live-evidence cite the F7 matrix; Outlook mail and calendar, which
    neither the F6 nor the F7 diff can touch, are cited from F5 and labelled so; F6 and F4 appear only labelled; the F6
    section carries the A30 correction; no older commit is presented as the current result."""
    ev = " ".join(_read("docs/live-evidence.md").split())
    for old in ("## Current: F6", "## Current: F5", "## Current: F4"):
        assert old not in ev
    assert "## F7 `d1f7aae` (A31)" in ev and "## F6 `5df2609` (A29)" in ev and "## F5 `8fa1f63` (A27)" in ev
    assert "Cold bare-prompt result on `d1f7aae`: **2 of 2 runs**, first attempt" in ev
    assert "Not run on F7, because its diff changes no code they use: `calendar` and `outlook`" in ev
    assert "Correction (A30): on `5df2609` the hold was released by a second blocked retry without any read" in ev
    assert ev.index("## F7 `d1f7aae` (A31)") < ev.index("## F6 `5df2609` (A29)") < \
        ev.index("## F5 `8fa1f63` (A27)") < ev.index("## Bare prompts, cold: `0847e28` (A26)")
    readme = " ".join(_read("README.md").split())
    assert "shipped runtime 402f00e (F8, A46)" in readme
    assert "## Current: F8 `402f00e` (A45/A46)" in ev
    assert ev.index("## Current: F8") < ev.index("## F7")
    assert "Standalone Outlook/mail and calendar results and timing rows remain **F5 `8fa1f63`**" in ev
    assert "on 5df2609 (F6) — every app" not in readme
    assert "on 8fa1f63 (F5) — every app" not in readme and "on 0847e28 (F4) — every app" not in readme
    assert "Outlook (mail), F5 |" in readme and "Outlook calendar, F5 |" in readme
    for old in ("On 0847e28 (F4) the same cold bare prompts passed 1 of 4 runs", "the 3/3 over midnight on 0847e28 (F4)",
                "Outlook calendar on 8fa1f63 (F5): 1/1"):
        assert old in readme


def test_a_pageless_change_whose_get_context_failed_is_documented_as_held_not_settled():
    """A32 P1 / A33: when the preliminary get_context fails, no later read can settle the change; the docs say it stays
    held (not re-sent) and the user checks the page, instead of implying a read always settles it."""
    ev = " ".join(_read("docs/live-evidence.md").split())
    assert "if that preliminary `get_context` read itself fails" in ev
    assert "it stays held indefinitely" in ev and "the identical change is not re-sent" in ev
    assert "reload the OneNote tab and check the page themselves" in ev and "never a duplicate" in ev
    ref = " ".join(_read("docs/reference.md").split())
    assert "if that read fails, the change is sent with no page recorded and no later read can release it" in ref
    assert "the identical change is not re-sent" in ref
    readme = " ".join(_read("README.md").split())
    assert "Open page unknown: stays held, not re-sent; reload the tab, check the page." in readme


def test_real_desktop_session_findings_reach_the_user_guidance():
    """A34: the first real Hermes Desktop session (Studio, F7) showed what a person must do that the harness did for
    itself; the skill, README, demo script and live evidence say it in user words, and nothing unproven is promised."""
    skill = " ".join(_read("skills/office-desktop/SKILL.md").split())
    for needle in ("Outlook on the web has no pin for this pane: it stays on the message it was opened with",
                   "the user opens the email FIRST", "search All folders",
                   "opens Workbench from the DRAFT's Message tab > Apps (not Home > More apps)",
                   "While the pane is still on the read message, do not fill anything",
                   "direct notebook-list clicks have been observed opening a blocked new window", "Create classic notebook", "first Hermes Desktop run pending",
                   "`drive_preview` acts only on the ACTIVE tab", "Never open or reload an app over another app's tab",
                   "click the Hermes chat box before typing a prompt", "\"cell-editing mode\"",
                   "`uncertain` Excel result is checked with `excel_get_workbook_overview` before any retry",
                   "A pairing code lasts 10 minutes", "\"Hermes is not answering on this computer\"",
                   "every pane must be paired again", "never enter one yourself", "\"Hermes is working…\"",
                   "the Add-ins button does nothing for a few seconds"):
        assert needle in skill, needle
    assert "can pin it" not in skill
    readme = " ".join(_read("README.md").split())
    connect = readme.split("## First connect", 1)[1].split("## Limits", 1)[0]
    for needle in ("There is no pin", "draft's Message tab > Apps (not Home > More apps)", "Search All folders",
                   "does nothing for a few seconds after the file opens", "lasts 10 minutes",
                   "\"Hermes is not answering on this computer\" after a wait", "click the Hermes chat box before you type"):
        assert needle in connect, needle
    assert "and pin it" not in readme
    limits = readme.split("## Limits", 1)[1].split("## Privacy", 1)[0]
    for needle in ("Excel refuses changes while you edit a cell", "not yet verified live in Hermes Desktop",
                   "After Hermes restarts, pair every pane again", "stuck on \"Hermes is working…\""):
        assert needle in limits, needle
    demo = " ".join(_read("docs/demo-script.md").split())
    for needle in ("Set every app up off camera", "pair all five panes again right before the take",
                   "open the meeting-notes email first", "Click the Hermes chat box before typing each prompt",
                   "You click the app tabs", "self-learning off", "the profile is warm",
                   "open Workbench from the draft's Message tab > Apps"):
        assert needle in demo, needle
    ev = " ".join(_read("docs/live-evidence.md").split())
    assert "## Real Hermes Desktop session, Studio, F7" in ev and "**Warm, not cold evidence:**" in ev
    assert "**5 of 5 prompts passed in one chat.** Nothing was sent or saved" in ev
    assert "stopped mid-run by an external signal" in ev
    assert ev.index("## F7 `d1f7aae` (A31)") < ev.index("## Real Hermes Desktop session, Studio, F7") < \
        ev.index("## F6 `5df2609` (A29)")


def test_onenote_opening_uses_observed_my_files_not_sharing_onboarding():
    """Static instructions, not a model response or end-to-end notebook-open observation."""
    skill = " ".join(_read("skills/office-desktop/SKILL.md").split())
    for needle in ("My Files opens the file browser in a separate window",
                   "That separate window is intentional", "Do not generalize to all new windows",
                   "final opened-notebook browser address", "Never automatically revert to sharing instructions",
                   "Never change permissions or create a sharing link",
                   "No pairing at tab-open or while only the external notebook is open"):
        assert needle in skill, needle
    assert "People with existing access" not in skill
    assert "Otherwise create one" not in skill
    readme = " ".join(_read("README.md").split())
    assert "Pair only once it loads here" in readme
    assert "Only My Files window-opening is user-confirmed" in readme


def test_readme_says_how_to_stop_approval_prompts_and_what_yolo_gives_up():
    """The README offers "always for this session" first (destructive tools keep asking) and says plainly that
    /yolo turns off every approval in the chat, those included."""
    readme = " ".join(_read("README.md").split())
    safety = readme.split("## Privacy and safety", 1)[1]
    assert "always for this session" in safety and "can't be undone still ask" in safety
    assert "`/yolo` turns off every approval in that chat" in safety and "Hermes' own command approvals" in safety
