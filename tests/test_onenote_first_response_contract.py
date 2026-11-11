# SPDX-License-Identifier: Apache-2.0
"""Static instruction contract only: these tests do not invoke or observe a fresh model.

The first response is the acceptance boundary, not a later correction. User-observed
My Files window opening is not evidence of end-to-end open/copy/paste/pair success.
"""
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent


def _read(rel):
    return " ".join((ROOT / rel).read_text().split())


def test_first_response_contract_is_in_the_open_flow_before_pairing():
    flow = _read("skills/office-desktop/SKILL.md").split("## Flow", 1)[1].split("## Tabs, codes", 1)[0]
    assert "FIRST assistant response" in flow
    assert '"open OneNote tab"' in flow
    assert '"Opened the OneNote tab. Which notebook would you like to use?' in flow
    assert "Click My Files to open the file browser in a separate window, open your notebook there," in flow
    assert "then copy its browser address and paste it into this OneNote tab’s address bar." in flow
    assert "Once the notebook opens here, we’ll connect it to Hermes." in flow
    assert flow.index("FIRST assistant response") < flow.index('"Connect to Excel')
    assert "Do not call the notebook ready" in flow
    assert "Do not offer or create a new notebook by default" in flow
    assert "Only pair OneNote after" in flow


def test_onboarding_adapts_to_observed_state_and_preserves_address_boundary():
    skill = _read("skills/office-desktop/SKILL.md")
    for needle in (
        "Do not ask for the notebook name again if the user supplied it",
        "If the actual notebook is already open, skip onboarding",
        "Do not claim My Files exists without observing it",
        "If My Files is absent or the landing page differs, inspect and report",
        "final opened-notebook browser address",
        "not a file-list address, an `onenote:` app link, or a sharing link",
        "If paste fails, capture the actual error",
        "offer only a verified fallback",
        "Never automatically revert to sharing instructions",
        "Never change permissions or create a sharing link",
        "Only if the user explicitly asks to create a notebook",
    ):
        assert needle in skill, needle
    assert "People with existing access" not in skill
    assert "https://onedrive.live.com/?view=0" not in skill
    assert "Otherwise create one" not in skill


def test_docs_use_my_files_without_claiming_end_to_end_success():
    for rel in ("README.md", "docs/demo-script.md", "docs/harness.md"):
        text = _read(rel)
        assert "My Files" in text, rel
        assert "separate window" in text, rel
        assert "People with existing access" not in text, rel
        assert "Copy link" not in text, rel
    readme = _read("README.md")
    assert "first response" in readme and "which notebook" in readme
    assert "not yet verified live in Hermes Desktop" in readme
    demo = _read("docs/demo-script.md")
    assert "first response confirms the tab opened" in demo
    assert "does not create one by default" in demo
    assert "not end-to-end" in demo


def test_live_acceptance_is_explicitly_separate_from_static_contract():
    text = _read("docs/harness.md")
    for needle in (
        "## OneNote first-response acceptance (manual, fresh chat)",
        "test_onenote_first_response_contract.py",
        "not observed fresh-model behavior",
        "open OneNote tab",
        "FIRST assistant response",
        "Do not count a later correction as a pass",
        "only My Files opening a separate window is user-confirmed",
        "open/copy/paste/pair is not end-to-end proven",
        "do not compete for the mouse",
    ):
        assert needle in text, needle
