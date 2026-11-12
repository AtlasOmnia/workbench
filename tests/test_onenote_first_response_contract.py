# SPDX-License-Identifier: Apache-2.0
"""Static instruction contract only: these tests do not invoke or observe a fresh model.

The first response is the acceptance boundary, not a later correction. User-observed
My Files window opening is not evidence of end-to-end open/copy/paste/pair success.
"""
from pathlib import Path

from test_manifests import connect, make  # noqa: F401  (make: the fixture that registers the plugin for office_connect)

ROOT = Path(__file__).resolve().parent.parent


def _read(rel):
    return " ".join((ROOT / rel).read_text(encoding="utf-8").split())


def test_first_response_contract_is_in_the_open_flow_before_pairing():
    flow = _read("skills/office-desktop/SKILL.md").split("## Flow", 1)[1].split("## Tabs, codes", 1)[0]
    assert "FIRST assistant response" in flow
    assert '"open OneNote tab"' in flow
    assert '"Opened the OneNote tab. Which notebook would you like to use?' in flow
    # W4: said up front, in plain words: a notebook-list click opens a window the pane blocks; the link goes in the chat
    assert "Clicking a notebook in this list tries to open a new window, which this pane blocks." in flow
    assert "Instead, click My Files (it opens in a separate window), open your notebook there," in flow
    assert "copy its link from that window’s address bar and paste it here in the chat." in flow
    assert "I’ll open it in this OneNote tab, then we’ll connect it to Hermes.\"" in flow
    assert "OneNote tab’s address bar" not in flow
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
        "open it yourself with `desktop_preview` `open` in the OneNote tab",
        "If the pasted link does not open the notebook, capture the actual error",
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
        assert "OneNote tab's address bar" not in text, rel  # W4: the link goes into the chat; Hermes opens it
    readme = _read("README.md")
    assert "first response" in readme and "which notebook" in readme
    assert "A list click opens a blocked window" in readme
    assert "paste its address in the chat; Hermes opens it" in readme
    assert "not yet verified live in Hermes Desktop" in readme
    demo = _read("docs/demo-script.md")
    assert "first response confirms the tab opened" in demo
    assert "paste it into the chat; Hermes opens it in the Workbench OneNote tab" in demo
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
        "paste it into the chat so Hermes opens it in the Workbench OneNote tab",
        "open/copy/paste/pair is not end-to-end proven",
        "do not compete for the mouse",
    ):
        assert needle in text, needle


def test_office_connect_for_onenote_gives_the_route_up_front(make, monkeypatch):
    """W4: the tool result itself (office_connect's OneNote firstTimeSetup) says the same thing in plain words."""
    monkeypatch.delenv("HERMES_CRON_SESSION", raising=False)
    monkeypatch.delenv("HERMES_SESSION_PLATFORM", raising=False)
    setup = " ".join(connect(make({}), "onenote")["firstTimeSetup"])
    for needle in ("Clicking a notebook in the notebook list tries to open a new window, which the pane blocks",
                   "click My Files (it opens in a separate window)", "paste it into the chat",
                   "Hermes opens it in the OneNote tab"):
        assert needle in setup, needle
    assert "Open or create the notebook" not in setup and "Copy link" not in setup
