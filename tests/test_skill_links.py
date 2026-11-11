# SPDX-License-Identifier: Apache-2.0
"""A22: every app office_connect accepts has a direct https link in the skill, and nothing sends the user to the
Microsoft 365 app launcher (it opens new windows, which the Desktop preview pane blocks)."""
import re
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
SKILL = ROOT / "skills" / "office-desktop" / "SKILL.md"
ROW = re.compile(r"^\| [^|]+\(`(\w+)`\) \| (\S+) \| (\S+) \|$", re.M)
# Phrases that send the user somewhere the pane cannot follow; a line may name them only to forbid them.
LAUNCHER = re.compile(r"app launcher|waffle|open in new window|open in desktop app", re.I)
FORBIDS = re.compile(r"\b(never|not|don't|do not|blocks?)\b", re.I)


def _apps():
    import bridge_server
    return set(bridge_server.APPS)


def _links():
    rows = ROW.findall(SKILL.read_text())
    assert rows, "no direct-link table in the skill"
    return rows


def test_every_supported_app_has_https_direct_links_for_both_account_types():
    rows = _links()
    assert {app for app, _, _ in rows} == _apps()
    for app, personal, work in rows:
        for url in (personal, work):
            assert re.fullmatch(r"https://[a-z0-9.-]+\.(?:microsoft|com)(?:/[\w./-]*)?", url), (app, url)


def test_outlook_links_match_the_account_type():
    by = {}
    for app, personal, work in _links():
        by.setdefault(app, []).append((personal, work))
    for personal, work in by["outlook"]:
        assert personal.startswith("https://outlook.live.com/") and work.startswith("https://outlook.office.com/")
    assert {p.split("/")[3] for p, _ in by["outlook"]} == {"mail", "calendar"}


def test_links_match_the_live_harness():
    """The personal links are the ones the live-test harness opens."""
    import importlib.util
    spec = importlib.util.spec_from_file_location("hol_office_ui", ROOT / "scripts" / "livetest" / "office_ui.py")
    ui = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(ui)
    personal = {p for _, p, _ in _links()}
    assert set(ui.HOME_URLS.values()) | {ui.ONENOTE_HOME, ui.CAL_URL} <= personal
    assert "https://outlook.live.com/mail/0/" in personal


def test_skill_readme_and_demo_never_recommend_the_app_launcher():
    for path in (SKILL, ROOT / "README.md", ROOT / "docs" / "demo-script.md", ROOT / "docs" / "reference.md"):
        text = path.read_text()
        # judge each sentence that names the launcher: it must forbid it or say the pane blocks it
        for sentence in re.split(r"(?<=[.:;])\s+", re.sub(r"\s+", " ", text)):
            if LAUNCHER.search(sentence):
                assert FORBIDS.search(sentence), (path.name, sentence)
    skill = SKILL.read_text()
    assert "desktop_preview` `action: open`" in skill and '"+"' in skill


def test_a_picture_named_by_file_name_is_looked_up_in_the_hermes_image_folders_never_invented():
    """A26: "put team.png ... on the title slide" names a file; the skill sends Hermes to its image folders first and
    forbids inventing or generating a stand-in unless asked."""
    text = " ".join(SKILL.read_text(encoding="utf-8").split())
    section = text.split("## Pictures: `office_insert_image`", 1)[1].split("## ", 1)[0]
    named = section.split("- **a named file**", 1)[1].split("- **", 1)[0]
    assert "the plugin itself looks in this profile's Hermes image folders first" in named and "`$HERMES_HOME/cache/images`" in named
    assert '`"path": "team.png"`' in named and "`source.path`" in named  # A27: the plugin resolves the bare name
    assert "never invent, generate or download a picture in its place unless the user asks for that" in named
    assert section.index("- **a named file**") < section.index('- **"generate"**')
