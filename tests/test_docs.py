# SPDX-License-Identifier: Apache-2.0
"""The README stays scannable and every relative link in it and in docs/ resolves (A17)."""
import re
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
LINK = re.compile(r"!?\[[^\]]*\]\(([^)\s]+)\)")


def _slug(heading: str) -> str:  # GitHub's anchor rule, enough for these headings
    s = re.sub(r"[^\w\- ]", "", heading.strip().lower())
    return s.replace(" ", "-")


def test_readme_is_short_and_starts_with_the_one_sentence():
    text = (ROOT / "README.md").read_text()
    assert len(text.encode()) < 13 * 1024  # A34: +1 KiB for the real-Desktop first-connect steps and limits
    lines = text.splitlines()
    assert lines[1] == "# Workbench by Atlas Omnia (0.3.0)" and lines[2] == ""
    assert lines[3] == ("Open and edit Word, Excel, PowerPoint, OneNote and Outlook in a Hermes Desktop tab, with your "
                        "agent working alongside you.")
    assert "shipped runtime 402f00e (F8, A46): all four apps 1/1 with images" in text
    assert "Standalone" not in text or "F5" in text
    assert "Also from Atlas Omnia: Hermes for LibreOffice (Calc, Writer, Impress) — separate repo." in text


def test_readme_gif_is_small():
    gif = ROOT / "assets" / "demo-highlight.gif"
    assert gif.read_bytes()[:6] == b"GIF89a" and gif.stat().st_size < 3 * 1024 * 1024


def test_relative_links_resolve():
    bad = []
    for md in [ROOT / "README.md", *sorted((ROOT / "docs").glob("*.md"))]:
        for target in LINK.findall(md.read_text()):
            if re.match(r"[a-z]+:", target):
                continue
            path, _, anchor = target.partition("#")
            dest = (md.parent / path).resolve() if path else md
            if not dest.exists():
                bad.append(f"{md.name}: {target} (missing file)")
            elif anchor and dest.suffix == ".md":
                heads = {_slug(h) for h in re.findall(r"^#+ (.+)$", dest.read_text(), re.M)}
                if anchor not in heads:
                    bad.append(f"{md.name}: {target} (missing anchor)")
    assert not bad, bad


def test_windows_tab_in_front_tip_in_skill_and_reference():
    """F8: live on Windows, writes from a background tab were slow and came back uncertain."""
    tip = "On Windows, keep the app's tab in front while Hermes writes to it; background tabs can make writes slow and uncertain."
    assert tip in (ROOT / "skills" / "office-desktop" / "SKILL.md").read_text()
    assert tip in (ROOT / "docs" / "reference.md").read_text()
