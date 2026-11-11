# SPDX-License-Identifier: Apache-2.0
"""Portable guidance contracts, not native-host OOXML acceptance tests."""
import json
import re
import shutil
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parent.parent
SKILL = ROOT / "skills/office-desktop/SKILL.md"
REFS = ("word-design-and-ooxml.md", "powerpoint-design.md")


@pytest.mark.parametrize("name", REFS)
def test_registered_plugin_bundle_serves_linked_reference(plugin, tmp_path, name):
    # Exercise the real plugin registration, then relocate its complete skill bundle.
    from test_onenote import FakeCtx
    ctx = FakeCtx()
    plugin.register(ctx)
    registered = ctx.skills["office-desktop"]
    target = "references/" + name
    assert f"]({target})" in registered.read_text()
    relocated = tmp_path / "bundle/skills/office-desktop"
    shutil.copytree(registered.parent, relocated)
    serving = pytest.importorskip("tools.skills_tool_plugin")
    linked = serving._plugin_skill_linked_files(relocated)
    assert target in linked["references"]
    reply = json.loads(serving._serve_skill_file(relocated, target, "hermes-office-desktop:office-desktop"))
    assert reply["success"] is True
    assert reply["content"] == (registered.parent / target).read_text()


@pytest.mark.parametrize("name", REFS)
def test_guidance_is_portable_and_names_only_registered_tools(name):
    text = (SKILL.parent / "references" / name).read_text()
    assert not re.search(r"/Users/|[A-Z]:\\\\|(?:https?://)?[\w.-]+\.ts\.net|127\.0\.0\.1|localhost", text)
    assert not re.search(r"https?://(?!schemas\.|www\.w3\.org/)[^\s`<>]+", text)
    tools = {s["name"] for host in ("word", "powerpoint")
             for s in json.loads((ROOT / f"schemas/{host}.json").read_text())}
    assert set(re.findall(r"\b(?:word|powerpoint)_[a-z_]+\b", text)) <= tools
    assert "native-host" in text and "readback" in text


def test_word_design_selection_and_recovery_contract():
    text = (SKILL.parent / "references" / REFS[0]).read_text()
    for phrase in ("native text/style", "Title", "Heading 1", "Normal", "spacing", "palette",
                   "word_read_selection", "word_get_document_structure", "current selection",
                   "start", "end", "selection", "FIRST failed OOXML insertion", "stop",
                   "GeneralException", "not proof that nothing changed", "no automatic flatten",
                   "exact failing payload", "unresolved", "one write", "readback"):
        assert phrase in text
    assert "targetText" in text and "not supported" in text


def test_ooxml_contract_and_evidence_boundary():
    text = (SKILL.parent / "references" / REFS[0]).read_text()
    for phrase in ("Word document XML", "Flat OPC", "pkg:package", "w:document", "w:body",
                   "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                   "http://schemas.microsoft.com/office/2006/xmlPackage",
                   "relationships", "styles.xml", "escape", "HTML", "regex",
                   "not XML or host validity", "static", "not host-tested"):
        assert phrase in text
    # Exact example is separately namespace/package linted, never host-qualified.
    assert "```xml" in text and "STATIC" in text


def test_powerpoint_design_uses_supported_layout_and_formatting():
    text = (SKILL.parent / "references" / REFS[1]).read_text()
    for phrase in ("hierarchy", "contrast", "font", "margins", "one message", "points",
                   "powerpoint_create_text_slide", "powerpoint_read_slide_content",
                   "powerpoint_format_text", "shapeId", "zero-based", "Unsupported",
                   "speaker notes", "clipping", "visual"):
        assert phrase in text


def test_static_native_text_example_matches_current_schema(plugin):
    text = (SKILL.parent / "references" / REFS[0]).read_text()
    match = re.search(r"```json\n(.*?)\n```", text, re.S)
    assert match is not None
    example = json.loads(match.group(1))
    from hermes_office_desktop import validation
    schema = next(s for s in json.loads((ROOT / "schemas/word.json").read_text())
                  if s["name"] == example["tool"])
    validation.validate(example["arguments"], schema["parameters"])
    assert example["arguments"]["location"] == "end"
    assert "\n" in example["arguments"]["text"]


def test_static_ooxml_example_matches_current_schema(plugin):
    text = (SKILL.parent / "references" / REFS[0]).read_text()
    xml = re.search(r"```xml\n(.*?)\n```", text, re.S).group(1)
    from hermes_office_desktop import validation
    schema = next(s for s in json.loads((ROOT / "schemas/word.json").read_text()) if s["name"] == "word_insert_ooxml")
    validation.validate({"ooxml": xml, "location": "end"}, schema["parameters"])
