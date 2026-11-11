# SPDX-License-Identifier: Apache-2.0
"""Authoring-example lint only: not runtime preflight or native acceptance."""
from pathlib import Path
import re
import xml.etree.ElementTree as ET
import posixpath
import pytest
ROOT = next(p for p in Path(__file__).resolve().parents if (p / "plugin.yaml").exists() or (p / "plugins").exists())
WB = (ROOT / "schemas/word.json").exists()
REF = ROOT / ("skills/office-desktop/references/word-design-and-ooxml.md" if WB else "plugins/atlas-sidecar-word/skills/office-guidance/references/word-workflow.md")
P = "http://schemas.microsoft.com/office/2006/xmlPackage"
W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R = "http://schemas.openxmlformats.org/package/2006/relationships"
D = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
def template():
    text = REF.read_text()
    assert "STATIC" in text and "not host-tested" in text
    return re.search(r"```xml\n(.*?)\n```", text, re.S).group(1)
def lint(xml):
    assert len(xml.encode("utf-16-le")) // 2 <= 20000
    assert not re.search(r"<!DOCTYPE|<!ENTITY|```", xml, re.I)
    root = ET.fromstring(xml)
    assert root.tag == f"{{{P}}}package"
    parts = {p.attrib[f"{{{P}}}name"]: p for p in root}
    assert len(parts) == len(root)
    assert {"/_rels/.rels", "/word/document.xml"} <= parts.keys()
    rels = {}
    for name, part in parts.items():
        assert part.tag == f"{{{P}}}part" and part.attrib[f"{{{P}}}contentType"]
        data = part.find(f"{{{P}}}xmlData")
        assert data is not None and len(data) == 1
        if name.endswith(".rels"):
            assert data[0].tag == f"{{{R}}}Relationships"
            owner = "/" if name == "/_rels/.rels" else name.replace("/_rels/", "/").removesuffix(".rels")
            rels[owner] = {}
            for rel in data[0]:
                assert rel.tag == f"{{{R}}}Relationship"
                assert rel.get("TargetMode") != "External"
                target = rel.attrib["Target"]
                assert not re.match(r"[a-z]+:", target, re.I)
                resolved = posixpath.normpath(posixpath.join(posixpath.dirname(owner), target))
                assert resolved in parts
                assert rel.attrib["Id"] not in rels[owner]
                rels[owner][rel.attrib["Id"]] = (rel.attrib["Type"], resolved)
    assert list(rels["/"].values()) == [(D + "/officeDocument", "/word/document.xml")]
    doc = parts["/word/document.xml"].find(f"{{{P}}}xmlData")[0]
    assert doc.tag == f"{{{W}}}document" and doc.find(f"{{{W}}}body") is not None
    assert not doc.findall(f".//{{{W}}}sectPr")
    for node in doc.iter():
        assert node.tag.startswith("{" + W + "}")
        for key, value in node.attrib.items():
            if key.startswith("{" + D + "}"):
                assert value in rels.get("/word/document.xml", {})
    if doc.findall(f".//{{{W}}}pStyle"):
        assert "/word/styles.xml" in parts
        assert (D + "/styles", "/word/styles.xml") in rels.get("/word/document.xml", {}).values()
    return doc

def test_static_complete_package_example():
    xml = template(); doc = lint(xml)
    assert set(p.attrib[f"{{{P}}}name"] for p in ET.fromstring(xml)) == {"/_rels/.rels", "/word/document.xml"}
    assert "Sample Bakery" in "".join(doc.itertext())

def test_source_copy_is_serialized_as_data():
    root = ET.fromstring(template())
    node = root.find(f".//{{{W}}}t")
    copy = 'A & B <Bread> "Special" > ordinary \U0001f950'
    node.text = copy
    xml = ET.tostring(root, encoding="unicode")
    assert "&amp;" in xml and "&lt;" in xml and "&gt;" in xml
    assert lint(xml).find(f".//{{{W}}}t").text == copy

@pytest.mark.parametrize("bad", ["<w:p/>", '<w:document xmlns:w="'+W+'"><w:body/></w:document>', '<p/> <p/>', '<html/>', '<!DOCTYPE x><x/>'])
def test_non_package_examples_are_rejected_by_example_lint(bad):
    with pytest.raises((AssertionError, ET.ParseError)): lint(bad)

@pytest.mark.parametrize("kind", ["missing-part", "external", "style", "section", "dangling", "length", "ampersand"])
def test_bad_example_variants(kind):
    xml = template()
    if kind == "missing-part": xml = xml.replace('Target="word/document.xml"', 'Target="word/missing.xml"')
    if kind == "external": xml = xml.replace('Target="word/document.xml"', 'Target="https://example.invalid/a" TargetMode="External"')
    if kind == "style": xml = xml.replace('<w:pPr>', '<w:pPr><w:pStyle w:val="Title"/>', 1)
    if kind == "section": xml = xml.replace('</w:body>', '<w:sectPr/></w:body>')
    if kind == "dangling": xml = xml.replace('<w:r>', '<w:r xmlns:r="'+D+'" r:id="missing">', 1)
    if kind == "length": xml = xml.replace('Sample Bakery', '\U0001f950' * 10000)
    if kind == "ampersand": xml = xml.replace('Sample Bakery', 'A & B')
    with pytest.raises((AssertionError, ET.ParseError)): lint(xml)

def test_guidance_operational_contract():
    text = REF.read_text()
    for phrase in ["COMPLETE Flat OPC", "never a bare fragment", "UTF-16", "20,000", "location", "REPLACES", "styles.xml", "single root", "sectPr"]: assert phrase in text
    if WB:
        for phrase in ["255", "5,000", "debugInfo", "elapsedMs", "roundTripMs", "newest-first", "PRECEDED"]: assert phrase in text
