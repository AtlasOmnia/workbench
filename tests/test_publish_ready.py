# SPDX-License-Identifier: Apache-2.0
"""Publish-ready names are not claims of a public-origin acceptance run."""
from pathlib import Path

import yaml

ROOT = Path(__file__).resolve().parent.parent


def test_catalog_and_manifest_keep_identity_and_enable_both_platforms():
    manifest = yaml.safe_load((ROOT / "plugin.yaml").read_text())
    catalog = yaml.safe_load((ROOT / "catalog/hermes-office-desktop.yaml").read_text())
    assert manifest["name"] == catalog["name"] == "hermes-office-desktop"
    assert manifest["platforms"] == catalog["platforms"] == ["macos", "windows"]
    assert catalog["repo"] == "https://github.com/AtlasOmnia/workbench"
    assert catalog["sha"] == "TODO"
    assert "filled by the publish runbook" in (ROOT / "catalog/hermes-office-desktop.yaml").read_text()
    assert len(catalog["description"]) < 2500
    assert "real Hermes Desktop sessions" in catalog["description"]
    assert any("two Hermes profiles at once are not live-tested" in x for x in catalog["known_issues"])


def test_public_origin_smoke_is_explicitly_pending_and_complete():
    readme = (ROOT / "README.md").read_text()
    evidence = (ROOT / "docs/live-evidence.md").read_text()
    sentence = ("Live runs used a private add-in site; the plugin-served add-in (the default since 0.4.0) is not "
                "live-tested yet")
    assert sentence in readme and sentence in evidence
    smoke = evidence.split("### Post-publish smoke test", 1)[1].split("## F7", 1)[0]
    for needle in ("clean Hermes profile", "leave `addin_base` and `allowed_origin` unset", "connect Excel",
                   "person type", "one cell write", "`office_undo`", "not run yet", "/workbench trust",
                   "https://localhost:<port>/office/live/"):
        assert needle in smoke
    assert 'folder name "Omnia Works"' in readme and "before the rename" in readme


def test_legacy_origin_override_survives_default_slug_change(bridge_mod):
    assert bridge_mod.resolve_addin_base(allowed_origin="https://studio.example.test") == \
        "https://studio.example.test/office/live/"
    assert bridge_mod.resolve_addin_base(bridge_mod.DEFAULT_ADDIN_BASE, "https://studio.example.test") == \
        "https://studio.example.test/office/live/"
