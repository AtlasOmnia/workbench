# SPDX-License-Identifier: Apache-2.0
"""F2: the add-in origin has one source of truth (since A24 the add-in base, addin_base / allowed_origin), and office_connect writes the five
manifests for that origin into a user-visible folder instead of telling the user to download them."""
import hashlib
import importlib.util
import json
import os
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parent.parent
OLD_HOST = "hermesagent" + "tools"  # the retired default host (A24): never in a manifest, setup step or doc
ORIGIN = "https://studio.example.test"
BASE = ORIGIN + "/office/live/"
APPS = ("excel", "word", "powerpoint", "onenote", "outlook")
SENTENCE = ("{start} > My Add-ins > Manage My Add-ins > Upload My Add-in > Browse, choose manifest-{app}.xml in "
            "Documents > Omnia Works > add-in manifests, then Upload.")


def sentence(app):
    return SENTENCE.format(app=app, start="Insert > Office Add-ins" if app == "onenote" else "Home > Add-ins > More Add-ins")


def _build_addin():
    spec = importlib.util.spec_from_file_location("hod_build_addin_test", ROOT / "scripts" / "build_addin.py")
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    return mod


class Ctx:
    def __init__(self, config):
        self.tools, self.hooks, self.config = {}, {}, config

    def register_tool(self, **kw):
        self.tools[kw["name"]] = kw

    def register_hook(self, name, fn):
        self.hooks[name] = fn

    def register_skill(self, name, path, **metadata):
        pass

    def get_config(self, key, default=None):
        return self.config.get(key, default)


@pytest.fixture(autouse=True)
def _attended(monkeypatch):
    monkeypatch.delenv("HERMES_CRON_SESSION", raising=False)
    monkeypatch.delenv("HERMES_SESSION_PLATFORM", raising=False)


@pytest.fixture
def make(plugin):
    made = []

    def _make(config, port=19900):
        if plugin._server:
            plugin._server.stop()
        plugin._server = None
        ctx = Ctx(config)
        ctx._test_ports = tuple(range(port, port + 5))
        plugin.register(ctx)
        made.append(ctx)
        return ctx
    yield _make
    if plugin._server:
        plugin._server.stop()
    plugin._server = None


def connect(ctx, app="excel"):
    out = json.loads(ctx.tools["office_connect"]["handler"]({"app": app}, session_id="chat-1"))
    assert out["ok"], out
    return out["data"]


def test_default_folder_gets_all_five_manifests_for_the_configured_origin(make, manifest_home):
    d = connect(make({"allowed_origin": ORIGIN}))
    folder = manifest_home
    expected = _build_addin().render_manifests(BASE)
    assert sorted(expected) == sorted(APPS)
    for app in APPS:
        assert (folder / f"manifest-{app}.xml").read_text(encoding="utf-8") == expected[app]
    assert d["manifestPath"] == str(folder / "manifest-excel.xml") and Path(d["manifestPath"]).is_absolute()
    text = json.dumps(d)
    assert OLD_HOST not in text and "download" not in text.lower()
    assert sentence("excel") in d["firstTimeSetup"]


def test_every_app_names_its_own_manifest(make, manifest_home):
    ctx = make({"allowed_origin": ORIGIN})
    for app in ("word", "powerpoint", "onenote"):
        d = connect(ctx, app)
        assert sentence(app) in d["firstTimeSetup"]
        assert d["manifestPath"] == str(manifest_home / f"manifest-{app}.xml")
    d = connect(ctx, "outlook")
    setup = " ".join(d["firstTimeSetup"])
    assert "manifest-outlook.xml in Documents > Omnia Works > add-in manifests" in setup and "Custom Addins" in setup
    assert OLD_HOST not in json.dumps(d) and d["manifestPath"].endswith("manifest-outlook.xml")


def test_default_base_is_the_plugin_served_site(make, manifest_home):
    connect(make({}))
    assert "https://localhost:" in (manifest_home / "manifest-excel.xml").read_text(encoding="utf-8")
    assert "/office/live/excel/taskpane.html" in \
        (manifest_home / "manifest-excel.xml").read_text(encoding="utf-8")


def test_idempotent_no_rewrite(make, manifest_home):
    ctx = make({"allowed_origin": ORIGIN})
    connect(ctx)
    before = {p.name: p.stat().st_mtime_ns for p in manifest_home.iterdir()}
    os.utime(manifest_home / "manifest-word.xml", ns=(10**9, 10**9))  # 1 s: NTFS keeps 100 ns ticks (1 ns reads as 0)
    connect(ctx)
    assert (manifest_home / "manifest-word.xml").stat().st_mtime_ns == 10**9  # unchanged content is not rewritten
    after = {p.name: p.stat().st_mtime_ns for p in manifest_home.iterdir()}
    assert set(after) == set(before)


def test_user_changed_file_is_never_overwritten(make, manifest_home):
    ctx = make({"allowed_origin": ORIGIN})
    connect(ctx)
    mine = manifest_home / "manifest-excel.xml"
    mine.write_text("<!-- my edit -->", encoding="utf-8")
    d = connect(ctx)
    assert mine.read_text(encoding="utf-8") == "<!-- my edit -->"
    alt = manifest_home / "manifest-excel (Omnia Works).xml"
    assert alt.read_text(encoding="utf-8") == _build_addin().render_manifests(BASE)["excel"]
    assert d["manifestPath"] == str(alt)
    assert "choose manifest-excel (Omnia Works).xml in Documents > Omnia Works > add-in manifests" in \
        " ".join(d["firstTimeSetup"])


def test_origin_change_replaces_plugin_written_files(make, manifest_home):
    connect(make({"allowed_origin": "https://old.example.test"}))
    d = connect(make({"allowed_origin": ORIGIN}, port=19910))
    assert d["manifestPath"] == str(manifest_home / "manifest-excel.xml")
    assert (manifest_home / "manifest-excel.xml").read_text(encoding="utf-8") == \
        _build_addin().render_manifests(BASE)["excel"]
    assert not (manifest_home / "manifest-excel (Omnia Works).xml").exists()


def test_configured_manifest_dir_with_tilde(make, tmp_path, monkeypatch):
    monkeypatch.setenv("HOME", str(tmp_path))
    monkeypatch.setenv("USERPROFILE", str(tmp_path))  # Windows expands ~ from USERPROFILE, not HOME
    d = connect(make({"allowed_origin": ORIGIN, "manifest_dir": "~/Desk/manifests"}))
    want = tmp_path / "Desk" / "manifests" / "manifest-excel.xml"
    assert d["manifestPath"] == str(want) and want.is_file()
    assert f"choose manifest-excel.xml in {want.parent}, then Upload." in " ".join(d["firstTimeSetup"])


def test_unwritable_folder_falls_back_to_the_configured_origin_url(make, tmp_path):
    blocker = tmp_path / "a-file"
    blocker.write_text("x")
    d = connect(make({"allowed_origin": ORIGIN, "manifest_dir": str(blocker / "sub")}))
    assert "manifestPath" not in d
    setup = " ".join(d["firstTimeSetup"])
    assert f"{BASE}manifest-excel.xml" in setup and OLD_HOST not in json.dumps(d)


def test_atomic_writes_leave_no_temp_files(make, manifest_home):
    connect(make({"allowed_origin": ORIGIN}))
    names = sorted(p.name for p in manifest_home.iterdir() if not p.name.startswith("."))
    assert names == sorted(f"manifest-{a}.xml" for a in APPS)


def test_plugin_reuses_build_addin_instead_of_forking_it():
    src = (ROOT / "__init__.py").read_text(encoding="utf-8")
    assert "render_manifests" in src and "{{BASE}}" not in src
    assert OLD_HOST not in src


def test_image_user_agent_derives_from_the_configured_origin(plugin, make, monkeypatch):
    images = __import__("sys").modules["hermes_office_desktop.images"]
    assert images.user_agent(BASE) == \
        f"HermesOfficeDesktop (office_insert_image; +{ORIGIN}/office/live/support.html)"
    seen = {}

    def fake(url, user_agent=None):
        seen["ua"] = user_agent
        raise images.ImageError("Network", "stop here")
    monkeypatch.setattr(images, "fetch_url", fake)
    ctx = make({"allowed_origin": ORIGIN})
    with pytest.raises(images.ImageError):
        plugin._image_payload(ctx, "excel", "url", {"url": "https://img.example/a.png"}, "a")
    assert seen.pop("ua") == images.user_agent(BASE)

    def fake_sanitize(raw, target, alt, fetch=None):
        fetch("https://img.example/b.png")
    monkeypatch.setattr(images, "sanitize_html", fake_sanitize)
    with pytest.raises(images.ImageError):
        plugin._image_payload(ctx, "word", "html", {"html": "<img src='https://img.example/b.png'>"}, "a")
    assert seen.pop("ua") == images.user_agent(BASE)


def test_skill_and_readme_relay_the_local_path_not_the_website():
    skill = (ROOT / "skills" / "office-desktop" / "SKILL.md").read_text(encoding="utf-8")
    assert "manifestPath" in skill and OLD_HOST not in skill
    readme = (ROOT / "README.md").read_text(encoding="utf-8")
    connect_sec = readme.split("## First connect (about 2 minutes per app, once)", 1)[1].split("\n## ", 1)[0]
    assert "Documents > Omnia Works > add-in manifests" in connect_sec and OLD_HOST not in connect_sec
