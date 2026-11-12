# SPDX-License-Identifier: Apache-2.0
"""F3 (A24): ONE source of truth for the add-in BASE URL.

Setting ``addin_base`` (default the plugin-served loopback site) gives the bridge's allowed origin (scheme + host[:port]) and every
manifest / task-pane / icon / support URL. The older ``allowed_origin`` setting still works: an origin alone means
base = origin + "/office/live/" (the private test profiles use exactly that).
"""
import importlib.util
import json
import re
import subprocess
import sys
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parent.parent
PAGES = "https://atlasomnia.github.io/workbench/office/live/"  # the optional, manual-only Pages site
LOOPBACK = "https://localhost:47615/office/live/"  # the default: the bundled add-in, served by the bridge
LOOPBACK_ORIGIN = "https://localhost:47615"
STUDIO = "https://studio.example.test"
PREFIXED = "https://addin.example.test:8443/team/tools/office/live/"
SLUG = "AtlasOmnia/workbench"
OLD_HOST = "hermesagent" + "tools"  # the retired default host; must appear nowhere in the repo


def _build_addin():
    spec = importlib.util.spec_from_file_location("hod_build_addin_base", ROOT / "scripts" / "build_addin.py")
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
    def _make(config, port=19940):
        if plugin._server:
            plugin._server.stop()
        plugin._server = None
        ctx = Ctx(config)
        ctx._test_ports = tuple(range(port, port + 5))
        plugin.register(ctx)
        return ctx
    yield _make
    if plugin._server:
        plugin._server.stop()
    plugin._server = None


def connect(ctx, app="excel"):
    out = json.loads(ctx.tools["office_connect"]["handler"]({"app": app}, session_id="chat-1"))
    assert out["ok"], out
    return out["data"]


# ---- resolution ----------------------------------------------------------------------------------------------------
def test_default_base_is_the_loopback_site_and_the_origin_derives_from_it(plugin, bridge_mod):
    assert bridge_mod.DEFAULT_ADDIN_BASE == LOOPBACK
    assert bridge_mod.DEFAULT_ORIGIN == LOOPBACK_ORIGIN
    assert bridge_mod.resolve_addin_base() == LOOPBACK
    assert plugin.addin_base(Ctx({})) == LOOPBACK and plugin.addin_origin(Ctx({})) == LOOPBACK_ORIGIN


def test_origin_alone_keeps_working(plugin, bridge_mod):
    assert bridge_mod.resolve_addin_base(allowed_origin=STUDIO) == STUDIO + "/office/live/"
    ctx = Ctx({"allowed_origin": STUDIO})
    assert plugin.addin_base(ctx) == STUDIO + "/office/live/" and plugin.addin_origin(ctx) == STUDIO


def test_path_prefixed_base_gives_scheme_host_port_origin(plugin, bridge_mod):
    assert bridge_mod.origin_of(PREFIXED) == "https://addin.example.test:8443"
    ctx = Ctx({"addin_base": PREFIXED})
    assert plugin.addin_base(ctx) == PREFIXED and plugin.addin_origin(ctx) == "https://addin.example.test:8443"
    # an explicit base and a matching origin agree; a mismatch is a configuration error, never a silent pick
    assert plugin.addin_base(Ctx({"addin_base": PREFIXED, "allowed_origin": "https://addin.example.test:8443"})) == PREFIXED
    with pytest.raises(bridge_mod.BridgeError) as e:
        plugin.addin_base(Ctx({"addin_base": PREFIXED, "allowed_origin": STUDIO}))
    assert e.value.category == "Configuration"
    # the default value written out explicitly does not override an older allowed_origin profile
    assert plugin.addin_base(Ctx({"addin_base": LOOPBACK, "allowed_origin": STUDIO})) == STUDIO + "/office/live/"


@pytest.mark.parametrize("bad", ["http://x.test/office/live/", "https://x.test/office/live", "https://x.test/a/?q=1",
                                 "https://x.test/a/#f", "https://user@x.test/a/", "office/live/", "https:///a/"])
def test_bad_bases_are_refused(bridge_mod, bad):
    with pytest.raises(bridge_mod.BridgeError):
        bridge_mod.resolve_addin_base(addin_base=bad)


@pytest.mark.parametrize("bad", ["http://x.test", "https://x.test/office", "*"])
def test_bad_origins_are_refused(bridge_mod, bad):
    with pytest.raises(bridge_mod.BridgeError):
        bridge_mod.resolve_addin_base(allowed_origin=bad)


# ---- bridge + manifests from the one base --------------------------------------------------------------------------
def test_bridge_origin_check_uses_the_derived_origin(plugin, make):
    ctx = make({"addin_base": PREFIXED})
    assert plugin.get_server(ctx).allowed_origin == "https://addin.example.test:8443"
    ctx = make({}, port=19950)  # the default: the five loopback origins of its own ports, nothing else
    server = plugin.get_server(ctx)
    assert server.allowed_origin is None and server.loopback
    assert server.allowed_origins == frozenset(f"https://localhost:{p}" for p in range(19950, 19955))


def test_default_profile_manifests_point_at_the_bridge_itself(make, manifest_home):
    d = connect(make({}))
    xml = (manifest_home / "manifest-excel.xml").read_text(encoding="utf-8")
    base = "https://localhost:19940/office/live/"  # the port this bridge bound (make's first test port)
    assert d["addinBase"] == base and f'<SourceLocation DefaultValue="{base}excel/taskpane.html"/>' in xml
    assert "<AppDomain>https://localhost:19940</AppDomain>" in xml and OLD_HOST not in xml and "github.io" not in xml


def test_prefixed_base_manifests_written_by_office_connect(make, manifest_home):
    d = connect(make({"addin_base": PREFIXED}), "outlook")
    xml = (manifest_home / "manifest-outlook.xml").read_text(encoding="utf-8")
    assert f"{PREFIXED}outlook/taskpane.html" in xml and f"{PREFIXED}assets/icon-80.png" in xml
    assert f'<SupportUrl DefaultValue="{PREFIXED}support.html"/>' in xml
    assert d["manifestPath"].endswith("manifest-outlook.xml")


def _urls(xml):
    return re.findall(r'(?:DefaultValue="|<AppDomain>)(https://[^"<]+)', xml)


@pytest.mark.parametrize("base", [PAGES, PREFIXED])
def test_build_addin_output_for_path_prefixed_bases(tmp_path, base):
    out = tmp_path / "site"
    r = subprocess.run([sys.executable, str(ROOT / "scripts" / "build_addin.py"), "--base", base, "--out", str(out)],
                       capture_output=True, text=True)
    assert r.returncode == 0, r.stderr
    origin = re.match(r"https://[^/]+", base).group(0)
    for app in ("excel", "word", "powerpoint", "onenote", "outlook"):
        xml = (out / f"manifest-{app}.xml").read_text(encoding="utf-8")
        urls = _urls(xml)
        assert urls and all(u.startswith(base) or u == origin for u in urls), urls
        assert f'<SourceLocation DefaultValue="{base}{app}/taskpane.html"/>' in xml
        for size in (16, 32, 80):
            assert f"{base}assets/icon-{size}.png" in xml
            assert (out / "assets" / f"icon-{size}.png").is_file()
        assert f'<SupportUrl DefaultValue="{base}support.html"/>' in xml
        assert f"<AppDomain>{origin}</AppDomain>" in xml
        assert (out / app / "taskpane.html").is_file()
    assert (out / "support.html").is_file()


def test_build_addin_default_base_is_the_bridge_default():
    import ast
    assert _build_addin().DEFAULT_BASE == LOOPBACK
    tree = ast.parse((ROOT / "scripts" / "build_addin.py").read_text(encoding="utf-8"))
    literals = [n.value for n in ast.walk(tree) if isinstance(n, ast.Constant) and isinstance(n.value, str)]
    assert LOOPBACK not in literals  # loaded from bridge_server.py (one source), not a second literal


def test_image_user_agent_names_the_base_support_page(plugin, bridge_mod):
    images = sys.modules["hermes_office_desktop.images"]
    assert images.user_agent(PREFIXED) == f"HermesOfficeDesktop (office_insert_image; +{PREFIXED}support.html)"
    # the loopback default's support page is unreachable from an image host: the public repository instead
    assert images.USER_AGENT == f"HermesOfficeDesktop (office_insert_image; +https://github.com/{SLUG})"


# ---- repo: no retired host, one slug, the optional Pages workflow ------------------------------------------------------------
def _tracked():
    """Every file of this tree: git's list in a checkout; every file (no caches) in an export (git archive), so the
    check never passes vacuously on an export inside some other repository."""
    top = subprocess.run(["git", "-C", str(ROOT), "rev-parse", "--show-toplevel"], capture_output=True, text=True)
    if top.returncode == 0 and Path(top.stdout.strip()).resolve() == ROOT.resolve():
        out = subprocess.run(["git", "-C", str(ROOT), "ls-files", "-z"], capture_output=True, check=True).stdout
        files = [ROOT / p for p in out.decode().split("\0") if p]
    else:
        skip = {".git", "__pycache__", "node_modules", "dist", ".pytest_cache"}
        files = [p for p in ROOT.rglob("*") if p.is_file() and not skip & set(p.relative_to(ROOT).parts)]
    assert len(files) > 50  # a real tree, not an empty listing
    return files


def test_the_retired_host_appears_nowhere_in_the_repo():
    hits = []
    for f in _tracked():
        if f.is_file():
            try:
                if OLD_HOST in f.read_text(encoding="utf-8"):
                    hits.append(str(f.relative_to(ROOT)))
            except UnicodeDecodeError:
                continue
    assert hits == []


def test_one_repo_slug_everywhere():
    readme = (ROOT / "README.md").read_text(encoding="utf-8")
    assert f"hermes plugins install https://github.com/{SLUG}" in readme
    assert f"https://github.com/{SLUG}/actions/workflows/tests.yml/badge.svg" in readme
    cat = (ROOT / "catalog" / "hermes-office-desktop.yaml").read_text(encoding="utf-8")
    assert f"\nrepo: https://github.com/{SLUG}\n" in cat and f"\ndocs_url: https://github.com/{SLUG}#readme\n" in cat
    others = []
    for f in _tracked():
        if f.suffix in (".md", ".yaml", ".yml", ".py", ".json", ".html", ".js") and f.is_file():
            for m in re.findall(r"github\.com/AtlasOmnia/[A-Za-z0-9._-]+", f.read_text(encoding="utf-8")):
                if m != f"github.com/{SLUG}":
                    others.append((str(f.relative_to(ROOT)), m))
    assert others == []
    assert (ROOT / "plugin.yaml").read_text(encoding="utf-8").startswith("# SPDX-License-Identifier: Apache-2.0\nname: hermes-office-desktop\n")


def test_pages_workflow_builds_the_optional_pages_base_by_hand_with_pinned_actions():
    from ruamel.yaml import YAML
    wf = YAML(typ="safe").load((ROOT / ".github" / "workflows" / "pages.yml").read_text(encoding="utf-8"))
    on = wf.get("on", wf.get(True))
    assert set(on) == {"workflow_dispatch"}  # never redeployed on push: the default add-in is the plugin's own
    steps = [s for job in wf["jobs"].values() for s in job.get("steps", [])]
    uses = [s["uses"] for s in steps if "uses" in s]
    assert uses and all(re.fullmatch(r"[\w.-]+/[\w.-]+@[0-9a-f]{40}", u) for u in uses), uses
    assert {u.split("@")[0] for u in uses} >= {"actions/checkout", "actions/upload-pages-artifact", "actions/deploy-pages"}
    runs = " ".join(s.get("run", "") for s in steps)
    assert f"scripts/build_addin.py --base {PAGES}" in runs
    assert wf["permissions"] == {"contents": "read", "pages": "write", "id-token": "write"}
