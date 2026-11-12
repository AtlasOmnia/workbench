# SPDX-License-Identifier: Apache-2.0
"""The default add-in comes from the reviewed plugin tree, not a site that redeploys on every push.

The bridge itself serves the bundled ``addin/`` files over HTTPS on loopback (a per-install local CA the user trusts
once; the plugin never changes OS trust), only names listed in ``addin/SHA256SUMS`` whose sha256 matches, so the pane
and the bridge are one origin. A self-hosted ``addin_base`` keeps the strict exact-Origin check.
"""
import argparse
import hashlib
import http.client
import json
import os
import re
import socket
import ssl
import sys
from pathlib import Path

import pytest

import https_client

ROOT = Path(__file__).resolve().parent.parent
ADDIN = ROOT / "addin"
PREFIX = "/office/live/"
LOOPBACK = "https://localhost:47615/office/live/"
SELF_HOSTED = "https://addin.example.test:8443/team/office/live/"
SELF_ORIGIN = "https://addin.example.test:8443"
APPS = ("excel", "word", "powerpoint", "onenote", "outlook")


def sha(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()


def listed(site=ADDIN):
    out = {}
    for line in (site / "SHA256SUMS").read_text(encoding="utf-8").splitlines():
        digest, _, name = line.partition("  ")
        out[name] = digest
    return out


def fetch(server, path, method="GET", host=None, origin=None, headers=None, tls=None, body=None):
    c = https_client.conn(server.port, timeout=15, tls=tls)
    h = {"Host": host or f"localhost:{server.port}"}
    if origin is not None:
        h["Origin"] = origin
    h.update(headers or {})
    if body is not None:
        h["Content-Type"] = "application/json"
        h["Content-Length"] = str(len(body))
    c.request(method, path, body=body, headers=h)
    r = c.getresponse()
    data = r.read()
    out = (r.status, data, {k.lower(): v for k, v in r.getheaders()})
    c.close()
    return out


@pytest.fixture()
def loopback(bridge_mod, tls):
    """The plugin default: no allowed_origin configured, the bundled add-in served from the pinned tree."""
    s = bridge_mod.BridgeServer(ports=range(19970, 19975), tls=tls, site=bridge_mod.SITE_ROOT)
    s.ensure_started()
    yield s
    s.stop()


@pytest.fixture()
def tiny_site(tmp_path):
    site = tmp_path / "site"
    (site / "excel").mkdir(parents=True)
    page = b"<!doctype html><p>ok</p>\n"
    (site / "excel" / "taskpane.html").write_bytes(page)
    (site / "unlisted.js").write_bytes(b"export const x = 1\n")
    (site / "changed.js").write_bytes(b"export const y = 2\n")  # listed with the hash of other bytes
    outside = tmp_path / "outside.js"
    outside.write_bytes(b"export const secret = 1\n")
    entries = {"excel/taskpane.html": sha(page), "changed.js": sha(b"export const y = 1\n")}
    if hasattr(os, "symlink"):
        try:
            (site / "link.js").symlink_to(outside)
            (site / "linkdir").symlink_to(site / "excel", target_is_directory=True)
            entries["link.js"] = sha(outside.read_bytes())  # hash matches: refused for being a symlink
            entries["linkdir/taskpane.html"] = sha(page)
        except OSError:
            pass
    (site / "SHA256SUMS").write_text("".join(f"{d}  {n}\n" for n, d in sorted(entries.items())), encoding="utf-8")
    return site


@pytest.fixture()
def tiny(bridge_mod, tls, tiny_site):
    s = bridge_mod.BridgeServer(ports=range(19975, 19980), tls=tls, site=tiny_site)
    s.ensure_started()
    yield s
    s.stop()


# ---- the default is the plugin-served loopback site -----------------------------------------------------------------
def test_default_base_is_the_plugin_served_loopback_site(plugin, bridge_mod):
    assert bridge_mod.DEFAULT_ADDIN_BASE == LOOPBACK and bridge_mod.ADDIN_PATH == PREFIX
    assert bridge_mod.resolve_addin_base() == LOOPBACK
    assert "github.io" not in bridge_mod.DEFAULT_ADDIN_BASE
    assert bridge_mod.SITE_ROOT == ADDIN


def test_bundle_sums_list_every_served_file_with_its_hash_and_never_dev():
    sums = listed()
    expected = sorted(p.relative_to(ADDIN).as_posix() for sub in APPS + ("shared", "assets")
                      for p in (ADDIN / sub).rglob("*") if p.is_file()) + ["support.html"]
    assert sorted(sums) == sorted(expected)
    assert not [n for n in sums if n.startswith("dev/") or "template" in n or n in ("package.json", "SHA256SUMS")]
    for name, digest in sums.items():
        assert re.fullmatch(r"[0-9a-f]{64}", digest) and sha((ADDIN / name).read_bytes()) == digest, name
    for app in APPS:
        assert f"{app}/taskpane.html" in sums
    assert {f"assets/icon-{n}.png" for n in (16, 32, 80)} <= set(sums)


def test_bundle_is_what_build_addin_generates(plugin):
    build = plugin._build_addin()
    assert (ADDIN / "SHA256SUMS").read_text(encoding="utf-8") == build.bundle_sums_text()
    for size in (16, 32, 80):
        assert (ADDIN / "assets" / f"icon-{size}.png").read_bytes() == build.png(size)
    info = (ADDIN / "shared" / "build-info.js").read_text(encoding="utf-8")
    assert info == build.bundle_build_info_text() and '"version": "0.4.0"' in info and '"commit": null' in info


# ---- static serving: only listed files with a matching hash ---------------------------------------------------------
def test_every_listed_file_is_served_with_its_hash_and_safe_headers(loopback, bridge_mod):
    for name, digest in listed().items():
        status, body, h = fetch(loopback, PREFIX + name)
        assert status == 200 and sha(body) == digest, name
        assert h["content-type"] == bridge_mod.STATIC_MIME_TYPES[Path(name).suffix], name
        assert h["x-content-type-options"] == "nosniff" and h["content-security-policy"] == bridge_mod.STATIC_CSP
        assert "frame-ancestors" not in h["content-security-policy"]  # Office for the web frames the pane
        assert h["cache-control"] == "no-cache" and "access-control-allow-origin" not in h
    status, body, h = fetch(loopback, PREFIX + "excel/taskpane.html", method="HEAD")
    assert status == 200 and body == b"" and int(h["content-length"]) == (ADDIN / "excel/taskpane.html").stat().st_size


def test_the_office_js_mock_and_unlisted_tree_files_are_never_served(loopback):
    assert (ADDIN / "dev" / "office-mock.js").is_file()
    for name in ("dev/office-mock.js", "dev/demo-notebook.json", "manifest.template.xml", "package.json",
                 "SHA256SUMS", "../bridge_server.py", "excel/", "excel"):
        assert fetch(loopback, PREFIX + name)[0] == 404, name


def test_unlisted_changed_and_symlinked_files_are_refused(tiny, tiny_site):
    assert fetch(tiny, PREFIX + "excel/taskpane.html")[0] == 200
    assert fetch(tiny, PREFIX + "unlisted.js")[0] == 404   # in the folder, not in SHA256SUMS
    assert fetch(tiny, PREFIX + "changed.js")[0] == 404    # listed, but its bytes no longer match
    if (tiny_site / "link.js").is_symlink():
        assert fetch(tiny, PREFIX + "link.js")[0] == 404            # a symlinked file, even with a matching hash
        assert fetch(tiny, PREFIX + "linkdir/taskpane.html")[0] == 404  # a symlinked folder on the way
    (tiny_site / "excel" / "taskpane.html").write_bytes(b"tampered\n")
    assert fetch(tiny, PREFIX + "excel/taskpane.html")[0] == 404  # checked on every request, not cached


@pytest.mark.parametrize("path", ["../bridge_server.py", "%2e%2e/bridge_server.py", "excel/../../bridge_server.py",
                                  "excel%2ftaskpane.html", "excel%2Ftaskpane.html", "excel%5ctaskpane.html",
                                  "excel\\taskpane.html", "excel/taskpane.html%00", "%00", "/excel/taskpane.html",
                                  "excel//taskpane.html", "", "%ff", "dev%2foffice-mock.js", "./excel/taskpane.html"])
def test_traversal_and_encoded_separators_are_refused(loopback, path):
    assert fetch(loopback, PREFIX + path)[0] == 404, path


def test_static_files_answer_get_and_head_only(loopback):
    for method in ("POST", "PUT", "DELETE", "OPTIONS", "PATCH"):
        status, _, h = fetch(loopback, PREFIX + "excel/taskpane.html", method=method, body=b"{}")
        assert status == 405, method


def test_static_host_must_be_localhost_and_the_bound_port(loopback):
    for host in (f"127.0.0.1:{loopback.port}", f"localhost:{loopback.port + 1}", "localhost", "evil.example",
                 f"LOCALHOST:{loopback.port}", f"localhost.evil:{loopback.port}"):
        assert fetch(loopback, PREFIX + "excel/taskpane.html", host=host)[0] == 421, host


def test_static_foreign_origin_is_refused(loopback):
    page = PREFIX + "excel/taskpane.html"
    assert fetch(loopback, page)[0] == 200  # an iframe navigation carries no Origin
    assert fetch(loopback, page, origin=f"https://localhost:{loopback.port}")[0] == 200
    for origin in ("https://evil.example", f"http://localhost:{loopback.port}", "null", "https://localhost:9"):
        assert fetch(loopback, page, origin=origin)[0] == 403, origin


# ---- the bridge API on the same origin -----------------------------------------------------------------------------
def test_same_origin_pane_reaches_the_api_and_foreign_origins_do_not(loopback):
    own = f"https://localhost:{loopback.port}"
    status, body, h = fetch(loopback, "/v1/hello")  # same-origin GET: browsers send no Origin
    assert status == 200 and json.loads(body)["service"] == "hermes-office-desktop"
    assert fetch(loopback, "/v1/hello", headers={"Sec-Fetch-Site": "same-origin"})[0] == 200
    assert fetch(loopback, "/v1/hello", origin=own)[0] == 200
    for site in ("cross-site", "same-site", "none"):
        assert fetch(loopback, "/v1/hello", headers={"Sec-Fetch-Site": site})[0] == 403, site
    for origin in ("https://evil.example", f"http://localhost:{loopback.port}", "https://localhost:9", "null",
                   "https://atlasomnia.github.io"):
        assert fetch(loopback, "/v1/hello", origin=origin)[0] == 403, origin
    assert fetch(loopback, "/v1/pair", method="POST", body=b'{"code":"x","app":"excel"}')[0] == 403  # POST: no Origin
    assert fetch(loopback, "/v1/hello", host=f"127.0.0.1:{loopback.port}", origin=own)[0] == 421


def test_a_pane_from_another_workbench_port_reaches_this_bridge(loopback):
    """A14/profiles: the pane looks for its bridge on all five ports; each is a Workbench bridge serving the same
    hash-checked files. Only those five loopback origins are accepted, nothing else on localhost."""
    sibling = f"https://localhost:{loopback.ports[-1]}"
    status, _, h = fetch(loopback, "/v1/hello", origin=sibling)
    assert status == 200 and h["access-control-allow-origin"] == sibling and h["vary"] == "Origin"
    code, _ = loopback.registry.mint_code("chat-1", "excel")
    status, body, _ = fetch(loopback, "/v1/pair", method="POST", origin=sibling,
                            body=json.dumps({"code": code, "app": "excel"}).encode())
    assert status == 200 and json.loads(body)["ok"]
    assert loopback.allowed_origins == frozenset(f"https://localhost:{p}" for p in loopback.ports)


def test_plain_http_gets_no_bridge(loopback):
    c = http.client.HTTPConnection("127.0.0.1", loopback.port, timeout=10)
    try:
        c.request("GET", "/v1/hello", headers={"Host": f"127.0.0.1:{loopback.port}", "Origin": "https://evil.example"})
        status = c.getresponse().status
    except (OSError, http.client.HTTPException):
        status = None
    finally:
        c.close()
    assert status != 200


def test_a_certificate_from_another_ca_is_not_trusted(loopback, tmp_path, tls_mod):
    other = tls_mod.LocalTLS(tmp_path / "other-profile")
    other.ensure()
    with pytest.raises(ssl.SSLError):
        fetch(loopback, "/v1/hello", tls=other)


# ---- a self-hosted site keeps the strict exact-Origin check --------------------------------------------------------
def test_self_hosted_base_keeps_the_exact_origin_and_serves_no_site(bridge_mod, tls):
    s = bridge_mod.BridgeServer(allowed_origin=SELF_ORIGIN, ports=range(19980, 19985), tls=tls)
    s.ensure_started()
    try:
        status, _, h = fetch(s, "/v1/hello", origin=SELF_ORIGIN)
        assert status == 200 and h["access-control-allow-origin"] == SELF_ORIGIN
        for origin in (None, f"https://localhost:{s.port}", f"https://localhost:{s.ports[1]}", "https://evil.example"):
            assert fetch(s, "/v1/hello", origin=origin)[0] == 403, origin
        assert fetch(s, PREFIX + "excel/taskpane.html")[0] == 404
        assert s.allowed_origins == frozenset({SELF_ORIGIN})
    finally:
        s.stop()


# ---- the plugin: manifests for the bound port, trust step, manifest command ---------------------------------------
class Ctx:
    def __init__(self, config=None):
        self.tools, self.hooks, self.commands, self.cli = {}, {}, {}, {}
        self.config = config or {}

    def register_tool(self, **kw):
        self.tools[kw["name"]] = kw

    def register_hook(self, name, fn):
        self.hooks[name] = fn

    def register_skill(self, name, path, **metadata):
        pass

    def register_command(self, name, handler, description="", args_hint=""):
        self.commands[name] = handler

    def register_cli_command(self, name, help, setup_fn, handler_fn=None, description=""):
        self.cli[name] = (setup_fn, handler_fn)

    def get_config(self, key, default=None):
        return self.config.get(key, default)


@pytest.fixture(autouse=True)
def _attended(monkeypatch):
    monkeypatch.delenv("HERMES_CRON_SESSION", raising=False)
    monkeypatch.delenv("HERMES_SESSION_PLATFORM", raising=False)


@pytest.fixture
def make(plugin):
    def _make(config=None, port=19985):
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


def _urls(xml):
    return re.findall(r'(?:DefaultValue="|<AppDomain>)(https://[^"<]+)', xml)


def test_default_manifests_point_every_url_at_the_bound_port(make, manifest_home, plugin):
    blocker = socket.socket()
    blocker.bind(("127.0.0.1", 19985))
    blocker.listen()
    try:
        ctx = make()
        out = json.loads(ctx.tools["office_connect"]["handler"]({"app": "excel"}, session_id="chat-1"))
        assert out["ok"], out
    finally:
        blocker.close()
    port = plugin._server.port
    assert port == 19986  # the first port was taken: the base follows the port actually bound
    base = f"https://localhost:{port}/office/live/"
    for app in APPS:
        xml = (manifest_home / f"manifest-{app}.xml").read_text(encoding="utf-8")
        urls = _urls(xml)
        assert urls and all(u.startswith(base) or u == f"https://localhost:{port}" for u in urls), (app, urls)
        assert f'<SourceLocation DefaultValue="{base}{app}/taskpane.html"/>' in xml
    d = out["data"]
    assert d["addinBase"] == base
    trust = d["trust"]
    assert Path(trust["caCertificate"]).is_file() and trust["caCertificate"] in trust["command"]
    assert any("certificate authority" in step and trust["command"] in step for step in d["firstTimeSetup"])
    status, body, _ = fetch(plugin._server, "/office/live/excel/taskpane.html")
    assert status == 200 and sha(body) == listed()["excel/taskpane.html"]


def test_self_hosted_addin_base_still_works_through_the_plugin(make, manifest_home, plugin):
    ctx = make({"addin_base": SELF_HOSTED})
    out = json.loads(ctx.tools["office_connect"]["handler"]({"app": "word"}, session_id="chat-1"))
    assert out["ok"] and out["data"]["addinBase"] == SELF_HOSTED and "trust" in out["data"]
    xml = (manifest_home / "manifest-word.xml").read_text(encoding="utf-8")
    assert all(u.startswith(SELF_HOSTED) or u == SELF_ORIGIN for u in _urls(xml))
    assert plugin._server.allowed_origins == frozenset({SELF_ORIGIN})
    assert fetch(plugin._server, "/v1/hello", origin=SELF_ORIGIN)[0] == 200
    assert fetch(plugin._server, "/v1/hello")[0] == 403


def test_workbench_manifest_command_writes_manifests_for_the_running_bridge(make, manifest_home, plugin):
    ctx = make(port=19990)
    slash = ctx.commands["workbench"]
    text = slash("manifest")
    port = plugin._server.port
    assert port == 19990 and str(manifest_home / "manifest-excel.xml") in text
    assert f"https://localhost:{port}/office/live/" in (manifest_home / "manifest-onenote.xml").read_text()
    assert "add-trusted-cert" in text or "certutil" in text or "trust" in text.lower()
    trust = slash("trust")
    assert str(plugin.local_tls(ctx).ca_path) in trust and "never" in trust
    assert "manifest" in slash("help")


def test_cli_manifest_probes_the_running_bridge_and_binds_nothing(make, manifest_home, plugin, bridge_mod):
    ctx = make(port=19990)
    setup, run = ctx.cli["workbench"]
    parser = argparse.ArgumentParser()
    setup(parser)
    ports = ctx._test_ports
    assert run(parser.parse_args(["manifest"])) == 1  # no bridge running: says so, starts nothing
    assert bridge_mod.busy_ports(ports) == [] and (plugin._server is None or plugin._server.port is None)
    json.loads(ctx.tools["office_connect"]["handler"]({"app": "excel"}, session_id="chat-1"))
    running, plugin._server = plugin._server, None  # a separate `hermes workbench` process has no bridge of its own
    try:
        (manifest_home / "manifest-excel.xml").unlink()
        assert run(parser.parse_args(["manifest"])) == 0
        assert plugin._server is None and bridge_mod.busy_ports(ports) == [running.port]
        assert f"https://localhost:{running.port}/office/live/" in (manifest_home / "manifest-excel.xml").read_text()
        assert plugin.find_running_port(ctx) == running.port
    finally:
        plugin._server = running


# ---- the local CA: per install, never trusted by the plugin itself -------------------------------------------------
@pytest.fixture(scope="session")
def tls_mod(plugin):
    return sys.modules["hermes_office_desktop.bridge_tls"]


def test_local_ca_is_name_constrained_owner_only_and_reused(tls_mod, tmp_path):
    from cryptography import x509
    t = tls_mod.LocalTLS(tmp_path)
    t.ensure()
    first = t.ca_path.read_bytes()
    t.ensure()
    assert t.ca_path.read_bytes() == first  # one CA per install: the user trusts it once
    ca = x509.load_pem_x509_certificate(first)
    nc = ca.extensions.get_extension_for_class(x509.NameConstraints)
    assert nc.critical and {str(n.value) for n in nc.value.permitted_subtrees} == {"localhost", "127.0.0.1/32"}
    assert ca.extensions.get_extension_for_class(x509.BasicConstraints).value.path_length == 0
    leaf = x509.load_pem_x509_certificate(t.cert_path.read_bytes())
    assert leaf.extensions.get_extension_for_class(x509.SubjectAlternativeName).value.get_values_for_type(
        x509.DNSName) == ["localhost"]
    if os.name != "nt":
        assert t.directory.stat().st_mode & 0o777 == 0o700
        assert all(p.stat().st_mode & 0o777 == 0o600 for p in (t.ca_path, t.cert_path, t.key_path, t.ca_key_path))
    mac = t.trust_instructions("darwin")
    assert "security add-trusted-cert" in mac and str(t.ca_path) in mac and "never" in mac
    assert "certutil -user -addstore Root" in t.trust_instructions("win32")


def test_the_plugin_never_changes_os_trust_itself():
    src = (ROOT / "bridge_tls.py").read_text(encoding="utf-8")
    assert "subprocess" not in src and "os.system" not in src and "popen" not in src.lower()
    for f in ("__init__.py", "bridge_server.py"):
        text = (ROOT / f).read_text(encoding="utf-8")
        assert "add-trusted-cert" not in text and "certutil" not in text, f


def test_cryptography_is_a_declared_dependency():
    import yaml
    manifest = yaml.safe_load((ROOT / "plugin.yaml").read_text(encoding="utf-8"))
    assert "cryptography>=43" in manifest["python_dependencies"]
    assert manifest["config_schema"]["addin_base"]["default"] == LOOPBACK


def test_pages_workflow_is_manual_only():
    from ruamel.yaml import YAML
    wf = YAML(typ="safe").load((ROOT / ".github" / "workflows" / "pages.yml").read_text(encoding="utf-8"))
    on = wf.get("on", wf.get(True))
    assert set(on) == {"workflow_dispatch"}


def test_addin_files_are_checked_out_as_lf_on_every_os():
    """Git for Windows (core.autocrlf=true) would otherwise write CRLF: every text file's sha256 would differ from
    SHA256SUMS and the bridge would refuse the whole add-in (404)."""
    rules = (ROOT / ".gitattributes").read_text(encoding="utf-8").splitlines()
    assert "addin/** text=auto eol=lf" in rules


def test_a_crlf_checkout_is_served_as_the_pinned_lf_bytes_and_nothing_else(tiny, tiny_site):
    """Git for Windows can leave text files with CRLF (core.autocrlf; a changed .gitattributes does not rewrite
    files git has already checked out). The bridge then serves the LF form, and only when that form is exactly the
    pinned content: what reaches the pane always matches SHA256SUMS."""
    page = b"<!doctype html><p>ok</p>\n"
    (tiny_site / "excel" / "taskpane.html").write_bytes(page.replace(b"\n", b"\r\n"))
    status, body, _ = fetch(tiny, PREFIX + "excel/taskpane.html")
    assert status == 200 and body == page and sha(body) == listed(tiny_site)["excel/taskpane.html"]
    (tiny_site / "excel" / "taskpane.html").write_bytes(b"<!doctype html><p>tampered</p>\r\n")
    assert fetch(tiny, PREFIX + "excel/taskpane.html")[0] == 404
    (tiny_site / "changed.js").write_bytes(b"export const y = 2\r\n")  # CRLF form of bytes that were never pinned
    assert fetch(tiny, PREFIX + "changed.js")[0] == 404
