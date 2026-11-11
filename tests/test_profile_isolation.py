# SPDX-License-Identifier: Apache-2.0
"""One Hermes process serving two profiles: each profile's OneNote index is its own brain.db.

Hermes imports a directory plugin once per served profile (its own module namespace,
hermes_cli/plugins_loader.py::_directory_module_name) and runs register(ctx) inside that profile's home
scope (_plugin_home_scope). These tests do the same with two module copies and two temporary homes, and
send page visits through the real bridge HTTP server, whose request threads carry no Hermes scope (the
case that wrote every gateway-served profile's rows into the launch profile's file).
"""
import http.client
import importlib.util
import json
import sqlite3
import sys
from contextlib import contextmanager
from pathlib import Path

import pytest

hermes_constants = pytest.importorskip("hermes_constants")
pytest.importorskip("plugins.plugin_storage")

ROOT = Path(__file__).resolve().parent.parent
ORIGIN = "https://addin.example"
DB = Path("plugin-data") / "hermes-office-desktop" / "brain.db"


class Ctx:
    MANIFEST_DIR = None  # set per test by _manifests_in_tmp: office_connect never writes the real ~/Documents
    def __init__(self, ports):
        self.tools, self.hooks, self.config = {}, {}, {"allowed_origin": ORIGIN, "manifest_dir": Ctx.MANIFEST_DIR}
        self._test_ports = ports

    def register_tool(self, **kw):
        self.tools[kw["name"]] = kw

    def register_hook(self, name, fn):
        self.hooks[name] = fn

    def get_config(self, key, default=None):
        return self.config.get(key, default)


def _load_copy(name):
    spec = importlib.util.spec_from_file_location(name, ROOT / "__init__.py", submodule_search_locations=[str(ROOT)])
    mod = importlib.util.module_from_spec(spec)
    sys.modules[name] = mod
    spec.loader.exec_module(mod)
    return mod


@pytest.fixture(autouse=True)
def _manifests_in_tmp(tmp_path, monkeypatch):
    monkeypatch.setattr(Ctx, "MANIFEST_DIR", str(tmp_path / "manifests"))


@contextmanager
def home_scope(home):
    token = hermes_constants.set_hermes_home_override(str(home))
    try:
        yield
    finally:
        hermes_constants.reset_hermes_home_override(token)


def _rows(db, sql):
    c = sqlite3.connect(f"file:{db}?mode=ro", uri=True)
    try:
        return c.execute(sql).fetchall()
    finally:
        c.close()


def _post_event(server, token, body):
    port = server.port
    c = http.client.HTTPConnection("127.0.0.1", port, timeout=10)
    data = json.dumps(body).encode()
    c.request("POST", "/v1/event", body=data, headers={"Origin": ORIGIN, "Host": f"127.0.0.1:{port}",
                                                       "Content-Type": "application/json",
                                                       "Content-Length": str(len(data)),
                                                       "Authorization": f"Bearer {token}"})
    r = c.getresponse()
    out = json.loads(r.read())
    c.close()
    return out


@pytest.fixture()
def two(tmp_path, monkeypatch):
    """Profiles A and B served by one process whose launch profile is a third home."""
    launch = tmp_path / "launch"
    launch.mkdir()
    monkeypatch.setenv("HERMES_HOME", str(launch))  # what an unscoped thread resolves (the gateway's own)
    for k in ("HERMES_CRON_SESSION", "HERMES_SESSION_PLATFORM"):
        monkeypatch.delenv(k, raising=False)
    out = {"launch": launch}
    for i, name in enumerate(("a", "b")):
        home = tmp_path / f"profile-{name}"
        home.mkdir()
        mod = _load_copy(f"hermes_plugins_iso_{name}_{tmp_path.name}".replace("-", "_"))
        ctx = Ctx(tuple(range(19900 + 10 * i, 19905 + 10 * i)))
        with home_scope(home):  # what hermes_cli/plugins_loader.py::_plugin_home_scope does around register()
            mod.register(ctx)
        out[name] = (mod, ctx, home)
    yield out
    for name in ("a", "b"):
        mod = out[name][0]
        if mod._server is not None:
            mod._server.stop()
        for b in list(mod._on._brains.values()):
            b.close()
        for k in [k for k in sys.modules if k == mod.__name__ or k.startswith(mod.__name__ + ".")]:
            sys.modules.pop(k, None)


def _connect_and_pair(mod, ctx, home, notebook, session="chat-1"):
    """office_connect as a tool call inside the profile's turn scope, then the pane pairs."""
    with home_scope(home):
        res = json.loads(ctx.tools["office_connect"]["handler"]({"app": "onenote"}, session_id=session))
    assert res["ok"], res
    server = mod._server
    code = res["data"]["pairingCode"]
    return server, server.registry.pair(code, "onenote", {"id": notebook, "name": "n"}, {}, None)


def _visit(notebook, pid, title):
    return {"kind": "page-visit", "document": {"id": notebook},
            "page": {"pageId": pid, "title": title, "notebookId": notebook,
                     "outlines": [{"id": "o1", "paragraphs": [{"id": "p1", "text": f"{title} notes"}]}]}}


def test_two_profiles_in_one_process_write_two_brain_dbs(two):
    (ma, ca, ha), (mb, cb, hb) = two["a"], two["b"]
    with home_scope(ha):
        ma._on.get_brain(ca).set_meta("last_notebook", "nb-A")
    with home_scope(hb):
        mb._on.get_brain(cb).set_meta("last_notebook", "nb-B")
    assert _rows(ha / DB, "select value from meta where key='last_notebook'") == [("nb-A",)]
    assert _rows(hb / DB, "select value from meta where key='last_notebook'") == [("nb-B",)]
    assert not (two["launch"] / DB).exists()


def test_bridge_thread_event_for_profile_b_lands_in_b_db(two):
    (ma, ca, ha), (mb, cb, hb) = two["a"], two["b"]
    sa, bind_a = _connect_and_pair(ma, ca, ha, "nb-A")
    sb, bind_b = _connect_and_pair(mb, cb, hb, "nb-B")
    assert sa.port != sb.port
    # The event arrives on the bridge server's own request thread: no Hermes scope is bound there.
    out = _post_event(sb, bind_b.token, _visit("nb-B", "page-b", "Beta plan"))
    assert out["accepted"] and out["indexed"], out
    assert not (two["launch"] / DB).exists()
    assert _rows(hb / DB, "select notebook_id, page_id, title from pages") == [("nb-B", "page-b", "Beta plan")]
    assert not (ha / DB).exists() or _rows(ha / DB, "select count(*) from pages") == [(0,)]
    assert bind_b.home == str(hb.resolve())
    out = _post_event(sa, bind_a.token, _visit("nb-A", "page-a", "Alpha plan"))
    assert out["accepted"] and out["indexed"], out
    assert _rows(ha / DB, "select page_id from pages") == [("page-a",)]
    assert _rows(hb / DB, "select page_id from pages") == [("page-b",)]


def test_profile_b_has_nothing_of_a_to_undo_or_replace(two):
    (ma, ca, ha), (mb, cb, hb) = two["a"], two["b"]
    with home_scope(ha):
        a = ma._on.get_brain(ca)
        a.begin_write("nb-1", "page-1", "op-a", "brain_links", "links_outline", "d", "A's block", None)
        a.finish_write("op-a", "done", page_id="page-1", created={"outlineId": "out-a"})
        assert a.last_undoable("nb-1")["op_id"] == "op-a"
    with home_scope(hb):
        b = mb._on.get_brain(cb)
        assert b.last_undoable("nb-1") is None
        assert b.writes_for_page("nb-1", "page-1") == []
        assert b.owned_outline_ids("nb-1", "page-1") == set()


def test_call_scoped_to_another_profile_is_refused(two):
    (ma, ca, ha), (mb, cb, hb) = two["a"], two["b"]
    with home_scope(hb):  # A's tool, but the call is bound to profile B
        res = json.loads(ca.tools["brain_resurface"]["handler"]({}, session_id="chat-1"))
    assert res["ok"] is False and res["error"]["category"] == "ProfileScope"
    assert not (ha / DB).exists() and not (hb / DB).exists() and not (two["launch"] / DB).exists()


def test_pairing_without_a_known_profile_refuses_events(two):
    ma, ca, ha = two["a"]
    ma.get_server(ca).ensure_started()
    code, _ = ma._server.registry.mint_code("chat-1", "onenote")  # no home stamped
    b = ma._server.registry.pair(code, "onenote", {"id": "nb-A", "name": "n"}, {}, None)
    out = _post_event(ma._server, b.token, _visit("nb-A", "page-a", "Alpha plan"))
    assert out.get("ok") is False and out["error"]["category"] == "ProfileScope", out
    assert not (ha / DB).exists() and not (two["launch"] / DB).exists()


def test_unknown_owner_refuses_the_index(tmp_path, monkeypatch):
    monkeypatch.setenv("HERMES_HOME", str(tmp_path))
    mod = _load_copy(f"hermes_plugins_iso_unknown_{tmp_path.name}".replace("-", "_"))
    try:
        with pytest.raises(mod.BridgeError) as e:
            mod._on.get_brain(Ctx(()))  # never registered: no owner recorded
        assert e.value.category == "ProfileScope"
        assert not (tmp_path / DB).exists()
    finally:
        for k in [k for k in sys.modules if k == mod.__name__ or k.startswith(mod.__name__ + ".")]:
            sys.modules.pop(k, None)
