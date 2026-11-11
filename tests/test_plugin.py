# SPDX-License-Identifier: Apache-2.0
"""Plugin registration, schema reuse, approval hook, unattended refusal, binding through tool handlers."""
import contextvars
import hashlib
import http.client
import json
import os
import re
import sys
import threading
import types
from pathlib import Path

import pytest
from ruamel.yaml import YAML  # ships with Hermes 0.23.0 (PyYAML does not)

ROOT = Path(__file__).resolve().parent.parent
ORIGIN = "https://addin.test:8444"
NATIVE_SHA = {  # native hermes-office-live schema artifacts (byte-identical copies)
    "excel": "0bce18cb0aa191b8e34a6107edda4bb98a910e391b6f567163999cb34284956f",
    "word": "0c3fc553d63f7ad6cafc0dad402c1ad2559e3aac2f90f4f704f7f201f3de2bbc",
    "powerpoint": "f1c0e19a2081b6d24b7d5a9ea8dc185e67408e339af676bd1477b0fbb1d41728",
}


class FakeCtx:
    def __init__(self, config=None):
        self.tools, self.hooks, self.config = {}, {}, config or {}

    def register_tool(self, **kw):
        assert kw["override"] is False
        self.tools[kw["name"]] = kw

    def register_hook(self, name, fn):
        self.hooks[name] = fn

    def register_skill(self, name, path, **metadata):
        self.skills = {**getattr(self, "skills", {}), name: path}

    def get_config(self, key, default=None):
        return self.config.get(key, default)


def fresh(plugin, port):
    if plugin._server:
        plugin._server.stop()
    plugin._server = None
    ctx = FakeCtx({"allowed_origin": ORIGIN})
    ctx._test_ports = tuple(range(port, port + 5))
    plugin.register(ctx)
    return ctx


def call(ctx, name, args, sid="chat-1"):
    return json.loads(ctx.tools[name]["handler"](args, session_id=sid))


@pytest.fixture(autouse=True)
def _no_unattended(monkeypatch):
    monkeypatch.delenv("HERMES_CRON_SESSION", raising=False)
    monkeypatch.delenv("HERMES_SESSION_PLATFORM", raising=False)


def test_schemas_are_byte_identical_to_native_bridge():
    for app, sha in NATIVE_SHA.items():
        assert hashlib.sha256((ROOT / "schemas" / f"{app}.json").read_bytes()).hexdigest() == sha


def test_registration_counts_and_manifest(plugin):
    ctx = fresh(plugin, 19765)
    by_set = {}
    for t in ctx.tools.values():
        by_set.setdefault(t["toolset"], []).append(t["name"])
    assert {k: len(v) for k, v in by_set.items()} == {
        "office-desktop-excel": 27, "office-desktop-word": 12, "office-desktop-powerpoint": 18, "office-desktop": 5,
        "office-desktop-onenote": 8, "office-desktop-brain": 9, "office-desktop-outlook": 7}
    assert "excel_execute_plan" not in ctx.tools
    manifest = YAML(typ="safe", pure=True).load((ROOT / "plugin.yaml").read_text())
    assert sorted(manifest["provides_tools"]) == sorted(ctx.tools)
    assert manifest["provides_hooks"] == ["pre_tool_call"]
    assert ctx.skills["office-desktop"].exists()
    for t in ctx.tools.values():
        assert "bridgeSessionId" not in t["schema"]["parameters"].get("properties", {})


@pytest.mark.parametrize("app", ["excel", "word", "powerpoint"])
def test_pane_executor_implements_every_registered_tool(plugin, app):
    """Every registered <app>_* tool exists in that app's executor table (and nothing is silently missing)."""
    src = (ROOT / "addin" / app / "executor.js").read_text()
    implemented = set(re.findall(r"^  async ([a-z_]+)\(", src, re.M))
    registered = {s["name"].split("_", 1)[1] for s in plugin.SCHEMAS[app]}
    # office_insert_image is one plugin-level tool; every pane implements its two commands.
    assert registered | {"insert_image", "image_status"} == implemented


def test_approval_requested_for_mutations_only(plugin):
    gate = fresh(plugin, 19765).hooks["pre_tool_call"]
    for ro in ("excel_read_range", "word_read_selection", "powerpoint_list_slide_layouts", "office_connect"):
        assert gate(tool_name=ro, args={}) is None
    for mut in ("excel_write_range", "excel_run_calculation", "word_insert_text", "powerpoint_delete_slide", "office_undo"):
        d = gate(tool_name=mut, args={})
        assert d["action"] == "approve" and d["rule_key"] == plugin.APPROVAL_RULE_KEY, mut
    assert gate(tool_name="terminal", args={}) is None  # never touches other tools
    assert gate(tool_name="excel_execute_plan", args={}) is None  # not ours


@pytest.mark.parametrize("env", [{"HERMES_CRON_SESSION": "1"}, {"HERMES_SESSION_PLATFORM": "webhook"},
                                 {"HERMES_SESSION_PLATFORM": "api_server"}])
def test_unattended_runs_are_refused_even_if_approval_policy_would_allow(plugin, monkeypatch, env):
    ctx = fresh(plugin, 19765)
    for k, v in env.items():
        monkeypatch.setenv(k, v)
    gate = ctx.hooks["pre_tool_call"]
    for name in ("excel_write_range", "excel_read_range", "office_connect", "office_undo", "word_insert_text"):
        d = gate(tool_name=name, args={})
        assert d["action"] == "block" and "refused" in d["message"], name
    out = call(ctx, "office_connect", {"app": "excel"})
    assert out["error"]["category"] == "Unattended" and out["error"]["hint"]
    out = call(ctx, "excel_read_range", {"sheet": "S", "address": "A1"})
    assert out["error"]["category"] == "Unattended"
    assert call(ctx, "office_status", {})["data"]["unattended"]
    assert gate(tool_name="terminal", args={}) is None



class CronCtx(FakeCtx):
    """FakeCtx with the public ctx.current_cron_execution() seam, backed by a ContextVar like Hermes'."""
    def __init__(self, var, config=None):
        super().__init__(config)
        self._cron_var = var

    def current_cron_execution(self):
        return self._cron_var.get()


def _register(plugin, ctx, port):
    if plugin._server:
        plugin._server.stop()
    plugin._server = None
    ctx._test_ports = tuple(range(port, port + 5))
    plugin.register(ctx)
    return ctx


def _fake_session_context(monkeypatch):
    """A stand-in gateway.session_context whose get_session_env reads per-turn ContextVars, as Hermes' does."""
    cvars = {n: contextvars.ContextVar(n, default=None) for n in ("HERMES_CRON_SESSION", "HERMES_SESSION_PLATFORM")}

    def get_session_env(name, default=""):
        var = cvars.get(name)
        value = var.get() if var is not None else None
        return os.getenv(name, default) if value is None else value
    pkg = types.ModuleType("gateway")
    pkg.__path__ = []
    mod = types.ModuleType("gateway.session_context")
    mod.get_session_env = get_session_env
    pkg.session_context = mod
    monkeypatch.setitem(sys.modules, "gateway", pkg)
    monkeypatch.setitem(sys.modules, "gateway.session_context", mod)
    return cvars


def _assert_refused(ctx):
    gate = ctx.hooks["pre_tool_call"]
    for name in ("excel_write_range", "excel_read_range", "office_connect", "onenote_get_context"):
        d = gate(tool_name=name, args={})
        assert d and d["action"] == "block" and "refused" in d["message"], name
    assert call(ctx, "office_connect", {"app": "excel"})["error"]["category"] == "Unattended"
    assert call(ctx, "office_status", {})["data"]["unattended"]


def test_gateway_hosted_cron_is_refused_via_ctx_current_cron_execution(plugin):
    """A cron fire hosted in the gateway binds only a ContextVar; os.environ says nothing."""
    var = contextvars.ContextVar("hermes_cron_execution", default=None)
    ctx = _register(plugin, CronCtx(var, {"allowed_origin": ORIGIN}), 19765)
    gate = ctx.hooks["pre_tool_call"]
    assert gate(tool_name="excel_write_range", args={})["action"] == "approve"

    def fire():
        var.set(object())
        assert os.getenv("HERMES_CRON_SESSION") is None and os.getenv("HERMES_SESSION_PLATFORM") is None
        _assert_refused(ctx)
    contextvars.copy_context().run(fire)
    assert gate(tool_name="excel_write_range", args={})["action"] == "approve"  # the binding never leaks


@pytest.mark.parametrize("name,value", [("HERMES_CRON_SESSION", "1"), ("HERMES_SESSION_PLATFORM", "webhook"),
                                        ("HERMES_SESSION_PLATFORM", "api_server")])
def test_per_turn_session_markers_are_read_through_get_session_env(plugin, monkeypatch, name, value):
    cvars = _fake_session_context(monkeypatch)
    ctx = fresh(plugin, 19765)
    gate = ctx.hooks["pre_tool_call"]

    def turn():
        cvars[name].set(value)
        assert os.getenv(name) is None
        _assert_refused(ctx)
    contextvars.copy_context().run(turn)
    assert gate(tool_name="excel_write_range", args={})["action"] == "approve"


def test_session_markers_fall_back_to_os_environ_without_gateway(plugin, monkeypatch):
    monkeypatch.setitem(sys.modules, "gateway.session_context", None)  # import raises ImportError
    monkeypatch.setenv("HERMES_CRON_SESSION", "1")
    _assert_refused(fresh(plugin, 19765))


def test_real_hermes_session_context_cron_binding_is_refused(plugin):
    """Against the installed Hermes, when importable: the scheduler's per-turn cron marker."""
    sc = pytest.importorskip("gateway.session_context")
    ctx = fresh(plugin, 19765)

    def turn():
        sc.set_session_vars(platform="cron", cron_session="1")
        assert os.getenv("HERMES_CRON_SESSION") is None
        _assert_refused(ctx)
    contextvars.copy_context().run(turn)
    assert ctx.hooks["pre_tool_call"](tool_name="excel_write_range", args={})["action"] == "approve"


def test_connect_description_routes_browser_hosted_desktop_to_the_app(plugin):
    desc = plugin.CONNECT_SCHEMA["description"]
    assert "surface 'iframe'" in desc and "Hermes Desktop app" in desc and "do not call office_connect" in desc

def test_interactive_platforms_are_not_refused(plugin, monkeypatch):
    ctx = fresh(plugin, 19765)
    for p in ("cli", "desktop", "telegram", "tui"):
        monkeypatch.setenv("HERMES_SESSION_PLATFORM", p)
        assert ctx.hooks["pre_tool_call"](tool_name="excel_read_range", args={}) is None


def test_tools_fail_cleanly_with_hints(plugin):
    ctx = fresh(plugin, 19765)
    out = json.loads(ctx.tools["excel_read_range"]["handler"]({"sheet": "S", "address": "A1"}))
    assert out["ok"] is False and out["error"]["category"] == "Validation"  # no session id (e.g. odd caller)
    out = call(ctx, "excel_read_range", {"sheet": "S", "address": "A1"}, sid="20261105_101010_abc")
    assert out["error"]["category"] == "NotPaired" and "office_connect" in out["error"]["hint"]
    assert call(ctx, "excel_read_range", {"sheet": "S"})["error"]["category"] == "Validation"
    assert call(ctx, "excel_read_range", {"sheet": "S", "address": "A1", "bridgeSessionId": "x"})["error"]["category"] == "Validation"
    assert call(ctx, "word_read_document_text", {"maxLength": 0})["error"]["category"] == "Validation"
    assert call(ctx, "office_connect", {"app": "gmail"})["error"]["category"] == "Validation"
    assert call(ctx, "office_undo", {"app": "word"})["error"]["category"] == "NotPaired"


def _http(port, method, path, body=None, token=None):
    c = http.client.HTTPConnection("127.0.0.1", port, timeout=30)
    h = {"Origin": ORIGIN, "Host": f"127.0.0.1:{port}"}
    data = json.dumps(body).encode() if body is not None else None
    if data:
        h.update({"Content-Type": "application/json", "Content-Length": str(len(data))})
    if token:
        h["Authorization"] = f"Bearer {token}"
    c.request(method, path, body=data, headers=h)
    r = c.getresponse()
    v = json.loads(r.read())
    c.close()
    return v


@pytest.mark.parametrize("app,tool,canonical", [("excel", "excel_read_range", "read_range"),
                                                ("word", "word_read_selection", "read_selection"),
                                                ("powerpoint", "powerpoint_read_all_slides", "read_all_slides")])
def test_connect_binds_that_chat_per_app_and_routes_undo(plugin, app, tool, canonical):
    ctx = fresh(plugin, 19780)
    try:
        out = call(ctx, "office_connect", {"app": app})
        assert out["ok"] and out["data"]["app"] == app
        code = out["data"]["pairingCode"]
        port = int(out["data"]["bridge"].split(":")[1])
        assert len(code) == 9 and code[4] == "-"
        assert f"manifest-{app}.xml" in " ".join(out["data"]["firstTimeSetup"])
        token = _http(port, "POST", "/v1/pair", {"code": code, "app": app, "document": {"name": "Q4"},
                                                 "supportedTools": [canonical, "undo_last"]})["paneToken"]
        status = call(ctx, "office_status", {"app": app})["data"]
        assert status["paired"] and status["document"] == "Q4"
        assert status["live"] is True and 0 <= status["lastPollAgeMs"] < 5000 and status["paneState"] == "live"
        assert not call(ctx, "office_status", {"app": app}, sid="chat-2")["data"]["paired"]
        seen = []

        def pane(n):
            for _ in range(n):
                cmd = _http(port, "GET", "/v1/next?wait=5", token=token)["command"]
                seen.append(cmd["tool"])
                res = {"ok": True, "data": {"echo": cmd["tool"]}} if cmd["tool"] != "undo_last" else \
                    {"ok": False, "error": {"category": "NothingToUndo", "message": "none"}}
                _http(port, "POST", "/v1/result", {"requestId": cmd["requestId"], "result": res}, token=token)
        t = threading.Thread(target=pane, args=(2,), daemon=True)
        t.start()
        args = {"sheet": "S", "address": "A1"} if app == "excel" else {}
        res = call(ctx, tool, args)
        assert res["ok"] and res["data"] == {"echo": canonical} and isinstance(res["roundTripMs"], int)
        undo = call(ctx, "office_undo", {"app": app})
        assert undo["error"]["category"] == "NothingToUndo" and undo["error"]["hint"]
        t.join(5)
        assert seen == [canonical, "undo_last"]
        # Another chat cannot drive this pane; the pane's own app is separate from the others.
        assert call(ctx, tool, args, sid="chat-2")["error"]["category"] == "NotPaired"
        assert call(ctx, "office_disconnect", {"app": app})["data"]["disconnected"]
    finally:
        plugin._server.stop()
        plugin._server = None


def test_plugin_holds_no_credentials():
    text = "\n".join(p.read_text() for p in ROOT.glob("*.py"))
    for forbidden in (".env", "auth.json", "get_secret", "keychain", "os.environ"):
        assert forbidden not in text
    # The only environment reads are the two non-secret Hermes session markers.
    assert set(re.findall(r"os\.getenv\(([^,)]+)", text)) == {"name"}
    init = (ROOT / "__init__.py").read_text()
    assert set(re.findall(r'_session_marker\("([A-Z_]+)"\)', init)) == {"HERMES_CRON_SESSION", "HERMES_SESSION_PLATFORM"}


def test_status_reports_an_unloaded_pane_as_not_live(plugin, bridge_mod):
    """F3/F4: paired but silent for > 45 s reads as unloaded, with the reopen instruction."""
    ctx = fresh(plugin, 19790)
    try:
        code = call(ctx, "office_connect", {"app": "excel"})["data"]["pairingCode"]
        reg = plugin.get_server(ctx).registry
        b = reg.pair(code, "excel", {"name": "Q4"}, {}, [])
        b.last_poll -= 50
        st = call(ctx, "office_status", {"app": "excel"})["data"]
        assert st["paired"] is True and st["live"] is False and st["lastPollAgeMs"] >= 50000
        assert st["paneState"] == "unloaded"
        assert "pane unloaded (switched chat or closed); reopen Workbench in the ribbon" in st["tellUser"]
        err = call(ctx, "excel_read_range", {"sheet": "S", "address": "A1"})["error"]
        assert err["category"] == "Unavailable" and "reopen Workbench in the ribbon" in err["message"]
        assert "one chat per pane" in err["hint"].lower()
        none = call(ctx, "office_status", {"app": "excel"}, sid="chat-9")["data"]
        assert none["paired"] is False and none["live"] is False and none["paneState"] == "unpaired"
    finally:
        plugin._server.stop()
        plugin._server = None


def test_connect_and_skill_say_read_the_pane_before_opening_office(plugin):
    """F2: never navigate the pane away from an open document."""
    d = plugin.CONNECT_SCHEMA["description"]
    assert "desktop_preview" in d and "read" in d and "only if no Office document is showing" in d
    skill = (ROOT / "skills" / "office-desktop" / "SKILL.md").read_text()
    assert "action: read" in skill and "only if no Office document is showing" in skill
    assert "one chat per pane" in skill.lower()  # the runtime hints' wording, explained in the skill
    readme = (ROOT / "README.md").read_text().lower()
    assert "all your office panes can share one chat" in readme and "one chat per pane" not in readme


def test_manifests_auto_show_the_task_pane_with_the_document():
    """F3: Office reopens the Hermes pane with a document that was paired before."""
    import subprocess, tempfile, sys as _sys
    with tempfile.TemporaryDirectory() as d:
        subprocess.run([_sys.executable, str(ROOT / "scripts" / "build_addin.py"), "--out", d], check=True,
                       capture_output=True)
        for app in ("excel", "word", "powerpoint", "onenote"):
            xml = (Path(d) / f"manifest-{app}.xml").read_text()
            assert "<TaskpaneId>Office.AutoShowTaskpaneWithDocument</TaskpaneId>" in xml


def _chat_hints():
    import importlib.util
    spec = importlib.util.spec_from_file_location("hol_chat", ROOT / "scripts" / "livetest" / "chat.py")
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    return mod.OFFICE_RULE_HINTS


def test_approval_prompt_names_app_document_operation_and_target(plugin):
    """E1-3: informed consent. Wording only; the rule_key is the same constant for every prompt."""
    ctx = fresh(plugin, 19795)
    try:
        gate = ctx.hooks["pre_tool_call"]
        code = call(ctx, "office_connect", {"app": "excel"})["data"]["pairingCode"]
        plugin.get_server(ctx).registry.pair(code, "excel", {"name": "HOL live test"}, {}, [])
        d = gate(tool_name="excel_write_range", args={"sheet": "Decision tracker", "address": "A1:F4", "values": [[1]]},
                 session_id="chat-1")
        assert d["message"] == "Excel · HOL live test · write Decision tracker!A1:F4 (undoable) · Workbench"
        assert d["rule_key"] == plugin.APPROVAL_RULE_KEY
        # another chat has no pane: no document title is borrowed from chat-1
        other = gate(tool_name="excel_write_range", args={"sheet": "S", "address": "B2"}, session_id="chat-2")
        assert other["message"] == "Excel · write S!B2 (undoable) · Workbench"
        assert gate(tool_name="excel_delete_sheet", args={"name": "Old"})["message"] == \
            "Excel · delete sheet 'Old' (not undoable) · Workbench"
        assert "Word · insert text at end (undoable)" in gate(tool_name="word_insert_text",
                                                               args={"text": "x", "location": "end"})["message"]
        assert "PowerPoint · create text slide 'Sources' (undoable)" in gate(
            tool_name="powerpoint_create_text_slide", args={"title": "Sources"})["message"]
        assert gate(tool_name="office_undo", args={"app": "word"})["message"] == \
            "Word · undo the last Hermes change · Workbench"
        assert "OneNote · create page 'Daily 2026-11-05' (a page: undo cannot delete it, you can)" in gate(
            tool_name="onenote_create_page", args={"title": "Daily 2026-11-05", "html": "<p/>"})["message"]
        assert "OneNote · write the Links & Backlinks block on the open page (undoable)" in gate(
            tool_name="brain_links", args={})["message"]
        assert "build Excel tracker, Word brief from the decision page" in gate(
            tool_name="brain_make_deliverables", args={"decisionPageId": "p", "hosts": ["excel", "word"]})["message"]
        # hostile names are flattened and capped, never multi-line
        evil = gate(tool_name="excel_write_range", args={"sheet": "A\nB\x1b[31m" + "x" * 300, "address": "A1"})["message"]
        assert "\n" not in evil and "\x1b" not in evil and len(evil) < 160
    finally:
        plugin._server.stop()
        plugin._server = None


def test_every_approval_prompt_still_matches_the_live_test_approver(plugin):
    """scripts/livetest/chat.py approves only prompts that name Office/OneNote or carry the rule key."""
    hints = _chat_hints()
    gate = fresh(plugin, 19765).hooks["pre_tool_call"]
    names = sorted(plugin.DOCUMENT_TOOLS - plugin.READ_ONLY) + ["office_undo", "onenote_append", "brain_decision_page",
                                                                "brain_make_deliverables", "brain_links"]
    for name in names:
        d = gate(tool_name=name, args={})
        assert d and d["rule_key"] == plugin.APPROVAL_RULE_KEY, name
        assert any(h in d["message"] for h in hints), (name, d["message"])


def test_release_copy_is_truthful_about_platforms_and_status():
    yaml = YAML(typ="safe", pure=True)
    assert yaml.load((ROOT / "plugin.yaml").read_text())["platforms"] == ["macos", "windows"]
    assert yaml.load((ROOT / "catalog" / "hermes-office-desktop.yaml").read_text())["platforms"] == ["macos", "windows"]
    readme = (ROOT / "README.md").read_text()
    assert "Windows: one real Hermes Desktop session passed the 5-prompt story on 2026-11-10 with a single profile." in readme and "## First connect (about 2 minutes per app, once)" in readme
    assert "not deployed, not published" not in readme and "verified only against a strict OneNoteApi 1.1 mock" not in readme
    assert "Hermes Office Live" not in readme


def test_build_writes_an_honest_build_identity_into_the_pane(tmp_path):
    """critic2 #7: the footer label = add-in version + content id of every other emitted file + source commit."""
    import hashlib, importlib.util, json, re, subprocess, sys as _sys
    subprocess.run([_sys.executable, str(ROOT / "scripts" / "build_addin.py"), "--out", str(tmp_path / "s")], check=True,
                   capture_output=True)
    site = tmp_path / "s"
    js = (site / "shared" / "build-info.js").read_text()
    info = json.loads(re.search(r"Object\.freeze\((\{.*\})\)", js).group(1))
    sums = [ln.split("  ", 1) for ln in (site / "SHA256SUMS").read_text().splitlines()]
    names = [n for _, n in sums]
    assert "shared/build-info.js" in names and "SHA256SUMS" not in names
    for h, n in sums:  # the checksum inventory stays valid, build-info included
        assert hashlib.sha256((site / n).read_bytes()).hexdigest() == h
    others = "".join(f"{h}  {n}\n" for h, n in sums if n != "shared/build-info.js").encode()
    assert info["build"] == hashlib.sha256(others).hexdigest()[:12]
    assert info["version"] == "0.3.0"
    head = subprocess.run(["git", "-C", str(ROOT), "rev-parse", "--short=12", "HEAD"], capture_output=True, text=True).stdout.strip()
    top = subprocess.run(["git", "-C", str(ROOT), "rev-parse", "--show-toplevel"], capture_output=True, text=True).stdout.strip()
    if top and Path(top).resolve() == ROOT.resolve():
        assert info["commit"] == head and isinstance(info["dirty"], bool)
    else:  # an exported copy (git archive) is never labelled with an enclosing repository's commit
        assert info["commit"] is None and info["dirty"] is None
    # an export of the tree placed inside some other git repository: no borrowed commit
    import shutil
    other_repo = tmp_path / "other"
    subprocess.run(["git", "init", "-q", str(other_repo)], check=True)
    (other_repo / "f").write_text("x")
    subprocess.run(["git", "-C", str(other_repo), "-c", "user.name=t", "-c", "user.email=t@t", "add", "f"], check=True)
    subprocess.run(["git", "-C", str(other_repo), "-c", "user.name=t", "-c", "user.email=t@t", "commit", "-qm", "x"], check=True)
    copy = other_repo / "export"
    shutil.copytree(ROOT / "addin", copy / "addin")
    shutil.copytree(ROOT / "scripts", copy / "scripts", ignore=shutil.ignore_patterns("__pycache__"))
    shutil.copy(ROOT / "bridge_server.py", copy / "bridge_server.py")  # the default add-in base lives there (A24)
    subprocess.run([_sys.executable, str(copy / "scripts" / "build_addin.py"), "--out", str(tmp_path / "x")], check=True,
                   capture_output=True)
    js3 = (tmp_path / "x" / "shared" / "build-info.js").read_text()
    exported = json.loads(re.search(r"Object\.freeze\((\{.*\})\)", js3).group(1))
    assert exported["commit"] is None and exported["dirty"] is None and exported["build"]
    for app in ("excel", "word", "powerpoint", "onenote", "outlook"):
        assert '<p id="build"' in (site / app / "taskpane.html").read_text()
    # a different site (another base URL) is a different build id: a cached older pane is distinguishable
    spec = importlib.util.spec_from_file_location("ba", ROOT / "scripts" / "build_addin.py")
    ba = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(ba)
    other = ba.build("https://addin.test:8444/office/live/", tmp_path / "o", False)
    js2 = (tmp_path / "o" / "shared" / "build-info.js").read_text()
    assert json.loads(re.search(r"Object\.freeze\((\{.*\})\)", js2).group(1))["build"] != info["build"]
    assert {n for _, n in other} == set(names)


def test_display_name_is_omnia_works_and_ids_are_unchanged(plugin, tmp_path):
    """Rename: display names only. Add-in ids, plugin id, ribbon/harness selectors and approval suffix stay in sync."""
    import importlib.util, re, subprocess, sys as _sys
    subprocess.run([_sys.executable, str(ROOT / "scripts" / "build_addin.py"), "--out", str(tmp_path / "s")], check=True,
                   capture_output=True)
    site = tmp_path / "s"
    ids = {"excel": "6b0f3c1e-2d4a-4f7e-9c5b-8a1d2e3f4a51", "word": "b6744181-1c6b-4301-b139-19cd2f7af1ec",
           "powerpoint": "516118d2-cf21-497f-a118-28595d71bbba", "onenote": "3f6c2a8e-5b1d-4e7a-9c40-7d2e1b9a6c53",
           "outlook": "6601c392-1a11-41f3-bd86-7d4263d9a1c5"}
    labels = {"excel": "Excel", "word": "Word", "powerpoint": "PowerPoint", "onenote": "OneNote", "outlook": "Outlook"}
    descriptor = "works with Excel, Word, PowerPoint, OneNote and Outlook (mail and calendar) on the web"
    spec = importlib.util.spec_from_file_location("hol_ui", ROOT / "scripts" / "livetest" / "office_ui.py")
    ui = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(ui)
    for app, label in labels.items():
        xml = (site / f"manifest-{app}.xml").read_text()
        assert f"<Id>{ids[app]}</Id>" in xml  # changing an add-in id would orphan installed add-ins
        assert f'<DisplayName DefaultValue="Workbench for {label}"/>' in xml
        desc = re.search(r'<Description DefaultValue="([^"]*)"', xml).group(1)
        assert desc.startswith("Workbench: " + descriptor) and len(desc) <= 250
        button = re.search(r'<bt:String id="Hermes.Button.Label" DefaultValue="([^"]*)"', xml).group(1)
        assert button == "Workbench" and ui.RIBBON_LABELS[0] == button  # the harness clicks what Office shows
        assert "Hermes Desktop for" not in xml and "Office Desktop" not in xml
        html = (site / app / "taskpane.html").read_text()
        assert f"<title>Workbench for {label}</title>" in html and "<h1>Workbench</h1>" in html and descriptor in html
    assert plugin.PRODUCT == "Workbench" and plugin.PRODUCT in _chat_hints()
    cat = (ROOT / "catalog" / "hermes-office-desktop.yaml").read_text()
    assert 'name: hermes-office-desktop\ntitle: "Workbench by Atlas Omnia"' in cat
    man = (ROOT / "plugin.yaml").read_text()
    assert man.startswith("# SPDX-License-Identifier: Apache-2.0\nname: hermes-office-desktop\n") and "Workbench by Atlas Omnia — " + descriptor in man
    assert (ROOT / "README.md").read_text().splitlines()[1] == "# Workbench by Atlas Omnia (0.3.0)"
    assert "Copyright 2026 Atlas Omnia" in (ROOT / "NOTICE").read_text()  # legal author unchanged


def test_outlook_manifest_is_a_mailapp_with_read_and_compose_and_no_send_permission(tmp_path):
    import subprocess, sys as _sys
    subprocess.run([_sys.executable, str(ROOT / "scripts" / "build_addin.py"), "--out", str(tmp_path / "s")], check=True,
                   capture_output=True)
    xml = (tmp_path / "s" / "manifest-outlook.xml").read_text()
    assert 'xsi:type="MailApp"' in xml and '<Host Name="Mailbox"/>' in xml and "{{" not in xml
    assert '<Form xsi:type="ItemRead">' in xml and '<Form xsi:type="ItemEdit">' in xml
    assert "MessageReadCommandSurface" in xml and "MessageComposeCommandSurface" in xml
    # Calendar in the SAME add-in (one id): appointment read (attendee) and compose (organizer) forms, both overrides.
    assert '<Rule xsi:type="ItemIs" ItemType="Appointment" FormType="ReadOrEdit"/>' in xml
    assert xml.count('xsi:type="AppointmentAttendeeCommandSurface"') == 2 and xml.count('xsi:type="AppointmentOrganizerCommandSurface"') == 2
    assert xml.count("<Id>") == 1 and "ReadWriteMailbox" not in xml
    assert "<Permissions>ReadWriteItem</Permissions>" in xml  # the open item only, never ReadWriteMailbox
    assert "<SupportsPinning>true</SupportsPinning>" in xml and "VersionOverridesV1_1" in xml
    assert "https://atlasomnia.github.io/workbench/office/live/outlook/taskpane.html" in xml
    assert not (tmp_path / "s" / "dev").exists() and (tmp_path / "s" / "outlook" / "executor.js").exists()


def test_connect_routing_office_connect_is_named_and_status_says_call_office_connect(plugin, bridge_mod):
    """A19 (seen live): 'connect to Outlook' went to Hermes core's manage_connections, or stopped at office_status.
    office_connect's description claims that phrase and rules out manage_connections; office_status never pairs and,
    with nothing paired, says exactly 'call office_connect' as the next step."""
    ctx = fresh(plugin, 19796)
    try:
        conn = ctx.tools["office_connect"]["schema"]["description"]
        stat = ctx.tools["office_status"]["schema"]["description"]
        assert "connect to Excel" in conn and "manage_connections" in conn and "office_status" in conn
        assert "call office_connect" in stat and "never pairs" in stat
        for app in ("excel", "word", "outlook"):
            st = call(ctx, "office_status", {"app": app}, sid="chat-route")["data"]
            assert st["paneState"] == "unpaired"
            assert f"call office_connect with app '{app}'" in st["next"] and "manage_connections" in st["next"]
        code = call(ctx, "office_connect", {"app": "excel"}, sid="chat-route")["data"]["pairingCode"]
        plugin.get_server(ctx).registry.pair(code, "excel", {"name": "Q4"}, {}, [])
        assert "next" not in call(ctx, "office_status", {"app": "excel"}, sid="chat-route")["data"]
    finally:
        plugin.get_server(ctx).stop()
        plugin._server = None
