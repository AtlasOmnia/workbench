# SPDX-License-Identifier: Apache-2.0
"""A14: an idle profile gives its bridge port back, so a sixth profile can pair without a gateway restart.

One process, several plugin copies (what a multiplexed gateway loads: one module namespace per served
profile, each register(ctx) inside its own home scope), one shared five-port range, real bridge HTTP
threads. The pane side is driven over real HTTP the way addin/shared/bridge-client.js does it: find the
bridge by its /v1/hello instance id on whichever port it holds, then poll with the pane token.
A shared fake clock stands in for the 10-minute idle period; the reaper thread is tested on its own.
"""
import http.client
import importlib.util
import json
import socket
import sys
import threading
import time
from contextlib import contextmanager
from pathlib import Path

import pytest

import https_client

hermes_constants = pytest.importorskip("hermes_constants")

ROOT = Path(__file__).resolve().parent.parent
ORIGIN = "https://addin.example"
PORTS = tuple(range(19950, 19955))


class Clock:
    def __init__(self):
        self.t = 1000.0

    def __call__(self):
        return self.t


class Ctx:
    MANIFEST_DIR = None  # set per test by _manifests_in_tmp: office_connect never writes the real ~/Documents
    def __init__(self):
        self.tools, self.hooks, self.config = {}, {}, {"allowed_origin": ORIGIN, "manifest_dir": Ctx.MANIFEST_DIR}
        self._test_ports = PORTS

    def register_tool(self, **kw):
        self.tools[kw["name"]] = kw

    def register_hook(self, name, fn):
        self.hooks[name] = fn

    def get_config(self, key, default=None):
        return self.config.get(key, default)


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


def _load_copy(name):
    spec = importlib.util.spec_from_file_location(name, ROOT / "__init__.py", submodule_search_locations=[str(ROOT)])
    mod = importlib.util.module_from_spec(spec)
    sys.modules[name] = mod
    spec.loader.exec_module(mod)
    mod._tls_dir = lambda ctx=None: https_client.DATA  # one test CA for every profile copy (TLS is tested elsewhere)
    return mod


class Profile:
    def __init__(self, mod, ctx, home, clock):
        self.mod, self.ctx, self.home = mod, ctx, home
        server = mod.get_server(ctx)
        server.reap_interval = 0  # releases are driven by the test (the reaper has its own test)
        server.registry._clock = clock
        server.registry._last_activity = clock()  # the idle clock starts on the fake clock
        self.server = server

    def release(self):
        """maybe_release once the test's own finished requests have left their handlers (a handler counts as
        busy until it returns, slightly after the client has read its response)."""
        assert _until(lambda: self.server._requests == 0)
        return self.server.maybe_release()

    def tool(self, name, args=None, session="chat-1"):
        with home_scope(self.home):
            return json.loads(self.ctx.tools[name]["handler"](args or {}, session_id=session))


@pytest.fixture()
def world(tmp_path, monkeypatch):
    for port in PORTS:
        with socket.socket() as s:
            s.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEADDR, 1)
            try:
                s.bind(("127.0.0.1", port))
            except OSError:
                pytest.skip(f"test port {port} is in use on this machine")
    for k in ("HERMES_CRON_SESSION", "HERMES_SESSION_PLATFORM"):
        monkeypatch.delenv(k, raising=False)
    clock = Clock()
    made = []

    def make(n):
        for i in range(n):
            home = tmp_path / f"profile-{len(made)}"
            home.mkdir()
            mod = _load_copy(f"hermes_plugins_ports_{len(made)}_{tmp_path.name}".replace("-", "_"))
            ctx = Ctx()
            with home_scope(home):
                mod.register(ctx)
            made.append(Profile(mod, ctx, home, clock))
        return made[-n:]

    yield clock, make
    for p in made:
        p.server.stop()
        for b in list(p.mod._on._brains.values()):
            b.close()
        for k in [k for k in sys.modules if k == p.mod.__name__ or k.startswith(p.mod.__name__ + ".")]:
            sys.modules.pop(k, None)


# --- the pane, over real HTTP ---------------------------------------------------------------------------
def req(port, method, path, body=None, token=None, conn=None):
    c = conn or https_client.conn(port, timeout=10)
    headers = {"Origin": ORIGIN, "Host": f"localhost:{port}"}
    data = None
    if token:
        headers["Authorization"] = f"Bearer {token}"
    if body is not None:
        data = json.dumps(body).encode()
        headers.update({"Content-Type": "application/json", "Content-Length": str(len(data))})
    c.request(method, path, body=data, headers=headers)
    r = c.getresponse()
    raw = r.read()
    if conn is None:
        c.close()
    return r.status, (json.loads(raw) if raw else None)


def hello(port):
    try:
        status, body = req(port, "GET", "/v1/hello")
    except OSError:
        return None
    return body.get("instance") if status == 200 and body else None


def find_bridge(instance):
    """bridge-client.js _findBridge: hello on every port, no token sent."""
    return next((p for p in PORTS if hello(p) == instance), None)


def pane_pair(port, code, app="excel"):
    status, body = req(port, "POST", "/v1/pair", {"code": code, "app": app, "document": {"name": "Book.xlsx"},
                                                   "supportedTools": []})
    assert status == 200 and body["ok"], body
    return body["paneToken"]


def poll(port, token, wait=0):
    return req(port, "GET", f"/v1/next?wait={wait}", token=token)


class Pane(threading.Thread):
    """A pane that keeps polling its bridge and answers commands; relocates by instance id when its port goes."""

    def __init__(self, port, token, instance):
        super().__init__(daemon=True)
        self.port, self.token, self.instance = port, token, instance
        self.stop_flag = threading.Event()
        self.ran, self.tokens_sent_to = [], set()

    def run(self):
        while not self.stop_flag.is_set():
            try:
                if hello(self.port) != self.instance:  # released, or another profile's bridge took the port
                    raise OSError("not our bridge here")
                self.tokens_sent_to.add(self.port)
                status, body = poll(self.port, self.token)
            except OSError:
                found = find_bridge(self.instance)
                if found:
                    self.port = found
                time.sleep(0.05)
                continue
            cmd = (body or {}).get("command") if status == 200 else None
            if cmd:
                self.ran.append((self.port, cmd["tool"]))
                req(self.port, "POST", "/v1/result", {"requestId": cmd["requestId"],
                                                       "result": {"ok": True, "data": {"tool": cmd["tool"]}}},
                     token=self.token)
            time.sleep(0.03)


def connect(profile, app="excel", session="chat-1"):
    res = profile.tool("office_connect", {"app": app}, session=session)
    assert res["ok"], res
    return int(res["data"]["bridge"].rsplit(":", 1)[1]), res["data"]["pairingCode"]


def port_is_free(port):
    with socket.socket() as s:
        s.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEADDR, 1)  # as the bridge binds off Windows
        try:
            s.bind(("127.0.0.1", port))
            return True
        except OSError:
            return False


# --- tests -----------------------------------------------------------------------------------------------
def test_sixth_profile_is_told_exactly_why_then_pairs_after_an_idle_profile_releases(world):
    clock, make = world
    profiles = make(6)
    ports, panes = [], {}
    for i, p in enumerate(profiles[:5]):
        port, code = connect(p)
        ports.append(port)
        if i:  # profiles 1-4 have a pane open; profile 0 only ever minted a code
            panes[i] = (port, pane_pair(port, code))
    assert sorted(ports) == list(PORTS)

    sixth = profiles[5]
    res = sixth.tool("office_connect", {"app": "excel"})
    bridge_mod = sys.modules[sixth.mod.__name__ + ".bridge_server"]
    assert res["ok"] is False and res["error"]["category"] == "PortsBusy"
    msg = res["error"]["message"]
    assert msg.startswith(bridge_mod.ports_busy_message(PORTS))
    assert "five Hermes profiles are using Workbench right now" in msg
    assert "Close a Workbench pane" in msg and "wait up to 10 minutes" in msg
    assert "127.0.0.1:19950-19954" in msg and res["error"]["hint"]
    status = sixth.tool("office_status", {"app": "excel"})["data"]
    assert status["bridgeListening"] is False and status["bridgePortsFree"] == 0
    assert status["bridgeBusy"]["category"] == "PortsBusy" and status["tellUser"] == bridge_mod.ports_busy_message(PORTS)

    # Ten idle minutes later: profile 0's code expired and it never had a pane; profiles 1-4 keep polling.
    clock.t += 601
    for port, token in panes.values():
        assert poll(port, token)[0] == 200
    released = [p.release() for p in profiles[:5]]
    assert released == [True, False, False, False, False]
    assert profiles[0].server.port is None and port_is_free(ports[0])
    # profile 0's status starts nothing for an unpaired chat (scripts/livetest/warm.py relies on it)
    assert profiles[0].tool("office_status", {"app": "excel"})["data"]["bridgeListening"] is False
    assert port_is_free(ports[0])

    port, code = connect(sixth)
    assert port == ports[0]
    token = pane_pair(port, code)
    pane = Pane(port, token, hello(port))
    pane.start()
    try:
        out = sixth.tool("excel_get_workbook_overview", {})
        assert out["ok"] is True and out["data"] == {"tool": "get_workbook_overview"}
    finally:
        pane.stop_flag.set()
        pane.join(5)
    assert pane.ran == [(port, "get_workbook_overview")]


def test_a_live_pairing_or_unfinished_work_is_never_released(world):
    clock, make = world
    (p,) = make(1)
    reg = p.server.registry
    port, code = connect(p)

    # a pending pairing code: someone may be typing it right now
    clock.t += 500
    assert p.release() is False
    token = pane_pair(port, code)

    # a live pane (it keeps starting polls), however long the bridge has existed
    for _ in range(4):
        clock.t += 601
        assert poll(port, token)[0] == 200
        assert p.release() is False

    # a held long-poll is a request being handled: never cut off, even once the pane looks stale
    held = {}
    t = threading.Thread(target=lambda: held.setdefault("r", poll(port, token, wait=20)), daemon=True)
    t.start()
    assert _until(lambda: p.server._requests == 1)
    clock.t += 601
    assert p.server.maybe_release() is False  # direct: the held request must be what keeps it
    assert poll(port, token)[0] == 200  # a newer poll supersedes the held one, which returns empty
    t.join(5)
    assert held["r"][0] == 200 and held["r"][1]["command"] is None

    # a command the pane picked up but has not answered (the edit may be running in Office)
    b = reg.binding_for("chat-1", "excel")
    calls = {}
    worker = threading.Thread(target=lambda: calls.setdefault("r", p.tool("excel_get_workbook_overview", {})),
                              daemon=True)
    worker.start()
    assert _until(lambda: bool(b.queue))
    status, body = poll(port, token)
    rid = body["command"]["requestId"]
    clock.t += 601  # the pane is stale now, but it holds an unanswered command
    assert p.release() is False
    assert req(port, "POST", "/v1/result", {"requestId": rid, "result": {"ok": True, "data": {}}}, token=token)[0] == 200
    worker.join(5)
    assert calls["r"]["ok"] is True

    # a tool call in progress (a OneNote run between two pane commands, a journal write) holds the bridge
    clock.t += 601
    with p.server.in_use():
        assert p.release() is False
    assert p.release() is True and p.server.port is None
    # the pairing itself outlives the release: only the port was given back
    assert reg.binding_for("chat-1", "excel") is b and not b.closed


def test_pane_finds_its_bridge_on_the_port_it_takes_again(world):
    clock, make = world
    a, b = make(2)
    old_port, code = connect(a)
    assert old_port == PORTS[0]
    token = pane_pair(old_port, code)
    instance = hello(old_port)
    assert instance and instance == a.server.instance

    # The user leaves (tab in the background, laptop asleep): the pane stops polling; ten minutes pass.
    clock.t += 601
    assert a.release() is True and a.server.port is None and a.server.releases == 1
    # Another profile pairs and takes the freed port.
    b_port, _ = connect(b)
    assert b_port == old_port and hello(old_port) != instance

    # The pane comes back and searches; nothing answers with its instance until A's next Office tool call.
    pane = Pane(old_port, token, instance)
    pane.start()
    try:
        time.sleep(0.3)
        assert find_bridge(instance) is None and pane.ran == []
        status = a.tool("office_status", {"app": "excel"})["data"]  # takes a port again, waits for the pane
        assert status["bridgeListening"] is True and status["paneState"] == "live", status
        assert a.server.port not in (None, old_port) and a.server.instance == instance
        out = a.tool("excel_get_workbook_overview", {})
        assert out["ok"] is True and out["data"] == {"tool": "get_workbook_overview"}
    finally:
        pane.stop_flag.set()
        pane.join(5)
    assert pane.ran == [(a.server.port, "get_workbook_overview")]
    assert old_port not in pane.tokens_sent_to  # the token never went to profile B's bridge


def test_a_tool_call_on_a_released_bridge_with_no_pairing_starts_nothing(world):
    clock, make = world
    (p,) = make(1)
    port, _ = connect(p)
    clock.t += 601
    assert p.release() is True
    out = p.tool("excel_get_workbook_overview", {})
    assert out["ok"] is False and out["error"]["category"] == "NotPaired"
    assert p.server.port is None and port_is_free(port)


def test_pairing_is_refused_while_a_release_is_closing_the_listener(world):
    clock, make = world
    (p,) = make(1)
    port, code = connect(p)
    reg = p.server.registry
    with reg._lock:
        reg._draining = True
    status, body = req(port, "POST", "/v1/pair", {"code": code, "app": "excel", "document": {}})
    assert status == 400 and body["error"]["category"] == "Unavailable" and "paneToken" not in body
    with reg._lock:
        reg._draining = False
    assert pane_pair(port, code)  # the code was not spent by the refused attempt


def test_a_lingering_keep_alive_connection_cannot_reach_a_released_bridge(world):
    clock, make = world
    a, b = make(2)
    port, _ = connect(a)
    conn = https_client.conn(port, timeout=10)
    assert req(port, "GET", "/v1/hello", conn=conn)[0] == 200  # keep-alive connection, accepted by A
    clock.t += 601
    assert a.release() is True
    assert connect(b)[0] == port
    try:
        status, _ = req(port, "GET", "/v1/hello", conn=conn)
    except (OSError, http.client.HTTPException):
        status = None  # the old handler may also just close the socket
    assert status in (503, None)
    conn.close()
    assert hello(port) == b.server.instance


def test_reaper_thread_releases_an_idle_bridge_and_it_starts_again(bridge_mod):
    s = bridge_mod.BridgeServer(allowed_origin=ORIGIN, ports=range(19960, 19963), idle_release_seconds=0.3,
                                reap_interval=0.05, tls=https_client.local_tls())
    try:
        port = s.ensure_started()
        assert _until(lambda: s.releases == 1, 5)
        assert s.port is None and port_is_free(port)
        assert s.ensure_started() == port
    finally:
        s.stop()


def test_a_port_another_bridge_listens_on_is_never_bound_a_second_time(bridge_mod):
    """W1: on Windows, SO_REUSEADDR let a second bridge bind a port another bridge was already listening on (both binds
    succeeded, and either could answer the pane). A taken port must fail the bind: the next bridge takes the next port,
    and once every port is taken it says PortsBusy."""
    ports = (19963, 19964)
    a, b, c = (bridge_mod.BridgeServer(allowed_origin=ORIGIN, ports=ports, reap_interval=0, tls=https_client.local_tls())
               for _ in range(3))
    try:
        assert a.ensure_started() == 19963
        assert b.ensure_started() == 19964
        with pytest.raises(bridge_mod.BridgeError) as busy:
            c.ensure_started()
        assert busy.value.category == bridge_mod.PORTS_BUSY and c.port is None
        assert hello(19963) == a.instance and hello(19964) == b.instance
        with socket.socket() as s:  # nor can any other socket that asks for address reuse
            s.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEADDR, 1)
            with pytest.raises(OSError):
                s.bind(("127.0.0.1", 19963))
    finally:
        for s in (a, b, c):
            s.stop()


def test_busy_message_names_the_ports_and_the_wait(bridge_mod):
    msg = bridge_mod.ports_busy_message(bridge_mod.DEFAULT_PORTS)
    assert msg == ("All five Workbench bridge ports (127.0.0.1:47615-47619) are in use: five Hermes profiles are "
                   "using Workbench right now (or another program holds a port). Close a Workbench pane that "
                   "another profile's chat paired, or wait up to 10 minutes: a profile with no open pane gives its "
                   "port back after 10 minutes. Then call office_connect again.")


def _until(fn, timeout=5.0):
    end = time.monotonic() + timeout
    while time.monotonic() < end:
        if fn():
            return True
        time.sleep(0.01)
    return False
