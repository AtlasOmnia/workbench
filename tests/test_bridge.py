# SPDX-License-Identifier: Apache-2.0
"""Bridge server: pairing, origin/host gating, command round trip, failure shapes."""
import http.client
import json
import threading
import time

import pytest

ORIGIN = "https://addin.test:8444"


@pytest.fixture()
def server(bridge_mod):
    s = bridge_mod.BridgeServer(allowed_origin=ORIGIN, ports=range(18765, 18775))
    s.ensure_started()
    yield s
    s.stop()


def req(server, method, path, body=None, origin=ORIGIN, token=None, host=None, ctype="application/json"):
    c = http.client.HTTPConnection("127.0.0.1", server.port, timeout=30)
    headers = {"Host": host or f"127.0.0.1:{server.port}"}
    if origin is not None:
        headers["Origin"] = origin
    if token:
        headers["Authorization"] = f"Bearer {token}"
    data = None
    if body is not None:
        data = json.dumps(body).encode()
        headers["Content-Type"] = ctype
        headers["Content-Length"] = str(len(data))
    c.request(method, path, body=data, headers=headers)
    r = c.getresponse()
    raw = r.read()
    out = (r.status, json.loads(raw) if raw else None, dict(r.getheaders()))
    c.close()
    return out


def test_hello_requires_exact_origin_and_host(server):
    s, body, h = req(server, "GET", "/v1/hello")
    assert s == 200 and body["service"] == "hermes-office-desktop"
    assert h["Access-Control-Allow-Origin"] == ORIGIN
    assert req(server, "GET", "/v1/hello", origin="https://evil.example")[0] == 403
    assert req(server, "GET", "/v1/hello", origin=None)[0] == 403
    assert req(server, "GET", "/v1/hello", host="localhost.evil:80")[0] == 421  # DNS-rebinding guard


def test_preflight_answers_private_network(server):
    c = http.client.HTTPConnection("127.0.0.1", server.port)
    c.request("OPTIONS", "/v1/pair", headers={"Host": f"127.0.0.1:{server.port}", "Origin": ORIGIN,
                                               "Access-Control-Request-Method": "POST",
                                               "Access-Control-Request-Private-Network": "true"})
    r = c.getresponse()
    r.read()
    assert r.status == 204
    assert r.getheader("Access-Control-Allow-Private-Network") == "true"
    assert r.getheader("Access-Control-Allow-Origin") == ORIGIN


def test_pairing_is_single_use_and_bound_to_minting_session(server):
    code, ttl = server.registry.mint_code("sess-A", "excel")
    assert len(code) == 8 and ttl == 600
    s, body, _ = req(server, "POST", "/v1/pair", {"code": f"{code[:4]}-{code[4:].lower()}", "app": "excel",
                                                   "document": {"name": "Budget.xlsx"},
                                                   "supportedTools": ["read_range"]})
    assert s == 200 and body["ok"] and len(body["paneToken"]) >= 40
    assert server.registry.binding_for("sess-A", "excel").document["name"] == "Budget.xlsx"
    assert server.registry.binding_for("sess-B", "excel") is None
    s2, body2, _ = req(server, "POST", "/v1/pair", {"code": code, "app": "excel"})
    assert s2 == 401 and body2["error"]["category"] == "Authentication"


def test_wrong_codes_burn_live_code_and_rate_limit(server, bridge_mod):
    code, _ = server.registry.mint_code("sess-A", "excel")
    for _ in range(bridge_mod.MAX_CODE_FAILURES):
        assert req(server, "POST", "/v1/pair", {"code": "ZZZZZZZZ", "app": "excel"})[0] == 401
    assert req(server, "POST", "/v1/pair", {"code": code, "app": "excel"})[0] == 401  # burned
    for _ in range(bridge_mod.MAX_GLOBAL_FAILURES_PER_MINUTE):
        req(server, "POST", "/v1/pair", {"code": "ZZZZZZZZ", "app": "excel"})
    assert req(server, "POST", "/v1/pair", {"code": "ZZZZZZZZ", "app": "excel"})[0] == 429


def test_expired_code_rejected(bridge_mod):
    now = [1000.0]
    reg = bridge_mod.Registry(clock=lambda: now[0])
    code, ttl = reg.mint_code("s", "excel")
    now[0] += ttl + 1
    with pytest.raises(bridge_mod.BridgeError):
        reg.pair(code, "excel", {}, {}, [])


def _pane(server, token, handler, stop):
    while not stop.is_set():
        s, body, _ = req(server, "GET", "/v1/next?wait=2", token=token)
        if s != 200:
            return
        cmd = body["command"]
        if cmd:
            req(server, "POST", "/v1/result", {"requestId": cmd["requestId"], "result": handler(cmd)}, token=token)


def test_command_round_trip_and_isolation(server):
    code, _ = server.registry.mint_code("sess-A", "excel")
    token = req(server, "POST", "/v1/pair", {"code": code, "app": "excel"})[1]["paneToken"]
    stop = threading.Event()
    t = threading.Thread(target=_pane, args=(server, token, lambda c: {"ok": True, "data": {"echo": c["tool"], "args": c["arguments"]}}, stop), daemon=True)
    t.start()
    out = server.registry.call("sess-A", "excel", "read_range", {"sheet": "S", "address": "A1"}, timeout=10)
    assert isinstance(out.pop("roundTripMs"), int)
    assert out == {"ok": True, "data": {"echo": "read_range", "args": {"sheet": "S", "address": "A1"}}}
    with pytest.raises(Exception) as e:
        server.registry.call("sess-B", "excel", "read_range", {}, timeout=1)
    assert e.value.category == "NotPaired"
    # Pane cannot use a forged token, nor complete someone else's request.
    assert req(server, "GET", "/v1/next?wait=0", token="forged")[0] == 401
    assert req(server, "POST", "/v1/result", {"requestId": "nope", "result": {"ok": True}}, token=token)[0] == 400
    stop.set()
    t.join(5)


def test_unsupported_tool_fails_fast(server):
    code, _ = server.registry.mint_code("s", "excel")
    req(server, "POST", "/v1/pair", {"code": code, "app": "excel", "supportedTools": ["read_range"]})
    with pytest.raises(Exception) as e:
        server.registry.call("s", "excel", "delete_sheet", {}, timeout=1)
    assert e.value.category == "Unsupported"


def test_timeout_distinguishes_never_delivered_from_uncertain(server):
    code, _ = server.registry.mint_code("s", "excel")
    token = req(server, "POST", "/v1/pair", {"code": code, "app": "excel"})[1]["paneToken"]
    with pytest.raises(Exception) as e:
        server.registry.call("s", "excel", "read_range", {}, timeout=0.3)
    assert e.value.category == "Timeout"  # never picked up: safe, not executed
    result = {}

    def call():
        try:
            server.registry.call("s", "excel", "write_range", {}, timeout=1.5)
        except Exception as exc:  # noqa: BLE001
            result["cat"] = exc.category
    th = threading.Thread(target=call)
    th.start()
    time.sleep(0.2)
    s, body, _ = req(server, "GET", "/v1/next?wait=1", token=token)  # picked up, never answered
    assert body["command"]["tool"] == "write_range"
    th.join(5)
    assert result["cat"] == "ExecutionUncertain"


@pytest.mark.parametrize("app,tool,thing", [
    ("excel", "write_range", "workbook"), ("word", "insert_text", "document"),
    ("powerpoint", "create_text_slide", "presentation"), ("onenote", "append", "page"),
    ("outlook", "draft_reply", "draft"), ("outlook", "update_event", "event")])
def test_uncertain_edit_names_what_to_check_for_each_app(server, app, tool, thing):
    """F8 D4: an edit picked up but never answered said "Check the workbook before retrying." in every app."""
    code, _ = server.registry.mint_code("s", app)
    token = req(server, "POST", "/v1/pair", {"code": code, "app": app})[1]["paneToken"]
    result = {}

    def call():
        try:
            server.registry.call("s", app, tool, {}, timeout=1.0)
        except Exception as exc:  # noqa: BLE001
            result["cat"], result["msg"] = exc.category, str(exc)
    th = threading.Thread(target=call)
    th.start()
    time.sleep(0.2)
    assert req(server, "GET", "/v1/next?wait=1", token=token)[1]["command"]["tool"] == tool  # picked up, never answered
    th.join(5)
    assert result["cat"] == "ExecutionUncertain" and "may or may not have applied" in result["msg"]
    assert f"Check the {thing} before retrying." in result["msg"], result["msg"]
    if app != "excel":
        assert "workbook" not in result["msg"]


def test_a_long_read_is_pane_busy_not_an_uncertain_edit_and_the_next_command_says_why(server):
    """A26 cold run: onenote_list_pages(all) ran past 60 s and came back as "the edit may or may not have applied.
    Check the workbook" (a read, in OneNote); the next command then timed out undelivered without saying why."""
    code, _ = server.registry.mint_code("s", "onenote")
    token = req(server, "POST", "/v1/pair", {"code": code, "app": "onenote"})[1]["paneToken"]
    result = {}

    def call():
        try:
            server.registry.call("s", "onenote", "list_pages", {"all": True}, timeout=1.0)
        except Exception as exc:  # noqa: BLE001
            result["cat"], result["msg"] = exc.category, str(exc)
    th = threading.Thread(target=call)
    th.start()
    time.sleep(0.2)
    body = req(server, "GET", "/v1/next?wait=1", token=token)[1]  # picked up; the pane is still reading
    assert body["command"]["tool"] == "list_pages"
    th.join(5)
    assert result["cat"] == "PaneBusy" and "nothing was changed" in result["msg"] and "workbook" not in result["msg"]
    with pytest.raises(Exception) as e:
        server.registry.call("s", "onenote", "create_page", {"title": "t"}, timeout=0.3)
    assert e.value.category == "Timeout" and "still busy with an earlier request" in str(e.value)

def test_repair_replaces_previous_pane(server):
    c1, _ = server.registry.mint_code("s", "excel")
    t1 = req(server, "POST", "/v1/pair", {"code": c1, "app": "excel"})[1]["paneToken"]
    c2, _ = server.registry.mint_code("s", "excel")
    t2 = req(server, "POST", "/v1/pair", {"code": c2, "app": "excel"})[1]["paneToken"]
    assert t1 != t2
    assert req(server, "GET", "/v1/next?wait=0", token=t1)[0] == 401
    assert req(server, "GET", "/v1/next?wait=0", token=t2)[0] == 200


def test_body_limits_and_content_type(server, bridge_mod):
    assert req(server, "POST", "/v1/pair", {"code": "x"}, ctype="text/plain")[0] == 400
    c = http.client.HTTPConnection("127.0.0.1", server.port)
    c.putrequest("POST", "/v1/pair", skip_host=True)
    c.putheader("Host", f"127.0.0.1:{server.port}")
    c.putheader("Origin", ORIGIN)
    c.putheader("Content-Type", "application/json")
    c.putheader("Content-Length", str(bridge_mod.MAX_BODY_BYTES + 1))
    c.endheaders()
    r = c.getresponse()
    assert r.status == 413


def test_rejects_non_https_origin_config(bridge_mod):
    for bad in ("http://atlasomnia.github.io", "https://atlasomnia.github.io/workbench", "*"):
        with pytest.raises(bridge_mod.BridgeError):
            bridge_mod.BridgeServer(allowed_origin=bad)


def _pair(server, sess, app="excel", **extra):
    code, _ = server.registry.mint_code(sess, app)
    s, body, _ = req(server, "POST", "/v1/pair", {"code": code, "app": app, **extra})
    assert s == 200, body
    return body["paneToken"]


def test_all_five_apps_pair_independently(server):
    s, body, _ = req(server, "GET", "/v1/hello")
    assert body["apps"] == ["excel", "word", "powerpoint", "onenote", "outlook"]
    tokens = {app: _pair(server, "chat", app) for app in ("excel", "word", "powerpoint", "onenote", "outlook")}
    assert len(set(tokens.values())) == 5
    for app in tokens:
        assert server.registry.binding_for("chat", app).app == app
    # A Word code cannot pair an Excel pane.
    code, _ = server.registry.mint_code("chat", "word")
    assert req(server, "POST", "/v1/pair", {"code": code, "app": "excel"})[0] == 401


def test_reload_resumes_same_chat_and_fails_inflight_fast(server):
    token = _pair(server, "chat-A", supportedTools=["read_range"])
    result = {}

    def call():
        out = server.registry.call("chat-A", "excel", "read_range", {}, timeout=20)
        result["cat"] = out["error"]["category"]
    th = threading.Thread(target=call)
    th.start()
    time.sleep(0.2)
    assert req(server, "GET", "/v1/next?wait=1", token=token)[1]["command"]["tool"] == "read_range"
    t0 = time.monotonic()
    # The page reloads mid-command and resumes with its tab-scoped token.
    s, body, _ = req(server, "POST", "/v1/resume", {"app": "excel", "supportedTools": ["read_range", "write_range"],
                                                    "document": {"name": "Book.xlsx"}}, token=token)
    assert s == 200 and body["ok"]
    th.join(5)
    assert result["cat"] == "ExecutionUncertain" and time.monotonic() - t0 < 3
    b = server.registry.binding_for("chat-A", "excel")
    assert b.resumes == 1 and "write_range" in b.supported_tools and b.document["name"] == "Book.xlsx"
    # Still bound to the same chat; still the same token.
    assert req(server, "GET", "/v1/next?wait=0", token=token)[0] == 200
    # Resume with an unknown token or the wrong app is refused.
    assert req(server, "POST", "/v1/resume", {"app": "excel"}, token="forged")[0] == 401
    assert req(server, "POST", "/v1/resume", {"app": "word"}, token=token)[0] == 401


def test_one_pane_one_chat(server):
    """A pane re-pairing with another chat (sending its previous token) leaves the first chat unpaired."""
    t1 = _pair(server, "chat-A")
    t2 = _pair(server, "chat-B", previousToken=t1)
    assert server.registry.binding_for("chat-A", "excel") is None
    assert server.registry.binding_for("chat-B", "excel").token == t2
    assert req(server, "GET", "/v1/next?wait=0", token=t1)[0] == 401
    with pytest.raises(Exception) as e:
        server.registry.call("chat-A", "excel", "read_range", {}, timeout=1)
    assert e.value.category == "NotPaired"


def test_idle_binding_expires(bridge_mod):
    now = [1000.0]
    reg = bridge_mod.Registry(clock=lambda: now[0])
    code, _ = reg.mint_code("s", "excel")
    b = reg.pair(code, "excel", {}, {}, [])
    now[0] += bridge_mod.BINDING_IDLE_SECONDS + 1
    with pytest.raises(bridge_mod.BridgeError) as e:
        reg.authenticate(b.token)
    assert e.value.category == "Authentication"
    assert reg.binding_for("s", "excel") is None


def test_stale_pane_is_unavailable_not_hanging(bridge_mod):
    now = [1000.0]
    reg = bridge_mod.Registry(clock=lambda: now[0])
    code, _ = reg.mint_code("s", "excel")
    reg.pair(code, "excel", {}, {}, [])
    now[0] += bridge_mod.PANE_STALE_SECONDS + 1
    with pytest.raises(bridge_mod.BridgeError) as e:
        reg.call("s", "excel", "read_range", {}, timeout=1)
    assert e.value.category == "Unavailable"


def test_abandoned_long_poll_never_takes_a_command(server):
    """A reloaded page leaves its old /v1/next open; that zombie poll must not swallow the next command."""
    token = _pair(server, "chat-Z", supportedTools=["read_range"])
    zombie = {}

    def old_page_poll():
        zombie["resp"] = req(server, "GET", "/v1/next?wait=8", token=token)[1]
    tz = threading.Thread(target=old_page_poll, daemon=True)
    tz.start()
    time.sleep(0.3)
    # The page reloads and resumes; its new loop polls.
    assert req(server, "POST", "/v1/resume", {"app": "excel"}, token=token)[0] == 200
    got = {}

    def new_page_poll():
        cmd = req(server, "GET", "/v1/next?wait=8", token=token)[1]["command"]
        got["tool"] = cmd["tool"]
        req(server, "POST", "/v1/result", {"requestId": cmd["requestId"], "result": {"ok": True, "data": {}}}, token=token)
    tn = threading.Thread(target=new_page_poll, daemon=True)
    tn.start()
    time.sleep(0.3)
    out = server.registry.call("chat-Z", "excel", "read_range", {}, timeout=5)
    assert isinstance(out.pop("roundTripMs"), int)
    assert out == {"ok": True, "data": {}}
    tn.join(5)
    tz.join(10)
    assert got["tool"] == "read_range"
    assert zombie["resp"]["command"] is None


def test_results_carry_pane_elapsed_and_plugin_round_trip(server):
    """F1: every pane-executed result reports elapsedMs (pane) and roundTripMs (plugin)."""
    token = _pair(server, "chat-T", supportedTools=["read_range", "write_range"])
    replies = iter([{"ok": True, "data": {"v": 1}, "elapsedMs": 12.6},
                    {"ok": True, "data": {"v": 2}, "elapsedMs": "fast"},  # not a number: dropped
                    {"ok": False, "error": {"category": "Office", "message": "x"}, "elapsedMs": -3}])
    stop = threading.Event()
    t = threading.Thread(target=_pane, args=(server, token, lambda c: next(replies), stop), daemon=True)
    t.start()
    out = server.registry.call("chat-T", "excel", "read_range", {}, timeout=10)
    assert out["data"] == {"v": 1} and out["elapsedMs"] == 13
    assert isinstance(out["roundTripMs"], int) and out["roundTripMs"] >= 0
    out = server.registry.call("chat-T", "excel", "write_range", {}, timeout=10)
    assert "elapsedMs" not in out and isinstance(out["roundTripMs"], int)
    out = server.registry.call("chat-T", "excel", "read_range", {}, timeout=10)
    assert out["ok"] is False and "elapsedMs" not in out and isinstance(out["roundTripMs"], int)
    stop.set()
    t.join(5)


def test_liveness_counts_poll_arrivals_not_held_polls(bridge_mod):
    """F4: a pane that stopped polling is not live, even while its binding still exists."""
    now = [1000.0]
    reg = bridge_mod.Registry(clock=lambda: now[0])
    code, _ = reg.mint_code("s", "excel")
    b = reg.pair(code, "excel", {}, {}, [])
    assert reg.liveness(b) == {"lastPollAgeMs": 0, "live": True}
    reg.next_command(b, 0)  # a poll arrives
    now[0] += 30
    assert reg.liveness(b) == {"lastPollAgeMs": 30000, "live": True}
    now[0] += 16  # 46 s since the last poll arrived
    assert reg.liveness(b) == {"lastPollAgeMs": 46000, "live": False}
    with pytest.raises(bridge_mod.BridgeError) as e:
        reg.call("s", "excel", "read_range", {}, timeout=1)
    assert e.value.category == "Unavailable" and "switched chat or closed" in str(e.value)
    reg.next_command(b, 0)  # the pane is back (reopened, resumed)
    assert reg.liveness(b)["live"] is True


def test_held_long_poll_does_not_refresh_liveness(server):
    """The server-side wait loop of a poll whose page is gone must not keep the pane 'live'."""
    token = _pair(server, "chat-H")
    b = server.registry.binding_for("chat-H", "excel")
    req(server, "GET", "/v1/next?wait=0", token=token)
    arrived = b.last_poll
    t = threading.Thread(target=lambda: req(server, "GET", "/v1/next?wait=2", token=token), daemon=True)
    t.start()
    time.sleep(0.2)
    first = b.last_poll
    time.sleep(1.2)
    assert b.last_poll == first >= arrived  # held for 1.2 s: arrival time unchanged
    t.join(5)


def test_an_abandoned_command_keeps_the_pane_live_for_a_grace_period_only(server, bridge_mod):
    """A19: a picked-up command Hermes stopped waiting for is still running in the pane: busy, not gone. Its late
    result is accepted (no Protocol error), and after ABANDONED_GRACE_SECONDS it no longer counts."""
    code, _ = server.registry.mint_code("s", "excel")
    token = req(server, "POST", "/v1/pair", {"code": code, "app": "excel"})[1]["paneToken"]
    reg = server.registry
    th = threading.Thread(target=lambda: pytest.raises(Exception, reg.call, "s", "excel", "read_range", {}, timeout=0.6))
    th.start()
    time.sleep(0.2)
    cmd = req(server, "GET", "/v1/next?wait=1", token=token)[1]["command"]
    th.join(5)
    b = reg.binding_for("s", "excel")
    b.last_poll -= 100
    assert reg.liveness(b)["live"] is True
    s, body, _ = req(server, "POST", "/v1/result", {"requestId": cmd["requestId"], "result": {"ok": True, "data": {}}},
                     token=token)
    assert s == 200
    th2 = threading.Thread(target=lambda: pytest.raises(Exception, reg.call, "s", "excel", "read_range", {}, timeout=0.6))
    th2.start()
    time.sleep(0.2)
    req(server, "GET", "/v1/next?wait=1", token=token)
    th2.join(5)
    b.last_poll -= 100
    for c in b.inflight.values():
        c.abandoned_at -= bridge_mod.ABANDONED_GRACE_SECONDS + 1
    assert reg.liveness(b)["live"] is False and not b.inflight
