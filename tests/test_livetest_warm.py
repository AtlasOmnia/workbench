# SPDX-License-Identifier: Apache-2.0
"""run_macbook.sh's self-warming preflight (scripts/livetest/warm.py): decision logic only, no network or gateway."""
import importlib.util
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
LIVE = ROOT / "scripts" / "livetest"


def _warm():
    name = "hol_warm"
    if name in sys.modules:
        return sys.modules[name]
    sys.path.insert(0, str(LIVE))
    spec = importlib.util.spec_from_file_location(name, LIVE / "warm.py")
    mod = importlib.util.module_from_spec(spec)
    sys.modules[name] = mod
    spec.loader.exec_module(mod)
    return mod


# ---- self-warming preflight ------------------------------------------------------------------------------
def test_preflight_skips_warming_when_reachable():
    w = _warm()
    warmed = []
    assert w.preflight(lambda: True, lambda: warmed.append(1), sleep=lambda s: None)[:2] == (True, "reachable")
    assert warmed == []


def test_preflight_warms_once_then_rechecks():
    w = _warm()
    answers = iter([False, False, True])
    warmed = []
    ok, how, msg = w.preflight(lambda: next(answers), lambda: warmed.append(1) or "tools=[office_status]", sleep=lambda s: None)
    assert (ok, how) == (True, "warmed") and warmed == [1] and "office_status" in msg


def test_preflight_fails_clearly_after_one_warm_or_a_failed_warm():
    w = _warm()
    warmed = []
    ok, how, msg = w.preflight(lambda: False, lambda: warmed.append(1) or "n", sleep=lambda s: None, rechecks=3)
    assert (ok, how) == (False, "unreachable") and warmed == [1] and "tunnel" in msg

    def boom():
        raise RuntimeError("gateway for this profile is not ready (absent)")
    ok, how, msg = w.preflight(lambda: False, boom, sleep=lambda s: None)
    assert (ok, how) == (False, "warm-failed") and "not ready" in msg


def test_hello_ok_and_warm_summary_never_echo_a_code():
    w = _warm()
    assert w.hello_ok(b'{"service":"hermes-office-desktop","v":1}')
    assert not w.hello_ok(b"<html>captive</html>") and not w.hello_ok(b'{"service":"other"}') and not w.hello_ok(b"")
    turn = {"text": "Code ABCD-EFGH", "tools": [
        {"name": "office_status", "result": {"ok": True, "data": {"bridgeListening": False}}},
        {"name": "office_connect", "result": {"ok": True, "data": {"pairingCode": "ABCD-EFGH"}}}]}
    s = w.summarize_warm(turn)
    assert "bridgeListening before=False" in s and "office_connect ok=True" in s and "ABCD" not in s
    assert "office_status" in w.WARM_PROMPT and "do NOT show" in w.WARM_PROMPT


def test_ssh_hello_uses_the_tunnel_and_survives_ssh_errors():
    w = _warm()
    seen = []

    class Out:
        stdout = b'{"service":"hermes-office-desktop"}'

    def runner(cmd, **kw):
        seen.append(cmd)
        return Out()
    assert w.ssh_hello("mac", "https://o", runner=runner)
    assert seen[0][:2] == ["ssh", "-o"] and "127.0.0.1:47615/v1/hello" in seen[0][-1] and "Origin: https://o" in seen[0][-1]

    def broken(cmd, **kw):
        raise OSError("no ssh")
    assert w.ssh_hello("mac", "https://o", runner=broken) is False


def test_the_preflight_finds_the_bridge_on_any_of_the_five_ports():
    """Seen live (A19, F df7b3a6): after another profile's idle bridge had held 47615, this profile's bridge listened
    on 47616 (A14: the first free port of five), and a 47615-only preflight stopped every later mode. The pane looks on
    all five ports, so the preflight does too, in one ssh call, and passes only on the plugin's own hello."""
    w = _warm()
    seen = []

    class Out:
        def __init__(self, stdout):
            self.stdout = stdout

    def runner(cmd, **kw):
        seen.append(cmd)
        return Out(b'\n{"service":"hermes-office-desktop","protocol":1}\n\n\n\n')
    assert w.ssh_hello("mac", "https://o", runner=runner)
    assert len(seen) == 1 and all(f"127.0.0.1:{p}/v1/hello" in seen[0][-1] for p in range(47615, 47620))
    assert w.ssh_hello("mac", "https://o", runner=lambda cmd, **kw: Out(b"<html>captive</html>\n\n\n\n\n")) is False
    assert w.ssh_hello("mac", "https://o", runner=lambda cmd, **kw: Out(b"")) is False
