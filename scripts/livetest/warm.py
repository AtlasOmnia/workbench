# SPDX-License-Identifier: Apache-2.0
"""Preflight for run_macbook.sh: is the Office bridge reachable from the client Mac, and if not, warm it once.

    PYTHONPATH=<hermes-agent> HERMES_HOME=<profile> <hermes venv python> scripts/livetest/warm.py \
        --mac <host> --origin https://office-addin.example.test

The plugin's bridge listener starts inside the gateway on the profile's first office_connect (office_status
alone reports bridgeListening=false and starts nothing; see __init__.py status_handler). After a gateway
restart the client Mac tunnel to 127.0.0.1:47615 therefore has nothing behind it. This script:
  1. checks /v1/hello through the tunnel (ssh to the client Mac, curl there);
  2. if unreachable, runs ONE chat turn in a new chat of the target profile, through the running gateway
     (chat.py: discovery only, never starts or restarts a gateway; credentials stay with Hermes): office_status
     first (read-only), and office_connect for Excel only if that reports bridgeListening false. The code it
     mints is never shown, typed or logged; it expires unused after 10 minutes;
  3. re-checks a few times, and only then fails with a message that says which of these steps failed.
Exit 0 = reachable (already or after warming), 2 = not reachable.
"""
from __future__ import annotations

import argparse
import asyncio
import json
import subprocess
import sys
import time
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

SERVICE = "hermes-office-desktop"
WARM_PROMPT = ("Live-test preflight. Call office_status for excel. If it reports bridgeListening false, call "
               "office_connect for excel exactly once and do NOT show me the pairing code (nobody will use it). "
               "Do nothing else. Reply with one line: bridgeListening before and whether you called office_connect.")


def hello_ok(body) -> bool:
    """True only for the plugin's own hello (not a captive page, an error page or another service)."""
    if isinstance(body, bytes):
        body = body.decode("utf-8", "replace")
    try:
        data = json.loads(body or "")
    except ValueError:
        return False
    return isinstance(data, dict) and data.get("service") == SERVICE


BRIDGE_PORTS = (47615, 47616, 47617, 47618, 47619)  # bridge_server.DEFAULT_PORTS / addin/shared/bridge-client.js


def ssh_hello(mac: str, origin: str, ports=BRIDGE_PORTS, runner=subprocess.run) -> bool:
    """Does the plugin's hello answer through the client Mac tunnel on any of the bridge ports? A19: a profile's bridge
    takes the first FREE port of five (A14), so it is not always 47615; the pane looks on all five, and so does this.
    One ssh call, one curl per port, one line of output each."""
    if isinstance(ports, int):
        ports = (ports,)
    # HTTPS since 0.4.0. -k: a hello probe only (no token or document data); the client Mac's curl does not have the
    # profile's local CA file, and the pane itself verifies the certificate.
    curls = "; ".join(f"curl -s -k -m5 -H 'Origin: {origin}' https://localhost:{p}/v1/hello; echo" for p in ports)
    cmd = ["ssh", "-o", "BatchMode=yes", "-o", "ConnectTimeout=10", mac, curls]
    try:
        out = runner(cmd, capture_output=True, timeout=60)
    except (OSError, subprocess.SubprocessError):
        return False
    body = out.stdout.decode("utf-8", "replace") if isinstance(out.stdout, bytes) else (out.stdout or "")
    return any(hello_ok(line) for line in body.splitlines())


def preflight(check, warm, sleep=time.sleep, rechecks: int = 5, wait: float = 3.0):
    """Decision logic (pure: check/warm/sleep are injected). Returns (ok, how, message).

    how: 'reachable' (no warm needed), 'warmed' (reachable after the one warm turn), 'warm-failed' (the warm
    turn itself failed), 'unreachable' (warm turn ran, bridge still not reachable through the tunnel).
    The warm turn runs at most once."""
    if check():
        return True, "reachable", "bridge reachable from the client Mac"
    try:
        note = warm()
    except Exception as exc:  # chat/gateway error: say so, do not retry the turn
        return False, "warm-failed", (f"bridge not reachable from the client Mac (127.0.0.1:47615-47619 tunnel), and the one warm "
                                      f"chat turn failed ({type(exc).__name__}: {str(exc)[:200]}); stopping. Is the "
                                      "gateway serving this profile?")
    for _ in range(max(1, rechecks)):
        sleep(wait)
        if check():
            return True, "warmed", f"bridge reachable after one warm chat turn ({note})"
    return False, "unreachable", (f"bridge still not reachable from the client Mac after one warm chat turn ({note}); the "
                                  "bridge may be listening on the Studio but the client Mac tunnel to 127.0.0.1:47615-47619 is "
                                  "down or points elsewhere. Stopping (the harness never touches tunnels).")


def summarize_warm(turn) -> str:
    """What the warm turn did, from its tool calls only (no reply text: it could contain a pairing code)."""
    names = [t.get("name") for t in turn.get("tools") or []]
    listening = None
    for t in turn.get("tools") or []:
        r = t.get("result")
        if t.get("name") == "office_status" and isinstance(r, dict):
            listening = (r.get("data") or {}).get("bridgeListening")
    connected = any(t.get("name") == "office_connect" and isinstance(t.get("result"), dict) and t["result"].get("ok")
                    for t in turn.get("tools") or [])
    return f"tools={names}, bridgeListening before={listening}, office_connect ok={connected}"


def gateway_warm(log=lambda obj: None) -> str:
    from chat import Chat, connect_live

    async def run():
        async with connect_live() as client:
            chat = await Chat.create(client, f"HOL preflight warm {time.strftime('%m%d-%H%M%S')}", log)
            return summarize_warm(await chat.turn(WARM_PROMPT, timeout=180))
    return asyncio.run(run())


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--mac", required=True)
    ap.add_argument("--origin", required=True)
    ap.add_argument("--port", type=int, default=None, help="check only this port (default: all five bridge ports)")
    a = ap.parse_args()

    def log(obj):  # tool names and timings only; reply text is never printed
        if obj.get("ev") in ("chat", "tool", "approval"):
            print("  warm:", json.dumps({k: v for k, v in obj.items() if k != "text"})[:300], flush=True)
    ports = (a.port,) if a.port else BRIDGE_PORTS
    ok, how, msg = preflight(lambda: ssh_hello(a.mac, a.origin, ports), lambda: gateway_warm(log))
    print(f"preflight: {how}: {msg}", file=sys.stdout if ok else sys.stderr, flush=True)
    return 0 if ok else 2


if __name__ == "__main__":
    sys.exit(main())
