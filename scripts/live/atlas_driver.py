# SPDX-License-Identifier: Apache-2.0
"""Live Office.js run on a Windows test machine: real Excel/Word/PowerPoint, this repo's real bridge.

    python scripts/live/atlas_driver.py --app excel --cdp-port 9333 --origin https://localhost:3010 --out run/results-excel.json

Preconditions (see docs/harness.md "Live testing on Windows"): the add-in build is served on the origin,
its manifest is sideloaded per user, and the host was started with WebView2 remote debugging on
--cdp-port so this script can type the pairing code into the pane and take document probes. No
Hermes is needed: the plugin module is loaded directly and its tool handlers are called as Hermes
would call them (approval prompts are Hermes' job and are not exercised here).
"""
from __future__ import annotations

import argparse
import importlib.util
import json
import subprocess
import sys
import time
from pathlib import Path

HERE = Path(__file__).resolve().parent
ROOT = HERE.parent.parent
sys.path.insert(0, str(HERE))
from plan import PLANS  # noqa: E402
from runner import Runner, summary  # noqa: E402


def load_plugin():
    spec = importlib.util.spec_from_file_location("hermes_office_desktop", ROOT / "__init__.py",
                                                  submodule_search_locations=[str(ROOT)])
    mod = importlib.util.module_from_spec(spec)
    sys.modules["hermes_office_desktop"] = mod
    spec.loader.exec_module(mod)
    return mod


class FakeCtx:
    def __init__(self, config):
        self.tools, self.hooks, self.config = {}, {}, config

    def register_tool(self, **kw):
        self.tools[kw["name"]] = kw

    def register_hook(self, name, fn):
        self.hooks[name] = fn

    def get_config(self, key, default=None):
        return self.config.get(key, default)


class Cdp:
    def __init__(self, port: int, match: str = "taskpane.html"):
        self.port, self.match = port, match
        self.tmp = ROOT / "build" / "cdp-expr.js"

    def evaluate(self, expr: str, timeout=150):
        self.tmp.parent.mkdir(parents=True, exist_ok=True)
        self.tmp.write_text(expr, encoding="utf-8")
        p = subprocess.run(["node", str(HERE / "cdp.mjs"), "eval", str(self.port), self.match, str(self.tmp)],
                           capture_output=True, text=True, timeout=timeout)
        out = json.loads(p.stdout.strip() or "null")
        if isinstance(out, dict) and "cdpError" in out:
            raise RuntimeError(out["cdpError"])
        return out

    def state(self):
        try:
            return self.evaluate("globalThis.__hermesPaneState || null", timeout=20)
        except Exception:  # noqa: BLE001  (page navigating)
            return None

    def wait_state(self, wanted, timeout=60):
        end = time.time() + timeout
        last = None
        while time.time() < end:
            last = self.state()
            if last in wanted:
                return last
            time.sleep(0.5)
        return last


def main(argv=None):
    ap = argparse.ArgumentParser()
    ap.add_argument("--app", required=True, choices=sorted(PLANS))
    ap.add_argument("--cdp-port", type=int, required=True)
    ap.add_argument("--origin", default="https://localhost:3010")
    ap.add_argument("--out", required=True)
    a = ap.parse_args(argv)
    cdp = Cdp(a.cdp_port)
    plugin = load_plugin()
    ctx = FakeCtx({"allowed_origin": a.origin})
    plugin.register(ctx)
    session = f"atlas-live-{a.app}-{int(time.time())}"
    log_lines: list[str] = []

    def log(s):
        print(s, flush=True)
        log_lines.append(s)

    if cdp.wait_state({"ready", "unpaired"}, timeout=60) not in ("ready", "unpaired"):
        raise SystemExit(f"pane not ready (state={cdp.state()})")
    env = cdp.evaluate("({ host: Office.context.host, platform: Office.context.platform, version: Office.context.diagnostics.version, "
                   "ua: navigator.userAgent, url: location.origin + location.pathname })")
    log(f"pane: {env}")
    t0 = time.perf_counter()
    conn = json.loads(ctx.tools["office_connect"]["handler"]({"app": a.app}, session_id=session))
    assert conn["ok"], conn
    code = conn["data"]["pairingCode"]
    cdp.evaluate("(() => { const i = document.getElementById('code'); i.value = %s; i.dispatchEvent(new Event('input')); "
             "document.getElementById('pair-form').requestSubmit(); return true })()" % json.dumps(code))
    st = cdp.wait_state({"paired", "rejected", "no-bridge"}, timeout=30)
    pair_ms = round((time.perf_counter() - t0) * 1000, 1)
    log(f"pairing: state={st} in {pair_ms}ms")
    if st != "paired":
        raise SystemExit("pairing failed")
    status = json.loads(ctx.tools["office_status"]["handler"]({"app": a.app}, session_id=session))
    probe_src = (HERE / f"probe-{a.app}.js").read_text(encoding="utf-8")
    probe_ms: list[float] = []

    def probe():
        t = time.perf_counter()
        v = cdp.evaluate(probe_src)
        probe_ms.append(round((time.perf_counter() - t) * 1000, 1))
        return v

    def reload_pane():
        cdp.evaluate("(setTimeout(() => location.reload(), 50), true)")
        time.sleep(1.5)
        s = cdp.wait_state({"paired"}, timeout=60)
        if s != "paired":
            raise RuntimeError(f"pane did not resume after reload (state={s})")
        resumed = json.loads(ctx.tools["office_status"]["handler"]({"app": a.app}, session_id=session))["data"]
        log(f"reload: resumed, paneReloads={resumed['paneReloads']}")

    def run_js(step):
        return cdp.evaluate(step["live"])

    runner = Runner(ctx.tools, session, a.app, probe, reload_pane, run_js, log=log)
    results = runner.run(PLANS[a.app]())
    ctx.tools["office_disconnect"]["handler"]({"app": a.app}, session_id=session)
    out = {"app": a.app, "environment": env, "pairing": {"state": st, "ms": pair_ms, "status": status.get("data")},
           "summary": summary(results), "probeMs": {"n": len(probe_ms), "median": sorted(probe_ms)[len(probe_ms) // 2] if probe_ms else None},
           "results": results}
    Path(a.out).write_text(json.dumps(out, indent=2, ensure_ascii=False), encoding="utf-8")
    log(json.dumps(out["summary"]))


if __name__ == "__main__":
    main()
