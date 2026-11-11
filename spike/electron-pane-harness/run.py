# SPDX-License-Identifier: Apache-2.0
"""Run the Electron pane harness end to end (spike evidence, local only).

    uv run --no-project --python 3.12 --with pyyaml python spike/electron-pane-harness/run.py [variant ...]

    ... run.py [--hosts=excel,word,powerpoint] [variant ...]   (default: all three hosts, allow-default)

Each host runs the shared coverage plan (scripts/live/plan.py): every tool, undo round trips
checked against a whole-document fingerprint, a pane reload that must resume the same chat
binding, and a cross-chat refusal.

Variants: allow-default (Office iframe delegates local-network-access; no permission
handler, like the Desktop preview partition), noallow-default (no delegation, the
pre-fix / OWA shape), allow-deny (delegated, but a deny-all permission handler),
allow-log (delegated, grant + log permission names).

Topology: https://office.test:8443 (simulated Office for the web, PUBLIC) iframes
https://addin.test:8444/office/live/excel/taskpane.html?mock=1 (the real bridge
add-in build, PUBLIC) which must reach the real plugin bridge on 127.0.0.1 (LOOPBACK).
The only simulated parts are the Office host page and Office.js (strict mock).
"""
from __future__ import annotations

import importlib.util
import json
import os
import queue
import ssl
import subprocess
import sys
import threading
import time
from http.server import SimpleHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import parse_qs, urlsplit

ROOT = Path(__file__).resolve().parents[2]
HERE = Path(__file__).resolve().parent
ELECTRON = Path(os.environ.get("HARNESS_ELECTRON",
                               str(Path.home() / ".hermes/hermes-agent/node_modules/electron/dist/Electron.app/Contents/MacOS/Electron")))
SCRATCH = Path(os.environ.get("HARNESS_SCRATCH", str(Path.home() / ".hermes/profiles/omh/cache/scratch/office-desktop-harness")))
OFFICE_PORT, ADDIN_PORT = 8443, 8444
ADDIN_ORIGIN = f"https://addin.test:{ADDIN_PORT}"


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

    def register_skill(self, name, path):
        pass

    def register_automation_blueprint(self, key, **kw):
        pass

    def register_tool(self, **kw):
        self.tools[kw["name"]] = kw

    def register_hook(self, name, fn):
        self.hooks[name] = fn

    def get_config(self, key, default=None):
        return self.config.get(key, default)


def make_cert(out: Path):
    key, crt = out / "harness.key", out / "harness.crt"
    if not crt.exists():
        subprocess.run(["openssl", "req", "-x509", "-newkey", "rsa:2048", "-nodes", "-days", "2",
                        "-subj", "/CN=office.test", "-addext", "subjectAltName=DNS:office.test,DNS:addin.test",
                        "-keyout", str(key), "-out", str(crt)], check=True, capture_output=True)
    return key, crt


def office_handler_factory():
    class Office(SimpleHTTPRequestHandler):
        def log_message(self, *a):
            pass

        def do_GET(self):
            u = urlsplit(self.path)
            if u.path != "/office.html":
                self.send_error(404)
                return
            q = parse_qs(u.query)
            mode = q.get("iframe", ["allow"])[0]
            allow = {"allow": ' allow="local-network-access; loopback-network"',
                     "legacy": ' allow="local-network-access"'}.get(mode, "")
            addin = q.get("addin", [""])[0]
            body = f"""<!doctype html><html><head><title>Excel (simulated)</title>
<style>body{{margin:0;font:13px system-ui;display:flex;height:100vh}} .grid{{flex:1;padding:12px;background:#fff;color:#222}}
iframe{{width:340px;border:0;border-left:1px solid #ccc;height:100%}}</style></head>
<body><div class="grid"><b>Simulated Office for the web host page</b><p>Office.js is mocked inside the add-in frame.
This page only reproduces the iframe embedding and permissions-policy shape.</p></div>
<iframe src="{addin}"{allow} sandbox="allow-scripts allow-same-origin allow-forms allow-popups allow-popups-to-escape-sandbox"></iframe></body></html>""".encode()
            self.send_response(200)
            self.send_header("Content-Type", "text/html; charset=utf-8")
            self.send_header("Content-Length", str(len(body)))
            self.end_headers()
            self.wfile.write(body)
    return Office


def serve_https(port, handler, key, crt):
    httpd = ThreadingHTTPServer((os.environ.get("HARNESS_PAGE_IP", "127.0.0.1"), port), handler)
    ctx = ssl.SSLContext(ssl.PROTOCOL_TLS_SERVER)
    ctx.load_cert_chain(str(crt), str(key))
    httpd.socket = ctx.wrap_socket(httpd.socket, server_side=True)
    threading.Thread(target=httpd.serve_forever, daemon=True).start()
    return httpd


sys.path.insert(0, str(ROOT / "scripts" / "live"))
from plan import PLANS  # noqa: E402
from runner import Runner, summary  # noqa: E402


def run_host(app: str, variant: str, plugin, ctx, out: Path) -> dict:
    session_id = f"desktop-chat-{app}-{variant}"
    connect = json.loads(ctx.tools["office_connect"]["handler"]({"app": app}, session_id=session_id))
    assert connect["ok"], connect
    code = connect["data"]["pairingCode"]
    vout = out / f"{app}-{variant}"
    vout.mkdir(parents=True, exist_ok=True)
    env = dict(os.environ, HARNESS_VARIANT=variant, HARNESS_APP=app, HARNESS_CODE=code, HARNESS_OUT=str(vout),
               HARNESS_OFFICE_PORT=str(OFFICE_PORT), HARNESS_ADDIN_PORT=str(ADDIN_PORT))
    proc = subprocess.Popen([str(ELECTRON), str(HERE)], env=env, stdin=subprocess.PIPE, stdout=subprocess.PIPE,
                            stderr=subprocess.DEVNULL, text=True)
    events: "queue.Queue[dict]" = queue.Queue()
    log: list[dict] = []

    def reader():
        for line in proc.stdout:
            try:
                ev = json.loads(line)
            except ValueError:
                continue
            log.append(ev)
            events.put(ev)
    threading.Thread(target=reader, daemon=True).start()

    def send(obj):
        proc.stdin.write(json.dumps(obj) + "\n")
        proc.stdin.flush()

    def wait_for(name, timeout=40, pred=lambda ev: True):
        end = time.time() + timeout
        while time.time() < end:
            try:
                ev = events.get(timeout=0.5)
            except queue.Empty:
                continue
            if ev.get("ev") == name and pred(ev):
                return ev
        return None

    counter = [0]

    def pane_eval(js):
        counter[0] += 1
        i = counter[0]
        send({"cmd": "eval", "id": i, "js": js})
        ev = wait_for("eval", 30, lambda e: e.get("id") == i)
        return ev and ev.get("value")

    def reload_pane():
        send({"cmd": "reload"})
        ev = wait_for("reloaded", 60)
        if not ev or ev.get("state") != "paired":
            raise RuntimeError(f"pane did not resume after reload: {ev}")

    result = {"app": app, "variant": variant}
    settled = wait_for("pairing-settled")
    result["paneState"] = settled and settled.get("state")
    if result["paneState"] == "paired":
        send({"cmd": "screenshot", "path": str(vout / "before.png")})
        wait_for("screenshot")
        runner = Runner(ctx.tools, session_id, app, lambda: pane_eval("globalThis.__hermesMockFingerprint()"),
                        reload_pane, lambda step: pane_eval(step["mock"]), log=lambda line: print(f"  [{app}] {line}"))
        result["results"] = runner.run(PLANS[app]())
        result["summary"] = summary(result["results"])
        send({"cmd": "screenshot", "path": str(vout / "after.png")})
        wait_for("screenshot")
    result["events"] = [e for e in log if e.get("ev") in ("office-page", "pane-ready", "pane-state", "permission-request", "permission-check", "reloaded")
                        or (e.get("ev") == "console" and str(e.get("level")) in ("warning", "error", "2", "3"))]
    try:
        send({"cmd": "quit"})
        proc.wait(10)
    except Exception:  # noqa: BLE001
        proc.kill()
    ctx.tools["office_disconnect"]["handler"]({"app": app}, session_id=session_id)
    return result


def main(argv):
    global SCRATCH
    if argv and argv[0].startswith("--features="):
        os.environ["HARNESS_FEATURES"] = argv.pop(0).split("=", 1)[1]
        SCRATCH = SCRATCH.with_name(SCRATCH.name + "-forced-lna")
    if os.environ.get("HARNESS_PAGE_IP", "127.0.0.1") != "127.0.0.1":
        SCRATCH = SCRATCH.with_name(SCRATCH.name + "-lan-page")
    hosts = ["excel", "word", "powerpoint", "onenote"]
    if argv and argv[0].startswith("--hosts="):
        hosts = argv.pop(0).split("=", 1)[1].split(",")
    variants = argv or ["allow-default"]
    SCRATCH.mkdir(parents=True, exist_ok=True)
    key, crt = make_cert(SCRATCH)
    site = SCRATCH / "site"
    subprocess.run([sys.executable, str(ROOT / "scripts" / "build_addin.py"), "--base",
                    f"{ADDIN_ORIGIN}/office/live/", "--out", str(site / "office" / "live"), "--include-dev"],
                   check=True, capture_output=True)

    class Addin(SimpleHTTPRequestHandler):
        def __init__(self, *a, **kw):
            super().__init__(*a, directory=str(site), **kw)

        def log_message(self, *a):
            pass

        def end_headers(self):
            self.send_header("Cache-Control", "no-store")
            super().end_headers()
    Addin.extensions_map = {**SimpleHTTPRequestHandler.extensions_map, ".js": "text/javascript", ".mjs": "text/javascript"}
    serve_https(OFFICE_PORT, office_handler_factory(), key, crt)
    serve_https(ADDIN_PORT, Addin, key, crt)
    plugin = load_plugin()
    ctx = FakeCtx({"allowed_origin": ADDIN_ORIGIN})  # real default ports 47615-47619
    # Brain index in a scratch SQLite file (never a real Hermes profile's plugin-data).
    import sqlite3
    db = SCRATCH / "brain.db"
    for f in (db, db.with_name("brain.db-wal"), db.with_name("brain.db-shm")):
        f.unlink(missing_ok=True)
    ctx._test_brain = sys.modules["hermes_office_desktop.brain"].Brain(lambda: sqlite3.connect(str(db), check_same_thread=False))
    plugin.register(ctx)
    results = [run_host(app, v, plugin, ctx, SCRATCH) for v in variants for app in hosts]
    (SCRATCH / "results.json").write_text(json.dumps(results, indent=2))
    for r in results:
        print(f"{r['app']:10} {r['variant']:16} pane={r['paneState']!s:8} {json.dumps(r.get('summary'))}")
    print("results:", SCRATCH / "results.json")


if __name__ == "__main__":
    main(sys.argv[1:])
