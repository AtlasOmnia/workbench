# SPDX-License-Identifier: Apache-2.0
"""Positive control: does real Google Chrome enforce Local Network Access on the same topology?

Serves the same simulated Office page (LAN address => 'local' space) iframing a probe page that
fetches the real plugin bridge on 127.0.0.1 and reports the outcome back to its own origin.
Runs headless Chrome with a throwaway profile; touches no user browser profile.
"""
from __future__ import annotations

import importlib.util
import json
import os
import subprocess
import sys
import tempfile
import threading
import time
from http.server import SimpleHTTPRequestHandler
from pathlib import Path
from urllib.parse import parse_qs, urlsplit

HERE = Path(__file__).resolve().parent
spec = importlib.util.spec_from_file_location("harness_run", HERE / "run.py")
run = importlib.util.module_from_spec(spec)
spec.loader.exec_module(run)

CHROME = os.environ.get("CONTROL_CHROME", "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome")
PAGE_IP = os.environ.get("HARNESS_PAGE_IP", "10.0.0.1")
REPORTS: list[dict] = []

PROBE = b"""<!doctype html><meta charset=utf-8><script type=module>
const r = { policyLoopback: document.featurePolicy ? document.featurePolicy.allowsFeature('loopback-network') : null }
try {
  const res = await fetch('http://127.0.0.1:47615/v1/hello', { targetAddressSpace: 'loopback', mode: 'cors', credentials: 'omit' })
  r.status = res.status; r.body = await res.json()
} catch (e) { r.error = String(e) }
await fetch('/report?variant=' + new URLSearchParams(location.search).get('v') + '&r=' + encodeURIComponent(JSON.stringify(r)))
</script>"""


class Addin(SimpleHTTPRequestHandler):
    def log_message(self, *a):
        pass

    def do_GET(self):
        u = urlsplit(self.path)
        if u.path == "/probe.html":
            body = PROBE
        elif u.path == "/report":
            q = parse_qs(u.query)
            REPORTS.append({"variant": q["variant"][0], **json.loads(q["r"][0])})
            body = b"ok"
        else:
            self.send_error(404)
            return
        self.send_response(200)
        self.send_header("Content-Type", "text/html" if u.path.endswith(".html") else "text/plain")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)


def main():
    os.environ["HARNESS_PAGE_IP"] = PAGE_IP
    out = Path(tempfile.mkdtemp(prefix="lna-control-", dir=os.environ.get("TMPDIR")))
    key, crt = run.make_cert(out)
    run.serve_https(run.OFFICE_PORT, run.office_handler_factory(), key, crt)
    run.serve_https(run.ADDIN_PORT, Addin, key, crt)
    plugin = run.load_plugin()
    ctx = run.FakeCtx({"allowed_origin": run.ADDIN_ORIGIN})
    plugin.register(ctx)
    plugin.get_server(ctx).ensure_started()
    for variant in os.environ.get("CONTROL_VARIANTS", "allow,legacy,none").split(","):
        addin = f"https://addin.test:{run.ADDIN_PORT}/probe.html?v={variant}"
        url = f"https://office.test:{run.OFFICE_PORT}/office.html?iframe={variant}&addin={addin}"
        profile = out / f"chrome-{variant}"
        p = subprocess.Popen([CHROME, "--headless=new", f"--user-data-dir={profile}", "--no-first-run",
                              "--no-default-browser-check", "--disable-extensions", "--ignore-certificate-errors",
                              f"--host-resolver-rules=MAP office.test {PAGE_IP}, MAP addin.test {PAGE_IP}"]
                             + ([f"--ip-address-space-overrides=127.0.0.1:{run.OFFICE_PORT}=public,127.0.0.1:{run.ADDIN_PORT}=public"]
                                if PAGE_IP == "127.0.0.1" else []) + [url],
                             stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        end = time.time() + 20
        while time.time() < end and not any(r["variant"] == variant for r in REPORTS):
            time.sleep(0.3)
        p.terminate()
        p.wait(10)
    version = subprocess.run([CHROME, "--version"], capture_output=True, text=True).stdout.strip()
    print(json.dumps({"chrome": version, "pageIp": PAGE_IP, "reports": REPORTS}, indent=1))
    plugin.get_server(ctx).stop()


if __name__ == "__main__":
    sys.exit(main())
