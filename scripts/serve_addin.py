# SPDX-License-Identifier: Apache-2.0
"""Serve a built add-in directory over HTTPS on loopback (local live testing only; never a deploy).

    python scripts/serve_addin.py --dir build/atlas --port 3010 --cert <crt> --key <key>

Binds 127.0.0.1 and ::1 only. Sends no X-Frame-Options / frame-ancestors (Office must
frame the task pane), `Cache-Control: no-store`, and module-compatible MIME types.
The certificate and key are read by the TLS stack and never printed or copied.
"""
from __future__ import annotations

import argparse
import socket
import ssl
import sys
import threading
from functools import partial
from http.server import SimpleHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path


class Handler(SimpleHTTPRequestHandler):
    extensions_map = {**SimpleHTTPRequestHandler.extensions_map, ".js": "text/javascript", ".mjs": "text/javascript",
                      ".xml": "application/xml", ".css": "text/css", ".png": "image/png", ".html": "text/html"}

    def log_message(self, fmt, *args):  # request lines only, no bodies or headers
        sys.stderr.write("%s %s\n" % (self.log_date_time_string(), fmt % args))

    def end_headers(self):
        self.send_header("Cache-Control", "no-store")
        self.send_header("X-Content-Type-Options", "nosniff")
        super().end_headers()


class V6Server(ThreadingHTTPServer):
    address_family = socket.AF_INET6


def main(argv=None):
    ap = argparse.ArgumentParser()
    ap.add_argument("--dir", required=True)
    ap.add_argument("--port", type=int, default=3010)
    ap.add_argument("--cert", required=True)
    ap.add_argument("--key", required=True)
    a = ap.parse_args(argv)
    root = Path(a.dir).resolve()
    if not (root / "SHA256SUMS").exists():
        raise SystemExit("--dir must be a build_addin.py output (SHA256SUMS missing)")
    ctx = ssl.SSLContext(ssl.PROTOCOL_TLS_SERVER)
    ctx.minimum_version = ssl.TLSVersion.TLSv1_2
    ctx.load_cert_chain(a.cert, a.key)
    handler = partial(Handler, directory=str(root))
    servers = [ThreadingHTTPServer(("127.0.0.1", a.port), handler)]
    try:
        servers.append(V6Server(("::1", a.port), handler))
    except OSError:
        pass
    for s in servers:
        s.daemon_threads = True
        s.socket = ctx.wrap_socket(s.socket, server_side=True)
        threading.Thread(target=s.serve_forever, daemon=True).start()
    print(f"serving {root} on https://localhost:{a.port}/ ({len(servers)} listeners)", flush=True)
    threading.Event().wait()


if __name__ == "__main__":
    main()
