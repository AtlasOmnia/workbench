# SPDX-License-Identifier: Apache-2.0
"""Test-only HTTPS client for the bridge: trusts only the session's local CA (never the OS store) and checks the
certificate for the name ``localhost`` while connecting to 127.0.0.1 (the bridge binds IPv4 loopback only)."""
import http.client
import socket
import sys

DATA = None  # conftest: the session's plugin data folder; the bridge's TLS identity lives in DATA / "tls"


def local_tls():
    return sys.modules["hermes_office_desktop.bridge_tls"].LocalTLS(DATA)


class Conn(http.client.HTTPSConnection):
    def __init__(self, port, timeout, context):
        super().__init__("localhost", port, timeout=timeout, context=context)
        self.tls_context = context

    def connect(self):
        sock = socket.create_connection(("127.0.0.1", self.port), self.timeout)
        self.sock = self.tls_context.wrap_socket(sock, server_hostname="localhost")


def conn(port, timeout=30, tls=None):
    return Conn(port, timeout, (tls or local_tls()).client_context())
