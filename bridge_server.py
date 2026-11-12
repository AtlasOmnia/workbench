# SPDX-License-Identifier: Apache-2.0
"""Loopback bridge between Hermes tool calls and the Office for the web task pane.

The bridge serves HTTPS on 127.0.0.1 with this install's local TLS identity (bridge_tls.py). By default it also
serves the task pane itself: the bundled add-in files of the pinned plugin tree (addin/, only names listed in
addin/SHA256SUMS whose sha256 matches) under /office/live/, so the pane, loaded into Office for the web through
"Upload My Add-in", and the bridge are one origin, https://localhost:<port>. It is a transport, not an agent: it
never calls Office or Microsoft services itself.

Security model (see docs/reference.md "Bridge security"):
- binds 127.0.0.1 only, HTTPS only; rejects any Host header other than localhost:<bound port> (421);
- every /v1 route requires an allowed Origin: by default one of the five Workbench loopback origins
  https://localhost:47615-47619 (a same-origin GET carries no Origin and is accepted when Sec-Fetch-Site is absent
  or same-origin); with a self-hosted add-in site, exactly that site's origin;
- static files: GET/HEAD only, listed and hash-checked on every request, no symlinks, no path escapes, never
  addin/dev (the Office.js mock), fixed MIME types, nosniff and a Content-Security-Policy;
- a pane becomes trusted only by presenting a short-lived, single-use pairing
  code that a Hermes chat session minted (office_connect). The code is shown to
  the human in that chat and typed by the human into the pane; that is the bind
  between a Desktop chat session and one pane;
- after pairing the pane holds a 256-bit bearer token kept in memory only;
- bounded bodies, bounded waits, no redirects, no retries, no persistence.
"""
from __future__ import annotations

import hashlib
import hmac
import json
import secrets
import socket
import threading
import time
import uuid
from dataclasses import dataclass, field
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from typing import Any, Optional
from urllib.parse import parse_qs, unquote, urlsplit

PROTOCOL = 1
SERVICE = "hermes-office-desktop"
DEFAULT_PORTS = (47615, 47616, 47617, 47618, 47619)  # must match addin/shared/bridge-client.js
# The add-in site's BASE URL is the one source of truth (plugin setting addin_base): the bridge's allowed origin
# (scheme + host[:port]) and every manifest / task-pane / icon / support URL derive from it. Default: this plugin
# serves the bundled add-in itself on its loopback HTTPS port; at runtime the base follows the port actually bound
# (loopback_base). A self-hosted site is optional (addin_base); the older setting allowed_origin (an origin alone)
# means origin + ADDIN_PATH.
DEFAULT_ADDIN_BASE = "https://localhost:47615/office/live/"
ADDIN_PATH = "/office/live/"
SITE_ROOT = Path(__file__).resolve().parent / "addin"  # the bundled add-in; only names in SITE_ROOT/SHA256SUMS
SITE_SUMS = "SHA256SUMS"
STATIC_MIME_TYPES = {
    ".html": "text/html; charset=utf-8",
    ".js": "text/javascript; charset=utf-8",
    ".css": "text/css; charset=utf-8",
    ".png": "image/png",
    ".svg": "image/svg+xml",
    ".ico": "image/x-icon",
    ".json": "application/json; charset=utf-8",
    ".txt": "text/plain; charset=utf-8",
}
# The task panes' own meta policy (addin/<app>/taskpane.html), sent as a header for every static file. No
# frame-ancestors: Office for the web frames the pane from several Microsoft hosts that differ by account type.
STATIC_CSP = ("default-src 'self'; script-src 'self' https://appsforoffice.microsoft.com; connect-src 'self' "
              "https://localhost:*; style-src 'self'; img-src 'self' data:; object-src 'none'; base-uri 'none'; "
              "form-action 'none'")
TLS_HANDSHAKE_SECONDS = 10.0
APPS = ("excel", "word", "powerpoint", "onenote", "outlook")
CODE_ALPHABET = "23456789ABCDEFGHJKMNPQRSTVWXYZ"  # no 0/O/1/I/L/U
CODE_LENGTH = 8  # 30^8 ~ 6.6e11 (~39 bits); TTL + attempt caps bound online guessing
CODE_TTL_SECONDS = 600
MAX_CODE_FAILURES = 5
MAX_GLOBAL_FAILURES_PER_MINUTE = 20
MAX_BODY_BYTES = 8 * 1024 * 1024
MAX_EVENT_BYTES = 256 * 1024  # one captured page per visit event
POLL_WAIT_MAX = 25.0
# A live pane starts a new /v1/next at least every ~20 s (its long-poll wait). No poll ARRIVING for
# this long means the page is gone (chat switched, pane closed); a held poll never counts as alive.
PANE_STALE_SECONDS = 45.0
# A19: a command the pane picked up but did not answer in time (Hermes gave up waiting) still keeps the pane counted
# as live for this long: the pane is busy running it (seen live: brain_sync's list_pages in a hidden OneNote tab), not
# gone. Its late result is accepted and dropped.
ABANDONED_GRACE_SECONDS = 90.0
PANE_UNLOADED = "pane unloaded (switched chat or closed); reopen Workbench in the ribbon"
MAX_ELAPSED_MS = 10_000_000
BINDING_IDLE_SECONDS = 15 * 60  # a pane gone this long (closed tab, sleep) must re-pair
DEFAULT_CALL_TIMEOUT = 60.0
# Pane commands that only read (OneNote): a timeout on one of these changed nothing (A27).
READ_COMMANDS = frozenset({"list_pages", "list_sections", "get_page", "get_context", "image_status", "op_status"})
# A14: a profile's bridge gives its port back after this long with nothing to serve (no live pane, no
# pending code, no command or tool call in flight, no request being handled) and takes a port again on
# its next Office tool call. Shorter than BINDING_IDLE_SECONDS so a pane that comes back can still resume.
IDLE_RELEASE_SECONDS = 10 * 60
REAP_INTERVAL_SECONDS = 30.0
# F8 D4: what an uncertain edit tells the agent to check, per app (it said "workbook" in every app).
EDIT_TARGETS = {"excel": "workbook", "word": "document", "powerpoint": "presentation", "onenote": "page"}


def _edit_target(app: str, tool: str) -> str:
    if app == "outlook":  # one pane for mail drafts and calendar event forms
        return "draft or event" if tool == "undo_last" else "event" if "event" in tool else "draft"
    return EDIT_TARGETS.get(app, "document")


def _check_base(base) -> str:
    parts = urlsplit(base) if isinstance(base, str) else None
    if parts is None or parts.scheme != "https" or not parts.hostname or "@" in parts.netloc \
            or not parts.path.endswith("/") or parts.query or parts.fragment or base.endswith("//"):
        raise BridgeError("Configuration", "addin_base must be an https URL ending in / (no query, fragment or user)")
    return base


def origin_of(base: str) -> str:
    """scheme://host[:port] of an add-in base URL (what the pane's requests carry as Origin)."""
    parts = urlsplit(_check_base(base))
    return f"{parts.scheme}://{parts.netloc}"


def _check_origin(origin) -> str:
    if not isinstance(origin, str) or not origin.startswith("https://") or origin.count("/") != 2 \
            or "@" in origin or len(origin) <= len("https://"):
        raise BridgeError("Configuration", "allowed_origin must be an https origin with no path")
    return origin


def resolve_addin_base(addin_base=None, allowed_origin=None) -> str:
    """The add-in base URL from the two settings: an explicit addin_base (other than the default) wins, and an
    allowed_origin set next to it must be its origin; else allowed_origin alone means origin + ADDIN_PATH; else the
    default. Never guesses between two disagreeing settings."""
    if addin_base and addin_base != DEFAULT_ADDIN_BASE:
        base = _check_base(addin_base)
        if allowed_origin and _check_origin(allowed_origin) != origin_of(base):
            raise BridgeError("Configuration", f"addin_base {base} and allowed_origin {allowed_origin} disagree; set "
                                               "only addin_base")
        return base
    if allowed_origin:
        return _check_origin(allowed_origin) + ADDIN_PATH
    return DEFAULT_ADDIN_BASE


DEFAULT_ORIGIN = origin_of(DEFAULT_ADDIN_BASE)  # derived, never configured separately


def loopback_base(port: int) -> str:
    """The plugin-served add-in base for the port this bridge actually bound."""
    return f"https://localhost:{int(port)}{ADDIN_PATH}"


def _site_sums(site: Path) -> Optional[dict]:
    """name -> sha256 from the site's SHA256SUMS; None when it is missing or any line is malformed."""
    try:
        lines = (site / SITE_SUMS).read_text(encoding="utf-8").splitlines()
    except (OSError, UnicodeError):
        return None
    out = {}
    for line in lines:
        digest, sep, name = line.partition("  ")
        if not sep or not name or len(digest) != 64 or any(c not in "0123456789abcdef" for c in digest):
            return None
        out[name] = digest
    return out


def read_site_file(site: Optional[Path], encoded: str) -> Optional[tuple[bytes, str]]:
    """(content, MIME type) of a bundled add-in file, or None. ``encoded`` is the raw path after ADDIN_PATH.

    Served only when the name is listed in the site's SHA256SUMS and the file's sha256 matches it now (checked on
    every request). Refused: encoded separators, '..', backslashes, NUL, empty or doubled separators, anything under
    dev/ (the Office.js mock), any symlink on the way, and anything that resolves outside the site folder."""
    if site is None or "%2f" in encoded.lower():
        return None
    try:
        name = unquote(encoded, errors="strict")
    except UnicodeDecodeError:
        return None
    if (not name or ".." in name or "\\" in name or "\x00" in name or name.startswith("/") or "//" in name
            or name.split("/", 1)[0] == "dev"):
        return None
    sums = _site_sums(site)
    expected = sums.get(name) if sums else None
    if expected is None:
        return None
    candidate = site
    for part in name.split("/"):
        candidate = candidate / part
        if candidate.is_symlink():
            return None
    try:
        root = site.resolve(strict=True)
        resolved = candidate.resolve(strict=True)
        resolved.relative_to(root)
        if not resolved.is_file():
            return None
        content = resolved.read_bytes()
    except (OSError, ValueError):
        return None
    if not hmac.compare_digest(hashlib.sha256(content).hexdigest(), expected):
        return None
    return content, STATIC_MIME_TYPES.get(Path(name).suffix.lower(), "application/octet-stream")
# Keep-alive connections that sit idle this long are closed, so an old connection can never keep a
# released bridge reachable (must exceed POLL_WAIT_MAX: a held long-poll is not idle reading).
CONNECTION_IDLE_SECONDS = 60
_NUMBER_WORDS = {1: "one", 2: "two", 3: "three", 4: "four", 5: "five", 6: "six", 7: "seven", 8: "eight",
                 9: "nine", 10: "ten"}


class BridgeError(RuntimeError):
    def __init__(self, category: str, message: str):
        super().__init__(message)
        self.category = category


PORTS_BUSY = "PortsBusy"


def _port_span(ports) -> str:
    ports = tuple(ports)
    if ports and list(ports) == list(range(ports[0], ports[0] + len(ports))):
        return f"127.0.0.1:{ports[0]}-{ports[-1]}" if len(ports) > 1 else f"127.0.0.1:{ports[0]}"
    return "127.0.0.1 ports " + ", ".join(str(p) for p in ports)


def ports_busy_message(ports, idle_release_seconds: float = IDLE_RELEASE_SECONDS) -> str:
    """The exact text office_connect / office_status show when every bridge port is taken."""
    n = len(tuple(ports))
    word = _NUMBER_WORDS.get(n, str(n))
    minutes = max(1, int(round(idle_release_seconds / 60)))
    return (f"All {word} Workbench bridge ports ({_port_span(ports)}) are in use: {word} Hermes profiles are using "
            f"Workbench right now (or another program holds a port). Close a Workbench pane that another "
            f"profile's chat paired, or wait up to {minutes} minutes: a profile with no open pane gives its port "
            f"back after {minutes} minutes. Then call office_connect again.")


def busy_ports(ports, timeout: float = 0.25) -> list:
    """Ports of ``ports`` something is listening on (a TCP connect probe; binds nothing, sends nothing)."""
    out = []
    for port in ports:
        try:
            with socket.create_connection(("127.0.0.1", port), timeout=timeout):
                out.append(port)
        except OSError:
            pass
    return out


def normalize_code(raw: Any) -> str:
    if not isinstance(raw, str):
        return ""
    return "".join(ch for ch in raw.upper() if ch.isalnum())[:32]


def format_code(code: str) -> str:
    return f"{code[:4]}-{code[4:]}"


@dataclass
class PendingCode:
    code: str
    session_key: str
    app: str
    expires_at: float
    failures: int = 0
    home: Optional[str] = None  # owning Hermes profile home of the chat that minted it (profile_scope.py)


@dataclass
class Command:
    request_id: str
    tool: str
    arguments: dict
    done: threading.Event = field(default_factory=threading.Event)
    result: Optional[dict] = None
    delivered: bool = False
    abandoned_at: Optional[float] = None  # A19: Hermes stopped waiting; the pane may still be running it


@dataclass
class Binding:
    binding_id: str
    session_key: str
    app: str
    token: str
    paired_at: float
    last_seen: float
    document: dict
    host: dict
    supported_tools: frozenset
    queue: list = field(default_factory=list)
    inflight: dict = field(default_factory=dict)
    wake: threading.Condition = field(default_factory=threading.Condition)
    closed: bool = False
    resumes: int = 0
    last_poll: float = 0.0  # when the pane last STARTED a poll or posted a result (liveness)
    poll_epoch: int = 0  # newest /v1/next wins; an abandoned long-poll (reloaded page) never takes a command
    home: Optional[str] = None  # owning Hermes profile home, carried to out-of-turn event handlers


class Registry:
    """All mutable bridge state. Thread-safe; no persistence by design."""

    def __init__(self, clock=time.monotonic):
        self._lock = threading.RLock()
        self._clock = clock
        # Optional plugin callback for pane-pushed events (OneNote page visits): (binding, event) -> dict.
        self.on_event = None
        self._codes: dict[str, PendingCode] = {}
        self._bindings: dict[str, Binding] = {}  # token -> binding
        self._by_session: dict[tuple[str, str], str] = {}  # (session, app) -> token
        self._fail_times: list[float] = []
        # A14 idle release (BridgeServer.maybe_release): newest activity, tool calls in flight, and the
        # short window in which a release is shutting the listener (pairing refused, never half-made).
        self._last_activity = self._clock()
        self._calls_active = 0
        self._draining = False

    def _touch(self) -> None:
        self._last_activity = self._clock()

    def idle_for_release(self, idle_seconds: float) -> tuple[bool, float]:
        """(nothing to serve, seconds until a release is due). Caller holds no lock; this takes the registry's."""
        with self._lock:
            now = self._clock()
            self._sweep(now)
            if self._codes or self._calls_active:
                return False, float(idle_seconds)
            for b in self._bindings.values():
                if b.queue or self._busy(b, now) or self.liveness(b)["live"]:
                    return False, float(idle_seconds)
            return True, max(0.0, idle_seconds - (now - self._last_activity))

    # --- pairing ---------------------------------------------------------
    def mint_code(self, session_key: str, app: str, home: Optional[str] = None) -> tuple[str, int]:
        if app not in APPS:
            raise BridgeError("Validation", f"app must be one of: {', '.join(APPS)}")
        with self._lock:
            if self._draining:
                raise BridgeError("Unavailable", "The Workbench bridge is giving its idle port back; try again")
            now = self._clock()
            self._sweep(now)
            self._touch()
            # One live code per (session, app): minting again replaces the old one.
            for c in [c for c in self._codes.values() if c.session_key == session_key and c.app == app]:
                self._codes.pop(c.code, None)
            code = "".join(secrets.choice(CODE_ALPHABET) for _ in range(CODE_LENGTH))
            self._codes[code] = PendingCode(code, session_key, app, now + CODE_TTL_SECONDS, home=home)
            return code, CODE_TTL_SECONDS

    def pair(self, raw_code: Any, app: Any, document: Any, host: Any, supported: Any,
             previous_token: Any = None) -> Binding:
        code = normalize_code(raw_code)
        with self._lock:
            if self._draining:  # never hand out a token on a listener that is closing (A14)
                raise BridgeError("Unavailable", "This Hermes bridge is giving its idle port back; ask Hermes for a "
                                  "new code (office_connect)")
            now = self._clock()
            self._sweep(now)
            self._fail_times = [t for t in self._fail_times if now - t < 60]
            if len(self._fail_times) >= MAX_GLOBAL_FAILURES_PER_MINUTE:
                raise BridgeError("RateLimited", "Too many failed pairing attempts; wait a minute")
            match = None
            for candidate in self._codes.values():
                if len(code) == len(candidate.code) and hmac.compare_digest(candidate.code, code):
                    match = candidate
                    break
            if match is None or match.app != app:
                self._fail_times.append(now)
                # Count a failure against every live code of this app: bounds guessing per code.
                for c in list(self._codes.values()):
                    if c.app == app:
                        c.failures += 1
                        if c.failures >= MAX_CODE_FAILURES:
                            self._codes.pop(c.code, None)
                raise BridgeError("Authentication", "Pairing code is invalid or expired")
            self._codes.pop(match.code, None)  # single use
            previous = self._by_session.get((match.session_key, app))
            if previous:
                self._close(previous, "replaced by a newer pairing")
            # One pane serves one chat: a pane re-pairing to another chat drops its old binding.
            if isinstance(previous_token, str) and previous_token in self._bindings:
                self._close(previous_token, "re-paired with another chat")
            token = secrets.token_urlsafe(32)
            binding = Binding(
                binding_id=str(uuid.uuid4()), session_key=match.session_key, app=app, token=token,
                paired_at=now, last_seen=now, last_poll=now, document=_clean_doc(document), host=_clean_host(host),
                supported_tools=frozenset(t for t in (supported or []) if isinstance(t, str) and len(t) < 64),
                home=match.home,
            )
            self._bindings[token] = binding
            self._by_session[(match.session_key, app)] = token
            self._touch()
            return binding

    def authenticate(self, token: str) -> Binding:
        with self._lock:
            b = self._bindings.get(token or "") if isinstance(token, str) else None
            if b is None or b.closed:
                raise BridgeError("Authentication", "Pane is not paired")
            if self._clock() - b.last_seen > BINDING_IDLE_SECONDS:
                self._close(token, "idle too long")
                raise BridgeError("Authentication", "Pane pairing expired; ask Hermes for a new code")
            b.last_seen = self._clock()
            self._touch()
            return b

    def resume(self, token: str, document: Any, host: Any, supported: Any, app: Any = None) -> Binding:
        """A reloaded pane presents its tab-scoped token to keep the same chat binding.

        Commands the previous page instance had picked up can never be answered now,
        so they fail as ExecutionUncertain at once instead of waiting for the timeout.
        Queued (undelivered) commands stay queued for the resumed pane.
        """
        with self._lock:
            b = self.authenticate(token)
            if app is not None and app != b.app:
                raise BridgeError("Authentication", "Pane token belongs to another Office app")
            # The pairing was consent for ONE document (for OneNote: one notebook). A reloaded pane
            # that now shows a different one must pair again; its old binding is dropped.
            new_id = _clean_doc(document).get("id")
            if b.document.get("id") and new_id and new_id != b.document["id"]:
                self._close(token, "now shows a different document")
                raise BridgeError("Authentication", "IdentityMismatch: the pane now shows a different document; pair again")
            for cmd in list(b.inflight.values()):
                cmd.result = {"ok": False, "error": {"category": "ExecutionUncertain", "message":
                              "The Office pane reloaded while running this command; check the document before retrying."}}
                cmd.done.set()
            b.inflight.clear()
            b.poll_epoch += 1  # any long-poll still held open for the old page is now stale
            b.document = _clean_doc(document) or b.document
            b.host = _clean_host(host) or b.host
            if isinstance(supported, list):
                b.supported_tools = frozenset(t for t in supported if isinstance(t, str) and len(t) < 64)
            b.resumes += 1
            b.last_poll = self._clock()
        with b.wake:
            b.wake.notify_all()
        return b

    def unbind_session(self, session_key: str, app: str) -> bool:
        with self._lock:
            token = self._by_session.get((session_key, app))
            if not token:
                return False
            self._close(token, "disconnected from Hermes")
            return True

    def unbind_token(self, token: str) -> None:
        with self._lock:
            if token in self._bindings:
                self._close(token, "pane closed")

    def _close(self, token: str, reason: str) -> None:
        b = self._bindings.pop(token, None)
        if not b:
            return
        b.closed = True
        if self._by_session.get((b.session_key, b.app)) == token:
            self._by_session.pop((b.session_key, b.app), None)
        for cmd in list(b.queue) + list(b.inflight.values()):
            cmd.result = {"ok": False, "error": {"category": "Unavailable", "message": f"Pane {reason}"}}
            cmd.done.set()
        with b.wake:
            b.wake.notify_all()

    def _sweep(self, now: float) -> None:
        for code in [c.code for c in self._codes.values() if c.expires_at <= now]:
            self._codes.pop(code, None)

    # --- command flow ----------------------------------------------------
    def binding_for(self, session_key: str, app: str) -> Optional[Binding]:
        with self._lock:
            token = self._by_session.get((session_key, app))
            b = self._bindings.get(token) if token else None
            if b is not None and self._clock() - b.last_seen > BINDING_IDLE_SECONDS:
                self._close(token, "idle too long")
                return None
            return b

    def _busy(self, b: Binding, now: float) -> bool:
        """Is the pane running a command? Abandoned ones count for ABANDONED_GRACE_SECONDS, then are forgotten.
        Caller holds the registry lock."""
        for rid, cmd in list(b.inflight.items()):
            if cmd.abandoned_at is not None and now - cmd.abandoned_at > ABANDONED_GRACE_SECONDS:
                b.inflight.pop(rid, None)
        return bool(b.inflight)

    def liveness(self, b: Binding) -> dict:
        """F4: age of the newest poll the pane STARTED; a pane running a command counts as live."""
        with self._lock:
            now = self._clock()
            age = max(0.0, now - b.last_poll)
            busy = self._busy(b, now)
            return {"lastPollAgeMs": int(round(age * 1000)),
                    "live": (not b.closed) and (age <= PANE_STALE_SECONDS or busy)}

    def call(self, session_key: str, app: str, tool: str, arguments: dict,
             timeout: float = DEFAULT_CALL_TIMEOUT) -> dict:
        with self._lock:
            self._calls_active += 1
            self._touch()
        try:
            return self._call(session_key, app, tool, arguments, timeout)
        finally:
            with self._lock:
                self._calls_active -= 1
                self._touch()

    def has_binding(self, session_key: str) -> bool:
        with self._lock:
            return any(k[0] == session_key for k in self._by_session)

    def wait_live(self, b: Binding, timeout: float) -> bool:
        """After a released bridge took a port again: give the pane a moment to find it (A14)."""
        deadline = time.monotonic() + max(0.0, timeout)
        while True:
            if b.closed:
                return False
            if self.liveness(b)["live"]:
                return True
            remaining = deadline - time.monotonic()
            if remaining <= 0:
                return False
            with b.wake:
                b.wake.wait(min(remaining, 0.2))

    def _call(self, session_key: str, app: str, tool: str, arguments: dict, timeout: float) -> dict:
        started = time.monotonic()
        with self._lock:
            b = self.binding_for(session_key, app)
            if b is None:
                raise BridgeError("NotPaired", "No Office pane is paired with this Hermes chat. "
                                  "Call office_connect and have the user enter the code in the pane.")
            live = self.liveness(b)
            if not live["live"]:
                raise BridgeError("Unavailable", f"The paired pane stopped polling {live['lastPollAgeMs'] // 1000} s "
                                  f"ago: {PANE_UNLOADED}. Nothing was sent.")
            if b.supported_tools and tool not in b.supported_tools:
                raise BridgeError("Unsupported", f"The paired pane does not implement {tool}")
            cmd = Command(request_id=str(uuid.uuid4()), tool=tool, arguments=arguments)
            busy_before = bool(b.inflight)
            b.queue.append(cmd)
        with b.wake:
            b.wake.notify_all()
        if not cmd.done.wait(timeout):
            with self._lock:
                undelivered = cmd in b.queue
                if undelivered:
                    b.queue.remove(cmd)
                elif cmd.request_id in b.inflight:
                    cmd.abandoned_at = self._clock()  # still running in the pane: busy, not gone (A19)
            if undelivered:
                raise BridgeError("Timeout", "Pane did not pick up the command; it was not executed"
                                  + (" (it was still busy with an earlier request)" if busy_before else ""))
            if tool in READ_COMMANDS:  # A27: a read that ran long changed nothing; never call it an edit
                raise BridgeError("PaneBusy", f"The pane did not finish this read ({tool}) within {int(timeout)} s; "
                                  "nothing was changed. It answers other requests only after it finishes.")
            raise BridgeError("ExecutionUncertain", "Pane did not answer in time; the edit may or may not "
                              f"have applied. Check the {_edit_target(app, tool)} before retrying.")
        assert cmd.result is not None
        return _timed(cmd.result, int(round((time.monotonic() - started) * 1000)))

    def next_command(self, b: Binding, wait: float) -> Optional[Command]:
        deadline = self._clock() + max(0.0, min(wait, POLL_WAIT_MAX))
        with self._lock:
            b.last_poll = self._clock()
            b.poll_epoch += 1
            epoch = b.poll_epoch
        with b.wake:
            b.wake.notify_all()  # let an older, abandoned poll notice it is superseded
        while True:
            with self._lock:
                if b.closed:
                    raise BridgeError("Authentication", "Pane is not paired")
                if b.poll_epoch != epoch:
                    return None  # superseded by a newer poll from the same pane (e.g. after a reload)
                b.last_seen = self._clock()  # keeps the binding from idling out; NOT liveness (see last_poll)
                if b.queue:
                    cmd = b.queue.pop(0)
                    cmd.delivered = True
                    b.inflight[cmd.request_id] = cmd
                    return cmd
                remaining = deadline - self._clock()
            if remaining <= 0:
                return None
            with b.wake:
                b.wake.wait(min(remaining, 1.0))

    def event(self, b: Binding, payload: Any) -> dict:
        """A paired pane reports something it observed (OneNote: the user opened a page).

        Never a command and never answered with one: the plugin may only index what it is told.
        Events for a document other than the paired one are dropped (identity check).
        """
        if not isinstance(payload, dict) or not isinstance(payload.get("kind"), str):
            raise BridgeError("Validation", "Event must be an object with a kind")
        if len(json.dumps(payload, ensure_ascii=False)) > MAX_EVENT_BYTES:
            raise BridgeError("Capacity", "Event too large")
        doc_id = _clean_doc(payload.get("document")).get("id")
        if b.document.get("id") and doc_id != b.document["id"]:
            return {"accepted": False, "reason": "IdentityMismatch"}
        handler = self.on_event
        if handler is None:
            return {"accepted": False, "reason": "NoHandler"}
        try:
            out = handler(b, payload)
        except BridgeError:
            raise
        except Exception:  # noqa: BLE001 - never leak plugin internals to the pane
            return {"accepted": False, "reason": "HandlerFailed"}
        return {"accepted": True, **(out if isinstance(out, dict) else {})}

    def complete(self, b: Binding, request_id: Any, payload: Any) -> None:
        with self._lock:
            b.last_poll = self._clock()
            cmd = b.inflight.pop(request_id if isinstance(request_id, str) else "", None)
            if cmd is None:
                raise BridgeError("Protocol", "Unknown or already-completed requestId")
            if not isinstance(payload, dict) or not isinstance(payload.get("ok"), bool):
                payload = {"ok": False, "error": {"category": "Protocol", "message": "Malformed pane result"}}
            cmd.result = payload
            cmd.done.set()

    def snapshot(self) -> dict:
        with self._lock:
            now = self._clock()
            return {
                "pendingCodes": len(self._codes),
                "panes": [
                    {"app": b.app, "document": b.document.get("name", ""),
                     "secondsSinceSeen": round(now - b.last_seen, 1)}
                    for b in self._bindings.values()
                ],
            }


def _timed(result: dict, round_trip_ms: int) -> dict:
    """F1: keep the pane's own execution time only if it is a sane number; add the plugin round trip."""
    out = {k: v for k, v in result.items() if k not in ("elapsedMs", "roundTripMs")}
    pane_ms = result.get("elapsedMs")
    if isinstance(pane_ms, (int, float)) and not isinstance(pane_ms, bool) and 0 <= pane_ms <= MAX_ELAPSED_MS:
        out["elapsedMs"] = int(round(pane_ms))
    out["roundTripMs"] = max(0, round_trip_ms)
    return out


def _clean_doc(doc: Any) -> dict:
    doc = doc if isinstance(doc, dict) else {}
    out = {}
    name = doc.get("name")
    if isinstance(name, str):
        out["name"] = name[:200]
    ident = doc.get("id")  # OneNote: the open notebook's id (identity for resume / event checks)
    if isinstance(ident, str) and 0 < len(ident) <= 300:
        out["id"] = ident
    return out


def _clean_host(host: Any) -> dict:
    host = host if isinstance(host, dict) else {}
    return {k: str(host[k])[:64] for k in ("host", "platform", "version") if k in host}


class _Handler(BaseHTTPRequestHandler):
    server_version = "HermesOfficeBridge/1"
    sys_version = ""
    protocol_version = "HTTP/1.1"
    timeout = CONNECTION_IDLE_SECONDS  # idle keep-alive sockets close; see _current() for released listeners

    def log_message(self, format, *args):  # noqa: A002 - quiet: never log request lines or bodies
        return

    # -- helpers --
    @property
    def bridge(self) -> "BridgeServer":
        return self.server.bridge  # type: ignore[attr-defined]

    def _origin(self) -> str:
        origins = self.headers.get_all("Origin", [])
        if origins:
            return origins[0] if len(origins) == 1 else ""
        # A same-origin GET/HEAD carries no Origin header (Fetch spec). When this bridge serves the pane itself (the
        # loopback default), such a request is the page's own origin unless the browser marks it as coming from
        # another site. Cross-origin fetches always send Origin, so they must still match.
        if self.command in ("GET", "HEAD") and self.bridge.loopback \
                and self.headers.get_all("Sec-Fetch-Site", []) in ([], ["same-origin"]):
            return f"https://localhost:{self.server.server_address[1]}"
        return ""

    def _origin_ok(self) -> bool:
        return self.bridge.origin_allowed(self._origin())

    def _host_ok(self) -> bool:
        return self.headers.get_all("Host", []) == [f"localhost:{self.server.server_address[1]}"]

    def _cors(self):
        origin = self._origin()
        if self.bridge.origin_allowed(origin):
            self.send_header("Access-Control-Allow-Origin", origin)
            self.send_header("Vary", "Origin")

    def _json(self, status: int, body: dict):
        data = json.dumps(body, separators=(",", ":"), ensure_ascii=False, allow_nan=False).encode()
        self.send_response(status)
        self._cors()
        self.send_header("Content-Type", "application/json")
        self.send_header("Cache-Control", "no-store")
        self.send_header("Content-Length", str(len(data)))
        self.end_headers()
        self.wfile.write(data)

    def _error(self, status: int, category: str, message: str):
        if status >= 400:
            self.close_connection = True  # an unread request body must not be parsed as the next request
        self._json(status, {"ok": False, "error": {"category": category, "message": message}})

    def _current(self) -> bool:
        """A request on a connection accepted by a listener that has since been released is refused (A14):
        the pane then finds the bridge's current port instead of a lingering keep-alive socket."""
        if self.bridge._httpd is self.server:
            return True
        self.close_connection = True
        self.send_response(503)
        self.send_header("Content-Length", "0")
        self.send_header("Connection", "close")
        self.end_headers()
        return False

    def _gate(self) -> bool:
        if not self._current():
            return False
        if not self._host_ok():
            self._error(421, "Validation", "Unexpected Host")
            return False
        if not self._origin_ok():
            self._error(403, "Authentication", "Origin not allowed")
            return False
        return True

    def _read_json(self) -> Any:
        ctype = self.headers.get("Content-Type", "").split(";", 1)[0].strip().lower()
        if ctype != "application/json":
            raise BridgeError("Validation", "Content-Type must be application/json")
        try:
            length = int(self.headers.get("Content-Length", ""))
        except ValueError:
            raise BridgeError("Validation", "Content-Length required") from None
        if length < 0 or length > MAX_BODY_BYTES:
            raise BridgeError("Capacity", "Body too large")
        try:
            return json.loads(self.rfile.read(length).decode("utf-8"))
        except (ValueError, UnicodeError):
            raise BridgeError("Validation", "Body must be JSON") from None

    def _bearer(self) -> str:
        auth = self.headers.get("Authorization", "")
        return auth[7:].strip() if auth.startswith("Bearer ") else ""

    def _plain(self, status: int, body: bytes = b"", mime: str = "application/octet-stream"):
        """A static-file response: fixed headers, never CORS (the pane loads its files same-origin)."""
        if status >= 400:
            self.close_connection = True
        self.send_response(status)
        self.send_header("Content-Type", mime)
        self.send_header("Content-Length", str(len(body)))
        self.send_header("X-Content-Type-Options", "nosniff")
        self.send_header("Content-Security-Policy", STATIC_CSP)
        self.send_header("Referrer-Policy", "no-referrer")
        self.send_header("Cache-Control", "no-cache")
        if status == 405:
            self.send_header("Allow", "GET, HEAD")
        self.end_headers()
        if self.command != "HEAD":
            self.wfile.write(body)

    def _static(self) -> bool:
        """The bundled add-in under ADDIN_PATH (see read_site_file). False: not a static path."""
        raw = urlsplit(self.path).path
        if not raw.startswith(ADDIN_PATH):
            return False
        if not self._current():
            return True
        if not self._host_ok():
            self._plain(421)
        elif self.command not in ("GET", "HEAD"):
            self._plain(405)
        elif self.headers.get_all("Origin", []) and not self.bridge.origin_allowed(self._origin()):
            self._plain(403)  # an iframe navigation or same-origin load carries none; a foreign page's fetch does
        else:
            item = read_site_file(self.bridge.site, raw[len(ADDIN_PATH):])
            if item is None:
                self._plain(404)
            else:
                self._plain(200, *item)
        return True

    # -- routes --
    def do_OPTIONS(self):
        with self.bridge._request():
            if not self._static():
                self._options()

    def do_GET(self):
        with self.bridge._request():
            if not self._static():
                self._get()

    def do_HEAD(self):
        with self.bridge._request():
            if not self._static():
                self._not_allowed()

    def do_POST(self):
        with self.bridge._request():
            if not self._static():
                self._post()

    def do_PUT(self):
        with self.bridge._request():
            if not self._static():
                self._not_allowed()

    do_DELETE = do_PATCH = do_PUT

    def _not_allowed(self):
        if self._gate():
            self._error(405, "Validation", "Method not allowed")

    def _options(self):
        if not self._current():
            return
        if not self._host_ok() or not self._origin_ok():
            self.send_response(403)
            self.send_header("Content-Length", "0")
            self.end_headers()
            return
        self.send_response(204)
        self._cors()
        self.send_header("Access-Control-Allow-Methods", "GET, POST")
        self.send_header("Access-Control-Allow-Headers", "Content-Type, Authorization")
        # Legacy Private Network Access preflight opt-in (pre-LNA Chromium).
        if self.headers.get("Access-Control-Request-Private-Network") == "true":
            self.send_header("Access-Control-Allow-Private-Network", "true")
        self.send_header("Access-Control-Max-Age", "600")
        self.send_header("Content-Length", "0")
        self.end_headers()

    def _get(self):
        if not self._gate():
            return
        url = urlsplit(self.path)
        try:
            if url.path == "/v1/hello":
                # instance: random, per profile bridge, stable across an idle release and re-acquire, so a pane
                # can find ITS bridge on whichever port it holds now (and never send its token to another's).
                return self._json(200, {"ok": True, "service": SERVICE, "protocol": PROTOCOL, "apps": list(APPS),
                                        "instance": self.bridge.instance})
            if url.path == "/v1/next":
                b = self.bridge.registry.authenticate(self._bearer())
                qs = parse_qs(url.query)
                try:
                    wait = float(qs.get("wait", ["20"])[0])
                except ValueError:
                    wait = 20.0
                cmd = self.bridge.registry.next_command(b, wait)
                if cmd is None:
                    return self._json(200, {"ok": True, "command": None})
                return self._json(200, {"ok": True, "command": {
                    "requestId": cmd.request_id, "tool": cmd.tool, "arguments": cmd.arguments}})
            return self._error(404, "Validation", "Not found")
        except BridgeError as exc:
            return self._error(401 if exc.category == "Authentication" else 400, exc.category, str(exc))

    def _post(self):
        if not self._gate():
            return
        path = urlsplit(self.path).path
        try:
            body = self._read_json()
            if not isinstance(body, dict):
                raise BridgeError("Validation", "Body must be an object")
            if path == "/v1/pair":
                b = self.bridge.registry.pair(body.get("code"), body.get("app"), body.get("document"),
                                              body.get("host"), body.get("supportedTools"),
                                              previous_token=body.get("previousToken"))
                return self._json(200, {"ok": True, "paneToken": b.token, "bindingId": b.binding_id,
                                        "app": b.app, "protocol": PROTOCOL})
            if path == "/v1/resume":
                b = self.bridge.registry.resume(self._bearer(), body.get("document"), body.get("host"),
                                                body.get("supportedTools"), app=body.get("app"))
                return self._json(200, {"ok": True, "bindingId": b.binding_id, "app": b.app, "protocol": PROTOCOL})
            if path == "/v1/result":
                b = self.bridge.registry.authenticate(self._bearer())
                self.bridge.registry.complete(b, body.get("requestId"), body.get("result"))
                return self._json(200, {"ok": True})
            if path == "/v1/event":
                if len(json.dumps(body)) > MAX_EVENT_BYTES:
                    raise BridgeError("Capacity", "Event too large")
                b = self.bridge.registry.authenticate(self._bearer())
                return self._json(200, {"ok": True, **self.bridge.registry.event(b, body)})
            if path == "/v1/bye":
                self.bridge.registry.unbind_token(self._bearer())
                return self._json(200, {"ok": True})
            return self._error(404, "Validation", "Not found")
        except BridgeError as exc:
            status = {"Authentication": 401, "RateLimited": 429, "Capacity": 413}.get(exc.category, 400)
            return self._error(status, exc.category, str(exc))


class _QuietServer(ThreadingHTTPServer):
    daemon_threads = True
    ssl_context = None  # set by BridgeServer.ensure_started: every connection is TLS
    # W1: on Windows, SO_REUSEADDR (http.server's default) lets a second socket bind a port that another bridge is
    # already listening on: both binds succeed and either can answer the pane. Bind exclusively there, so a taken port
    # fails and ensure_started moves on to the next one. On macOS/Linux SO_REUSEADDR never shares a listening port.
    allow_reuse_address = not hasattr(socket, "SO_EXCLUSIVEADDRUSE")

    def server_bind(self):
        if hasattr(socket, "SO_EXCLUSIVEADDRUSE"):
            self.socket.setsockopt(socket.SOL_SOCKET, socket.SO_EXCLUSIVEADDRUSE, 1)
        super().server_bind()

    def get_request(self):
        sock, address = super().get_request()
        if self.ssl_context is not None:
            try:  # the handshake runs in the connection's own thread (process_request_thread), never the accept loop
                sock = self.ssl_context.wrap_socket(sock, server_side=True, do_handshake_on_connect=False)
            except OSError:
                sock.close()
                raise
        return sock, address

    def process_request_thread(self, request, client_address):
        if self.ssl_context is not None:
            try:
                request.settimeout(TLS_HANDSHAKE_SECONDS)
                request.do_handshake()
            except (OSError, ValueError):  # plain HTTP, an untrusting client, a stalled peer: drop it quietly
                self.shutdown_request(request)
                return
        super().process_request_thread(request, client_address)

    def handle_error(self, request, client_address):  # peer resets are normal; never log request data
        return


class BridgeServer:
    """One profile's bridge: a ThreadingHTTPServer on the first free port of ``ports``.

    A14: the listener is not held for the process lifetime. With nothing to serve for
    ``idle_release_seconds`` (see Registry.idle_for_release, plus no tool call holding ``in_use`` and no
    request being handled) the port is given back; the registry (pairings, tokens, OneNote state) stays,
    and ``ensure_started`` takes the first free port again on the next Office tool call. The pane finds the
    bridge by its ``instance`` id on whichever port that is (addin/shared/bridge-client.js).

    HTTPS only: ``tls`` is the install's bridge_tls.LocalTLS (or an ssl.SSLContext). ``allowed_origin`` None is the
    plugin default (the bundled add-in from ``site``, origins https://localhost:<each of ports>); an https origin is a
    self-hosted add-in site, accepted exactly.
    """

    def __init__(self, allowed_origin: str | None = None, ports=DEFAULT_PORTS, registry: Registry | None = None,
                 idle_release_seconds: float = IDLE_RELEASE_SECONDS, reap_interval: float = REAP_INTERVAL_SECONDS,
                 tls: Any = None, site: Path | None = None):
        self.ports = tuple(ports)
        if allowed_origin is None:
            self.allowed_origin = None
            self.allowed_origins = frozenset(f"https://localhost:{p}" for p in self.ports)
        else:
            self.allowed_origin = _check_origin(allowed_origin)  # the add-in base's origin (resolve_addin_base)
            self.allowed_origins = frozenset({self.allowed_origin})
        self.tls = tls
        self.site = Path(site) if site is not None and allowed_origin is None else None
        self.registry = registry or Registry()
        self.idle_release_seconds = float(idle_release_seconds)
        self.reap_interval = float(reap_interval)
        self.instance = secrets.token_hex(8)
        self.releases = 0  # how many times the idle port was given back (status / tests)
        self.last_start_error: Optional[str] = None
        self._httpd: ThreadingHTTPServer | None = None
        self._thread: threading.Thread | None = None
        self._reap_stop: threading.Event | None = None
        self._lock = threading.Lock()  # start / release / stop
        self._count_lock = threading.Lock()  # holds and requests (never taken while waiting on _lock)
        self._holds = 0
        self._requests = 0

    @property
    def port(self) -> Optional[int]:
        httpd = self._httpd
        return httpd.server_address[1] if httpd else None

    @property
    def loopback(self) -> bool:
        """The plugin default: the pane comes from this bridge's own port (or a sibling Workbench port)."""
        return self.allowed_origin is None

    def origin_allowed(self, origin: str) -> bool:
        return isinstance(origin, str) and origin in self.allowed_origins

    def _tls_context(self):
        if self.tls is None:
            raise BridgeError("TLS", "The Workbench bridge has no TLS identity configured")
        if not hasattr(self.tls, "ensure"):
            return self.tls
        try:
            return self.tls.ensure()
        except Exception as exc:  # noqa: BLE001 - never echo parser text that might include key material
            raise BridgeError("TLS", "Workbench could not load its local HTTPS certificate "
                                     f"({type(exc).__name__}). Install cryptography>=43 and check this profile's "
                                     "plugin-data/hermes-office-desktop/tls folder.") from None

    # --- activity that must never be cut off by a release ---------------------------------------------
    def _request(self):
        return _Counter(self, "_requests")

    def in_use(self):
        """Context manager a tool call holds while it may use the bridge: no release while held."""
        return _Counter(self, "_holds")

    def ensure_started(self) -> int:
        with self._lock:
            if self._httpd:
                with self.registry._lock:
                    self.registry._touch()
                return self.port  # type: ignore[return-value]
            context = self._tls_context()
            last: Exception | None = None
            for port in self.ports:
                try:
                    httpd = _QuietServer(("127.0.0.1", port), _Handler)
                except OSError as exc:
                    last = exc
                    continue
                httpd.daemon_threads = True
                httpd.ssl_context = context
                httpd.bridge = self  # type: ignore[attr-defined]
                with self.registry._lock:
                    self.registry._draining = False
                    self.registry._touch()  # the idle clock starts now, not at the last pane activity
                self._httpd = httpd
                self.last_start_error = None
                self._thread = threading.Thread(target=httpd.serve_forever, name="office-desktop-bridge",
                                                daemon=True)
                self._thread.start()
                if self.reap_interval > 0 and self.idle_release_seconds > 0:
                    stop = threading.Event()
                    self._reap_stop = stop
                    threading.Thread(target=self._reap, args=(stop,), name="office-desktop-bridge-idle",
                                     daemon=True).start()
                return port
            self.last_start_error = ports_busy_message(self.ports, self.idle_release_seconds)
            raise BridgeError(PORTS_BUSY, self.last_start_error + f" ({type(last).__name__})")

    def _reap(self, stop: threading.Event) -> None:
        while not stop.wait(self.reap_interval):
            if self.maybe_release():
                return

    def release_in(self) -> Optional[float]:
        """Seconds until the idle port is given back, None while listening is not idle (or not listening)."""
        if self._httpd is None or self.idle_release_seconds <= 0:
            return None
        with self._count_lock:
            if self._holds or self._requests:
                return None
        idle, left = self.registry.idle_for_release(self.idle_release_seconds)
        return left if idle else None

    def maybe_release(self) -> bool:
        """Give the port back if nothing is using this bridge and it has been idle long enough.

        Race safety: the idle check and the registry's draining flag are set together under the registry
        lock while no tool call holds ``in_use`` and no request is being handled; while draining, a pairing
        attempt is refused (never answered with a token for a closing listener). A tool call that arrives
        meanwhile waits on ``_lock`` in ensure_started and simply takes a port again.
        """
        if self.idle_release_seconds <= 0:
            return False
        with self._lock:
            httpd = self._httpd
            if httpd is None:
                return False
            with self._count_lock:
                if self._holds or self._requests:
                    return False
                idle, left = self.registry.idle_for_release(self.idle_release_seconds)
                if not idle or left > 0:
                    return False
                with self.registry._lock:
                    self.registry._draining = True
            try:
                self._httpd = None  # requests on lingering keep-alive connections are refused from now on
                if self._reap_stop is not None:
                    self._reap_stop.set()
                    self._reap_stop = None
                httpd.shutdown()
                httpd.server_close()
                self.releases += 1
            finally:
                with self.registry._lock:
                    self.registry._draining = False
            return True

    def stop(self) -> None:
        with self._lock:
            if self._reap_stop is not None:
                self._reap_stop.set()
                self._reap_stop = None
            if self._httpd:
                httpd, self._httpd = self._httpd, None
                httpd.shutdown()
                httpd.server_close()


class _Counter:
    def __init__(self, server: BridgeServer, attr: str):
        self._server, self._attr = server, attr

    def __enter__(self):
        with self._server._count_lock:
            setattr(self._server, self._attr, getattr(self._server, self._attr) + 1)
        return self._server

    def __exit__(self, *exc):
        with self._server._count_lock:
            setattr(self._server, self._attr, getattr(self._server, self._attr) - 1)
        return False
