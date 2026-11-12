# SPDX-License-Identifier: Apache-2.0
"""Workbench: drive Excel, Word, PowerPoint, OneNote and Outlook open in the Hermes Desktop preview pane.

Agent half of a free plugin. The user chats in Hermes Desktop; the document is open
in Office for the web inside the Desktop preview pane; a free bridge add-in (task
pane served by this plugin's own loopback HTTPS bridge from the pinned ``addin/`` files, or from a self-hosted
``addin_base``) executes Office.js on that live document. This module registers the tools, the session<->pane
pairing, undo, the mutation approval hook and the ``/workbench`` command. No Microsoft app registration, Graph or
OAuth is used.
"""
from __future__ import annotations

import hashlib
import http.client
import importlib.util
import json
import os
import re
import socket
import tempfile
import threading
import uuid
from pathlib import Path

from . import images as _img
from . import onenote as _on
from . import outlook as _ol
from . import profile_scope as _scope
from .bridge_server import (APPS, DEFAULT_ADDIN_BASE, DEFAULT_PORTS, PANE_UNLOADED, PORTS_BUSY, SERVICE, SITE_ROOT,
                            BridgeError, BridgeServer, busy_ports, format_code, loopback_base, origin_of,
                            ports_busy_message, resolve_addin_base)
from .bridge_tls import LocalTLS
from .validation import validate

HERE = Path(__file__).parent
PLUGIN_ID = "hermes-office-desktop"
COMMAND = "workbench"  # /workbench and `hermes workbench`: manifests for the running bridge, the trust step, status
MANIFEST_RECORD = ".omnia-works-manifests.json"  # hashes of the manifests this plugin wrote (never a user's file)
WEB_URLS = {"excel": "https://excel.cloud.microsoft/", "word": "https://word.cloud.microsoft/",
            "powerpoint": "https://powerpoint.cloud.microsoft/", "onenote": "https://www.onenote.com/notebooks",
            "outlook": "https://outlook.live.com/mail/"}
LABELS = {"excel": "Excel", "word": "Word", "powerpoint": "PowerPoint", "onenote": "OneNote", "outlook": "Outlook"}
NOUNS = {"excel": "workbook", "word": "document", "powerpoint": "presentation", "onenote": "notebook", "outlook": "mailbox"}
# Hosts whose tool schemas are the native bridge's canonical artifacts (schemas/<app>.json).
# OneNote has no native counterpart; its schemas live in onenote.py.
NATIVE_APPS = ("excel", "word", "powerpoint")

# Canonical schemas reused byte-for-byte from the native bridge (schemas/<app>.json).
# excel_execute_plan is excluded on purpose: its trusted-turn, replay and fingerprint
# guards rely on a companion-side plan authority and host turn metadata that the
# browser pane cannot provide (README "Excel plans").
EXCLUDED = frozenset({"excel_execute_plan"})
# Explicit read-only allowlist (same as the native bridge): every other tool needs approval.
READ_ONLY = frozenset({
    "excel_get_workbook_overview", "excel_read_range", "excel_summarize_range_stats",
    "word_get_document_overview", "word_read_selection", "word_read_document_text", "word_get_document_structure",
    "powerpoint_get_presentation_overview", "powerpoint_read_selected_slides", "powerpoint_read_all_slides",
    "powerpoint_read_slide_content", "powerpoint_list_slide_layouts",
    "office_connect", "office_status", "office_disconnect",
})
APPROVAL_RULE_KEY = "hermes-office-desktop:document-mutation"

# Next step for the agent, per error category (kept short; surfaced as error.hint).
HINTS = {
    "NotPaired": "Call office_connect, show the user the code, and wait for them to type it into the Workbench task pane.",
    "Unavailable": ("The Workbench pane is not running: Hermes Desktop keeps one chat per pane, so switching chats or "
                    "closing the preview unloads it, and Office does not reopen task panes by itself. Ask the user to "
                    "come back to this chat, show the document in the preview pane and click Workbench in the "
                    "ribbon (it reconnects by itself), or call office_connect again."),
    "Timeout": "The pane never received the command, so nothing changed. Check the pane is open, then retry once.",
    "ExecutionUncertain": "Read the affected part of the document before deciding whether to retry; never retry blindly.",
    "Unsupported": "This Office version or the pane cannot do this; tell the user and offer an alternative.",
    "Validation": "Fix the arguments to match the tool schema and the document (sheet names, addresses, sizes).",
    "Office": "Office rejected the call; read the document state (overview/read tools) and adjust.",
    "Partial": ("Only part of the step applied: the message lists what was restored or changed and what was NOT. Read the "
                "item, then tell the user exactly which fields to fix by hand (the message gives the values); do not retry "
                "the undo, it is no longer in the undo history."),
    "Conflict": "Read the current document state; the requested change conflicts with it.",
    "NothingToUndo": "There is no Hermes change left to undo in this pane (undo history clears when the pane reloads).",
    "Unattended": "Office tools need the user watching the pane; do this from an interactive Hermes Desktop chat.",
    "Blocked": ("This image source is not allowed (see the message). Use a file in the Hermes image folders or a "
                "public https URL; do not try another route around the check."),
    "WrongType": "Only PNG and JPEG images can be inserted; convert or pick another image.",
    "Capacity": "Use a smaller image (4 MiB at most) or a shorter request.",
    "Network": "The image could not be fetched; nothing changed. Tell the user, or pick another image.",
    "PaneReload": ("OneNote stopped answering the Workbench pane (every request fails until the pane reloads). Do not "
                   "call more OneNote tools now: tell the user to reload the OneNote browser tab (or close and reopen "
                   "the Workbench pane; it reconnects to this chat by itself), then try again once. If it was a "
                   "change, it may or may not have been made: read the page first; an identical change is held back "
                   "until a successful read of that page (or, for a new page, a complete page listing) has shown "
                   "whether it landed."),
    "AlreadyApplied": ("A read after the earlier identical change showed it is already there, so nothing was sent now. "
                       "Tell the user it is done; do not call it again."),
    "PaneBusy": ("The pane is still working on an earlier read and could not answer in time; nothing was changed. Do not "
                 "repeat a whole-notebook read: list one section (sectionId), wait a moment, then retry once."),
    _scope.CATEGORY: _scope.HINT,
    PORTS_BUSY: ("Tell the user exactly what the message says (which ports, close a pane in another profile's chat or "
                 "wait); do not retry in a loop. Nothing was changed."),
    "NeedsRaster": ("Render the HTML in the browser tool, take a PNG screenshot of just that element, save it into one of "
                    "needsRaster.saveTo, then call office_insert_image again with path."),
    "TLS": ("Tell the user the message: the bridge's local HTTPS certificate could not be prepared, so nothing was "
            "started or changed. Do not change certificate trust yourself."),
}

_server_lock = threading.Lock()
_server: BridgeServer | None = None
# After an idle release (bridge_server.IDLE_RELEASE_SECONDS) a tool call takes a port again; give a pane that
# is still open this long to find the bridge's new port before the call is judged (its retry is <= 8 s).
REACQUIRE_PANE_WAIT = 10.0


def _json(obj) -> str:
    return json.dumps(obj, separators=(",", ":"), ensure_ascii=False, allow_nan=False)


def _err(category: str, message: str) -> str:
    error = {"category": category, "message": message}
    if category in HINTS:
        error["hint"] = HINTS[category]
    return _json({"ok": False, "error": error})


def _with_hint(result: dict) -> dict:
    if isinstance(result, dict) and result.get("ok") is False and isinstance(result.get("error"), dict):
        cat = result["error"].get("category")
        if cat in HINTS and "hint" not in result["error"]:
            result = {**result, "error": {**result["error"], "hint": HINTS[cat]}}
    return result


def load_schemas() -> dict[str, list[dict]]:
    """Canonical schemas per app, minus EXCLUDED; refuses anything malformed or identity-bearing."""
    out: dict[str, list[dict]] = {}
    seen: set[str] = set()
    for app in NATIVE_APPS:
        value = json.loads((HERE / "schemas" / f"{app}.json").read_text(encoding="utf-8"))
        if not isinstance(value, list):
            raise RuntimeError("Schema artifact must be an array")
        kept = []
        for s in value:
            name = s.get("name") if isinstance(s, dict) else None
            if not isinstance(name, str) or not name.startswith(app + "_") or not isinstance(s.get("parameters"), dict):
                raise RuntimeError("Invalid canonical schema")
            if "bridgeSessionId" in s["parameters"].get("properties", {}):
                raise RuntimeError("Session identity must never be a model argument")
            if name in seen:
                raise RuntimeError("Duplicate canonical tool name")
            seen.add(name)
            if name not in EXCLUDED:
                kept.append(s)
        out[app] = kept
    return out


SCHEMAS = load_schemas()
DOCUMENT_TOOLS = frozenset(s["name"] for schemas in SCHEMAS.values() for s in schemas)


def session_key(session_id) -> str:
    """Hermes session ids are opaque strings (not UUIDs); bound and validate only."""
    if not isinstance(session_id, str) or not 1 <= len(session_id) <= 200 \
            or any(not 33 <= ord(c) <= 126 for c in session_id):
        raise BridgeError("Validation", "This tool needs a Hermes chat session (no session id was supplied)")
    return session_id


def _session_marker(name: str) -> str:
    """A non-secret Hermes session marker (cron / platform) for the current turn.

    Public surface only (catalog rule 9): ``gateway.session_context.get_session_env``, which
    reads the per-turn ContextVar a gateway-hosted turn binds and otherwise falls back to the
    process environment. Without the gateway package (a standalone run) it is ``os.getenv``.
    """
    try:
        from gateway.session_context import get_session_env
    except ImportError:
        return os.getenv(name, "") or ""
    return get_session_env(name, "") or ""


UNATTENDED_PLATFORMS = frozenset({"webhook", "msgraph_webhook", "api_server", "cron"})
# ctx.current_cron_execution, bound by register(): the cron fire this code runs inside, else None.
_cron_execution = None
_plugin_ctx = None  # the registering ctx: lets the approval prompt read this profile's index (two-way links)


def unattended_reason() -> str | None:
    """Rule 12: Office tools act on a document a person is looking at; never in cron or headless runs.

    Unattended means any of: ``ctx.current_cron_execution()`` reports a cron fire (a job hosted
    inside the gateway, which exports nothing to the process environment); the ``HERMES_CRON_SESSION``
    marker is set; or the session platform is webhook, msgraph_webhook, api_server or cron
    (both markers read per turn through ``get_session_env``, else the process environment).
    Hermes' own approval gate already refuses unattended mutations unless an operator set
    ``approvals.cron_mode``/``unattended_mode`` to approve; for those runs this refusal holds
    even then, and it covers read-only Office tools too.
    """
    if _cron_execution is not None and _cron_execution() is not None:
        return "a scheduled (cron) run"
    if _session_marker("HERMES_CRON_SESSION").strip().lower() in ("1", "true", "yes", "on"):
        return "a scheduled (cron) run"
    platform = _session_marker("HERMES_SESSION_PLATFORM").strip().lower()
    if platform in UNATTENDED_PLATFORMS:
        return f"an unattended {platform} session"
    return None


def _refuse_if_unattended() -> None:
    reason = unattended_reason()
    if reason:
        raise BridgeError("Unattended", f"Office tools are refused in {reason}: no one is watching the document.")


def addin_base(ctx) -> str:
    """The one source of truth: the add-in site's base URL (plugin setting addin_base; default this plugin's own
    loopback site, DEFAULT_ADDIN_BASE, whose port follows the bridge (served_base); an older allowed_origin alone
    means origin + /office/live/). The manifests, the image User-Agent and every URL the plugin mentions use it."""
    if not ctx:
        return resolve_addin_base()
    return resolve_addin_base(ctx.get_config("addin_base", None), ctx.get_config("allowed_origin", None))


def addin_origin(ctx) -> str:
    return origin_of(addin_base(ctx))


def is_loopback(ctx) -> bool:
    """The default: the bridge serves the bundled add-in itself (no add-in site configured)."""
    return addin_base(ctx) == DEFAULT_ADDIN_BASE


def served_base(ctx, port: int) -> str:
    """The base the manifests must name now: the bound port's loopback site, or the self-hosted base."""
    return loopback_base(port) if is_loopback(ctx) else addin_base(ctx)


def _tls_dir(ctx) -> Path:
    """This profile's plugin data folder (the owning profile's home; never a guessed one). TLS goes to <it>/tls."""
    return _scope.owner(ctx) / "plugin-data" / PLUGIN_ID


def local_tls(ctx) -> LocalTLS:
    return LocalTLS(_tls_dir(ctx))


def get_server(ctx) -> BridgeServer:
    global _server
    with _server_lock:
        if _server is None:
            loopback = is_loopback(ctx)
            # Fixed port set: the pane probes exactly these (no user-tunable port to drift).
            # _test_ports exists only for the unit tests' isolated ranges.
            ports = getattr(ctx, "_test_ports", None) or DEFAULT_PORTS
            _server = BridgeServer(allowed_origin=None if loopback else addin_origin(ctx), ports=tuple(ports),
                                   tls=local_tls(ctx), site=SITE_ROOT if loopback else None)
            _server.registry.on_event = _on.make_event_handler(ctx)
        return _server


class _LoopbackHTTPS(http.client.HTTPSConnection):
    """https://localhost:<port> dialled on 127.0.0.1 (the bridge binds IPv4 loopback only); the certificate is still
    checked for the name localhost."""

    def __init__(self, port: int, context, timeout: float = 2.0):
        super().__init__("localhost", port, timeout=timeout, context=context)
        self._tls_context = context

    def connect(self):
        sock = socket.create_connection(("127.0.0.1", self.port), self.timeout)
        self.sock = self._tls_context.wrap_socket(sock, server_hostname="localhost")


def find_running_port(ctx):
    """The port of THIS profile's running bridge, for a separate `hermes workbench` process: a hello over HTTPS that
    trusts only this profile's CA (another profile's bridge fails the handshake). Binds and starts nothing."""
    tls = local_tls(ctx)
    if not tls.ca_path.is_file():
        return None
    try:
        context = tls.client_context()
    except (OSError, ValueError):
        return None
    for port in getattr(ctx, "_test_ports", None) or DEFAULT_PORTS:
        conn = _LoopbackHTTPS(port, context)
        try:
            conn.request("GET", "/v1/hello", headers={"Origin": f"https://localhost:{port}"})
            reply = conn.getresponse()
            body = reply.read(4096)
            if reply.status == 200 and json.loads(body.decode("utf-8")).get("service") == SERVICE:
                return port
        except (OSError, http.client.HTTPException, ValueError):
            pass
        finally:
            conn.close()
    return None


def _reacquire(server: BridgeServer, key: str, app: str, not_paired: str) -> None:
    """Make sure this profile's bridge listens for a chat that has a pairing (A14 lazy re-acquire).

    A chat with no pairing at all gets NotPaired (nothing is started for it). A chat whose pairing survived an
    idle release takes a port again (PortsBusy if none is free) and, if its pane is still open, waits briefly
    for the pane to find the new port. Call inside ``server.in_use()``.
    """
    was_listening = server.port is not None
    if not server.registry.has_binding(key):
        if was_listening:
            return  # unchanged path: the tool reports its own NotPaired for this chat
        raise BridgeError("NotPaired", not_paired)
    server.ensure_started()
    b = server.registry.binding_for(key, app)
    if not was_listening and b is not None:
        server.registry.wait_live(b, REACQUIRE_PANE_WAIT)


def _busy_report(server: BridgeServer) -> dict:
    """office_status: whether this profile could take a bridge port right now (probe only; binds nothing)."""
    taken = busy_ports(server.ports)
    out = {"bridgePortsFree": len(server.ports) - len(taken)}
    if len(taken) >= len(server.ports):
        msg = ports_busy_message(server.ports, server.idle_release_seconds)
        out["bridgeBusy"] = {"category": PORTS_BUSY, "message": msg, "hint": HINTS[PORTS_BUSY]}
        out["tellUser"] = msg
    return out


APP_PROP = {"app": {"type": "string", "enum": list(APPS),
                    "description": "Which Office app's pane: excel, word, powerpoint, onenote or outlook (default excel)."}}
CONNECT_SCHEMA = {
    "name": "office_connect",
    "description": (
        "THE tool for \"connect to Excel\" / \"connect to Word\" / PowerPoint / OneNote / Outlook / Outlook calendar: call it "
        "directly (not manage_connections, which is for hosted accounts such as Gmail and cannot reach Office for the web; "
        "not office_status, which only reports). "
        "Pair the Office document (Excel, Word, PowerPoint or OneNote for the web, or the Outlook on the web mailbox) open in "
        "the Hermes Desktop preview pane "
        "with THIS chat. Returns a one-time code: show it to the user and ask them to type it into the "
        "'Workbench' task pane in that document. Never enter the code yourself (not with drive_preview or any "
        "other tool). Before opening anything, check what the pane shows with desktop_preview (action read); open "
        "Office with desktop_preview (action open) only if no Office document is showing: opening the Office start "
        "page over an open document navigates away from it and unloads its task pane. One chat per pane: the pane "
        "unloads when the user switches to another chat. Does not wait; call office_status to confirm. Requires the "
        "Hermes Desktop app: if desktop_preview read reports surface 'iframe' (Desktop in a browser tab, the "
        "Webapp), Office for the web cannot run in that pane, so do not call office_connect; tell the user to "
        "open this chat in the Hermes Desktop app on the machine that runs Hermes."),
    "parameters": {"type": "object", "properties": APP_PROP, "additionalProperties": False},
}
STATUS_SCHEMA = {
    "name": "office_status",
    "description": ("Report whether an Office pane is paired with this chat and still live (it polled within the "
                    "last 45 s; lastPollAgeMs says when). paneState is live, unloaded or unpaired. It never pairs: "
                    "to connect an app, call office_connect (data.next says so when nothing is paired). bridgeBusy "
                    "(when present) says all bridge ports are taken by other profiles: tell the user its message. No "
                    "document access."),
    "parameters": {"type": "object", "properties": APP_PROP, "additionalProperties": False},
}
DISCONNECT_SCHEMA = {
    "name": "office_disconnect",
    "description": "Unpair the Office pane from this chat.",
    "parameters": {"type": "object", "properties": APP_PROP, "additionalProperties": False},
}
UNDO_SCHEMA = {
    "name": "office_undo",
    "description": (
        "Undo the most recent change Hermes made through the paired pane (one step per call, newest first). "
        "Each mutating tool result says whether it can be undone (data.undo.available). Undo history lives in the "
        "pane and is cleared when the pane reloads; it never undoes the user's own edits. For app 'onenote' undo "
        "walks the local change ledger instead (survives reloads): Hermes-added blocks are removed; a page Hermes "
        "created is marked for the user to delete and opened (add-ins cannot delete pages)."),
    "parameters": {"type": "object", "properties": APP_PROP, "additionalProperties": False},
}


def _app(args) -> str:
    return (args or {}).get("app", "excel")


# ---- add-in manifests in a user-visible folder (F2) -------------------------------------------------------
_build_mod = None


def _build_addin():
    """scripts/build_addin.py, loaded once: its render_manifests is the only manifest generator (never forked)."""
    global _build_mod
    if _build_mod is None:
        spec = importlib.util.spec_from_file_location("hermes_office_desktop_build_addin",
                                                      HERE / "scripts" / "build_addin.py")
        mod = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(mod)
        _build_mod = mod
    return _build_mod


def _default_manifest_dir() -> Path:
    return Path.home() / "Documents" / "Omnia Works" / "add-in manifests"


def _manifest_dir(ctx) -> Path:
    raw = ctx.get_config("manifest_dir", "") if ctx else ""
    if not raw:
        return _default_manifest_dir()
    folder = Path(os.path.expanduser(str(raw)))
    if not folder.is_absolute():
        raise OSError("manifest_dir must be an absolute folder")
    return folder


def _sha(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()


def _atomic_write(target: Path, data: bytes) -> None:
    fd, tmp = tempfile.mkstemp(dir=target.parent, prefix=".omnia-works-", suffix=".tmp")
    try:
        with os.fdopen(fd, "wb") as f:
            f.write(data)
            f.flush()
            os.fsync(f.fileno())
        os.replace(tmp, target)
    except BaseException:
        try:
            os.unlink(tmp)
        except OSError:
            pass
        raise


def write_manifests(folder: Path, base: str) -> dict:
    """Write the five manifests for ``base`` into ``folder``; returns app -> path of a file holding exactly that text.

    Idempotent (identical files are not rewritten), atomic, and never overwrites a file the user changed: a file is
    replaced only when its hash is the one this plugin last wrote there; otherwise the plugin writes
    ``manifest-<app> (Omnia Works).xml`` beside it. An app whose two names are both the user's is left out."""
    texts = _build_addin().render_manifests(base)
    folder.mkdir(parents=True, exist_ok=True)
    rec_path = folder / MANIFEST_RECORD
    try:
        record = json.loads(rec_path.read_text(encoding="utf-8"))
        record = record if isinstance(record, dict) else {}
    except (OSError, ValueError):
        record = {}
    before = dict(record)
    written = {}
    for app, text in texts.items():
        data = text.encode("utf-8")
        want = _sha(data)
        for name in (f"manifest-{app}.xml", f"manifest-{app} (Omnia Works).xml"):
            target = folder / name
            if target.exists() or target.is_symlink():
                current = _sha(target.read_bytes()) if target.is_file() and not target.is_symlink() else None
                if current == want:
                    written[app] = target
                    record[name] = want
                    break
                if current is None or current != record.get(name):
                    continue  # the user's file (changed, or never ours): leave it alone
            _atomic_write(target, data)
            record[name] = want
            written[app] = target
            break
    if record != before:
        _atomic_write(rec_path, (json.dumps(record, indent=1, sort_keys=True) + "\n").encode("utf-8"))
    return written


def _manifest_for(ctx, app: str, base: str):
    """(absolute manifest path, where-to-find-it text), or (None, None) when the folder could not be written."""
    try:
        folder = _manifest_dir(ctx)
        path = write_manifests(folder, base).get(app)
    except (OSError, SystemExit, ValueError):
        return None, None
    if path is None:
        return None, None
    where = "Documents > Omnia Works > add-in manifests" if folder == _default_manifest_dir() else str(folder)
    return path.resolve() if not path.is_absolute() else path, where


def _upload_step(ctx, app: str, path, where) -> str:
    if path is not None:
        choose = f"choose {path.name} in {where}"
    elif is_loopback(ctx):
        choose = (f"choose manifest-{app}.xml (the manifest folder could not be written: set manifest_dir to a "
                  f"writable folder, then run /{COMMAND} manifest)")
    else:
        choose = (f"choose manifest-{app}.xml saved from {addin_base(ctx)}manifest-{app}.xml (the manifest folder "
                  "could not be written)")
    if app == "outlook":
        return ("Once: in a message, More apps (or Apps) > Get add-ins > My add-ins > Custom Addins > Add a custom "
                f"add-in > Add from File..., {choose}, then Install.")
    # A24: the path that worked in every live run (the harness's first try, Home > Add-ins > More Settings /
    # Advanced, failed in every PowerPoint and one Word preparation). OneNote has its add-ins under Insert.
    start = "Insert > Office Add-ins" if app == "onenote" else "Home > Add-ins > More Add-ins"
    return f"{start} > My Add-ins > Manage My Add-ins > Upload My Add-in > Browse, {choose}, then Upload."


OUTLOOK_SETUP = [
    f"If the preview pane does not already show your mailbox, open {WEB_URLS['outlook']} there and sign in.",
    None,  # the upload step: _upload_step(ctx, "outlook", ...)
    "Open a message (or a draft), choose More apps (Apps in a draft) > Workbench, and type the pairing code. Pin the "
    "pane to keep it open as you move between messages.",
    "Calendar: the same add-in. Open an event (or its edit form) in Outlook's calendar, choose Apps (or ...) > Omnia "
    "Works, and type the pairing code there.",
]
OUTLOOK_PRIVACY = ("While paired, when you ask, the open message (subject, people, date, text and attachment names, never "
                   "attachment contents) or the open calendar event (subject, times, location, organizer and attendee "
                   "names with responses, text) is sent to Hermes on this computer; the calendar itself is never listed. "
                   + _ol.NEVER_SENDS + " " + _ol.NEVER_INVITES)


def connect_handler(ctx):
    def handler(args, session_id=None, **_):
        try:
            _refuse_if_unattended()
            key = session_key(session_id)
            validate(args or {}, CONNECT_SCHEMA["parameters"])
            app = _app(args)
            server = get_server(ctx)
            with server.in_use():  # no idle release between taking the port and minting the code
                port = server.ensure_started()
                code, ttl = server.registry.mint_code(key, app, home=_scope.try_call_home(ctx))
            label, noun = LABELS[app], NOUNS[app]
            base = served_base(ctx, port)
            manifest_path, where = _manifest_for(ctx, app, base)
            upload = _upload_step(ctx, app, manifest_path, where)
            trust = _trust(server.tls)
            setup = [x for s in OUTLOOK_SETUP for x in ((trust["step"], upload) if s is None else (s,))] \
                if app == "outlook" else [
                    f"If the preview pane does not already show the {noun}, open {WEB_URLS[app]} there and sign in "
                    "with your Microsoft account.",
                    f"Open or create the {noun}.",
                    trust["step"],
                    upload,
                    "Open the 'Workbench' task pane and type the pairing code.",
                    "One chat per pane: switching to another chat unloads the pane; come back and click Omnia "
                    "Works in the ribbon to reopen it."]
            return _json({"ok": True, "data": {
                "app": app, "pairingCode": format_code(code), "expiresInSeconds": ttl, "bridge": f"127.0.0.1:{port}",
                "addinBase": base,
                "trust": {k: v for k, v in trust.items() if k != "step"},
                "tellUser": (f"In the {label} task pane titled 'Workbench', enter code {format_code(code)}. "
                             "It expires in 10 minutes and works once."),
                **({"privacy": OUTLOOK_PRIVACY} if app == "outlook" else {}),
                **({"manifestPath": str(manifest_path)} if manifest_path is not None else {}),
                "firstTimeSetup": setup}})
        except BridgeError as exc:
            return _err(exc.category, str(exc))
    return handler


def _trust(tls: LocalTLS) -> dict:
    """The one-time manual trust step for this profile's local CA (the bridge serves https://localhost with it).
    The plugin never runs the command; the user reviews and runs it."""
    command = tls.trust_command()
    return {"caCertificate": str(tls.ca_path), "command": command, "tellUser": tls.trust_instructions(),
            "step": ("Once per Hermes profile, before the first upload: trust Workbench's local certificate authority "
                     "so Office can load the pane from https://localhost. The user reviews and runs this command "
                     f"themselves (never run it for them): {command}")}


def _manifests_text(ctx, base: str) -> str:
    folder = _manifest_dir(ctx)
    written = write_manifests(folder, base)
    lines = [f"Workbench add-in manifests for {base}:"]
    lines += [f"- {LABELS[app]}: {written[app]}" for app in APPS if app in written]
    kept = [LABELS[app] for app in APPS if app not in written]
    if kept:
        lines.append("- not written (both file names hold your own edited files): " + ", ".join(kept))
    lines.append("Upload the one for your app once: Office for the web > Add-ins > My Add-ins > Upload My Add-in "
                 "(Outlook: Get add-ins > My add-ins > Add a custom add-in > Add from File).")
    return "\n".join(lines)


def _status_text(ctx, port) -> str:
    tls = local_tls(ctx)
    site = (f"served by this plugin from its pinned add-in files ({loopback_base(port) if port else DEFAULT_ADDIN_BASE}"
            f"{'' if port else ', port follows the bridge'})" if is_loopback(ctx) else addin_base(ctx) + " (self-hosted)")
    return "\n".join([
        "Workbench:",
        f"- bridge: {'https://localhost:' + str(port) if port else 'not running'}",
        f"- add-in site: {site}",
        f"- local CA: {tls.ca_path}{'' if tls.ca_path.is_file() else ' (created when the bridge first starts)'}",
        f"Usage: /{COMMAND} [manifest|trust|status]"])


def workbench_command(ctx):
    """/workbench [manifest|trust|status], in the process that runs this profile's bridge."""
    def handler(raw_args: str = "") -> str:
        words = (raw_args or "").split()
        verb = words[0].lower() if words else "status"
        try:
            if verb == "manifest":
                server = get_server(ctx)
                port = None
                if is_loopback(ctx):
                    with server.in_use():
                        port = server.ensure_started()
                return _manifests_text(ctx, served_base(ctx, port)) + "\n\n" + server.tls.trust_instructions()
            if verb == "trust":
                tls = local_tls(ctx)
                tls.ensure()
                return tls.trust_instructions()
            if verb == "status":
                return _status_text(ctx, _server.port if _server else None)
            return f"Usage: /{COMMAND} [manifest|trust|status]: write the add-in manifests, show the trust step, status."
        except (BridgeError, OSError, ValueError, ImportError) as exc:
            return f"Workbench: {exc}"
    return handler


def _cli_setup(subparser) -> None:
    subparser.add_argument("action", nargs="?", default="status", choices=["status", "trust", "manifest"],
                           help="status, the local CA trust step, or write the add-in manifests for the running bridge")


def cli_handler(ctx):
    """`hermes workbench ...`: a separate process. It never starts a bridge; it finds this profile's running one."""
    def run(args) -> int:
        action = getattr(args, "action", None) or "status"
        try:
            if action == "trust":
                tls = local_tls(ctx)
                tls.ensure()
                print(tls.trust_instructions())
                return 0
            port = find_running_port(ctx) if is_loopback(ctx) else None
            if action == "manifest":
                if is_loopback(ctx) and port is None:
                    print("Workbench: no Workbench bridge of this Hermes profile is running. Ask Hermes to connect an "
                          f"Office app (office_connect starts the bridge and writes the manifests), or run /{COMMAND} "
                          "manifest in a chat.")
                    return 1
                print(_manifests_text(ctx, served_base(ctx, port)))
                print(local_tls(ctx).trust_instructions())
                return 0
            print(_status_text(ctx, port))
            return 0
        except (BridgeError, OSError, ValueError, ImportError) as exc:
            print(f"Workbench: {exc}")
            return 1
    return run


def status_handler(ctx):
    def handler(args, session_id=None, **_):
        try:
            key = session_key(session_id)
            validate(args or {}, STATUS_SCHEMA["parameters"])
            app = _app(args)
            server = get_server(ctx)
            extra: dict = {}
            with server.in_use():
                b = server.registry.binding_for(key, app)
                if b is not None and server.port is None:
                    # A14: this chat's pairing outlived an idle release; take a port again so its pane can find it.
                    # An unpaired chat starts nothing here (scripts/livetest/warm.py relies on that).
                    try:
                        _reacquire(server, key, app, "")
                    except BridgeError as exc:
                        if exc.category != PORTS_BUSY:
                            raise
                        extra.update(_busy_report(server))
                    b = server.registry.binding_for(key, app)
                elif server.port is None:
                    extra.update(_busy_report(server))
                live = server.registry.liveness(b) if b else {"lastPollAgeMs": None, "live": False}
            state = "unpaired" if b is None else ("live" if live["live"] else "unloaded")
            if state == "unloaded" and "tellUser" not in extra:
                extra["tellUser"] = f"The {LABELS[app]} {PANE_UNLOADED}."
            if state == "unpaired" and "bridgeBusy" not in extra:
                # A19: seen live, the model answered "connect to Outlook" with Hermes core's manage_connections or a
                # bare office_status. The next step is spelled out; the listener starts with that office_connect.
                extra["next"] = (f"Nothing is paired for {app}: call office_connect with app '{app}' now. It returns "
                                 "the pairing code for the user to type into the Workbench pane (and starts the "
                                 "bridge if it is not listening). Do not use manage_connections for Office.")
            release_in = server.release_in()
            if release_in is not None:
                extra["idleReleaseInSeconds"] = int(round(release_in))
            return _json({"ok": True, "data": {
                "app": app,
                "bridgeListening": server.port is not None,
                "paired": b is not None,
                "live": live["live"],
                "lastPollAgeMs": live["lastPollAgeMs"],
                "paneState": state,
                **extra,
                "document": (b.document.get("name", "") if b else ""),
                "host": (b.host if b else {}),
                "paneReloads": (b.resumes if b else 0),
                "unattended": unattended_reason(),
                "supportedTools": sorted(b.supported_tools) if b else []}})
        except BridgeError as exc:
            return _err(exc.category, str(exc))
    return handler


def disconnect_handler(ctx):
    def handler(args, session_id=None, **_):
        try:
            key = session_key(session_id)
            validate(args or {}, DISCONNECT_SCHEMA["parameters"])
            removed = get_server(ctx).registry.unbind_session(key, _app(args))
            return _json({"ok": True, "data": {"disconnected": removed}})
        except BridgeError as exc:
            return _err(exc.category, str(exc))
    return handler


def _dispatch(ctx, session_id, app: str, canonical: str, args: dict) -> str:
    _refuse_if_unattended()
    key = session_key(session_id)
    server = get_server(ctx)
    with server.in_use():
        _reacquire(server, key, app, "No Office pane is paired yet. Call office_connect first.")
        return _json(_with_hint(server.registry.call(key, app, canonical, args)))


def _env(ctx, session_id) -> "_on.Env":
    key = session_key(session_id)
    server = get_server(ctx)

    def host_call(app: str, canonical: str, args: dict) -> dict:
        if app == "outlook":
            params = _ol.PANE_ARGS.get(canonical)
            if params is None:
                raise BridgeError("Unsupported", f"outlook {canonical} is not a command of this plugin")
            validate(args, params)
            return _with_hint(server.registry.call(key, app, canonical, args))
        schema = next((s for s in SCHEMAS.get(app, []) if s["name"] == f"{app}_{canonical}"), None)
        if schema is None:
            raise BridgeError("Unsupported", f"{app}_{canonical} is not a tool of this plugin")
        validate(args, schema["parameters"])
        return _with_hint(server.registry.call(key, app, canonical, args))
    return _on.Env(ctx=ctx, server=server, key=key, brain=_on.get_brain(ctx), host_call=host_call)


def undo_handler(ctx):
    def handler(args, session_id=None, **_):
        try:
            validate(args or {}, UNDO_SCHEMA["parameters"])
            if _app(args) == "onenote":
                _refuse_if_unattended()
                return _json(_with_hint(_on.onenote_undo(_env(ctx, session_id))))
            return _dispatch(ctx, session_id, _app(args), "undo_last", {})
        except BridgeError as exc:
            return _err(exc.category, str(exc))
        except Exception:
            return _err("Protocol", "Office execution failed; details withheld")
    return handler


# ---- office_insert_image ---------------------------------------------------------------------------
_NUM = {"type": "number", "minimum": 0, "maximum": 10000}
_SIZE = {"type": "number", "minimum": 1, "maximum": 5000}
IMAGE_SCHEMA = {
    "name": "office_insert_image",
    "description": (
        "Insert ONE picture into the paired Excel, Word, PowerPoint or OneNote pane, from exactly one source: "
        "path (a PNG/JPEG the user's image_generate tool or a browser screenshot saved in the Hermes image cache, or "
        "a folder listed in the plugin setting image_dirs; path may be a file name in your Hermes images folder, such as "
        "team.png: the plugin looks it up in those folders and their subfolders, case-insensitively, and the receipt "
        "names the file it used), url (a public https PNG/JPEG; fetched by the plugin with "
        "SSRF guards, at most 4 MiB; pass credit for the author/licence) or html (Word and OneNote insert sanitized "
        "HTML natively; Excel and PowerPoint answer NeedsRaster: render it with the browser tool, screenshot it to a "
        "PNG and call again with path). alt is the picture's alternative text (defaults to prompt or credit; one of "
        "them is required). Placement: Excel sheet + cell (top-left anchor); Word location end|selection|start; "
        "PowerPoint slideIndex + left/top/width/height in points; OneNote pageId + left/top. One approval per insert. "
        "status verified = inserted and read back; applied = inserted, readback unavailable; partial = inserted but "
        "alt text could not be set; uncertain = check the document before doing anything else. Never retried; "
        "undo with office_undo for that app."),
    "parameters": {"type": "object", "additionalProperties": False, "required": ["app"], "properties": {
        "app": {"type": "string", "enum": [a for a in APPS if a != "outlook"], "description": "Which paired pane gets the picture."},
        "path": {"type": "string", "minLength": 1, "maxLength": 1024, "description": "Absolute path of a local PNG/JPEG, or just its file name (e.g. team.png) when it is in your Hermes images folder."},
        "url": {"type": "string", "minLength": 1, "maxLength": 2048, "description": "https URL of a PNG/JPEG."},
        "html": {"type": "string", "minLength": 1, "maxLength": _img.MAX_HTML_CHARS,
                 "description": "HTML to insert (Word/OneNote) or to render first (Excel/PowerPoint)."},
        "alt": {"type": "string", "minLength": 1, "maxLength": 500, "description": "Alternative text for the picture."},
        "prompt": {"type": "string", "maxLength": 1000, "description": "The image_generate prompt (default alt text)."},
        "credit": {"type": "string", "maxLength": 300, "description": "Author / licence line for a found image."},
        "sheet": {"type": "string", "minLength": 1, "maxLength": 100, "description": "Excel: worksheet (default active)."},
        "cell": {"type": "string", "minLength": 2, "maxLength": 10, "description": "Excel: anchor cell such as B2."},
        "location": {"type": "string", "enum": ["end", "selection", "start"], "description": "Word: where (default end)."},
        "slideIndex": {"type": "integer", "minimum": 0, "description": "PowerPoint: zero-based slide (default selected)."},
        "pageId": {"type": "string", "minLength": 1, "maxLength": 300, "description": "OneNote: the open page's id."},
        "left": _NUM, "top": _NUM, "width": _SIZE, "height": _SIZE}},
}
_SOURCES = ("path", "url", "html")
_HOST_ARGS = {"excel": {"sheet", "cell", "width", "height"}, "word": {"location", "width", "height"},
              "powerpoint": {"slideIndex", "left", "top", "width", "height"}, "onenote": {"pageId", "left", "top", "width", "height"}}
_COMMON = {"app", "path", "url", "html", "alt", "prompt", "credit"}
_CELL = re.compile(r"^[A-Za-z]{1,3}[1-9][0-9]{0,6}$")
RASTER_HOSTS = ("excel", "powerpoint")


def _image_source(args: dict) -> str:
    given = [k for k in _SOURCES if k in args]
    if len(given) != 1:
        raise BridgeError("Validation", "Give exactly one image source: path, url or html")
    app = _app(args)
    extra = set(args) - _COMMON - _HOST_ARGS[app]
    if extra:
        raise BridgeError("Validation", f"{', '.join(sorted(extra))} do not apply to {LABELS[app]}")
    if "cell" in args and not _CELL.match(args["cell"]):
        raise BridgeError("Validation", "cell must be one cell address such as B2")
    return given[0]


def _alt(args: dict) -> str:
    for k in ("alt", "prompt", "credit"):
        v = " ".join(str(args.get(k) or "").split())
        if v:
            return v[:500]
    raise BridgeError("Validation", "alt text is required: pass alt (or the image_generate prompt, or a credit line)")


def _alt_full(alt: str, args: dict, source: dict) -> str:
    bits = [alt]
    credit = " ".join(str(args.get("credit") or "").split())
    if credit and credit != alt:
        bits.append(credit)
    if source.get("type") == "url":
        bits.append(f"source: {source['url']}")
    return " — ".join(bits)[:1000]


def needs_raster(app: str, args: dict) -> dict:
    dirs = [str(d) for d in _img.allowed_dirs()]
    then = {"app": app, "path": "<absolute path of the PNG you saved>",
            **{k: args[k] for k in ("alt", "credit", *sorted(_HOST_ARGS[app])) if k in args}}
    return {"ok": False, "error": {"category": "NeedsRaster", "hint": HINTS["NeedsRaster"],
                                   "message": f"{LABELS[app]} cannot take HTML from an add-in; nothing was changed."},
            "needsRaster": {"app": app, "saveTo": dirs, "steps": [
                "Open the HTML in the browser tool (for example as a data: or file URL in a new tab).",
                "Take a PNG screenshot of just the rendered element (crop to it; keep it under 4 MiB).",
                f"Save the PNG into one of: {', '.join(dirs)} (or a folder the user listed in image_dirs).",
                "Call office_insert_image again with path set to that file."],
                "then": {"tool": "office_insert_image", "args": then}}}


def _image_payload(ctx, app: str, src: str, args: dict, alt: str):
    """(payload for the pane, source record, html report) after the plugin checked the bytes itself."""
    report = None
    if src == "html":
        ua = _img.user_agent(addin_base(ctx))
        clean = _img.sanitize_html(args["html"], app, alt, lambda u: _img.fetch_url(u, user_agent=ua))
        report = {"imagesInlined": clean.inlined, "imagesDropped": len(clean.images) - clean.inlined,
                  "images": clean.images, "removed": clean.dropped}
        source = {"type": "html"}
        if app == "onenote":
            body = clean.html
        else:
            return {"html": clean.html, "alt": alt}, source, report
    else:
        image = _img.load_path(args["path"], ctx.get_config("image_dirs", []) if ctx else []) if src == "path" \
            else _img.fetch_url(args["url"], user_agent=_img.user_agent(addin_base(ctx)))
        source = dict(image.source)
        if args.get("credit"):
            source["credit"] = " ".join(args["credit"].split())
        full = _alt_full(alt, args, source)
        if app != "onenote":
            return {"base64": image.base64, "format": image.kind, "alt": full, "title": alt[:255]}, source, report
        size = "".join(f' {k}="{int(args[k])}"' for k in ("width", "height") if k in args)
        caption = " — ".join(x for x in (source.get("credit"), source.get("url")) if x)
        body = (f'<p><img src="{image.data_uri}" alt="{_img.escape(full)}"{size}></p>'
                + (f"<p><i>{_img.escape(caption)}</i></p>" if caption else ""))
        return {"html": body, "alt": full}, source, report
    return {"html": body}, source, report


def _image_status(ack: dict, readback, err, report, app: str) -> tuple[str, str]:
    """verified / applied / partial / uncertain (e428328 vocabulary) and the exact next step."""
    undo = f"office_undo app='{app}'"
    if err is not None or not isinstance(readback, dict):
        return "applied", (f"The {LABELS[app]} pane acknowledged the insert but the readback was unavailable; look at the "
                           f"document before relying on it. Do not insert it again; undo with {undo} if unwanted.")
    expected = (report or {}).get("imagesInlined")
    count = readback.get("pictures", readback.get("images"))
    # A picture tool is verified only by at least one picture actually read back: an HTML insert must report a
    # count (Word pictures / OneNote images), and any reported count must be >= max(1, pictures inlined).
    need = max(1, expected or 0)
    if (readback.get("exists") is not True or (isinstance(count, int) and count < need)
            or (report is not None and not isinstance(count, int))):
        return "uncertain", (f"The {LABELS[app]} readback does not show the inserted picture"
                             + (f" ({count if isinstance(count, int) else 'no'} of {need} pictures)"
                                if report is not None or isinstance(count, int) else "")
                             + f". Read the document before doing anything else; do not retry. If it is wrong, undo once "
                             f"with {undo}.")
    if ack.get("altTextApplied") is False:
        return "partial", (f"The picture is in the {LABELS[app]} document but its alt text could not be set (this "
                           "Office version lacks the API). Ask the user to add alt text by hand, or undo with " + undo + ".")
    return "verified", f"Nothing to do; undo with {undo} (one call) if unwanted."


def _image_target(app: str, ack: dict, payload: dict) -> dict:
    if app == "excel":
        return {"sheet": ack.get("sheet"), "shapeId": ack.get("shapeId")}
    if app == "word":
        return {"tag": payload["tag"]}
    if app == "powerpoint":
        return {"slideId": ack.get("slideId"), "shapeId": ack.get("shapeId")}
    return {"pageId": ack.get("pageId"), "outlineId": ack.get("outlineId")}


def insert_image_handler(ctx):
    def handler(args, session_id=None, **_):
        try:
            if not isinstance(args, dict) or "bridgeSessionId" in args:
                raise BridgeError("Validation", "Arguments must be an object without session identity")
            validate(args, IMAGE_SCHEMA["parameters"])
            src = _image_source(args)
            app = _app(args)
            _refuse_if_unattended()
            key = session_key(session_id)
            if src == "html" and app in RASTER_HOSTS:
                return _json(needs_raster(app, args))
            server = get_server(ctx)
            not_paired = f"No {LABELS[app]} pane is paired with this chat. Call office_connect with app '{app}' first."
            if server.registry.binding_for(key, app) is None:
                raise BridgeError("NotPaired", not_paired)
            with server.in_use():
                _reacquire(server, key, app, not_paired)
                alt = _alt(args)
                payload, source, report = _image_payload(ctx, app, src, args, alt)
                for k in sorted(_HOST_ARGS[app]):
                    if k in args and not (app == "onenote" and k in ("width", "height")):
                        payload[k] = args[k]
                if app == "word":
                    payload.setdefault("location", "end")
                    payload["tag"] = "hermes-image-" + uuid.uuid4().hex[:16]
                return _json(_insert_and_verify(ctx, server, key, app, payload, source, report))
        except BridgeError as exc:
            return _err(exc.category, str(exc))
        except Exception:
            return _err("Protocol", "Image insert failed; details withheld")
    return handler


def _insert_and_verify(ctx, server, key, app, payload, source, report) -> dict:
    env = _env(ctx, key) if app == "onenote" else None
    try:
        if app == "onenote":
            ack = _on.mutate(env, "office_insert_image", "insert_image", payload, "image", payload.get("pageId"),
                             "inserted an image" + (f" from {source.get('url')}" if source.get("url") else ""))
            res = {"ok": True, "data": ack}
        else:
            res = server.registry.call(key, app, "insert_image", payload)
    except BridgeError as exc:  # Timeout / NotPaired / Unavailable changed nothing; ExecutionUncertain may have
        res = {"ok": False, "error": {"category": exc.category, "message": str(exc)}}
    if not res.get("ok"):
        err = res.get("error") or {}
        cat = err.get("category") or "Office"
        uncertain = cat in _on.WRITE_UNCERTAIN  # A28: PaneReload after the insert was sent may have landed too
        reload = ("OneNote stopped answering the pane: ask the user to reload the OneNote browser tab (or close and "
                  "reopen the Workbench pane), then " if cat == "PaneReload" else "")
        nxt = (f"The insert may or may not have landed; it was never retried. {reload or ''}"
               f"{'r' if reload else 'R'}ead the {LABELS[app]} document before "
               "doing anything else; if the picture is there and unwanted, undo once with office_undo app='" + app + "'."
               if uncertain else f"Nothing was changed in {LABELS[app]}. Fix the cause, then call office_insert_image once.")
        error = {"category": cat, "message": err.get("message") or "The pane refused the insert"}
        if cat in HINTS:
            error["hint"] = HINTS[cat]
        status = "uncertain" if uncertain else "failed"
        if cat == "AlreadyApplied":  # A30: a read of the page showed the earlier identical insert landed
            status, nxt = "applied", ("A read of the page after the earlier identical insert showed the picture there, "
                                      "so it was not inserted again. Its result was lost, so office_undo cannot remove "
                                      "it: if it is unwanted, the user deletes that picture by hand.")
        data = {}
        if app == "word" and cat in ("Office", "ExecutionUncertain") and payload.get("tag"):
            # A19 (seen live, Word run 3): Word for the web raised an internal error ("Cannot read properties of null
            # (reading 'message')") after the picture had landed, so "Nothing was changed" was false. The picture's tag
            # is known before the insert: read it back before saying what happened. Never retried.
            try:
                rb = server.registry.call(key, app, "image_status", {"tag": payload["tag"]})
            except BridgeError as exc:
                rb = {"ok": False, "error": {"category": exc.category, "message": str(exc)}}
            found = rb.get("data") if rb.get("ok") else None
            data["readback"] = found if found is not None else {"error": rb.get("error") or {}}
            if isinstance(found, dict) and found.get("exists"):
                status = "uncertain"
                nxt = (f"Office reported an error, but the readback found the picture in the {LABELS[app]} document "
                       f"(content control titled 'Hermes image', tag {payload['tag']}). It was not retried. The pane's "
                       "undo history does not hold it, so office_undo cannot remove it: if it is unwanted, the user "
                       "deletes that picture by hand.")
            elif isinstance(found, dict):
                nxt = f"Nothing was changed in {LABELS[app]} (the readback found no picture). Fix the cause, then call office_insert_image once."
            else:
                status = "uncertain"
                nxt = (f"Office reported an error and the readback was unavailable, so the insert may or may not have "
                       f"landed; it was never retried. Read the {LABELS[app]} document before doing anything else.")
        return {"ok": False, "error": error, "data": {"status": status, "next": nxt, **data}}
    ack = res.get("data") or {}
    target = _image_target(app, ack, payload)
    if app == "onenote":
        rb = env.call("image_status", target)
    else:
        rb = server.registry.call(key, app, "image_status", target) if all(target.values()) else \
            {"ok": False, "error": {"category": "NotFound", "message": "the pane did not report the picture's id"}}
    readback, err = (rb.get("data"), None) if rb.get("ok") else (None, rb.get("error") or {})
    status, nxt = _image_status(ack, readback, err, report, app)
    undo = _on.undo_info("outline") if app == "onenote" else ack.get("undo")
    data = {"status": status, "app": app, **{k: v for k, v in ack.items() if k not in ("undo", "created")},
            **({"tag": payload["tag"]} if app == "word" else {}),
            "source": source, "readback": readback if readback is not None else {"error": err}, "undo": undo, "next": nxt,
            "statusMeaning": _on.STATUS_MEANING}
    if report is not None:
        data["html"] = report
    return {"ok": True, "data": data}


def tool_handler(ctx, name: str, schema: dict):
    app, canonical = name.split("_", 1)

    def handler(args, session_id=None, **_):
        try:
            if not isinstance(args, dict) or "bridgeSessionId" in args:
                raise BridgeError("Validation", "Arguments must be an object without session identity")
            validate(args, schema)
            return _dispatch(ctx, session_id, app, canonical, dict(args))
        except BridgeError as exc:
            return _err(exc.category, str(exc))
        except Exception:
            return _err("Protocol", "Office execution failed; details withheld")
    return handler


def onenote_handler(ctx, name: str, schema: dict):
    """OneNote host tools and pane-backed brain tools (never in unattended runs)."""
    def handler(args, session_id=None, **_):
        try:
            if not isinstance(args, dict) or "bridgeSessionId" in args:
                raise BridgeError("Validation", "Arguments must be an object without session identity")
            validate(args, schema)
            _refuse_if_unattended()
            env = _env(ctx, session_id)
            with env.server.in_use():  # a multi-step OneNote/brain run (ledger, receipts) is never cut off
                _reacquire(env.server, env.key, "onenote",
                           "No Office pane is paired yet. Call office_connect with app 'onenote' first.")
                if name in _on.ONENOTE_TOOLS:
                    return _on.json_dumps(_with_hint(_on.host_tool(env, name, dict(args))))
                return _on.json_dumps(_with_hint(_on.BRAIN_HANDLERS[name](env, dict(args))))
        except BridgeError as exc:
            return _err(exc.category, str(exc))
        except Exception:
            return _err("Protocol", "OneNote execution failed; details withheld")
    return handler


def outlook_handler(ctx, name: str, schema: dict):
    """Outlook host tools and brain_from_email (never in unattended runs; never send, never recipients)."""
    def handler(args, session_id=None, **_):
        try:
            if not isinstance(args, dict) or "bridgeSessionId" in args:
                raise BridgeError("Validation", "Arguments must be an object without session identity")
            validate(args, schema)
            _refuse_if_unattended()
            env = _env(ctx, session_id)
            with env.server.in_use():
                _reacquire(env.server, env.key, "outlook",
                           "No Office pane is paired yet. Call office_connect with app 'outlook' first.")
                if name == "brain_from_email":
                    return _on.json_dumps(_with_hint(_ol.brain_from_email(env, dict(args))))
                return _on.json_dumps(_with_hint(_ol.host_tool(env, name, dict(args))))
        except BridgeError as exc:
            return _err(exc.category, str(exc))
        except Exception:
            return _err("Protocol", "Outlook execution failed; details withheld")
    return handler


def index_handler(ctx, name: str, schema: dict):
    """Brain tools that read only the local index: no pane, no writes, allowed in scheduled runs."""
    def handler(args, session_id=None, **_):
        try:
            validate(args if isinstance(args, dict) else None, schema)
            server = _server if _server is not None else None
            key = session_id if isinstance(session_id, str) and session_id else "-"
            env = _on.Env(ctx=ctx, server=server, key=key, brain=_on.get_brain(ctx), host_call=lambda *a: {})
            return _on.json_dumps(_on.BRAIN_HANDLERS[name](env, dict(args)))
        except BridgeError as exc:
            return _err(exc.category, str(exc))
        except Exception:
            return _err("Protocol", "Brain lookup failed; details withheld")
    return handler


# Tools whose change Office.js cannot reverse (or only sometimes): the approval says so up front.
NOT_UNDOABLE = frozenset({"excel_delete_sheet"})
MAYBE_UNDOABLE = frozenset({"excel_protect_sheet", "excel_unprotect_sheet", "powerpoint_delete_slide",
                            "powerpoint_update_slide_text", "powerpoint_add_image"})
PRODUCT = "Workbench"


def _short_text(v, n=60) -> str:
    s = " ".join(str(v).split())
    s = "".join(c for c in s if c.isprintable())
    return s if len(s) <= n else s[:n - 1] + "…"


def _doc_title(session_id, app: str):
    """The paired pane's document name for this chat, if any (read-only; never starts the bridge)."""
    if _server is None or not isinstance(session_id, str) or not session_id:
        return None
    try:
        b = _server.registry.binding_for(session_id, app)
    except Exception:
        return None
    name = (b.document or {}).get("name") if b is not None else None
    return _short_text(name) if name else None


def _target(args: dict) -> str:
    sheet, addr = args.get("sheet"), args.get("address") or args.get("dataRange") or args.get("range")
    if sheet and addr:
        return f" {_short_text(sheet, 40)}!{_short_text(addr, 30)}"
    if sheet or args.get("name"):
        return f" '{_short_text(sheet or args.get('name'), 40)}'"
    if args.get("title"):
        return f" '{_short_text(args['title'], 50)}'"
    if isinstance(args.get("slideIndex"), int):
        return f" on slide {args['slideIndex'] + 1}"
    if args.get("location"):
        return f" at {_short_text(args['location'], 20)}"
    return ""


def _operation(tool_name: str, args: dict, session_id=None) -> tuple[str, str]:
    """(app, 'operation target (undo note)') for one mutating call, from its arguments (and, for two-way links,
    the local index: the prompt names every page the run will open)."""
    if tool_name == "brain_link_both_ways":
        return "onenote", _both_ways_operation(args, session_id)
    if tool_name in _ol.TOOLS:
        return "outlook", _ol.operation(tool_name, args)
    if tool_name == "brain_from_email":
        return "outlook", _ol.operation(tool_name, args)
    if tool_name == "office_undo":
        return _app(args), "undo the last Hermes change"
    if tool_name == "office_insert_image":
        app = args.get("app") if args.get("app") in LABELS else ""
        return app, f"insert image{_image_where(app, args)} (undoable)"
    if tool_name == "brain_make_deliverables":
        hosts = args.get("hosts") or ["excel", "word", "powerpoint"]
        return "onenote", ("build " + ", ".join(f"{LABEL_OF.get(h, h)} {WHAT_OF.get(h, 'output')}" for h in hosts if isinstance(h, str))
                           + " from the decision page (undoable per step with office_undo)")
    if tool_name == "brain_decision_page":
        return "onenote", f"create decision page{_target(args)} (a page: undo cannot delete it, you can)"
    if tool_name == "brain_links":
        return "onenote", "write the Links & Backlinks block on the open page (undoable)"
    if tool_name == "brain_sync":
        return "onenote", f"open up to {int(args.get('maxPages') or 50)} pages one by one to index them (reads only)"
    if tool_name == "onenote_create_page":
        return "onenote", f"create page{_target(args)} (a page: undo cannot delete it, you can)"
    if tool_name == "onenote_append":
        return "onenote", "add a block to the open page (undoable)"
    if tool_name == "onenote_insert_table":
        rows = args.get("rows")
        size = f"{len(rows)}-row " if isinstance(rows, list) else ""
        return "onenote", f"add a {size}table to the open page (undoable)"
    app, canonical = tool_name.split("_", 1)
    note = ("not undoable" if tool_name in NOT_UNDOABLE else
            "may not be undoable" if tool_name in MAYBE_UNDOABLE else "undoable")
    verb = canonical[:-len("_range")] if canonical.endswith("_range") else canonical  # write_range -> write
    return app, f"{verb.replace('_', ' ')}{_target(args)} ({note})"


def _both_ways_operation(args: dict, session_id) -> str:
    key = session_key(session_id)
    nb, brain = None, None
    if _server is not None and _plugin_ctx is not None:
        b = _server.registry.binding_for(key, "onenote")
        nb = (b.document or {}).get("id") if b is not None else None
        brain = _on.get_brain(_plugin_ctx) if nb else None
    return _on.both_ways_approval(key, nb, brain, args)


def _image_where(app: str, args: dict) -> str:
    if app == "excel":
        sheet, cell = args.get("sheet"), args.get("cell")
        if sheet and cell:
            return f" at {_short_text(sheet, 40)}!{_short_text(cell, 12)}"
        if cell:
            return f" at {_short_text(cell, 12)}"
        return f" on sheet '{_short_text(sheet, 40)}'" if sheet else " on the active sheet"
    if app == "word":
        return {"start": " at the start", "selection": " at the cursor"}.get(args.get("location"), " at the end")
    if app == "powerpoint":
        i = args.get("slideIndex")
        return f" on slide {i + 1}" if isinstance(i, int) and not isinstance(i, bool) else " on the selected slide"
    return " on the open page" if app == "onenote" else ""


LABEL_OF = {"excel": "Excel", "word": "Word", "powerpoint": "PowerPoint", "outlook": "Outlook", "calendar": "Outlook calendar"}
WHAT_OF = {"excel": "tracker", "word": "brief", "powerpoint": "deck", "outlook": "reply draft (never sent)",
           "calendar": "follow-up meeting draft (never saved or sent)"}


def approval_message(tool_name: str, args, session_id=None) -> str:
    """'<App> · <document title, when the paired pane reported one> · <operation target> (<undo note>) · product'.

    Wording only: the rule_key stays one constant, so "always for this session" and the live-test approver
    (which keys on the rule and the product name) behave exactly as before."""
    a = args if isinstance(args, dict) else {}
    try:
        app, op = _operation(tool_name, a, session_id)
    except Exception:
        app, op = "", tool_name
    label = LABELS.get(app, "Office")
    title = _doc_title(session_id, app) if app in LABELS else None
    return " · ".join(x for x in (label, title, op, PRODUCT) if x)


def approval_gate(tool_name="", args=None, session_id=None, **_):
    """Request (never grant) human approval for document mutations via pre_tool_call.

    Hermes owns the prompt, session grants and unattended refusal (cron/gateway deny).
    One rule_key lets the user choose "always" once per session instead of per edit.
    In an unattended run the call is blocked outright, whatever the approval policy says.
    """
    if tool_name in _ol.TOOLS or tool_name == "brain_from_email":
        reason = unattended_reason()
        if reason:
            return {"action": "block", "message": f"BLOCKED: {tool_name} is refused in {reason}; Office tools need "
                                                  "a person watching the document in Hermes Desktop."}
        if not _ol.needs_approval(tool_name, args):
            return None
        return {"action": "approve", "message": approval_message(tool_name, args, session_id),
                "rule_key": APPROVAL_RULE_KEY}
    pane_tool = tool_name in _on.ONENOTE_TOOLS or tool_name in _on.BRAIN_PANE_TOOLS
    if tool_name not in DOCUMENT_TOOLS and not pane_tool \
            and tool_name not in ("office_undo", "office_connect", "office_insert_image"):
        return None  # includes brain_ask / brain_why / brain_resurface: local index reads only
    reason = unattended_reason()
    if reason:
        return {"action": "block", "message": f"BLOCKED: {tool_name} is refused in {reason}; Office tools need "
                                              "a person watching the document in Hermes Desktop."}
    if pane_tool:
        if not _on.needs_approval(tool_name, args):
            return None
        return {"action": "approve", "message": approval_message(tool_name, args, session_id),
                "rule_key": APPROVAL_RULE_KEY}
    if tool_name in READ_ONLY:
        return None
    if tool_name == "office_insert_image" and isinstance(args, dict) and "html" in args \
            and args.get("app") in RASTER_HOSTS:
        return None  # answered with NeedsRaster; nothing is changed
    return {"action": "approve", "message": approval_message(tool_name, args, session_id),
            "rule_key": APPROVAL_RULE_KEY}


def register(ctx):
    global _cron_execution, _plugin_ctx
    _plugin_ctx = ctx
    _scope.bind(ctx)  # Hermes runs register() in the owning profile's home scope; storage follows it
    _cron_execution = getattr(ctx, "current_cron_execution", None)
    for app, schemas in SCHEMAS.items():
        for schema in schemas:
            name = schema["name"]
            ctx.register_tool(name=name, toolset=f"office-desktop-{app}", schema=schema,
                              handler=tool_handler(ctx, name, schema["parameters"]),
                              check_fn=lambda: True, override=False, description=schema["description"])
    for schema, factory in ((CONNECT_SCHEMA, connect_handler), (STATUS_SCHEMA, status_handler),
                            (DISCONNECT_SCHEMA, disconnect_handler), (UNDO_SCHEMA, undo_handler),
                            (IMAGE_SCHEMA, insert_image_handler)):
        ctx.register_tool(name=schema["name"], toolset="office-desktop", schema=schema,
                          handler=factory(ctx), check_fn=lambda: True, override=False,
                          description=schema["description"])
    for schema in _on.HOST_SCHEMAS:
        ctx.register_tool(name=schema["name"], toolset="office-desktop-onenote", schema=schema,
                          handler=onenote_handler(ctx, schema["name"], schema["parameters"]),
                          check_fn=lambda: True, override=False, description=schema["description"])
    for schema in _ol.SCHEMAS:
        ctx.register_tool(name=schema["name"], toolset="office-desktop-outlook", schema=schema,
                          handler=outlook_handler(ctx, schema["name"], schema["parameters"]),
                          check_fn=lambda: True, override=False, description=schema["description"])
    ctx.register_tool(name=_ol.BRAIN_SCHEMA["name"], toolset="office-desktop-brain", schema=_ol.BRAIN_SCHEMA,
                      handler=outlook_handler(ctx, _ol.BRAIN_SCHEMA["name"], _ol.BRAIN_SCHEMA["parameters"]),
                      check_fn=lambda: True, override=False, description=_ol.BRAIN_SCHEMA["description"])
    for schema in _on.BRAIN_SCHEMAS:
        factory = index_handler if schema["name"] in _on.BRAIN_INDEX_ONLY else onenote_handler
        ctx.register_tool(name=schema["name"], toolset="office-desktop-brain", schema=schema,
                          handler=factory(ctx, schema["name"], schema["parameters"]),
                          check_fn=lambda: True, override=False, description=schema["description"])
    ctx.register_hook("pre_tool_call", approval_gate)
    if hasattr(ctx, "register_command"):
        ctx.register_command(COMMAND, workbench_command(ctx), description=(
            "Workbench: write the Office add-in manifests for the running bridge, show the one-time local "
            "certificate trust step, or status"), args_hint="[manifest|trust|status]")
    if hasattr(ctx, "register_cli_command"):
        ctx.register_cli_command(name=COMMAND, help="Workbench: add-in manifests for the running bridge, the local "
                                 "certificate trust step, status", setup_fn=_cli_setup, handler_fn=cli_handler(ctx))
    if hasattr(ctx, "register_skill"):
        for skill in ("office-desktop", "daily-note", "meeting-to-decisions"):
            skill_md = HERE / "skills" / skill / "SKILL.md"
            if skill_md.exists():
                if skill == "office-desktop":
                    ctx.register_skill(skill, skill_md, description=(
                        "Use for spreadsheet analysis, budgets and charts, Word menus and briefs, "
                        "PowerPoint decks and slides, and live Office documents in Workbench."))
                else:
                    ctx.register_skill(skill, skill_md)
    if hasattr(ctx, "register_automation_blueprint"):
        ctx.register_automation_blueprint(
            "resurface-notes", title="Resurface notes worth revisiting",
            description=("Each morning, list a few OneNote pages worth another look (decision pages first). Reads the local "
                         "Hermes index only; never opens or writes OneNote or any document."),
            schedule_template="{minute} {hour} * * 1-5",
            prompt_template=("Call brain_resurface with count {count}. Reply with the page titles as links and one line each on "
                             "why they matter, using only that tool result. Do not call any onenote_*, excel_*, word_*, "
                             "powerpoint_* or office_* tool."),
            category="productivity",
            slots=[{"name": "time", "type": "time", "label": "What time on weekdays?", "default": "08:00"},
                   {"name": "count", "type": "enum", "label": "How many notes?", "default": "3",
                    "options": ["1", "3", "5"]},
                   {"name": "deliver", "type": "enum", "label": "Where to deliver?", "default": "origin",
                    "options": ["origin", "local"], "strict": False}],
            tags=["onenote", "second-brain"])
