# SPDX-License-Identifier: Apache-2.0
"""Drive a Hermes chat the way Hermes Desktop does: the gateway's native session socket.

Runs with the Hermes venv python and HERMES_HOME set to the target profile (hermes-live-demo).
Discovery only: it attaches to the ready gateway that already serves the profile and never
starts, restarts or reconfigures one. Tool calls therefore execute inside that gateway, where
the Office bridge lives.

Approvals are answered through the gateway's own approval protocol (``approval.respond``)
with choice ``once`` for the Office mutation rule only; anything else is denied. Nothing
here enables yolo or changes approval policy.
"""
from __future__ import annotations

import asyncio
import json
import re
import sys
import time
import uuid
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
from redact import learn_names  # noqa: E402

# Approval prompts from this plugin end with its product name (plugin PRODUCT, "Workbench"); tests pin the two.
OFFICE_RULE_HINTS = ("Workbench", "Office", "OneNote", "office", "onenote")
CODE_RE = re.compile(r"\b([2-9A-HJ-NP-TV-Z]{4})-([2-9A-HJ-NP-TV-Z]{4})\b")


_SEEN: set = set()


class ChatError(RuntimeError):
    pass


def live_endpoint():
    from hermes_cli.gateway_runtime import discover_gateway_endpoint
    from hermes_constants import get_hermes_home
    found = discover_gateway_endpoint(get_hermes_home().resolve())
    if found.state != "ready":
        raise ChatError(f"gateway for this profile is not ready ({found.state}); not starting one")
    return found.endpoint


def connect_live():
    """Async context manager: the gateway client connected to the ready gateway of this profile."""
    from hermes_cli.gateway_client import connect_gateway
    return connect_gateway(live_endpoint())


def _parse(result):
    if isinstance(result, str):
        try:
            return json.loads(result)
        except ValueError:
            return result
    return result


class Chat:
    """One chat session. ``turn(text)`` returns the reply text and every tool call it made."""

    def __init__(self, client, snapshot, log):
        self.client = client
        self.sid = snapshot["stored_session_id"]
        self.generation = snapshot.get("execution_generation")
        self.log = log
        self.approvals = []
        self.deny_rules = []  # substrings of an Office approval text to DENY (the loop's one denial beat)

    @classmethod
    async def create(cls, client, title, log, source="cli"):
        contract = (await client.rpc("runtime.describe")).get("session_create", {})
        if source not in contract.get("sources", []):
            raise ChatError(f"gateway has no session source {source!r}")
        params = {"title": title} if "title" in contract.get("parameters", []) else {}
        snap = await client.rpc("session.create", request_id=uuid.uuid4().hex, source=source, **params)
        chat = cls(client, snap, log)
        # The gateway continues an existing session that has the same title; every run must be a new chat.
        if chat.sid in _SEEN:
            raise ChatError("session.create returned a session this harness already used")
        _SEEN.add(chat.sid)
        log({"ev": "chat", "session": chat.sid[:18]})
        return chat

    async def turn(self, text, timeout=600):
        tools, deltas, open_tools = [], [], {}
        t0 = time.monotonic()
        receipt = await self.client.rpc("prompt.submit", session_id=self.sid, input_id=uuid.uuid4().hex, text=text)
        admission = (receipt or {}).get("admission_id")
        self.log({"ev": "user", "text": text, "admission": admission})
        deadline = t0 + timeout
        while True:
            remaining = deadline - time.monotonic()
            if remaining <= 0:
                raise ChatError(f"turn timed out after {timeout}s")
            try:
                event = await asyncio.wait_for(self.client.events.get(), remaining)
            except asyncio.TimeoutError:
                raise ChatError(f"turn timed out after {timeout}s") from None
            if isinstance(event, Exception):
                raise ChatError(f"gateway event stream failed: {type(event).__name__}")
            p = event.get("params", {})
            if p.get("session_id") != self.sid:
                continue
            kind, payload = p.get("type"), p.get("payload", {}) or {}
            self.generation = p.get("execution_generation", self.generation)
            own = p.get("admission_id") or payload.get("admission_id")
            if own and admission and own != admission and kind not in ("approval.request", "clarify.request"):
                continue
            now = time.monotonic()
            if kind == "tool.start":
                key = payload.get("tool_call_id") or f"{payload.get('tool_name')}#{len(tools)}"
                open_tools[key] = {"name": payload.get("tool_name"), "args": payload.get("args"), "t0": now}
            elif kind == "tool.complete":
                key = payload.get("tool_call_id") or next((k for k, v in open_tools.items() if v["name"] == payload.get("tool_name")), None)
                rec = open_tools.pop(key, None) or {"name": payload.get("tool_name"), "args": payload.get("args"), "t0": now}
                rec.update(result=_parse(payload.get("result")), is_error=bool(payload.get("is_error")),
                           harnessMs=int(round((now - rec.pop("t0")) * 1000)))
                tools.append(rec)
                learn_names(rec["result"])  # names in this result are masked in the reply logged next (A20)
                self.log({"ev": "tool", "name": rec["name"], "harnessMs": rec["harnessMs"],
                          "ok": isinstance(rec["result"], dict) and rec["result"].get("ok")})
            elif kind == "approval.request":
                await self._answer(payload)
            elif kind == "clarify.request":
                await self.client.rpc("clarify.respond", session_id=self.sid,
                                      execution_generation=payload.get("execution_generation", self.generation),
                                      prompt_id=payload["prompt_id"], answer="Proceed as described in my message.")
            elif kind == "message.delta":
                deltas.append(payload.get("text") or payload.get("delta") or "")
            elif kind == "message.complete":
                reply = payload.get("text") or payload.get("content") or "".join(deltas)
                self.log({"ev": "agent", "text": reply[:2000], "ms": int((now - t0) * 1000)})
                return {"text": reply, "tools": tools, "ms": int((now - t0) * 1000),
                        "status": payload.get("status") or payload.get("outcome")}

    async def _answer(self, prompt):
        """Approve exactly the plugin's Office edit prompts, once each; deny anything else (and, for the
        denial beat, any Office prompt matching one of ``deny_rules``)."""
        choice, rule = approval_choice(prompt, self.deny_rules)
        await self.client.rpc("approval.respond", session_id=self.sid,
                              execution_generation=prompt.get("execution_generation", self.generation),
                              prompt_id=prompt["prompt_id"], choice=choice)
        rec = {"choice": choice, "description": str(prompt.get("description") or "")[:200]}
        if rule:
            rec["deniedByRule"] = rule
        self.approvals.append(rec)
        self.log({"ev": "approval", **rec})


def approval_choice(prompt, deny_rules=()):
    """(choice, matched deny rule or None). Office prompts get 'once' unless a deny rule matches; anything
    that is not the plugin's Office prompt is denied. Never 'always'/'session': no grant outlives one call."""
    desc = " ".join(str(prompt.get(k) or "") for k in ("description", "command", "question", "rule_key"))
    office = "hermes-office-desktop" in desc or any(h in desc for h in OFFICE_RULE_HINTS)
    rule = next((r for r in deny_rules if r and r in desc), None) if office else None
    choice = "once" if office and not rule else "deny"
    choices = prompt.get("choices") or []
    if choice not in choices and choices:
        choice = "deny" if "deny" in choices else choices[-1]
    return choice, rule


def pairing_code(reply_text, tools):
    """The code as the chat SHOWED it (reply text); the tool result is only a cross-check."""
    from_tool = None
    for t in tools:
        r = t.get("result")
        if t.get("name") == "office_connect" and isinstance(r, dict) and r.get("ok"):
            from_tool = r["data"]["pairingCode"]
    if from_tool is None:
        return None, "none (office_connect was not called)"
    codes = [f"{a}-{b}" for a, b in CODE_RE.findall(reply_text or "")]
    if from_tool in codes:
        return from_tool, "reply"
    return from_tool, "tool-result (reply did not show it)"
