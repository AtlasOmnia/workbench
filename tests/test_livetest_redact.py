# SPDX-License-Identifier: Apache-2.0
"""Harness redaction (A20): people's display names are masked like their addresses, in results.json, events.jsonl
and the on-screen chat lines. Seen live on 40f1b3e (story gw2/gw3): the Outlook sender's display name was written
unmasked in a tool result's from.name, in the agent's reply text ("the email ... from <name>") and in a reply
draft's quoted header ("From: <name> <<address>>"); only the address was masked. Pure logic, no Office."""
import json
from pathlib import Path

from test_livetest_story import LIVE, _driver, _load

NAME = "Pat Example"
OTHER = "Robin Sample"


def _redact():
    r = _load("redact", "redact.py")
    r.reset_names()
    return r


def test_every_person_field_has_its_name_masked():
    r = _redact()
    d = _driver()
    v = {"subject": "HOL live test: pilot sync",
         "from": {"name": NAME, "email": "pat@example.com"},
         "sender": {"displayName": NAME, "emailAddress": "pat@example.com"},
         "to": [{"name": OTHER, "email": "robin@example.org"}, {"name": "robin@example.org", "email": "robin@example.org"}],
         "cc": [{"name": OTHER}], "bcc": [OTHER + " <robin@example.org>"],
         "organizer": {"name": NAME, "address": "pat@example.com"},
         "attendees": [{"name": OTHER, "email": "robin@example.org", "response": "accepted"}],
         "requiredAttendees": [{"displayName": OTHER}], "optionalAttendees": [], "count": 2}
    out = d.mask_addresses(v)
    blob = json.dumps(out)
    assert NAME not in blob and OTHER not in blob and "@" not in blob
    assert out["from"] == {"name": "<name>", "email": "<address>"}
    assert out["sender"] == {"displayName": "<name>", "emailAddress": "<address>"}
    assert out["to"][0] == {"name": "<name>", "email": "<address>"}
    assert out["to"][1]["name"] == "<name>"  # a name field holding the address is still a name field
    assert out["cc"] == [{"name": "<name>"}] and out["bcc"] == ["<name> <<address>>"]
    assert out["organizer"] == {"name": "<name>", "address": "<address>"}
    assert out["attendees"][0]["response"] == "accepted" and out["attendees"][0]["name"] == "<name>"
    assert out["requiredAttendees"] == [{"displayName": "<name>"}]
    assert out["subject"] == "HOL live test: pilot sync" and out["count"] == 2
    assert r.known_names() >= {NAME, OTHER}


def test_a_name_seen_in_a_person_field_is_masked_in_free_text_later():
    _redact()
    d = _driver()
    d.mask_addresses({"from": {"name": NAME, "email": "pat@example.com"}})
    text = f'Source: the email "HOL live test: pilot sync" from {NAME}, 7 Nov 2026.'
    assert d.mask_addresses(text) == 'Source: the email "HOL live test: pilot sync" from <name>, 7 Nov 2026.'
    assert NAME not in d.ui_redact(text)  # chat lines on screen and user/agent events


def test_quoted_mail_headers_have_the_name_masked_even_if_never_seen():
    _redact()
    d = _driver()
    body = f"From: {OTHER} <robin@example.org>Sent: Saturday, November 7, 2026\nTo: {NAME} <pat@example.com>"
    out = d.mask_addresses(body)
    assert OTHER not in out and NAME not in out
    assert out.startswith("From: <name> <<address>>Sent:") and "To: <name> <<address>>" in out


def test_recorder_masks_names_in_events_and_results(tmp_path):
    _redact()
    d = _driver()
    rec = d.Recorder(tmp_path, identity={"runId": "t"})
    rec.add({"steps": [{"results": [{"data": {"from": {"name": NAME, "email": "pat@example.com"}}}]}]})
    rec.save()
    rec.log({"ev": "agent", "text": f"I indexed the email from {NAME} (pat@example.com)."})
    rec.log({"ev": "readback", "data": {"to": [{"name": OTHER, "email": "robin@example.org"}]}})
    rec.events.close()
    events = (tmp_path / "events.jsonl").read_text()
    results = (tmp_path / "results.json").read_text()
    for blob in (events, results):
        assert NAME not in blob and OTHER not in blob and "@example" not in blob
    assert "from <name> (<address>)" in events and '"name": "<name>"' in results


def test_chat_learns_names_from_tool_results_before_the_reply_is_logged():
    """The agent's reply event is logged inside Chat.turn, before any results pass through mask_addresses: the chat
    registers the names in each tool result as it completes, so the reply that follows is masked too."""
    src = (LIVE / "chat.py").read_text()
    branch = src[src.index('elif kind == "tool.complete":'):src.index('elif kind == "approval.request":')]
    assert "learn_names(rec[\"result\"])" in branch
    assert branch.index("learn_names(") < branch.index('self.log({"ev": "tool"')
