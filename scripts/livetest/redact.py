# SPDX-License-Identifier: Apache-2.0
"""Harness redaction for everything written to results.json, events.jsonl and the window's chat lines.

Addresses become <address>; people's display names become <name>. A name is known from any person field
(from/to/cc/bcc/sender/organizer/attendees ...) the harness has seen in this process, including tool results as
they complete (chat.py learns them before the agent's reply is logged), and is then masked in free text too. A
"Name <address>" pair after a mail header label (From:/To:/Cc: ..., or after ';' in a recipient list) is masked even
when the name was never seen. Only whole display names are masked, never single words (a fixture's 'Sam' stays)."""
from __future__ import annotations

import re

EMAIL_RE = re.compile(r"[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}")
PERSON_KEYS = frozenset({"from", "to", "cc", "bcc", "sender", "replyTo", "reply_to", "organizer", "attendees",
                         "requiredAttendees", "optionalAttendees", "required", "optional", "recipients"})
NAME_KEYS = frozenset({"name", "displayName", "display_name"})
HEADER_RE = re.compile(r"(?i)(\b(?:From|To|Cc|Bcc|Sender|Organizer|Reply-To|Attendees|Required|Optional)[ \t]*:[ \t]*"
                       r"|;[ \t]*)([^<>\r\n;:]+?)[ \t]*<([^<>\s@]+@[^<>\s]+)>")
PLACEHOLDERS = ("<name>", "<address>")

_names: set[str] = set()
_name_re = None


def reset_names():
    global _name_re
    _names.clear()
    _name_re = None


def known_names() -> set[str]:
    return set(_names)


def _remember(s):
    global _name_re
    if not isinstance(s, str):
        return
    s = EMAIL_RE.sub("", s)
    for p in PLACEHOLDERS:
        s = s.replace(p, "")
    s = s.strip(" \t\r\n<>\"',;()")
    if len(s) < 2 or not re.search(r"[^\W\d_]", s) or s in _names:
        return
    _names.add(s)
    alts = "|".join(re.escape(n) for n in sorted(_names, key=len, reverse=True))
    _name_re = re.compile(rf"(?<!\w)(?:{alts})(?!\w)", re.I)


def _learn_person(x):
    if isinstance(x, list):
        for y in x:
            _learn_person(y)
    elif isinstance(x, dict):
        for k, y in x.items():
            if k in NAME_KEYS:
                _remember(y)
            elif isinstance(y, (dict, list)):
                _learn_person(y)
    elif isinstance(x, str) and EMAIL_RE.search(x):  # 'Name <address>'; a bare string may be a time or a cell
        _remember(x)


def learn_names(v):
    """Register the display names in every person field of a value (a tool result, a readback). Never raises."""
    try:
        if isinstance(v, list):
            for x in v:
                learn_names(x)
        elif isinstance(v, dict):
            for k, x in v.items():
                if k in PERSON_KEYS:
                    _learn_person(x)
                elif isinstance(x, (dict, list)):
                    learn_names(x)
    except RecursionError:
        pass
    return v


def mask_text(s: str) -> str:
    s = HEADER_RE.sub(lambda m: f"{m.group(1)}<name> <{m.group(3)}>", s)
    s = EMAIL_RE.sub("<address>", s)
    if _name_re is not None:
        s = _name_re.sub("<name>", s)
    return s


def _mask_person(x):
    if isinstance(x, list):
        return [_mask_person(y) for y in x]
    if isinstance(x, dict):
        return {k: ("<name>" if k in NAME_KEYS and isinstance(y, str) and y else
                    _mask_person(y) if isinstance(y, (dict, list)) else _mask(y)) for k, y in x.items()}
    if isinstance(x, str):
        if not EMAIL_RE.search(x):
            return mask_text(x)
        rest = EMAIL_RE.sub("", x).strip(" \t\r\n<>\"',;()")
        return "<name> <<address>>" if rest else "<address>"
    return x


def _mask(v):
    if isinstance(v, str):
        return mask_text(v)
    if isinstance(v, list):
        return [_mask(x) for x in v]
    if isinstance(v, dict):
        return {k: (_mask_person(x) if k in PERSON_KEYS else _mask(x)) for k, x in v.items()}
    return v


def mask_people(v):
    """A copy of v with every address and every person's display name masked (learns the names in v first)."""
    learn_names(v)
    return _mask(v)
