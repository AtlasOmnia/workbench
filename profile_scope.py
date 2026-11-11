# SPDX-License-Identifier: Apache-2.0
"""Which Hermes profile owns this plugin instance, so its local index stays that profile's own.

One gateway process can serve many profiles. Hermes imports a directory plugin once per served profile
(its own module namespace) and runs ``register(ctx)`` inside that profile's home scope; a tool call in a
turn runs inside the turn's profile scope too. Code Hermes does not scope, such as the bridge server's
request threads (OneNote page-visit events), sees the LAUNCH profile's home through
``get_hermes_home()``. Opening the index from there put several profiles' rows into one shared file.

So ``register()`` records the owning home for its ctx, storage is always opened under that home, a tool
call must agree with it, and out-of-turn handlers use the home the pairing carried. When the owner cannot
be determined, or a call is scoped to another profile, the index is refused (never a guessed file).
"""
from __future__ import annotations

import threading
import weakref
from contextlib import contextmanager
from pathlib import Path
from typing import Any, Iterator, Optional

from .bridge_server import BridgeError

CATEGORY = "ProfileScope"
HINT = ("Hermes could not tell which profile this belongs to, so this profile's local OneNote index was not "
        "read or written. Tell the user; do not retry through another tool.")

_lock = threading.Lock()
_owners: "weakref.WeakKeyDictionary[Any, Path]" = weakref.WeakKeyDictionary()


def _key(path: Any) -> Path:
    return Path(str(path)).expanduser().resolve(strict=False)


def _ambient_home() -> Optional[Path]:
    """``get_hermes_home()`` now: the scope Hermes bound (load, turn), else the process home."""
    try:
        from hermes_constants import get_hermes_home  # public plugin surface
        return _key(get_hermes_home())
    except Exception:  # noqa: BLE001 - no Hermes on sys.path, or no home: owner stays unknown
        return None


def _bound_override() -> Optional[Path]:
    """The home Hermes bound for this context (turn, callback), or None when nothing is bound."""
    try:
        from hermes_constants import get_hermes_home_override
    except Exception:  # noqa: BLE001
        return None
    value = get_hermes_home_override()
    return _key(value) if value else None


def bind(ctx) -> Optional[Path]:
    """Called from ``register(ctx)``: Hermes loads each plugin inside its owning profile's home scope."""
    home = _ambient_home()
    if home is not None:
        with _lock:
            _owners[ctx] = home
    return home


def owner(ctx) -> Path:
    with _lock:
        home = _owners.get(ctx) if ctx is not None else None
    if home is None:
        raise BridgeError(CATEGORY, "The Hermes profile that loaded this plugin is unknown; the local index "
                                    "was not read or written.")
    return home


def call_home(ctx) -> Path:
    """For tool calls and hooks: the owner, which the call's own bound profile (when any) must match."""
    home = owner(ctx)
    bound = _bound_override()
    if bound is not None and bound != home:
        raise BridgeError(CATEGORY, "This call is scoped to a different Hermes profile than the one that "
                                    "loaded this plugin; the local index was not read or written.")
    return home


def carried_home(ctx, carried: Any) -> Path:
    """For out-of-turn handlers (bridge threads): the home the pairing carried; it must be the owner."""
    home = owner(ctx)
    if carried is None or _key(carried) != home:
        raise BridgeError(CATEGORY, "This pane's pairing does not name this plugin's Hermes profile; the local "
                                    "index was not read or written.")
    return home


def try_call_home(ctx) -> Optional[str]:
    """The home to stamp on a new pairing, or None (events from that pane are then refused)."""
    try:
        return str(call_home(ctx))
    except BridgeError:
        return None


@contextmanager
def scoped(home: Path) -> Iterator[None]:
    """Bind ``home`` as the Hermes home for this context (public ``hermes_constants`` API)."""
    from hermes_constants import reset_hermes_home_override, set_hermes_home_override
    token = set_hermes_home_override(str(home))
    try:
        yield
    finally:
        reset_hermes_home_override(token)
