#!/bin/bash
# SPDX-License-Identifier: Apache-2.0
# Turn self-learning off in a DISPOSABLE test profile, so bare-prompt evidence is cold (A26).
#
#   scripts/livetest/cold_profile.sh <profile-home>      e.g. ~/.hermes/profiles/office-live-n0847e28
#
# A profile's own agent learns between runs: the post-turn background review writes skills and MEMORY.md, and the
# curator maintains skills. The keys and their cold values come from ONE place, live_driver.py SELF_LEARNING_KEYS
# (real Hermes keys: curator.enabled, auxiliary.background_review.enabled, skills.creation_nudge_interval,
# skills.write_approval, memory.memory_enabled, memory.user_profile_enabled, memory.nudge_interval). Each is written
# with `hermes config set` into that profile's config.yaml only (HERMES_HOME=<profile-home>), then read back; the
# script fails unless every key reads cold. It does not delete skills or memories the profile already learned:
# create the profile fresh (docs/harness.md, "Cold test profiles"), or move them aside yourself.
# The driver records the effective values in every result's identity (identity.selfLearning) and story-bare refuses
# a profile that is not cold unless HOL_ALLOW_WARM=1.
set -euo pipefail

PH="${1:?usage: $0 <profile-home>}"
AGENT="${HOL_HERMES_AGENT:-${HOME}/.hermes/hermes-agent}"
HERMES="${HOL_HERMES_BIN:-hermes}"
HERE="$(cd "$(dirname "$0")" && pwd)"
[ -f "$PH/config.yaml" ] || { echo "no config.yaml in $PH; not a Hermes profile home" >&2; exit 2; }
case "$(cd "$PH" && pwd -P)" in
  */.hermes|*/.hermes/) echo "refusing the root Hermes home; give a disposable profile" >&2; exit 2 ;;
esac
PY=("$AGENT/venv/bin/python")
keys() { PYTHONPATH="$AGENT" "${PY[@]}" -c "
import sys; sys.path.insert(0, '$HERE')
from live_driver import SELF_LEARNING_KEYS
for k, (_d, c) in SELF_LEARNING_KEYS.items():
    print(k, str(c).lower())"; }
while read -r key value; do
  HERMES_HOME="$PH" "$HERMES" config set "$key" "$value" >/dev/null
done < <(keys)
PYTHONPATH="$AGENT" HERMES_HOME="$PH" "${PY[@]}" -c "
import json, sys; sys.path.insert(0, '$HERE')
from pathlib import Path
from live_driver import learned_state, profile_config, self_learning
sl = self_learning(profile_config(Path('$PH')))
print(json.dumps({'profile': Path('$PH').name, 'selfLearning': sl, 'learned': learned_state(Path('$PH'))}, indent=2))
sys.exit(0 if sl['cold'] else 1)"
