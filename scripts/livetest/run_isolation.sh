#!/bin/bash
# SPDX-License-Identifier: Apache-2.0
# Live two-profile isolation check (scripts/livetest/isolation.py) against the client Mac harness window.
#
#   HOL_PROFILE_A=<profile> HOL_PROFILE_B=<profile> HOL_OUT=<dir> scripts/livetest/run_isolation.sh
#
# Both profiles must already be served by the running gateway with this plugin enabled (never started or
# restarted here). The harness window must already be running and signed in (run_macbook.sh starts it).
set -euo pipefail
A="${HOL_PROFILE_A:?set HOL_PROFILE_A}"; B="${HOL_PROFILE_B:?set HOL_PROFILE_B}"
MAC="${HOL_MAC:-office-client}"
AGENT="${HOL_HERMES_AGENT:-${HOME}/.hermes/hermes-agent}"
LOCAL_PORT="${HOL_LOCAL_CONTROL_PORT:-19334}"
OUT="${HOL_OUT:?set HOL_OUT}"
HERE="$(cd "$(dirname "$0")" && pwd)"
PY=("$AGENT/venv/bin/python" "$HERE/isolation.py")
[ "$A" != "$B" ] || { echo "A and B must be different profiles" >&2; exit 64; }
ssh -o BatchMode=yes -o ConnectTimeout=10 "$MAC" 'curl -s -m3 -X POST -d "{\"cmd\":\"state\"}" http://127.0.0.1:9334/' >/dev/null \
  || { echo "harness window not running on the client Mac; run run_macbook.sh first" >&2; exit 2; }
ssh -o BatchMode=yes -o ExitOnForwardFailure=yes -N -L "127.0.0.1:$LOCAL_PORT:127.0.0.1:9334" "$MAC" &
FWD=$!
trap 'kill $FWD 2>/dev/null || true' EXIT
sleep 2
mkdir -p "$OUT"
run() { local prof=$1; shift; PYTHONPATH="$AGENT" HERMES_HOME="${HOME}/.hermes/profiles/$prof" "${PY[@]}" "$@" \
  --out "$OUT" --control "http://127.0.0.1:$LOCAL_PORT"; }
run "$A" root-counts
rc=0
run "$A" write || rc=1
[ $rc -eq 0 ] && { run "$B" probe || rc=1; }
run "$A" cleanup || rc=1   # always try to remove A's block again
run "$A" verify || rc=1
exit $rc
