#!/bin/bash
# SPDX-License-Identifier: Apache-2.0
# One command for the automated in-pane live test on the client Mac (client only).
#
#   scripts/livetest/run_macbook.sh [excel|word|powerpoint|onenote|outlook|calendar|all|f3|loop|story|story-bare|<host>,<host>] [runs]
#
#   outlook: its own harness tab and cookie jar (persist:hol-outlook), its own account; see live_driver.py outlook_run_once.
#   calendar: the same Outlook tab and account, on one harness-created event, deleted at the end even when a run fails; see
#             live_driver.py calendar_host / calendar_run_once.
#
#   story: the recorded demo story (docs/demo-script.md) in ONE chat with FIVE tabs: Outlook (persist:hol-outlook, on the
#          existing self-sent fixture email, never sending a new one) plus OneNote, Excel, Word, PowerPoint
#          (persist:hol-livetest); see live_driver.py story_once.
#
#   story-bare: the same story with the prompts EXACTLY as docs/demo-script.md lists them (no bracketed note naming
#          tools, titles, ids, paths or times), judged only on outcomes: what the documents and the index hold
#          afterwards and the receipts, never which tool was called; see live_driver.py STORY_BARE.
#
#   loop: the whole docs/demo.md loop in ONE chat with four Office tabs (OneNote, Excel, Word, PowerPoint) in the
#         harness window, all paired to that chat; see live_driver.py loop_once.
#
# Runs on the Studio (where the gateway and the Office bridge live). It:
#   1. preflights the client Mac over ssh: the bridge tunnel (127.0.0.1:47615-47619 there; the profile's bridge may be on any of the five) and the served add-in. A fresh
#      gateway has no bridge until the profile's first office_connect, so an unreachable bridge gets ONE warm chat
#      turn (scripts/livetest/warm.py: office_status, then office_connect only if the bridge is not listening)
#      before the preflight fails;
#   2. syncs the Electron 44.9.3 harness to ~/hermes-office-livetest/app and the manifests from the
#      served site dir to ~/hermes-office-livetest/manifests (nothing else on the client Mac is touched);
#   3. starts the harness window if it is not running (persistent profile under ~/hermes-office-livetest);
#   4. forwards its loopback control port over ssh -L, and runs scripts/livetest/live_driver.py, which
#      drives a real chat in the hermes-live-demo profile through the running gateway (never starts one).
# The first run stops at the one-time sign-in gate in the harness window until you sign in (30 min max).
# Results: $HOL_OUT (default ~/.hermes/profiles/omh/cache/scratch/hol-livetest/<timestamp>/).
set -euo pipefail

TARGET="${1:-all}"
RUNS="${2:-3}"
MAC="${HOL_MAC:-office-client}"
SITE="${HOL_SITE_DIR:-${HOME}/.hermes/cache/scratch/office-live-site/office/live}"
PROFILE="${HOL_PROFILE:-hermes-live-demo}"   # e.g. HOL_PROFILE=hermes-live-demo2
PROFILE_HOME="${HOL_PROFILE_HOME:-${HOME}/.hermes/profiles/$PROFILE}"
AGENT="${HOL_HERMES_AGENT:-${HOME}/.hermes/hermes-agent}"
ORIGIN="${HOL_ORIGIN:-https://office-addin.example.test}"
LOCAL_PORT="${HOL_LOCAL_CONTROL_PORT:-19334}"
OUT="${HOL_OUT:-${HOME}/.hermes/profiles/omh/cache/scratch/hol-livetest/$(date +%Y%m%d-%H%M%S)}"
HERE="$(cd "$(dirname "$0")" && pwd)"
SSH=(ssh -o BatchMode=yes -o ConnectTimeout=10 "$MAC")

LOOP=()
STORY=()
case "$TARGET" in
  all) HOSTS="excel,word,powerpoint,onenote"; F3=(--f3) ;;
  excel|word|powerpoint|onenote|outlook|calendar) HOSTS="$TARGET"; F3=() ;;
  *,*) HOSTS="$TARGET"; F3=() ;;  # e.g. word,onenote
  f3) HOSTS=""; F3=(--f3) ;;
  loop) HOSTS=""; F3=(); LOOP=(--loop) ;;
  story) HOSTS=""; F3=(); STORY=(--story) ;;
  story-bare) HOSTS=""; F3=(); STORY=(--story --bare) ;;  # the presenter's exact prompts, outcome-only checks
  *) echo "usage: $0 [excel|word|powerpoint|onenote|outlook|calendar|all|f3|loop|story|story-bare] [runs]" >&2; exit 64 ;;
esac

echo "== preflight"
[ -x "$AGENT/venv/bin/python" ] || { echo "hermes venv not found at $AGENT" >&2; exit 2; }
# Self-warming: a fresh gateway starts the bridge on the profile's first office_connect (warm.py explains).
PYTHONPATH="$AGENT" HERMES_HOME="$PROFILE_HOME" "$AGENT/venv/bin/python" "$HERE/warm.py" --mac "$MAC" --origin "$ORIGIN" \
  || { echo "bridge not reachable from the client Mac (127.0.0.1:47615-47619 tunnel), even after one warm chat turn; stopping" >&2; exit 2; }
"${SSH[@]}" "curl -s -m10 -o /dev/null -w '%{http_code}' $ORIGIN/office/live/manifest-excel.xml" | grep -q 200 \
  || { echo "add-in site not reachable from the client Mac; stopping" >&2; exit 2; }

echo "== sync harness and manifests"
"${SSH[@]}" 'mkdir -p ~/hermes-office-livetest/app ~/hermes-office-livetest/manifests ~/hermes-office-livetest/logs'
scp -q -o BatchMode=yes "$HERE/macbook/main.js" "$HERE/macbook/tabscript.js" "$HERE/macbook/package.json" "$MAC:hermes-office-livetest/app/"
scp -q -o BatchMode=yes "$SITE"/manifest-*.xml "$MAC:hermes-office-livetest/manifests/"
"${SSH[@]}" 'cd ~/hermes-office-livetest/app && if [ ! -x node_modules/electron/dist/Electron.app/Contents/MacOS/Electron ]; then
    npm_config_cache=$HOME/hermes-office-livetest/npm-cache npm install --no-audit --no-fund >/dev/null &&
    electron_config_cache=$HOME/hermes-office-livetest/electron-cache node node_modules/electron/install.js; fi;
  cat node_modules/electron/dist/version'

echo "== harness window"
# The window reports a sha256 over the harness files it started from (tabscript.js harnessHash: main.js then
# tabscript.js); any difference from the files just synced, an old window without a hash, or a window that does
# not answer `version` at all means it runs other code: quit it, and if it does not go, stop our Electron (only
# the one under ~/hermes-office-livetest/app; nothing else on the client Mac).
WANT="$(cat "$HERE/macbook/main.js" "$HERE/macbook/tabscript.js" | shasum -a 256 | cut -c1-12)"
HAVE="$("${SSH[@]}" 'curl -s -m3 -X POST -d "{\"cmd\":\"version\"}" http://127.0.0.1:9334/' 2>/dev/null || true)"
RUNNING="$("${SSH[@]}" 'pgrep -f "[h]ermes-office-livetest/app/node_modules/electron/dist/Electron.app/Contents/MacOS/Electron" || true')"
if [ -n "$RUNNING" ] && ! echo "$HAVE" | grep -q "\"hash\":\"$WANT\""; then
  echo "harness window runs other harness code (want $WANT); restarting it (it is ours)"
  "${SSH[@]}" 'curl -s -m3 -X POST -d "{\"cmd\":\"quit\"}" http://127.0.0.1:9334/' >/dev/null 2>&1 || true
  for _ in $(seq 1 10); do
    "${SSH[@]}" 'pgrep -f "[h]ermes-office-livetest/app/node_modules/electron/dist/Electron.app/Contents/MacOS/Electron" >/dev/null' || break
    sleep 1
  done
  "${SSH[@]}" 'pkill -f "[h]ermes-office-livetest/app/node_modules/electron/dist/Electron.app/Contents/MacOS/Electron"; sleep 2' || true
fi
if ! "${SSH[@]}" 'curl -s -m3 -X POST -d "{\"cmd\":\"state\"}" http://127.0.0.1:9334/' >/dev/null 2>&1; then
  "${SSH[@]}" 'open -n --stdout "$HOME/hermes-office-livetest/logs/electron.out" --stderr "$HOME/hermes-office-livetest/logs/electron.err" -a "$HOME/hermes-office-livetest/app/node_modules/electron/dist/Electron.app" --args "$HOME/hermes-office-livetest/app"'
  for _ in $(seq 1 30); do
    "${SSH[@]}" 'curl -s -m3 -X POST -d "{\"cmd\":\"state\"}" http://127.0.0.1:9334/' >/dev/null 2>&1 && break
    sleep 2
  done
fi

ssh -o BatchMode=yes -o ExitOnForwardFailure=yes -N -L "127.0.0.1:$LOCAL_PORT:127.0.0.1:9334" "$MAC" &
FWD=$!
trap 'kill $FWD 2>/dev/null || true' EXIT
sleep 2

echo "== driver -> $OUT"
mkdir -p "$OUT"
DRIVER_ARGS=(--control "http://127.0.0.1:$LOCAL_PORT" --runs "$RUNS" --out "$OUT" --signin-timeout "${HOL_SIGNIN_TIMEOUT:-1800}" ${F3[@]+"${F3[@]}"} ${LOOP[@]+"${LOOP[@]}"} ${STORY[@]+"${STORY[@]}"})
if [ -n "${HOL_IMAGES:-}" ]; then DRIVER_ARGS+=(--images); fi   # HOL_IMAGES=1: office_insert_image steps
[ -n "$HOSTS" ] && DRIVER_ARGS+=(--hosts "$HOSTS") || DRIVER_ARGS+=(--hosts "")
PYTHONPATH="$AGENT" HERMES_HOME="$PROFILE_HOME" "$AGENT/venv/bin/python" "$HERE/live_driver.py" "${DRIVER_ARGS[@]}"
