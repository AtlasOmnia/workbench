# SPDX-License-Identifier: Apache-2.0
"""Run a tool-coverage plan (plan.py) against a paired pane through the real plugin tool handlers.

The environment supplies three callbacks:
  probe()        -> str   a deterministic snapshot of the whole document (mock fingerprint or real Office.js)
  reload_pane()  -> None  reload the task pane page (it must resume the same chat binding)
  run_js(step)   -> Any   run a setup script in the pane (e.g. select text)
"""
from __future__ import annotations

import json
import statistics
import time


class Runner:
    def __init__(self, tools: dict, session: str, app: str, probe, reload_pane, run_js, log=print):
        self.tools, self.session, self.app = tools, session, app
        self.probe, self.reload_pane, self.run_js, self.log = probe, reload_pane, run_js, log
        self.state: dict = {}
        self.results: list[dict] = []

    def call(self, name, args, session=None):
        t0 = time.perf_counter()
        raw = self.tools[name]["handler"](args, session_id=session or self.session)
        ms = round((time.perf_counter() - t0) * 1000, 1)
        return json.loads(raw), ms

    def _args(self, step):
        a = step.get("args", {})
        return a(self.state) if callable(a) else dict(a)

    def step(self, step) -> dict:
        special = step.get("special")
        if special == "reload":
            t0 = time.perf_counter()
            self.reload_pane()
            return {"tool": "(pane reload + resume)", "mode": "special", "pass": True, "ms": round((time.perf_counter() - t0) * 1000, 1)}
        if special == "js":
            return {"tool": f"(setup: {step['label']})", "mode": "special", "pass": bool(self.run_js(step)), "ms": None}
        if special == "other_chat":
            res, ms = self.call(step["tool"], self._args(step), session="some-other-chat")
            cat = (res.get("error") or {}).get("category")
            return {"tool": step["tool"], "mode": "other_chat", "pass": cat == "NotPaired", "ms": ms, "error": res.get("error"),
                    "label": "another chat is refused"}
        tool, mode = step["tool"], step["mode"]
        args = self._args(step)
        out = {"tool": tool, "mode": mode, "label": step.get("label")}
        before = self.probe() if mode == "undo" else None
        res, ms = self.call(tool, args)
        out["ms"] = ms
        out["ok"] = res.get("ok")
        if not res.get("ok"):
            out["error"] = res.get("error")
        data = res.get("data") or {}
        if mode == "error":
            out["pass"] = res.get("ok") is False and (res.get("error") or {}).get("category") == step["expect"]
            return out
        passed = res.get("ok") is True
        if passed and "check" in step:
            try:
                passed = bool(step["check"](data, self.state))
            except Exception as exc:  # noqa: BLE001
                passed, out["checkError"] = False, repr(exc)
        if passed and "save" in step:
            step["save"](data, self.state)
        undo_info = data.get("undo") if isinstance(data, dict) else None
        if undo_info is not None:
            out["undoAvailable"] = undo_info.get("available")
            if undo_info.get("partial"):
                out["undoPartial"] = undo_info["partial"]
        if mode == "noundo":
            passed = passed and undo_info is not None and undo_info.get("available") is False
        if mode == "undo":
            mid = self.probe()
            out["changed"] = mid != before
            u, ums = self.call("office_undo", {"app": self.app})
            out["undoMs"] = ums
            after = self.probe()
            out["restored"] = after == before
            if not u.get("ok"):
                out["undoError"] = u.get("error")
            if not out["restored"]:
                out["diff"] = _diff(before, after)
            passed = passed and bool(undo_info and undo_info.get("available")) and out["changed"] and u.get("ok") is True and out["restored"]
        out["pass"] = passed
        if mode in ("read",) and isinstance(data, dict):
            out["sample"] = json.dumps(data, ensure_ascii=False)[:300]
        return out

    def run(self, plan) -> list[dict]:
        for st in plan:
            r = self.step(st)
            self.results.append(r)
            self.log(f"{'PASS' if r['pass'] else 'FAIL'} {r['mode']:9} {r['tool']:40} {r.get('ms')}ms"
                     + (f" undo={r.get('undoMs')}ms" if r.get('undoMs') is not None else "")
                     + ("" if r["pass"] else f"  {json.dumps({k: r.get(k) for k in ('error', 'undoError', 'changed', 'restored', 'diff', 'checkError') if r.get(k) is not None})[:600]}"))
        return self.results


def _diff(a: str, b: str) -> str:
    if a is None or b is None:
        return "probe missing"
    i = next((k for k in range(min(len(a), len(b))) if a[k] != b[k]), min(len(a), len(b)))
    return f"at {i}: before={a[max(0, i - 80):i + 120]!r} after={b[max(0, i - 80):i + 120]!r}"


def summary(results: list[dict]) -> dict:
    tool_ms = [r["ms"] for r in results if r.get("ms") is not None and not r["tool"].startswith("(")]
    undo_ms = [r["undoMs"] for r in results if r.get("undoMs") is not None]

    def stats(v):
        return {"n": len(v), "median": statistics.median(v), "p90": sorted(v)[int(0.9 * (len(v) - 1))], "max": max(v)} if v else {}
    return {"steps": len(results), "passed": sum(1 for r in results if r["pass"]), "failed": [r["tool"] for r in results if not r["pass"]],
            "toolMs": stats(tool_ms), "undoMs": stats(undo_ms)}
