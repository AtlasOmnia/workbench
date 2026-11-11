# SPDX-License-Identifier: Apache-2.0
"""Seed the fictional demo notebook (Lumen Bikes) into a real OneNote for the web notebook.

    python3 scripts/demo/demo_notebook.py --list      # sections, titles and the HTML of every page
    python3 scripts/demo/demo_notebook.py --prompt    # one Hermes chat message that creates them

The content lives in addin/dev/demo-notebook.json (the same file the Electron harness loads with
?seed=demo). It is invented: no real people, companies or work data. This script only prints; it
never talks to OneNote, Microsoft or Hermes. Create the notebook and its three sections by hand
first (the plugin has no section-creation tool), then paste the --prompt output into a Hermes chat
paired with that notebook. Each page creation asks for approval.
"""
from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

SEED = Path(__file__).resolve().parents[2] / "addin" / "dev" / "demo-notebook.json"


def load() -> dict:
    return json.loads(SEED.read_text(encoding="utf-8"))


def main(argv=None) -> int:
    ap = argparse.ArgumentParser()
    g = ap.add_mutually_exclusive_group(required=True)
    g.add_argument("--list", action="store_true")
    g.add_argument("--prompt", action="store_true")
    a = ap.parse_args(argv)
    nb = load()
    if a.list:
        print(f"Notebook: {nb['name']}")
        for s in nb["sections"]:
            print(f"\n== Section: {s['name']}")
            for p in s["pages"]:
                print(f"-- {p['title']}\n{''.join(p.get('html') or [])}")
        return 0
    lines = [f"Set up the demo notebook '{nb['name']}'. I created the sections already. Call onenote_list_sections, then "
             "call onenote_create_page once per page below with exactly this title and HTML and the sectionId of the "
             "named section. Do not change the text. After the last page, call brain_sync and tell me the coverage."]
    for s in nb["sections"]:
        lines.append(f"\nSection '{s['name']}':")
        for p in s["pages"]:
            lines.append(f"- title: {p['title']}\n  html: {''.join(p.get('html') or [])}")
    print("\n".join(lines))
    return 0


if __name__ == "__main__":
    sys.exit(main())
