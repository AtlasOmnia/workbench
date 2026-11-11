# SPDX-License-Identifier: Apache-2.0
from pathlib import Path
import json
import re
ROOT = Path(__file__).resolve().parents[1]
def test_excel_design_bundle_via_real_registration(plugin, tmp_path):
    from test_onenote import FakeCtx
    from tools.skills_tool_plugin import _plugin_skill_linked_files, _serve_skill_file
    import shutil
    ctx = FakeCtx(); plugin.register(ctx)
    entry = ctx.skills["office-desktop"]
    dest = tmp_path / "relocated"; shutil.copytree(entry.parent, dest)
    for ref in ["excel-workflow.md", "design-playbook.md", "provenance.md"]:
        rel = "references/" + ref
        assert "](" + rel + ")" in entry.read_text()
        assert rel in _plugin_skill_linked_files(dest)["references"]
        reply = json.loads(_serve_skill_file(dest, rel, "hermes-office-desktop:office-desktop"))
        assert reply["success"]
        assert reply["content"] == (entry.parent / rel).read_text()
    text = (dest / "references/excel-workflow.md").read_text()
    names = {s["name"] for s in json.loads((ROOT / "schemas/excel.json").read_text())}
    assert set(re.findall(r"\bexcel_[a-z_]+\b", text)) <= names
    for word in ["formulas", "currency", "assumptions", "reconciliation", "trendlines", "axes"]: assert word in text
    for word in ["spreadsheet", "menu", "deck"]: assert word in entry.read_text().split("version:")[0].lower()


def test_actual_host_metadata_routes_ordinary_requests(tmp_path):
    import os
    import subprocess
    import sys
    code = '''from pathlib import Path
import sys
from hermes_cli.plugins import PluginManager
from hermes_cli.plugins_manifest import parse_manifest_file
p = Path(sys.argv[1]); manager = PluginManager()
manager._load_plugin(parse_manifest_file(p / "plugin.yaml", p, "user", ""))
item = next(x for x in manager.list_plugin_skill_metadata() if x["name"] == "hermes-office-desktop:office-desktop")
for word in ["spreadsheet", "budget", "menu", "brief", "deck"]:
    assert word in item["description"].lower(), item
manager.unload()
'''
    env = dict(os.environ, HERMES_HOME=str(tmp_path / "home"), KANBAN_HOME=str(tmp_path / "kanban"), GATEWAY_LOCK_DIR=str(tmp_path / "locks"))
    env.pop("_HERMES_GATEWAY", None)
    result = subprocess.run([sys.executable, "-c", code, str(ROOT)], env=env, text=True, capture_output=True, timeout=90)
    assert result.returncode == 0, result.stdout + result.stderr
