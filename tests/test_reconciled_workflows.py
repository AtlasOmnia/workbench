# SPDX-License-Identifier: Apache-2.0
from pathlib import Path
ROOT = next(p for p in Path(__file__).resolve().parents if (p / "plugin.yaml").exists() or (p / "plugins").exists())
def test_research_reconciliation_contract():
    refs = ROOT / 'skills/office-desktop/references'
    design = (refs / "design-playbook.md").read_text()
    for phrase in ["claim → evidence", "keyboard", "blocker", "microtype", "CSS pixels"]: assert phrase in design
    text = (refs / "excel-workflow.md").read_text()
    for phrase in ["ReportedPL", "IncludeInPL", "Tolerance", "SUMIFS", "cash or accrual", "5,000 cells"]: assert phrase in text
    assert "registration excludes" in text
    ppt = (refs / "powerpoint-design.md").read_text()
    assert "office_insert_image" in ppt and "diagnostic-only" in ppt
