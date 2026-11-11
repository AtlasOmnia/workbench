# SPDX-License-Identifier: Apache-2.0
"""Tool-coverage plans shared by the Electron pane harness (mock Office.js) and the live
Windows run (real Excel / Word / PowerPoint). Each step calls one plugin tool handler.

Modes:
  read    call; must succeed
  keep    call a mutation and keep it (setup for later steps)
  undo    probe document; call; probe (must differ); office_undo; probe (must equal the first)
  noundo  mutation that must succeed and report undo.available == false
  error   call; must fail with ``expect`` category
  special harness actions: reload (pane reload + resume), other_chat (cross-chat refusal), js (setup script)
"""
from __future__ import annotations

PNG_1PX = ("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR4nGNgYGD4DwABBAEAX+XDSwAAAABJRU5ErkJggg==")
WORD_OOXML = (
    '<pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage">'
    '<pkg:part pkg:name="/_rels/.rels" pkg:contentType="application/vnd.openxmlformats-package.relationships+xml">'
    '<pkg:xmlData><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" '
    'Target="word/document.xml"/></Relationships></pkg:xmlData></pkg:part>'
    '<pkg:part pkg:name="/word/document.xml" '
    'pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"><pkg:xmlData>'
    '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r>'
    '<w:t>Inserted from OOXML</w:t></w:r></w:p></w:body></w:document></pkg:xmlData></pkg:part></pkg:package>')

DATA = [["Region", "Sales", "Units"], ["East", 100, 3], ["West", 300, 1], ["North", 200, 2], ["South", 50, 4]]


def S(st):
    return st.get("sheet", "Sheet1")


def excel_plan():
    return [
        {"tool": "excel_get_workbook_overview", "args": {}, "mode": "read",
         "save": lambda d, st: st.update(sheet=d["activeSheet"])},
        {"tool": "excel_write_range", "args": lambda st: {"sheet": S(st), "address": "A1:C5", "values": DATA}, "mode": "keep"},
        {"tool": "excel_read_range", "args": lambda st: {"sheet": S(st), "address": "A1:C5"}, "mode": "read",
         "check": lambda d, st: d["values"][2] == ["West", 300, 1]},
        {"special": "reload"},
        {"tool": "excel_read_range", "args": lambda st: {"sheet": S(st), "address": "A1:C5"}, "mode": "read", "label": "read after pane reload",
         "check": lambda d, st: d["values"][1] == ["East", 100, 3]},
        {"special": "other_chat", "tool": "excel_read_range", "args": lambda st: {"sheet": S(st), "address": "A1"}},
        {"tool": "excel_summarize_range_stats", "args": lambda st: {"sheet": S(st), "address": "B2:B5"}, "mode": "read",
         "check": lambda d, st: d["sum"] == 650 and d["numericCount"] == 4},
        {"tool": "excel_write_range", "args": lambda st: {"sheet": S(st), "address": "E1:F2", "values": [["x", "=B2*2"], [1, 2]]}, "mode": "undo"},
        {"tool": "excel_format_range", "args": lambda st: {"sheet": S(st), "address": "A1:C1", "format": {
            "bold": True, "fillColor": "#107C41", "fontColor": "#FFFFFF", "numberFormat": "@", "borders": True,
            "borderColor": "#333333", "columnAutofit": True, "horizontalAlignment": "Center"}}, "mode": "undo"},
        {"tool": "excel_conditional_format_range", "args": lambda st: {"sheet": S(st), "address": "B2:B5", "ruleType": "cellValue",
                                                                       "operator": "GreaterThan", "formula1": "150", "fillColor": "#FFC7CE"}, "mode": "undo"},
        {"tool": "excel_conditional_format_range", "args": lambda st: {"sheet": S(st), "address": "A2:A5", "ruleType": "containsText",
                                                                       "text": "st", "fontColor": "#9C0006"}, "mode": "undo"},
        {"tool": "excel_freeze_panes", "args": lambda st: {"sheet": S(st), "rows": 1, "columns": 1}, "mode": "undo"},
        {"tool": "excel_set_data_validation", "args": lambda st: {"sheet": S(st), "address": "D2:D4", "type": "list", "source": "Yes,No",
                                                                  "promptTitle": "Pick", "promptMessage": "Yes or No"}, "mode": "undo"},
        {"tool": "excel_create_sheet", "args": {"name": "HOD Summary"}, "mode": "undo"},
        {"tool": "excel_create_sheet", "args": {"name": "Scratch"}, "mode": "keep"},
        {"tool": "excel_activate_sheet", "args": {"name": "Scratch"}, "mode": "undo"},
        {"tool": "excel_rename_sheet", "args": {"oldName": "Scratch", "newName": "Scratch2"}, "mode": "undo"},
        {"tool": "excel_copy_sheet", "args": lambda st: {"sourceName": S(st), "newName": "Data copy"}, "mode": "undo"},
        {"tool": "excel_protect_sheet", "args": {"sheet": "Scratch", "allowSort": True}, "mode": "undo"},
        {"tool": "excel_protect_sheet", "args": {"sheet": "Scratch"}, "mode": "keep"},
        {"tool": "excel_unprotect_sheet", "args": {"sheet": "Scratch"}, "mode": "undo"},
        {"tool": "excel_delete_sheet", "args": {"name": "Scratch"}, "mode": "noundo"},
        {"tool": "excel_insert_table", "args": lambda st: {"sheet": S(st), "address": "A1:C5", "hasHeaders": True}, "mode": "undo"},
        {"tool": "excel_insert_table", "args": lambda st: {"sheet": S(st), "address": "A1:C5", "hasHeaders": True}, "mode": "keep",
         "save": lambda d, st: st.update(table=d["table"])},
        {"tool": "excel_sort_table", "args": lambda st: {"tableName": st["table"], "columnIndex": 1, "ascending": False}, "mode": "undo"},
        {"tool": "excel_filter_table", "args": lambda st: {"tableName": st["table"], "columnIndex": 1, "criterion1": ">150"}, "mode": "undo"},
        {"tool": "excel_create_chart", "args": lambda st: {"sheet": S(st), "dataRange": "A1:B5", "chartType": "ColumnClustered",
                                                           "title": "Sales by region"}, "mode": "undo"},
        {"tool": "excel_find_replace", "args": lambda st: {"sheet": S(st), "find": "est", "replace": "EST", "matchCase": True}, "mode": "undo"},
        {"tool": "excel_write_range", "args": lambda st: {"sheet": S(st), "address": "E10:F11", "values": [["keep", 1], ["me", 2]]}, "mode": "keep"},
        {"tool": "excel_format_range", "args": lambda st: {"sheet": S(st), "address": "E10:F10", "format": {"italic": True}}, "mode": "keep"},
        {"tool": "excel_clear_range", "args": lambda st: {"sheet": S(st), "address": "E10:F11"}, "mode": "undo"},
        {"tool": "excel_write_range", "args": lambda st: {"sheet": S(st), "address": "G2", "values": [["=B2*C2"]]}, "mode": "keep"},
        {"tool": "excel_autofill_formula", "args": lambda st: {"sheet": S(st), "sourceAddress": "G2", "targetAddress": "G2:G5"}, "mode": "undo"},
        {"tool": "excel_create_sheet", "args": {"name": "Pivot"}, "mode": "keep"},
        {"tool": "excel_create_pivot_table", "args": lambda st: {"sheet": "Pivot", "address": "A3", "sourceSheet": S(st), "sourceRange": "A1:C5",
                                                                 "rowFields": ["Region"], "dataFields": [{"name": "Sales", "summarizeBy": "sum"}]},
         "mode": "undo"},
        {"tool": "excel_run_calculation", "args": {}, "mode": "read"},
        {"tool": "excel_add_comment", "args": lambda st: {"sheet": S(st), "address": "B2", "text": "Check this number"}, "mode": "undo"},
        {"tool": "excel_add_comment", "args": lambda st: {"sheet": S(st), "address": "B3", "text": "Second note"}, "mode": "keep"},
        {"tool": "excel_clear_comments", "args": lambda st: {"sheet": S(st), "address": "A1:C5"}, "mode": "undo"},
        {"tool": "excel_import_csv_to_sheet", "args": lambda st: {"sheet": S(st), "address": "I1", "hasHeaders": True,
                                                                  "csv": 'name,score\nAda,"=HYPERLINK(1)"\n"Lovelace, A",42\n'}, "mode": "undo"},
        {"tool": "excel_read_range", "args": {"sheet": "No such sheet", "address": "A1"}, "mode": "error", "expect": "Office"},
        {"tool": "excel_write_range", "args": lambda st: {"sheet": S(st), "address": "A1:B2", "values": [[1, 2, 3]]}, "mode": "error", "expect": "Validation"},
    ]


def word_select(text):
    return {"special": "js", "label": f"select '{text}'",
            "live": f"Word.run(async c => {{ const r = c.document.body.search({text!r}, {{ matchCase: true }}); r.load('items'); await c.sync(); r.items[0].select(); await c.sync(); return r.items.length }})",
            "mock": f"(() => {{ const d = globalThis.__hermesMockDocument; const i = d.paragraphs.findIndex(p => p.text.includes({text!r})); const k = d.paragraphs[i].text.indexOf({text!r}); d.setSelection(i, k, k + {len(text)}); return 1 }})()"}


def word_plan():
    return [
        {"tool": "word_get_document_overview", "args": {}, "mode": "read"},
        {"tool": "word_insert_text", "args": {"text": "Quarterly report\nRevenue grew in the East region.\nCosts fell in the West region.", "location": "end"}, "mode": "keep"},
        {"tool": "word_read_document_text", "args": {"maxLength": 2000}, "mode": "read", "check": lambda d, st: "Revenue grew" in d["text"]},
        {"special": "reload"},
        {"tool": "word_get_document_structure", "args": {"maxParagraphs": 20}, "mode": "read", "label": "structure after pane reload",
         "check": lambda d, st: d["paragraphCount"] >= 3},
        {"special": "other_chat", "tool": "word_read_selection", "args": {}},
        word_select("Revenue"),
        {"tool": "word_read_selection", "args": {}, "mode": "read", "check": lambda d, st: d["text"] == "Revenue"},
        {"tool": "word_apply_edit_plan", "args": {"edits": [{"find": "grew", "replace": "rose", "highlightColor": "Yellow"},
                                                           {"find": "West", "replace": "Western", "reason": "style"}]}, "mode": "undo"},
        {"tool": "word_insert_text", "args": {"text": "Prepared by Hermes. ", "location": "start"}, "mode": "undo"},
        {"tool": "word_insert_text", "args": {"text": " End of report.", "location": "end"}, "mode": "undo"},
        word_select("Revenue"),
        {"tool": "word_insert_text", "args": {"text": " (draft)", "location": "selection"}, "mode": "undo"},
        word_select("Revenue"),
        {"tool": "word_replace_selection", "args": {"text": "Income", "highlightColor": "Yellow"}, "mode": "undo"},
        {"tool": "word_find_replace", "args": {"find": "region", "replace": "territory", "matchCase": True}, "mode": "undo"},
        word_select("Costs fell"),
        {"tool": "word_format_selection", "args": {"bold": True, "italic": True, "color": "#C00000", "fontSize": 16}, "mode": "undo"},
        word_select("Costs fell"),
        {"tool": "word_apply_style", "args": {"style": "Heading 2"}, "mode": "undo"},
        {"tool": "word_insert_ooxml", "args": {"ooxml": WORD_OOXML, "location": "end"}, "mode": "undo"},
        word_select("Revenue"),
        {"tool": "word_insert_comment", "args": {"text": "Source for this?"}, "mode": "undo"},
        {"tool": "word_find_replace", "args": {"find": "x" * 300, "replace": "y"}, "mode": "error", "expect": "Validation"},
    ]


def _layout(st, name):
    for m in st["layouts"]["slideMasters"]:
        for lay in m["layouts"]:
            if lay["name"].lower() == name.lower():
                return lay["id"]
    return st["layouts"]["slideMasters"][0]["layouts"][-1]["id"]


def powerpoint_plan():
    return [
        {"tool": "powerpoint_get_presentation_overview", "args": {}, "mode": "read"},
        {"tool": "powerpoint_list_slide_layouts", "args": {}, "mode": "read", "save": lambda d, st: st.update(layouts=d)},
        {"tool": "powerpoint_create_text_slide", "args": {"title": "Agenda", "bullets": ["Pair the pane", "Edit live", "Undo"]}, "mode": "keep",
         "save": lambda d, st: st.update(slide=d["slideIndex"])},
        {"tool": "powerpoint_read_all_slides", "args": {}, "mode": "read", "check": lambda d, st: any("Agenda" in " ".join(s["text"]) for s in d["slides"])},
        {"special": "reload"},
        {"tool": "powerpoint_read_slide_content", "args": lambda st: {"slideIndex": st["slide"]}, "mode": "read", "label": "slide content after pane reload"},
        {"special": "other_chat", "tool": "powerpoint_read_all_slides", "args": {}},
        {"tool": "powerpoint_read_selected_slides", "args": {}, "mode": "read"},
        {"tool": "powerpoint_create_slide", "args": lambda st: {"layoutId": _layout(st, "Blank")}, "mode": "undo"},
        {"tool": "powerpoint_create_text_slide", "args": {"title": "Next steps", "body": "Ship it"}, "mode": "undo"},
        {"tool": "powerpoint_add_text_box", "args": lambda st: {"text": "Temporary", "slideIndex": st["slide"], "left": 40, "top": 400, "width": 300, "height": 40}, "mode": "undo"},
        {"tool": "powerpoint_add_text_box", "args": lambda st: {"text": "Hermes note", "slideIndex": st["slide"], "left": 500, "top": 400, "width": 300, "height": 40},
         "mode": "keep", "save": lambda d, st: st.update(shape=d["shapeId"])},
        {"tool": "powerpoint_update_slide_text", "args": lambda st: {"text": "Updated by Hermes", "slideIndex": st["slide"], "shapeId": st["shape"]}, "mode": "undo"},
        {"tool": "powerpoint_add_speaker_notes", "args": lambda st: {"slideIndex": st["slide"], "text": "Say hi"}, "mode": "error", "expect": "Unsupported"},
        {"tool": "powerpoint_add_shape", "args": lambda st: {"shapeType": "Rectangle", "slideIndex": st["slide"], "left": 600, "top": 100, "width": 120, "height": 80}, "mode": "undo"},
        {"tool": "powerpoint_delete_slide", "args": {"slideIndex": 0}, "mode": "undo"},
        {"tool": "powerpoint_duplicate_slide", "args": lambda st: {"slideIndex": st["slide"]}, "mode": "undo"},
        {"tool": "powerpoint_reorder_slide", "args": lambda st: {"fromIndex": st["slide"], "toIndex": 0}, "mode": "undo"},
        {"tool": "powerpoint_find_and_replace_text", "args": lambda st: {"findText": "note", "replaceText": "NOTE", "slideIndex": st["slide"]}, "mode": "undo"},
        {"tool": "powerpoint_set_slide_layout", "args": lambda st: {"slideIndex": st["slide"], "layoutId": _layout(st, "Title Only")}, "mode": "undo"},
        {"tool": "powerpoint_format_text", "args": lambda st: {"slideIndex": st["slide"], "shapeId": st["shape"],
                                                               "font": {"bold": True, "italic": True, "color": "#C00000", "size": 28, "name": "Arial"}}, "mode": "undo"},
        {"tool": "powerpoint_add_image", "args": lambda st: {"base64Image": PNG_1PX, "slideIndex": st["slide"], "left": 50, "top": 50, "width": 60, "height": 60}, "mode": "undo"},
        {"tool": "powerpoint_reorder_slide", "args": {"fromIndex": 0, "toIndex": 99}, "mode": "error", "expect": "Validation"},
    ]


def _pid(st, title):
    return next(p["id"] for p in st["pages"] if p["title"] == title)


def _user_opens(title, wait_ms=2600):
    """The user clicks another page; the pane's visit watcher must index it with no Hermes call."""
    return {"special": "js", "label": f"user opens '{title}' (visit-indexing)",
            "mock": f"(async () => {{ const nb = globalThis.__hermesMockNotebook; nb.userOpen(nb.pageByTitle({title!r}).id); "
                    f"await new Promise(r => setTimeout(r, {wait_ms})); return 1 }})()",
            "live": None}


def onenote_plan():
    MEET, HUB, PRICING = "Launch meeting 2026-11-03", "Q4 Launch", "Pricing review 2026-10-20"
    return [
        {"tool": "onenote_get_context", "args": {}, "mode": "read", "check": lambda d, st: d["notebook"]["name"].startswith("Lumen")},
        {"tool": "onenote_list_sections", "args": {}, "mode": "read", "check": lambda d, st: len(d["sections"]) == 3},
        {"tool": "onenote_list_pages", "args": {"all": True}, "mode": "read", "save": lambda d, st: st.update(pages=d["pages"]),
         "check": lambda d, st: len(d["pages"]) == 5 and d["complete"]},
        {"tool": "brain_sync", "args": {}, "mode": "read", "check": lambda d, st: d["coverage"]["label"] == "indexed 1 of 5 pages"},
        {"tool": "onenote_get_page", "args": {}, "mode": "read", "check": lambda d, st: d["title"] == HUB and "[[Risks]]" in d["text"]},
        {"special": "reload"},
        {"tool": "onenote_get_context", "args": {}, "mode": "read", "label": "context after pane reload",
         "check": lambda d, st: d["page"]["title"] == HUB},
        {"special": "other_chat", "tool": "onenote_get_context", "args": {}},
        {"tool": "onenote_get_page", "args": lambda st: {"pageId": _pid(st, MEET)}, "mode": "error", "expect": "Conflict",
         "label": "another page cannot be read without opening it"},
        _user_opens(MEET),
        {"tool": "brain_ask", "args": {"question": "What did we decide about pricing?"}, "mode": "read",
         "label": "visit-indexed page answers a question",
         "check": lambda d, st: d["results"] and d["results"][0]["title"] == MEET and d["coverage"]["indexed"] == 2
         and d["results"][0]["open"]["args"]["pageId"] == _pid(st, MEET)},
        {"tool": "onenote_open_page", "args": lambda st: {"pageId": _pid(st, PRICING)}, "mode": "read",
         "check": lambda d, st: d["title"] == PRICING},
        {"tool": "onenote_get_page", "args": lambda st: {"pageId": _pid(st, PRICING)}, "mode": "read",
         "check": lambda d, st: "Rejected option B" in d["text"]},
        {"tool": "onenote_open_page", "args": lambda st: {"pageId": _pid(st, HUB)}, "mode": "read"},
        {"tool": "brain_links", "args": {}, "mode": "undo", "label": "Links & Backlinks block (undo removes it)"},
        {"tool": "brain_links", "args": {}, "mode": "keep",
         "check": lambda d, st: {b["title"] for b in d["backlinks"]} == {MEET, PRICING}
         and [x["label"] for x in d["links"]][:1] == ["October pricing review"]},
        {"tool": "brain_links", "args": {}, "mode": "keep", "label": "regenerate replaces Hermes' own block",
         "check": lambda d, st: d["replacedPreviousBlock"] is True},
        {"tool": "onenote_append", "args": {"html": "<p>Hermes: pilot review moved to Thursday.</p>"}, "mode": "undo"},
        {"tool": "onenote_insert_table", "args": {"rows": [["City", "Status"], ["Porto", "Booked"]], "title": "Pilot cities"},
         "mode": "undo"},
        {"tool": "onenote_create_page", "args": {"title": "Q4 Launch"}, "mode": "error", "expect": "Conflict",
         "label": "duplicate title refused"},
        {"tool": "onenote_create_page", "args": {"title": "Scratch page", "html": "<p>temporary</p>"}, "mode": "keep",
         "check": lambda d, st: "manual" in d["undo"]["effect"]},
        {"tool": "office_undo", "args": {"app": "onenote"}, "mode": "read", "label": "undo of a created page = manual delete + opened",
         "check": lambda d, st: d["manualDelete"] is True and d["opened"] is True},
        {"tool": "onenote_open_page", "args": lambda st: {"pageId": _pid(st, MEET)}, "mode": "read"},
        {"tool": "onenote_get_page", "args": lambda st: {"pageId": _pid(st, MEET)}, "mode": "read"},
        {"tool": "brain_decision_page", "args": lambda st: {
            "sourcePageId": _pid(st, MEET), "title": "Pilot decision 2026-11-03",
            "decisions": [{"text": "Pilot in two cities before any national launch",
                           "quote": "run a pilot in two cities before any national launch"},
                          {"text": "Keep $29 per month for the pilot", "quote": "keep the $29 per month plan for the pilot"}],
            "actions": [{"text": "Draft the pilot plan", "owner": "Priya", "due": "Nov 12", "quote": "Priya drafts the pilot plan by Nov 12"},
                        {"text": "Book the two pilot cities", "owner": "Sam", "quote": "Sam books the two pilot cities"},
                        {"text": "Check supplier lead times", "quote": "Someone needs to check supplier lead times"}]},
         "mode": "keep", "save": lambda d, st: st.update(decision=d["pageId"]),
         "check": lambda d, st: d["unverifiedQuotes"] == 0 and d["unknownOwners"] == 3},
        {"tool": "brain_make_deliverables", "args": lambda st: {"decisionPageId": st["decision"]}, "mode": "read",
         "label": "deliverables report hosts that are not paired (no retry)",
         "check": lambda d, st: {r["status"] for r in d["results"]} == {"not paired"}},
        {"tool": "brain_why", "args": lambda st: {"host": "onenote", "pageId": st["decision"]}, "mode": "read",
         "check": lambda d, st: d["writes"] and d["writes"][0]["kind"] == "create_page"},
        {"tool": "brain_resurface", "args": {"count": 3}, "mode": "read", "check": lambda d, st: len(d["pages"]) >= 1},
        {"tool": "brain_sync", "args": {"crawl": True, "maxPages": 10}, "mode": "read", "label": "approved crawl, back to start",
         "check": lambda d, st: d["coverage"]["indexed"] == d["coverage"]["discovered"] and d["crawl"]["returnedToStart"]},
        {"tool": "onenote_append", "args": {"html": "<p onclick=\"x()\">bad</p>"}, "mode": "error", "expect": "Validation"},
    ]


PLANS = {"excel": excel_plan, "word": word_plan, "powerpoint": powerpoint_plan, "onenote": onenote_plan}
