# SPDX-License-Identifier: Apache-2.0
"""Make a disposable Office document that opens a sideloaded add-in's task pane on launch.

Same transform as Microsoft's office-addin-dev-settings "sideload" (MIT): copy its
*WithTaskPane template and put the add-in id/version into the webextension part
(store="developer" storeType="Registry", i.e. the per-user WEF\\Developer entry).
The template is read from an existing local install; it is not shipped in this repo.

    python make_taskpane_doc.py <template.xlsx|.docx|.pptx> <addin-id> <version> <out>
"""
import sys
import zipfile

PARTS = {".xlsx": "xl/webextensions/webextension.xml", ".docx": "word/webextensions/webextension.xml",
         ".pptx": "ppt/webextensions/webextension.xml"}


def main(template, addin_id, version, out):
    part = PARTS[template[template.rfind("."):].lower()]
    with zipfile.ZipFile(template) as src, zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as dst:
        for info in src.infolist():
            data = src.read(info.filename)
            if info.filename == part:
                text = data.decode("utf-8-sig")
                if "00000000-0000-0000-0000-000000000000" not in text:
                    raise SystemExit("template has no placeholder id")
                data = ("\ufeff" + text.replace("00000000-0000-0000-0000-000000000000", addin_id)
                        .replace("1.0.0.0", version)).encode("utf-8")
            dst.writestr(info, data)
    print(out)


if __name__ == "__main__":
    main(*sys.argv[1:5])
