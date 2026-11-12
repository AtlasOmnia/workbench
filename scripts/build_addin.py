# SPDX-License-Identifier: Apache-2.0
"""Build the static bridge add-in site for a base URL (local only; never deploys), or refresh the plugin's bundle.

    python3 scripts/build_addin.py --bundle             # refresh addin/SHA256SUMS, icons and build-info (commit them)
    python3 scripts/build_addin.py --base https://<host>/<path>/ --out <dir>   # a site to self-host (optional)
    python3 scripts/build_addin.py --base https://addin.test:8444/office/live/ --out build/harness --include-dev

The plugin serves its bundle itself (bridge_server.read_site_file): exactly the files addin/SHA256SUMS lists, each
only while its sha256 matches. Production output and the bundle exclude addin/dev/ (the Office.js mock). A site build
prints a SHA-256 manifest of every emitted file so a later deploy can be verified byte-for-byte.
"""
from __future__ import annotations

import argparse
import ast
import hashlib
import json
import re
import shutil
import subprocess
import struct
import sys
import zlib
from pathlib import Path
from urllib.parse import urlsplit

ROOT = Path(__file__).resolve().parent.parent
ADDIN = ROOT / "addin"


def _bridge_default_base() -> str:
    """The plugin's default add-in base: the DEFAULT_ADDIN_BASE literal in bridge_server.py (one source, read from
    the plugin source without importing it)."""
    tree = ast.parse((ROOT / "bridge_server.py").read_text(encoding="utf-8"))
    for node in tree.body:
        if isinstance(node, ast.Assign) and any(getattr(t, "id", None) == "DEFAULT_ADDIN_BASE" for t in node.targets):
            return ast.literal_eval(node.value)
    raise SystemExit("bridge_server.py has no DEFAULT_ADDIN_BASE")


DEFAULT_BASE = _bridge_default_base()  # the plugin-served loopback site (its port follows the bridge)
# One add-in per Office host. Ids are fixed so re-uploading replaces the earlier copy.
HOSTS = {
    "excel": {"ID": "6b0f3c1e-2d4a-4f7e-9c5b-8a1d2e3f4a51", "LABEL": "Excel", "HOST": "Workbook",
              "SET": "ExcelApi", "SETVER": "1.7", "NOUN": "workbook"},
    "word": {"ID": "b6744181-1c6b-4301-b139-19cd2f7af1ec", "LABEL": "Word", "HOST": "Document",
             "SET": "WordApi", "SETVER": "1.3", "NOUN": "document"},
    "powerpoint": {"ID": "516118d2-cf21-497f-a118-28595d71bbba", "LABEL": "PowerPoint", "HOST": "Presentation",
                   "SET": "PowerPointApi", "SETVER": "1.3", "NOUN": "presentation"},
    # OneNote add-ins run only in OneNote for the web (OneNoteApi 1.1); the manifest host is "Notebook".
    "onenote": {"ID": "3f6c2a8e-5b1d-4e7a-9c40-7d2e1b9a6c53", "LABEL": "OneNote", "HOST": "Notebook",
                "SET": "OneNoteApi", "SETVER": "1.1", "NOUN": "notebook"},
    # Outlook on the web: a MailApp manifest of its own (read + compose forms, pinnable pane), not the TaskPaneApp one.
    "outlook": {"ID": "6601c392-1a11-41f3-bd86-7d4263d9a1c5", "TEMPLATE": "manifest-outlook.template.xml"},
}
DEFAULT_TEMPLATE = "manifest.template.xml"


def png(size: int) -> bytes:
    """Deterministic solid-glyph icon (no binary assets in git)."""
    rows = []
    for y in range(size):
        row = bytearray([0])
        for x in range(size):
            inside = abs(x - size / 2) < size * 0.38 and abs(y - size / 2) < size * 0.38
            bar = inside and (abs(x - size * 0.36) < size * 0.08 or abs(x - size * 0.64) < size * 0.08
                              or abs(y - size / 2) < size * 0.07)
            row += bytes((255, 255, 255, 255)) if bar else (bytes((16, 124, 65, 255)) if inside else bytes(4))
        rows.append(bytes(row))
    raw = zlib.compress(b"".join(rows), 9)

    def chunk(t, d):
        return struct.pack(">I", len(d)) + t + d + struct.pack(">I", zlib.crc32(t + d) & 0xFFFFFFFF)
    return b"\x89PNG\r\n\x1a\n" + chunk(b"IHDR", struct.pack(">IIBBBBB", size, size, 8, 6, 0, 0, 0)) \
        + chunk(b"IDAT", raw) + chunk(b"IEND", b"")


BUILD_INFO = "shared/build-info.js"


def source_revision() -> dict:
    """The plugin commit this site was built from, and whether the tree had uncommitted changes (never hidden)."""
    def git(*args):
        try:
            r = subprocess.run(["git", "-C", str(ROOT), *args], capture_output=True, text=True, timeout=10)
        except (OSError, subprocess.SubprocessError):
            return None
        return r.stdout.strip() if r.returncode == 0 else None
    commit = git("rev-parse", "--short=12", "HEAD")
    top = git("rev-parse", "--show-toplevel")
    if not top or Path(top).resolve() != ROOT.resolve():  # an export inside some other repository is not this source
        return {"commit": None, "dirty": None}
    status = git("status", "--porcelain")
    return {"commit": commit, "dirty": (bool(status) if status is not None else None) if commit else None}


def build_info(version: str, sums: list[tuple[str, str]], rev: dict) -> dict:
    """Build identity shown in the pane footer: the add-in version, a content id over every other emitted file
    (so a cached older pane shows a different id), and the source commit. A label, not an attestation."""
    content = "".join(f"{h}  {n}\n" for h, n in sums).encode("utf-8")
    return {"version": version, "build": hashlib.sha256(content).hexdigest()[:12], **rev}


# ---- the plugin-served bundle: the pinned tree's own addin/ files ----------------------------------------------------
BUNDLE_DIRS = tuple(HOSTS) + ("shared", "assets")
BUNDLE_FILES = ("support.html",)
BUNDLE_SUMS = "SHA256SUMS"


def bundle_names() -> list[str]:
    """Every file the bridge may serve, relative to addin/: the five task panes, shared code, icons, support page.
    Never dev/ (the Office.js mock), the manifest templates, package.json or SHA256SUMS itself."""
    names = [p.relative_to(ADDIN).as_posix() for sub in BUNDLE_DIRS for p in (ADDIN / sub).rglob("*") if p.is_file()]
    return sorted(names + list(BUNDLE_FILES))


def _sums(names) -> list[tuple[str, str]]:
    return [(hashlib.sha256((ADDIN / n).read_bytes()).hexdigest(), n) for n in names]


def bundle_build_info_text() -> str:
    """shared/build-info.js for the bundle: version and a content id over every other bundled file (no commit:
    a file cannot name the commit it is part of)."""
    template = (ADDIN / DEFAULT_TEMPLATE).read_text(encoding="utf-8")
    version = ".".join(re.search(r"<Version>([0-9.]+)</Version>", template).group(1).split(".")[:3])
    info = build_info(version, _sums(n for n in bundle_names() if n != BUILD_INFO), {"commit": None, "dirty": None})
    return ("// SPDX-License-Identifier: Apache-2.0\n"
            "// Generated by scripts/build_addin.py --bundle for the plugin-served add-in; do not edit.\n"
            f"export const BUILD = Object.freeze({json.dumps(info, sort_keys=True)})\n")


def bundle_sums_text() -> str:
    return "".join(f"{h}  {n}\n" for h, n in _sums(bundle_names()))


def write_bundle() -> str:
    """Refresh the committed bundle metadata in addin/: icons, shared/build-info.js, then SHA256SUMS."""
    (ADDIN / "assets").mkdir(exist_ok=True)
    for size in (16, 32, 80):
        (ADDIN / "assets" / f"icon-{size}.png").write_bytes(png(size))
    (ADDIN / BUILD_INFO).write_text(bundle_build_info_text(), encoding="utf-8")
    text = bundle_sums_text()
    (ADDIN / BUNDLE_SUMS).write_text(text, encoding="utf-8")
    return text


def render_manifests(base: str) -> dict[str, str]:
    """The five add-in manifests (app -> XML text) for a base URL. The plugin's office_connect writes these same
    texts to the user's manifest folder, so the site build and the plugin never diverge."""
    parts = urlsplit(base)
    if parts.scheme != "https" or not parts.hostname or "@" in parts.netloc or not base.endswith("/") \
            or parts.query or parts.fragment:
        raise SystemExit("--base must be an https URL ending in /")
    origin = f"{parts.scheme}://{parts.netloc}"
    template = (ADDIN / DEFAULT_TEMPLATE).read_text(encoding="utf-8")
    out = {}
    for app, h in HOSTS.items():
        own = (ADDIN / h["TEMPLATE"]).read_text(encoding="utf-8") if "TEMPLATE" in h else template
        manifest = own.replace("{{BASE}}", base).replace("{{ORIGIN}}", origin).replace("{{APP}}", app)
        for key, value in h.items():
            if key != "TEMPLATE":
                manifest = manifest.replace("{{" + key + "}}", value)
        if "{{" in manifest:
            raise SystemExit("unreplaced template token")
        out[app] = manifest
    return out


def build(base: str, out: Path, include_dev: bool) -> list[tuple[str, str]]:
    manifests = render_manifests(base)  # validates base first
    if out.exists():
        shutil.rmtree(out)
    out.mkdir(parents=True)
    for sub in tuple(HOSTS) + ("shared",) + (("dev",) if include_dev else ()):
        shutil.copytree(ADDIN / sub, out / sub)
    (out / "assets").mkdir()
    for size in (16, 32, 80):
        (out / "assets" / f"icon-{size}.png").write_bytes(png(size))
    template = (ADDIN / DEFAULT_TEMPLATE).read_text(encoding="utf-8")
    for app, manifest in manifests.items():
        (out / f"manifest-{app}.xml").write_text(manifest, encoding="utf-8")
    shutil.copy(ADDIN / "support.html", out / "support.html")

    def digest():
        return [(hashlib.sha256(f.read_bytes()).hexdigest(), f.relative_to(out).as_posix())
                for f in sorted(p for p in out.rglob("*") if p.is_file())]
    version = re.search(r"<Version>([0-9.]+)</Version>", template).group(1)
    version = ".".join(version.split(".")[:3])
    info = build_info(version, [x for x in digest() if x[1] != BUILD_INFO], source_revision())
    (out / BUILD_INFO).write_text("// SPDX-License-Identifier: Apache-2.0\n// Generated by scripts/build_addin.py; do not edit.\n"
                                  f"export const BUILD = Object.freeze({json.dumps(info, sort_keys=True)})\n", encoding="utf-8")
    sums = digest()
    (out / "SHA256SUMS").write_text("".join(f"{h}  {n}\n" for h, n in sums), encoding="utf-8")
    return sums


def main(argv=None):
    ap = argparse.ArgumentParser()
    ap.add_argument("--base", default=DEFAULT_BASE)
    ap.add_argument("--out", default=str(ROOT / "dist" / "office" / "live"))
    ap.add_argument("--include-dev", action="store_true", help="include the Office.js mock (never for deploys)")
    ap.add_argument("--bundle", action="store_true",
                    help="refresh addin/SHA256SUMS, addin/assets and addin/shared/build-info.js (the plugin-served add-in)")
    a = ap.parse_args(argv)
    if a.bundle:
        print(write_bundle(), end="")
        return
    for h, n in build(a.base, Path(a.out), a.include_dev):
        print(h, n)


if __name__ == "__main__":
    sys.exit(main())
