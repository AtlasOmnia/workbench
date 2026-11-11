# SPDX-License-Identifier: Apache-2.0
"""Build assets/demo-highlight.gif for the README from masked live-harness screenshots.

The frames are Excel run 1 of the live harness run ``hol-maint-d715acd-all-3`` (plugin ``d715acd``): the chat
asks for an overview, writes a range, formats its header, adds a SUM, draws a chart and undoes it; each change
lands in the open workbook and is selected (highlighted) there. The harness masks the Microsoft account area
in every screenshot ("account area hidden").

Only frames that were reviewed by eye for privacy (account area masked; no email address, person's name,
OneDrive or file path visible) are used, and each is pinned by sha256: a frame whose bytes differ is refused,
so a re-run of the harness cannot slip an unreviewed picture into the README. To use new frames, review each
one and update FRAMES.

    uv run --no-project --python 3.12 --with pillow==12.3.0 \
        python scripts/make_readme_gif.py --shots <HOL_OUT>/shots [--out assets/demo-highlight.gif]

Same inputs and the same Pillow version give the same bytes. The GIF is a stand-in: a capture from the
Hermes Desktop app itself will replace it.
"""
from __future__ import annotations

import argparse
import hashlib
import sys
from pathlib import Path

from PIL import Image, ImageDraw, ImageFont

ROOT = Path(__file__).resolve().parent.parent
MAX_BYTES = 3 * 1024 * 1024
WIDTH = 1200  # output width in px; height follows the screenshots' aspect ratio, plus the caption strip
CAPTION_H = 56

# (file, sha256 of the reviewed screenshot, caption, display ms)
FRAMES = [
    ("excel-r1-02-overview.png", "e7e717decce361886745a86583cc0b9a86040120741c23b5c908df6af91ce290",
     "1/6  Paired: Hermes reads the open workbook", 1500),
    ("excel-r1-03-write.png", "3c413781520572c76b38cd74d9a0ca91a5a8da18f6c434800b00eab7af6a1d05",
     "2/6  Approved edit lands: A1:B5 written and highlighted", 1800),
    ("excel-r1-04-format.png", "874bd140159a01a610fc0d68aec314f433a94e3ec7f514c112ccc2631fc05385",
     "3/6  Header formatted (bold, pale fill)", 1600),
    ("excel-r1-05-sum.png", "9c06c6144d4b9196bf7aa67c29a219cbe983ccf0ce8657659dc60f50bade9e8c",
     "4/6  Live =SUM total added and highlighted", 1600),
    ("excel-r1-06-chart.png", "d7fa6fc58a113a3979128b2d2be77c70c9333571d4bbb5bf2ecb1012284358b8",
     "5/6  Chart created from the range", 1700),
    ("excel-r1-07-undo.png", "5e242e3b2cd31a29174ea1f4c21ed8fc70fd8c79468f5d020043f862f593080b",
     "6/6  \"Undo that\": the chart is removed again", 2000),
]


def sha256(p: Path) -> str:
    return hashlib.sha256(p.read_bytes()).hexdigest()


def frame(src: Path, caption: str, font) -> Image.Image:
    im = Image.open(src).convert("RGB")
    h = round(im.height * WIDTH / im.width)
    im = im.resize((WIDTH, h), Image.Resampling.LANCZOS)
    out = Image.new("RGB", (WIDTH, h + CAPTION_H), (24, 24, 27))
    out.paste(im, (0, 0))
    d = ImageDraw.Draw(out)
    d.text((20, h + CAPTION_H // 2), caption, font=font, fill=(245, 245, 245), anchor="lm")
    d.text((WIDTH - 20, h + CAPTION_H // 2), "live harness, d715acd", font=font, fill=(150, 150, 160), anchor="rm")
    return out


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(description=(__doc__ or "").splitlines()[0])
    ap.add_argument("--shots", required=True, type=Path, help="the harness run's shots/ directory")
    ap.add_argument("--out", type=Path, default=ROOT / "assets" / "demo-highlight.gif")
    a = ap.parse_args(argv)
    bad = []
    for name, want, _, _ in FRAMES:
        p = a.shots / name
        if not p.is_file():
            bad.append(f"{name}: missing")
        elif sha256(p) != want:
            bad.append(f"{name}: sha256 differs from the reviewed frame")
    if bad:
        print("refusing to build (unreviewed or missing frames):\n  " + "\n  ".join(bad), file=sys.stderr)
        return 2
    font = ImageFont.load_default(size=24)
    rgb = [frame(a.shots / n, c, font) for n, _, c, _ in FRAMES]
    # One shared adaptive palette keeps colours stable across frames (no flicker) and the file small.
    strip = Image.new("RGB", (rgb[0].width, rgb[0].height * len(rgb)))
    for i, f in enumerate(rgb):
        strip.paste(f, (0, i * f.height))
    pal = strip.quantize(colors=128, method=Image.Quantize.MEDIANCUT, dither=Image.Dither.NONE)
    frames = [f.quantize(palette=pal, dither=Image.Dither.NONE) for f in rgb]
    a.out.parent.mkdir(parents=True, exist_ok=True)
    frames[0].save(a.out, save_all=True, append_images=frames[1:], duration=[d for *_, d in FRAMES],
                   loop=0, optimize=True, disposal=1)
    size = a.out.stat().st_size
    total = sum(d for *_, d in FRAMES) / 1000
    print(f"{a.out}: {len(frames)} frames, {total:.1f} s, {size} bytes, sha256 {sha256(a.out)}")
    if size >= MAX_BYTES:
        print(f"too large: {size} >= {MAX_BYTES}", file=sys.stderr)
        return 1
    return 0


if __name__ == "__main__":
    sys.exit(main())
