"""Rasterise the brand SVGs to PNG with headless Chrome.

Chrome is the only rasteriser on this machine (no rsvg/cairo/ImageMagick),
and it is also the one that matters: these are SVG files, so the browser is
the reference renderer. Usage: python3 rasterize.py
"""

import pathlib
import re
import subprocess
import sys

CHROME = "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome"
HERE = pathlib.Path(__file__).resolve().parent

# name -> target output width in px. The device scale factor is derived from it
# so the output keeps the SVG's aspect ratio exactly; scaling the Chrome window
# instead would letterbox rather than scale.
JOBS = {
    "board": 1600,
    "social-preview": 1280,
    "logo": 960,
    "mark": 256,
}


def viewbox_size(source: pathlib.Path):
    match = re.search(r'viewBox="0 0 (\d+) (\d+)"', source.read_text(encoding="utf-8"))
    if not match:
        raise SystemExit(f"{source.name}: no viewBox — cannot infer the pixel size")
    return int(match.group(1)), int(match.group(2))


def shoot(source: pathlib.Path, out: pathlib.Path, width: int, height: int, scale: int):
    result = subprocess.run(
        [
            CHROME,
            "--headless",
            "--disable-gpu",
            "--hide-scrollbars",
            "--no-sandbox",
            f"--force-device-scale-factor={scale}",
            f"--window-size={width},{height}",
            f"--screenshot={out}",
            f"file://{source}",
        ],
        capture_output=True,
        text=True,
        timeout=120,
    )
    if not out.exists():
        raise SystemExit(f"{source.name} -> {out.name}: chrome produced nothing\n{result.stderr[:400]}")


for name, target_width in JOBS.items():
    source = HERE / f"{name}.svg"
    if not source.exists():
        print(f"{name}.svg missing, skipped")
        continue
    vb_w, vb_h = viewbox_size(source)
    scale = target_width / vb_w
    if scale > 4 or scale < 1:
        raise SystemExit(
            f"{name}: scale {scale:.2f} is outside 1-4. "
            "Upscaling a layout past 4 softens the type; redraw at a bigger preset."
        )
    out = HERE / f"{name}.png"
    if out.exists():
        out.unlink()
    shoot(source, out, vb_w, vb_h, scale)
    print(
        f"{name}.svg {vb_w}x{vb_h} @ {scale:g}x -> {out.name} "
        f"{int(vb_w * scale)}x{int(vb_h * scale)}  {out.stat().st_size // 1024} KB"
    )
