"""Export a generated diagram HTML to a standalone SVG, per the skill's
references/export.md procedure. Diagram-only: the SVG node, no page chrome.

Usage: python3 export_svg.py <source.html> [...]
"""

import pathlib
import re
import sys

XML_HEADER = '<?xml version="1.0" encoding="UTF-8"?>\n'

# Mirrors the <link> in the source files. `&` must be `&amp;` because a
# standalone .svg is parsed as strict XML.
FONT_IMPORT = (
    "<style>@import url('https://fonts.googleapis.com/css2"
    "?family=Instrument+Serif:ital@0;1"
    "&amp;family=Geist:wght@400;500;600"
    "&amp;family=Geist+Mono:wght@400;500;600"
    "&amp;display=swap');</style>"
)

RGBA = re.compile(
    r'(fill|stroke)="rgba\(\s*(\d+)\s*,\s*(\d+)\s*,\s*(\d+)\s*,\s*(\d*\.?\d+)\s*\)"'
)
SVG_BLOCK = re.compile(r'<svg\b.*?</svg>', re.DOTALL)


def rgba_to_hex(match):
    prop, r, g, b, alpha = match.groups()
    return '{0}="#{1:02x}{2:02x}{3:02x}" {0}-opacity="{4}"'.format(
        prop, int(r), int(g), int(b), alpha
    )


def export(source: pathlib.Path) -> pathlib.Path:
    html = source.read_text(encoding="utf-8")

    match = SVG_BLOCK.search(html)
    if not match:
        raise SystemExit(f"no <svg> block in {source} — not a diagram file")
    svg = match.group(0)

    if 'xmlns=' not in svg.split('>', 1)[0]:
        raise SystemExit(f"{source}: opening <svg> has no xmlns")
    if 'viewBox=' not in svg:
        raise SystemExit(f"{source}: no viewBox — refusing to guess the size")

    if '<defs>' in svg:
        svg = svg.replace('<defs>', '<defs>\n        ' + FONT_IMPORT, 1)
    else:
        svg = re.sub(
            r'(</title>|<desc[^>]*>.*?</desc>)',
            r'\1\n  <defs>' + FONT_IMPORT + '</defs>',
            svg,
            count=1,
            flags=re.DOTALL,
        )

    svg = RGBA.sub(rgba_to_hex, svg)
    svg = re.sub(r'(fill|stroke)="transparent"', r'\1="none"', svg)

    out = source.with_suffix('.svg')
    out.write_text(XML_HEADER + svg + '\n', encoding='utf-8')
    return out


for arg in sys.argv[1:]:
    written = export(pathlib.Path(arg))
    print(f"{arg} -> {written}")
