"""Luma palette pass: theme, styles, sheets and drawings to the Luma Advisors system.

Usage: python tools/postbuild/luma_palette.py [workbook]

Palette (lumaadvisors.com.au):
  #5C2D91 brand purple — the one accent: titles, links, Financial tab
  #04001F near-black   — section headings
  #2B2733 dark neutral — emphasis font
  #B1AFAD warm grey    — input shading family
  #4F485E grey-violet  — structure tabs / secondary

Three sources merged:
  phase_c2 — explicit colour map, teal/salmon hue remap, font-family consolidation
  phase_e  — help-label greens/blues/maroon folded to brand purple / warning red
  phase_f  — mint help-block fill folded to pale lavender

Deliberately NOT remapped: the TOC length-guard conditional format keeps its
warning red (font C00000, fill FFCCCC). The pale-yellow FFFFCC fill is folded.

Idempotent on the current workbook: a second run finds nothing to do and writes
nothing. A fixed-point guard refuses to write if any colour in the result is still
remappable, so a future in-band addition to the palette cannot drift on reruns.

One honest limitation: the band remap is the v3.1.0 formula, which is not a fixed
point for high-saturation salmon inputs (the output can stay above the band's
saturation gate). Against a true v3.0.0-state workbook the pass reproduces v3.1.0
once; the guard then correctly blocks a second application. No v3.0.0 workbook is
available to test that path, so it is documented rather than asserted.

Pure XML surgery: no COM, no recalculation.
"""

from __future__ import annotations

import colorsys
import re
import sys
import zipfile
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from sanitise_workbook import write_deterministic

EXPLICIT = {
    "3F656F": "5C2D91",
    "66A1AF": "4F485E",
    "B1D2DA": "DED9E8",
    "92563E": "6E6862",
    "D7876B": "B1AFAD",
    "EFBEAB": "DCDAD7",
    "44546A": "04001F",
    "2E4950": "2B2733",
    "0563C1": "5C2D91",
    "954F72": "7A5AB5",
    "595959": "4F485E",
    "1F4E79": "2B1A54",
    "157A5F": "5C2D91",
    "2E75B6": "7A5AB5",
    "375623": "38206B",
}

# Colours the hue remap must never touch: the palette targets, deliberate
# warnings, neutrals, and the legacy accents that phase e/f own instead.
KEEP = {
    "FFCCCC", "C00000", "FF0000",  # warning family (TOC length guard)
    "000000", "FFFFFF", "A6A6A6", "D9D9D9", "F8F8F8", "FA7D00",
    "5C2D91", "04001F", "2B2733", "B1AFAD", "4F485E", "6E6862", "DED9E8",
    "DCDAD7", "7A5AB5", "2B1A54", "38206B", "F3F1F6", "F5F5F5",
    "006600", "00C000", "0000FF", "0000C0", "800000", "CCFFCC", "FFFFCC",
}

# phase_e: help-label fonts folded to one accent (or the existing warning red).
FONT_REMAP = {
    "FF006600": "FF5C2D91",
    "FF00C000": "FF5C2D91",
    "FF0000FF": "FF5C2D91",
    "FF0000C0": "FF5C2D91",
    "FF800000": "FFC00000",
}

# phase_f: legacy fills folded to neutrals already in the palette. The TOC
# warning fill (FFCCCC) is deliberately absent: it is intentional.
FILL_REMAP = {
    "FFCCFFCC": "FFF3F1F6",
    "FFFFFFCC": "FFF3F1F6",
}

FONT_FAMILIES = {
    '<name val="Roboto Condensed"/>': '<name val="Aptos Display"/>',
    '<name val="Bahnschrift SemiCondensed"/>': '<name val="Aptos"/>',
    '<name val="Arial"/>': '<name val="Aptos"/>',
    '<name val="Aptos Narrow"/>': '<name val="Aptos"/>',
}

# Rich-text runs in sharedStrings.xml use rFont, not name.
RFONT_FAMILIES = {
    '<rFont val="Roboto Condensed"/>': '<rFont val="Aptos Display"/>',
    '<rFont val="Bahnschrift SemiCondensed"/>': '<rFont val="Aptos"/>',
    '<rFont val="Arial"/>': '<rFont val="Aptos"/>',
    '<rFont val="Aptos Narrow"/>': '<rFont val="Aptos"/>',
}


def hue_map(hexrgb: str):
    r, g, b = (int(hexrgb[i : i + 2], 16) / 255 for i in (0, 2, 4))
    h, lightness, s = colorsys.rgb_to_hls(r, g, b)
    deg = h * 360
    if s < 0.03:
        return None
    if 150 <= deg <= 225 and s > 0.05:
        h2, s2 = 268 / 360, min(0.35, s * 0.45)
    elif 5 <= deg <= 40 and s > 0.10:
        h2, s2 = 28 / 360, s * 0.22
    else:
        return None
    r2, g2, b2 = colorsys.hls_to_rgb(h2, lightness, s2)
    return f"{round(r2 * 255):02X}{round(g2 * 255):02X}{round(b2 * 255):02X}"


def map_colour(hexrgb: str):
    up = hexrgb.upper()
    if up in KEEP or up in EXPLICIT.values():
        return None
    if up in EXPLICIT:
        return EXPLICIT[up]
    return hue_map(up)


def recolour(text: str) -> tuple[str, int]:
    count = 0

    def rep8(match):
        nonlocal count
        pre, argb = match.group(1), match.group(2)
        alpha, rgb = argb[:2], argb[2:]
        new = map_colour(rgb)
        if new is None:
            return match.group(0)
        count += 1
        return f'{pre}"{alpha}{new}"'

    def rep6(match):
        nonlocal count
        pre, rgb = match.group(1), match.group(2)
        new = map_colour(rgb)
        if new is None:
            return match.group(0)
        count += 1
        return f'{pre}"{new}"'

    text = re.sub(r'(rgb=)"([0-9A-Fa-f]{8})"', rep8, text)
    text = re.sub(r'(val=)"([0-9A-Fa-f]{6})"(?=/>|")', rep6, text)
    # Fixed-point guard: nothing in the result may still be remappable. The band
    # remap lowers saturation below the input gate by design, so a colour that is
    # still remappable after one pass would drift on every subsequent run.
    residual = re.findall(r'rgb="[0-9A-Fa-f]{8}"', text) + re.findall(
        r'val="[0-9A-Fa-f]{6}"(?=/>|")', text
    )
    for token in residual:
        hexv = token.split('"')[1]
        rgb = hexv[2:] if len(hexv) == 8 else hexv
        if map_colour(rgb) is not None:
            raise ValueError(
                f"remap is not a fixed point: {rgb} is still remappable after one pass"
            )
    return text, count


def _section(text: str, tag: str) -> tuple[str, str, str]:
    start = text.index(f"<{tag}")
    end = text.index(f"</{tag}>") + len(f"</{tag}>")
    return text[:start], text[start:end], text[end:]


def run(workbook: Path) -> list[str]:
    with zipfile.ZipFile(workbook) as archive:
        order = archive.namelist()
        parts = {n: archive.read(n) for n in order}

    changes = []

    # Main palette: theme, styles, every worksheet, drawing, and shared strings
    # (rich-text runs carry colour and rFont inside sharedStrings.xml).
    for name in ["xl/theme/theme1.xml", "xl/styles.xml", "xl/sharedStrings.xml"] + [
        n for n in order if re.match(r"xl/(worksheets/sheet|drawings/drawing)\d+\.xml$", n)
    ]:
        text = parts[name].decode("utf-8")
        recoloured, count = recolour(text)
        if recoloured != text:
            parts[name] = recoloured.encode("utf-8")
            changes.append(f"{name}: {count} colours remapped")

    # Font families + help-label fonts + help-block fill, in styles.xml only.
    styles = parts["xl/styles.xml"].decode("utf-8")

    for old, new in FONT_FAMILIES.items():
        if old in styles:
            styles = styles.replace(old, new)
            changes.append(f"font family {old} -> {new}")
    # Rich-text font names live in sharedStrings.xml as rFont.
    shared = parts["xl/sharedStrings.xml"].decode("utf-8")
    for old, new in RFONT_FAMILIES.items():
        if old in shared:
            shared = shared.replace(old, new)
            changes.append(f"rich-text font {old} -> {new}")
    parts["xl/sharedStrings.xml"] = shared.encode("utf-8")

    pre, fonts, post = _section(styles, "fonts")
    folded = 0
    for old, new in FONT_REMAP.items():
        hits = fonts.count(f'rgb="{old}"')
        if hits:
            fonts = fonts.replace(f'rgb="{old}"', f'rgb="{new}"')
            folded += hits
    if folded:
        changes.append(f"help-label fonts folded to palette ({folded} entries)")
    styles = pre + fonts + post

    pre, fills, post = _section(styles, "fills")
    folded = 0
    for old, new in FILL_REMAP.items():
        hits = fills.count(f'rgb="{old}"')
        if hits:
            fills = fills.replace(f'rgb="{old}"', f'rgb="{new}"')
            folded += hits
    # phase_f also caught the indexed mint fill left by earlier passes.
    hits = fills.count('indexed="42"')
    if hits:
        fills = fills.replace('<fgColor indexed="42"/>', '<fgColor rgb="FFF3F1F6"/>')
        folded += hits
    if folded:
        changes.append(f"help-block fills folded to neutrals ({folded} entries)")
    styles = pre + fills + post

    parts["xl/styles.xml"] = styles.encode("utf-8")

    # Guard: the TOC warning dxf must survive.
    dxfs = styles[styles.index("<dxfs") : styles.index("</dxfs>")]
    assert "FFFFCCCC" in dxfs and "FFC00000" in dxfs, "TOC warning red was remapped"

    if not changes:
        return []
    write_deterministic(workbook, parts)
    return changes


def main() -> None:
    if len(sys.argv) > 2:
        sys.exit("FAIL: usage: python tools/postbuild/luma_palette.py [workbook]")
    workbook = Path(sys.argv[1] if len(sys.argv) == 2 else "ozzit.xlsx")
    if not workbook.is_file():
        sys.exit(f"FAIL: no such workbook: {workbook}")
    try:
        changes = run(workbook)
    except (OSError, ValueError, AssertionError, zipfile.BadZipFile) as exc:
        sys.exit(f"FAIL: {exc}")
    if changes:
        for change in changes:
            print(change)
    else:
        print("Luma palette already applied; no changes")
    print(f"OK: {workbook}")


if __name__ == "__main__":
    main()
