"""Apply Ozzit's professional presentation rules to an existing workbook.

Usage: python tools/polish_workbook.py [workbook]

Pure, targeted OOXML surgery. Formula text, cached values, drawings, namespace
prefixes and the Advanced Formula Environment store are preserved. The pass is
idempotent and writes through the canonical archive implementation used by the
Excel-save sanitiser.
"""

from __future__ import annotations

import html
import re
import sys
import xml.etree.ElementTree as ET
import zipfile
from pathlib import Path

from sanitise_workbook import write_deterministic

MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
DOC_REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NS = {"m": MAIN_NS}

EXPECTED_STRAPLINES = {
    "oz.CountCλ": "Count the number of times one or more characters appear in a string",
    "oz.CountRowsλ": "Count the number of numbers in each row of an array",
    "oz.CountColsλ": "Count the number of numbers in each column of an array",
    "oz.CountARowsλ": "Count all non-empty cells in each row of an array",
    "oz.CountAColsλ": "Count all non-empty cells in each column of an array",
    "oz.IsBetweenEλ": "Determine if a value is between a lower and upper limit",
    "oz.RangeToDAEλ": "Convert a static range into a dynamic array",
    "oz.FinancialRatios": "Three dozen financial ratios",
}
COVER_OLD = (
    "Each function worksheet has green shaded cells. These cells contain the function "
    "whose name appears in the worksheet's title."
)
COVER_NEW = (
    "Each function worksheet has pale purple shaded cells. These cells contain the "
    "function whose name appears in the worksheet's title."
)


def load_parts(workbook: Path) -> dict[str, bytes]:
    with zipfile.ZipFile(workbook) as archive:
        return {name: archive.read(name) for name in archive.namelist()}


def sheet_paths(parts: dict[str, bytes]) -> dict[str, str]:
    workbook = ET.fromstring(parts["xl/workbook.xml"])
    rels = ET.fromstring(parts["xl/_rels/workbook.xml.rels"])
    targets = {rel.attrib["Id"]: rel.attrib["Target"] for rel in rels}
    rid = f"{{{DOC_REL_NS}}}id"
    sheets = workbook.find("m:sheets", NS)
    return {
        sheet.attrib["name"]: "xl/" + targets[sheet.attrib[rid]].lstrip("/")
        for sheet in sheets
    }


def shared_values(text: str) -> list[str]:
    root = ET.fromstring(text)
    return ["".join(item.itertext()) for item in root.findall("m:si", NS)]


def ensure_shared(text: str, values: list[str], value: str) -> tuple[str, int, bool]:
    if value in values:
        return text, values.index(value), False
    escaped = html.escape(value, quote=False)
    text = text.replace("</sst>", f"<si><t>{escaped}</t></si></sst>")
    values.append(value)
    text = re.sub(
        r'(<sst\b[^>]*\buniqueCount=")\d+("[^>]*>)',
        lambda match: f'{match.group(1)}{len(values)}{match.group(2)}',
        text,
        count=1,
    )
    return text, len(values) - 1, True


def set_existing_shared_cell(text: str, address: str, index: int) -> tuple[str, bool]:
    pattern = re.compile(
        rf'(<c r="{re.escape(address)}"[^>]*\bt="s"[^>]*>\s*<v>)\d+(</v>\s*</c>)'
    )
    match = pattern.search(text)
    if not match:
        raise ValueError(f"shared-string cell {address} not found")
    current = int(re.search(r"\d+", match.group(0).split("<v>", 1)[1]).group())
    if current == index:
        return text, False
    return pattern.sub(rf"\g<1>{index}\2", text, count=1), True


def add_financial_ratios_strapline(text: str, index: int) -> tuple[str, bool]:
    if re.search(r'<c r="A2"[^>]*>', text):
        return set_existing_shared_cell(text, "A2", index)
    # Financial Ratios' sheet begins at row 1, so the first closing row is row 1.
    row = (
        '<row r="2" spans="1:11" x14ac:dyDescent="0.25">'
        f'<c r="A2" s="288" t="s"><v>{index}</v></c></row>'
    )
    text, count = re.subn(r"</row>", lambda match: match.group(0) + row, text, count=1)
    if count != 1:
        raise ValueError("could not insert oz.FinancialRatios row 2")
    if '<mergeCell ref="A2:K2"' not in text:
        match = re.search(r'<mergeCells count="(\d+)">(.*?)</mergeCells>', text, re.DOTALL)
        if not match:
            raise ValueError("oz.FinancialRatios mergeCells not found")
        merged = (
            f'<mergeCells count="{int(match.group(1)) + 1}">'
            f'{match.group(2)}<mergeCell ref="A2:K2"/></mergeCells>'
        )
        text = text[: match.start()] + merged + text[match.end() :]
    return text, True


def update_shared_and_sheets(
    parts: dict[str, bytes], paths: dict[str, str]
) -> list[str]:
    shared = parts["xl/sharedStrings.xml"].decode("utf-8")
    values = shared_values(shared)
    changes = []

    if COVER_OLD in shared:
        shared = shared.replace(COVER_OLD, COVER_NEW)
        values[values.index(COVER_OLD)] = COVER_NEW
        changes.append("updated Cover formula-highlight wording")

    added_references = 0
    for sheet_name, strapline in EXPECTED_STRAPLINES.items():
        shared, index, _added = ensure_shared(shared, values, strapline)
        path = paths[sheet_name]
        sheet = parts[path].decode("utf-8")
        if sheet_name == "oz.FinancialRatios":
            sheet, changed = add_financial_ratios_strapline(sheet, index)
            if changed:
                added_references += 1
        else:
            sheet, changed = set_existing_shared_cell(sheet, "A2", index)
        if changed:
            parts[path] = sheet.encode("utf-8")
            changes.append(f"set {sheet_name}!A2")

    if added_references:
        shared = re.sub(
            r'(<sst\b[^>]*\bcount=")(\d+)(")',
            lambda match: (
                f'{match.group(1)}{int(match.group(2)) + added_references}{match.group(3)}'
            ),
            shared,
            count=1,
        )
    parts["xl/sharedStrings.xml"] = shared.encode("utf-8")
    return changes


def update_styles(parts: dict[str, bytes]) -> list[str]:
    styles = parts["xl/styles.xml"].decode("utf-8")
    original = styles
    changes = []

    mint_count = styles.count('<fgColor indexed="42"/>')
    if mint_count:
        styles = styles.replace(
            '<fgColor indexed="42"/>', '<fgColor rgb="FFDED9E8"/>'
        )
        changes.append(f"replaced {mint_count} mint fill definition")

    fonts = re.search(r"<fonts\b.*?</fonts>", styles, re.DOTALL)
    if not fonts:
        raise ValueError("styles.xml fonts section missing")
    font_text = fonts.group(0)
    grey_count = font_text.count('<color rgb="FF808080"/>')
    if grey_count:
        new_fonts = font_text.replace(
            '<color rgb="FF808080"/>', '<color rgb="FF6E6862"/>'
        )
        styles = styles[: fonts.start()] + new_fonts + styles[fonts.end() :]
        changes.append(f"raised {grey_count} strapline font colour")

    xfs = re.search(r"<cellXfs\b.*?</cellXfs>", styles, re.DOTALL)
    if not xfs:
        raise ValueError("styles.xml cellXfs section missing")
    xf_text = xfs.group(0)
    date_style_count = xf_text.count('numFmtId="14"')
    if date_style_count:
        new_xfs = xf_text.replace('numFmtId="14"', 'numFmtId="169"')
        styles = styles[: xfs.start()] + new_xfs + styles[xfs.end() :]
        changes.append(f"made {date_style_count} date styles locale-safe")

    if styles != original:
        parts["xl/styles.xml"] = styles.encode("utf-8")
    return changes


def polish(workbook: Path) -> list[str]:
    parts = load_parts(workbook)
    before = dict(parts)
    paths = sheet_paths(parts)
    changes = update_shared_and_sheets(parts, paths)
    changes.extend(update_styles(parts))
    if parts == before:
        return ["already polished"]
    write_deterministic(workbook, parts)
    return changes


def main() -> None:
    if len(sys.argv) > 2:
        sys.exit("FAIL: usage: python tools/polish_workbook.py [workbook]")
    workbook = Path(sys.argv[1] if len(sys.argv) == 2 else "ozzit.xlsx")
    if not workbook.is_file():
        sys.exit(f"FAIL: no such workbook: {workbook}")
    try:
        changes = polish(workbook)
    except (
        OSError,
        KeyError,
        UnicodeDecodeError,
        ValueError,
        ET.ParseError,
        zipfile.BadZipFile,
    ) as exc:
        sys.exit(f"FAIL: cannot polish {workbook}: {exc}")
    for change in changes:
        print(change)
    print(f"OK: {workbook} polished, {workbook.stat().st_size:,} bytes")


if __name__ == "__main__":
    main()
