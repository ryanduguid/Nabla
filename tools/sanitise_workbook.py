"""Strip the parts Excel adds on save and rebuild the zip deterministically.

Saving a workbook through Excel (a COM script, or a manual open and save)
adds parts that do not belong in a distributed file:

- one xl/printerSettings/printerSettingsN.bin per worksheet, pinned to the
  printer installed where the save happened,
- <x15ac:absPath> in xl/workbook.xml, recording the directory the file was
  saved from,
- ca="1"/aca="1" always-calculate flags on cells that compute nothing
  volatile,
- an xl/worksheets/_rels/sheetN.xml.rels left empty for any worksheet whose
  only relationship was its printer settings,
- a [Content_Types].xml rewritten so the .bin default points at printer
  settings and every customProperty*.bin gets its own Override.

All of it is removed here and the archive is rewritten sorted, at a fixed
timestamp, at deflate level 9, so two builds of the same content produce the
same bytes. verify_workbook.py fails the file on absPath or a stray
always-calculate flag regardless; this tool exists so a save through Excel
does not have to be reverted.

Run after any edit session that ended in an Excel save:

    python tools/sanitise_workbook.py ozzit.xlsx

Then re-run the verify gates, and open the result in Excel once: zip surgery
can pass the gates while Excel still objects to the file.
"""

from __future__ import annotations

import re
import sys
import zipfile
from pathlib import Path

CUSTOMPROP_CT = (
    "application/vnd.openxmlformats-officedocument.spreadsheetml.customProperty"
)
FIXED_DATE = (2026, 1, 1, 0, 0, 0)
CELL_RE = re.compile(r"<c\b.*?(?:/>|</c>)", re.DOTALL)
EMPTY_RELS = re.compile(rb"<Relationships[^>]*>\s*</Relationships>")


def sanitise(workbook: Path) -> list[str]:
    parts: dict[str, bytes] = {}
    with zipfile.ZipFile(workbook) as z:
        for n in z.namelist():
            parts[n] = z.read(n)

    log: list[str] = []

    # Printer settings: parts, then the relationships, then the pageSetup
    # attributes that referenced them.
    ps = [n for n in parts if n.startswith("xl/printerSettings/")]
    for n in ps:
        del parts[n]
    if ps:
        log.append(f"removed {len(ps)} printerSettings parts")

    for name in [n for n in parts if n.endswith(".rels")]:
        txt = parts[name].decode("utf-8")
        new = re.sub(
            r'<Relationship[^>]*Target="[^"]*printerSettings[^"]*"[^>]*/>',
            "",
            txt,
        )
        if new != txt:
            parts[name] = new.encode("utf-8")

    sheets = [n for n in parts if n.startswith("xl/worksheets/sheet")]
    cleared = 0
    for name in sheets:
        txt = parts[name].decode("utf-8")
        if "r:id=" not in txt:
            continue
        new = re.sub(r'(<pageSetup\b[^>]*?)\s+r:id="[^"]*"', r"\1", txt)
        if new != txt:
            parts[name] = new.encode("utf-8")
            cleared += 1
    if cleared:
        log.append(f"cleared pageSetup r:id on {cleared} sheets")

    # Content types: printer Overrides gone, .bin default back at custom
    # properties, per-part customProperty Overrides dropped.
    ct = parts["[Content_Types].xml"].decode("utf-8")
    before = len(ct)
    ct = re.sub(
        r'<Override[^>]*PartName="/xl/printerSettings/[^"]*"[^>]*/>', "", ct
    )
    if '<Default Extension="bin"' in ct:
        ct = re.sub(
            r'<Default Extension="bin"[^>]*/>',
            f'<Default Extension="bin" ContentType="{CUSTOMPROP_CT}"/>',
            ct,
        )
        ct = re.sub(
            r'<Override[^>]*PartName="/xl/customProperty\d+\.bin"[^>]*/>',
            "",
            ct,
        )
    parts["[Content_Types].xml"] = ct.encode("utf-8")
    if len(ct) != before:
        log.append(f"[Content_Types].xml {before} -> {len(ct)} bytes")

    # The save path.
    wb = parts["xl/workbook.xml"].decode("utf-8")
    n_abs = len(re.findall(r"<x15ac:absPath[^>]*/>", wb))
    if n_abs:
        wb = re.sub(r"<x15ac:absPath[^>]*/>", "", wb)
        wb = re.sub(
            r"<mc:AlternateContent[^>]*>\s*<mc:Choice[^>]*>\s*</mc:Choice>"
            r"\s*</mc:AlternateContent>",
            "",
            wb,
        )
        parts["xl/workbook.xml"] = wb.encode("utf-8")
        log.append(f"removed {n_abs} x15ac:absPath")

    # Always-calculate flags belong on CELL() formulas and nowhere else.
    total_ca = 0
    for name in sheets:
        txt = parts[name].decode("utf-8")
        if 'ca="1"' not in txt and 'aca="1"' not in txt:
            continue
        hits = 0

        def scrub(m: re.Match) -> str:
            nonlocal hits
            cell = m.group(0)
            if "CELL(" in cell:
                return cell
            if 'ca="1"' not in cell and 'aca="1"' not in cell:
                return cell
            hits += 1
            return cell.replace(' aca="1"', "").replace(' ca="1"', "")

        new = CELL_RE.sub(scrub, txt)
        if hits:
            parts[name] = new.encode("utf-8")
            total_ca += hits
    if total_ca:
        log.append(f"cleared always-calculate flags on {total_ca} cells")

    # Rels parts with nothing left in them.
    empty = [
        n
        for n, b in parts.items()
        if n.startswith("xl/worksheets/_rels/") and EMPTY_RELS.search(b)
    ]
    for n in empty:
        del parts[n]
    if empty:
        log.append(f"dropped {len(empty)} empty worksheet rels")

    if not log:
        return ["already clean"]

    tmp = workbook.with_suffix(".xlsx.tmp")
    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED, compresslevel=9) as z:
        for name in sorted(parts):
            zi = zipfile.ZipInfo(name, date_time=FIXED_DATE)
            zi.compress_type = zipfile.ZIP_DEFLATED
            zi.external_attr = 0o600 << 16
            z.writestr(zi, parts[name])
    tmp.replace(workbook)
    return log


def main() -> None:
    if len(sys.argv) != 2:
        sys.exit("usage: python tools/sanitise_workbook.py ozzit.xlsx")
    workbook = Path(sys.argv[1])
    if not workbook.is_file():
        sys.exit(f"no such workbook: {workbook}")
    for line in sanitise(workbook):
        print(line)
    print(f"OK: {workbook} sanitised, {workbook.stat().st_size:,} bytes")


if __name__ == "__main__":
    main()
