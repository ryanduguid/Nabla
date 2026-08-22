"""Sheet-name pass: the oz.SumContains tab gains the λ every sibling carries.

Usage: python tools/postbuild/sheet_names.py [workbook]

Every single-function demonstration sheet is named after its function,
λ included; oz.SumContains was the one exception (its title drawing already
reads SumContainsλ). The rename must land in five parts together: the sheet
element in xl/workbook.xml, the two table-of-contents hyperlinks on sheet2,
the cached CELL("filename") title on the sheet itself (sheet25), the titles
list in docProps/app.xml, and the table-of-contents row text in
xl/sharedStrings.xml.

Anchors are delimiter-bounded so oz.SumContainsλ (the function name, which is
everywhere) never matches. Each part carries an asserted hit count. A second
run reports "already applied" and writes nothing. Pure text surgery: no COM,
no recalculation.
"""

from __future__ import annotations

import sys
import zipfile
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from sanitise_workbook import write_deterministic

OLD = "oz.SumContains"
NEW = "oz.SumContainsλ"

# part -> list of (old anchor, new anchor, expected count)
EDITS = {
    "xl/workbook.xml": [
        (f'<sheet name="{OLD}"', f'<sheet name="{NEW}"', 1),
    ],
    "xl/worksheets/sheet2.xml": [
        (f"'{OLD}'!", f"'{NEW}'!", 4),
    ],
    "xl/worksheets/sheet25.xml": [
        (f"<v>{OLD}</v>", f"<v>{NEW}</v>", 1),
    ],
    "docProps/app.xml": [
        (f"<vt:lpstr>{OLD}</vt:lpstr>", f"<vt:lpstr>{NEW}</vt:lpstr>", 1),
    ],
    "xl/sharedStrings.xml": [
        (f"<t>{OLD}</t>", f"<t>{NEW}</t>", 1),
    ],
}


def run(workbook: Path) -> list[str]:
    with zipfile.ZipFile(workbook) as archive:
        parts = {n: archive.read(n) for n in archive.namelist()}

    failures: list[str] = []
    texts: dict[str, str] = {}
    for part, edits in EDITS.items():
        if part not in parts:
            failures.append(f"missing part {part}")
            continue
        text = parts[part].decode("utf-8")
        texts[part] = text
        for old, new, expected in edits:
            counts = {"pre-rename": text.count(old), "renamed": text.count(new)}
            if sum(counts.values()) != expected:
                failures.append(
                    f"{part}: {old!r} expected {expected} recognised anchors, "
                    f"got pre-rename={counts['pre-rename']} renamed={counts['renamed']}"
                )
    if failures:
        raise ValueError("; ".join(failures))

    changed = False
    for part, edits in EDITS.items():
        text = texts[part]
        for old, new, _expected in edits:
            text = text.replace(old, new)
        if text != texts[part]:
            parts[part] = text.encode("utf-8")
            changed = True
    if changed:
        write_deterministic(workbook, parts)
        return ["workbook"]
    return []


def main() -> None:
    if len(sys.argv) > 2:
        sys.exit("FAIL: usage: python tools/postbuild/sheet_names.py [workbook]")
    workbook = Path(sys.argv[1] if len(sys.argv) > 1 else "ozzit.xlsx")
    if not workbook.is_file():
        sys.exit(f"FAIL: no such workbook: {workbook}")
    try:
        changed = run(workbook)
    except (OSError, ValueError, zipfile.BadZipFile) as exc:
        sys.exit(f"FAIL: {exc}")
    if changed:
        print("applied sheet names to: workbook")
    else:
        print("sheet names already applied; no changes")
    print(f"OK: {workbook}")


if __name__ == "__main__":
    main()
