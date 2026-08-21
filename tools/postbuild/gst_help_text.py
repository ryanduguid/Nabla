"""GST help-text pass: legislative scope NOTES! in GSTAddλ and GSTExtractλ.

Usage: python tools/postbuild/gst_help_text.py [workbook] [src dir]

This pass starts from the committed ozzit.xlsx and src/ (the post-v3.0.0 input
recorded in ATTRIBUTION.md). It does not read the upstream workbook and does not
go through transform_from_upstream.py.

It inserts a NOTES! block after DESCRIPTION and before WEBPAGE in the two GST
helpers' inline help, PeriodDiffλ style, inside the existing quoted DESCRIPTION
string so the help stays one Excel string with ¶ rows. Anchors are help-only
(`DESCRIPTION:   →…¶"`): the same one-line blurbs also live in AboutFinancialλ
and in comment=, and those must not be rewritten.

Every swap carries an asserted hit count in both workbook.xml and src/. A second
run reports "already applied" and writes nothing. A store whose anchors do not
match fails loudly instead of leaving the workbook and source views out of sync.

Pure text surgery: no COM, no recalculation.
"""

from __future__ import annotations

import sys
import zipfile
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from sanitise_workbook import write_deterministic

MODULES = ("Dates", "Essentials", "Financial", "Ratios", "Utilities", "Debt")

# 15-character labels, matching PeriodDiffλ.
NOTES_LABEL = "NOTES!         →"
CONT_LABEL = "               →"
assert len(NOTES_LABEL) - 1 == 15, "NOTES! label must pad to 15 characters"
assert len(CONT_LABEL) - 1 == 15, "continuation label must pad to 15 characters"

# Short Australian scope note, not advice. No raw "&" (workbook.xml would see it
# as the start of an entity); write "and" and "ss 9-70". Keep the first published
# branch version as an accepted migration state so the pass can repair that
# version as well as the pre-note committed input.
LEGACY_NOTE_ROWS = (
    "The 10% default and one-eleventh extraction reflect the basic rule",
    "for a taxable supply in A New Tax System (Goods and Services Tax)",
    "Act 1999 ss 9-70 and 9-75 (source checked 20 August 2026).",
    "Recheck the current Act. These helpers apply arithmetic only: they",
    "do not decide whether a supply is taxable, GST-free, input taxed,",
    "or subject to a special rule.",
)
NOTE_ROWS = LEGACY_NOTE_ROWS[:-1] + (
    "outside the GST system, or subject to a special rule.",
)
assert all("&" not in row for row in NOTE_ROWS), (
    "raw & is illegal in workbook.xml text; write 'and'"
)


def _notes(rows: tuple[str, ...]) -> str:
    return NOTES_LABEL + rows[0] + "¶" + "".join(
        CONT_LABEL + row + "¶" for row in rows[1:]
    )


LEGACY_NOTES = _notes(LEGACY_NOTE_ROWS)
NOTES = _notes(NOTE_ROWS)

# (old, new, workbook hits, src hits). The closing quote is part of the anchor
# so old is not a prefix of new: a second run cannot insert twice.
SWAPS = [
    (
        'DESCRIPTION:   →Adds GST to one or more GST-exclusive amounts.¶"',
        'DESCRIPTION:   →Adds GST to one or more GST-exclusive amounts.¶' + NOTES + '"',
        1,
        1,
    ),
    (
        'DESCRIPTION:   →Returns the GST contained in one or more GST-inclusive amounts.¶"',
        'DESCRIPTION:   →Returns the GST contained in one or more GST-inclusive amounts.¶'
        + NOTES
        + '"',
        1,
        1,
    ),
]


def _legacy_replacement(old: str) -> str:
    return old[:-1] + LEGACY_NOTES + '"'


def validate_store(text: str, label: str, failures: list[str]) -> None:
    """Require every target exactly once in this store in a recognised state."""
    for old, new, wb_expected, src_expected in SWAPS:
        expected = wb_expected if label == "workbook" else src_expected
        legacy = _legacy_replacement(old)
        counts = {
            "pre-note": text.count(old),
            "legacy-note": text.count(legacy),
            "current-note": text.count(new),
        }
        hits = sum(counts.values())
        if hits != expected:
            states = ", ".join(f"{state}={count}" for state, count in counts.items())
            failures.append(
                f"{label}: {old[:50]!r} expected {expected} recognised anchor, "
                f"got {hits} ({states})"
            )


def apply_swaps(text: str) -> str:
    for old, new, _wb_expected, _src_expected in SWAPS:
        text = text.replace(old, new)
        text = text.replace(_legacy_replacement(old), new)
    return text


def _read_text(path: Path) -> str:
    with open(path, encoding="utf-8", newline="") as handle:
        return handle.read()


def _write_text(path: Path, text: str) -> None:
    with open(path, "w", encoding="utf-8", newline="") as handle:
        handle.write(text)


def run(workbook: Path, src_dir: Path) -> list[str]:
    failures: list[str] = []

    src_originals = {}
    for module in MODULES:
        path = src_dir / f"{module}.txt"
        if not path.is_file():
            failures.append(f"src: missing {path}")
            continue
        original = _read_text(path)
        src_originals[module] = original

    with zipfile.ZipFile(workbook) as archive:
        parts = {n: archive.read(n) for n in archive.namelist()}
    book = parts["xl/workbook.xml"].decode("utf-8")
    validate_store(book, "workbook", failures)
    validate_store("\n".join(src_originals.values()), "src", failures)

    if failures:
        raise ValueError("; ".join(failures))

    parts["xl/workbook.xml"] = apply_swaps(book).encode("utf-8")
    src_texts = {
        module: apply_swaps(original) for module, original in src_originals.items()
    }

    changed = []
    with zipfile.ZipFile(workbook) as archive:
        original_book = archive.read("xl/workbook.xml")
    if parts["xl/workbook.xml"] != original_book:
        write_deterministic(workbook, parts)
        changed.append("workbook")
    for module, text in src_texts.items():
        path = src_dir / f"{module}.txt"
        if _read_text(path) != text:
            _write_text(path, text)
            changed.append(f"src/{module}.txt")

    return changed


def main() -> None:
    if len(sys.argv) > 3:
        sys.exit("FAIL: usage: python tools/postbuild/gst_help_text.py [workbook] [src dir]")
    workbook = Path(sys.argv[1] if len(sys.argv) > 1 else "ozzit.xlsx")
    src_dir = Path(sys.argv[2] if len(sys.argv) > 2 else "src")
    if not workbook.is_file():
        sys.exit(f"FAIL: no such workbook: {workbook}")
    try:
        changed = run(workbook, src_dir)
    except (OSError, ValueError, zipfile.BadZipFile) as exc:
        sys.exit(f"FAIL: {exc}")
    if changed:
        print(f"applied GST help text to: {', '.join(changed)}")
    else:
        print("GST help text already applied; no changes")
    print(f"OK: {workbook}")


if __name__ == "__main__":
    main()
