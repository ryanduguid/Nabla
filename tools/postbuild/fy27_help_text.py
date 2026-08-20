"""FY27 help-text pass: dated examples in defined names and src/.

Usage: python tools/postbuild/fy27_help_text.py [workbook] [src dir]

This is the text layer of the v3.1.0 release: every worked example's start date
moves to 1 July 2026 (the start of FY27), plus the caption corrections proven in
Excel. Every swap is length-preserving and carries an asserted hit count, so a
second run reports "already applied" and changes nothing, and a workbook whose
anchors do not match fails loudly instead of silently corrupting.

Pure text surgery: no COM, no recalculation. The workbook's cached values are
refreshed separately by the FY27 date pass (tools/postbuild/fy27_dates.ps1).
"""

from __future__ import annotations

import sys
import zipfile
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from sanitise_workbook import write_deterministic

MODULES = ("Dates", "Essentials", "Financial", "Ratios", "Utilities", "Debt")

# (old, new, workbook hits, src hits) — both stores sit inside string literals,
# so doubled quotes appear in both. Counts are from the v3.0.0 -> v3.1.0 commit.
SWAPS = [
    ('SEQUENCE(,90,""1/1/2025"")', 'SEQUENCE(,90,""1/7/2026"")', 3, 3),
    ('""1/1/2025"", ""13/3/2025""', '""1/7/2026"", ""10/9/2026""', 3, 3),
    ('""17/1/2025"", ""25/1/2025"", ""7/1/2025"", ""18/1/2025""',
     '""17/7/2026"", ""25/7/2026"", ""7/7/2026"", ""18/7/2026""', 1, 1),
    ('DATEVALUE(""26/2/2025"")', 'DATEVALUE(""26/8/2026"")', 4, 4),
    ('2025-Feb-26    →', '2026-Aug-26    →', 1, 1),
    ('2023-Feb       →', '2026-Aug       →', 1, 1),
    ('2025:Q1       →', '2026:Q3       →', 1, 1),
    ('2023           →=oz.PeriodLabelλ', '2026           →=oz.PeriodLabelλ', 1, 1),
    ('""31/3/2025"", ""15/5/2025""', '""31/7/2026"", ""15/9/2026""', 1, 1),
    ('""15/5/2025"", ""31/3/2025""', '""15/9/2026"", ""31/7/2026""', 1, 1),
    ('""15/1/2026"", ""16/1/2025""', '""15/7/2027"", ""16/7/2026""', 2, 2),
    ('"-53            →=oz.Periodsλ', '"-52            →=oz.Periodsλ', 1, 1),
    ('Timelineλ(""1/1/2025"",5,""W""),{45662; 45670; 45676}',
     'Timelineλ(""1/7/2026"",5,""W""),{46208; 46216; 46222}', 1, 1),
    ('Timelineλ( ""1/1/2025"")', 'Timelineλ( ""1/7/2026"")', 2, 2),
    ('Timelineλ( ""1/1/2025"",,,FALSE)', 'Timelineλ( ""1/7/2026"",,,FALSE)', 1, 1),
    ('Timelineλ(""1/1/2025""), Timelineλ(""1/1/2025"",,,FALSE), {1;3}, {""1/1/2025""; ""15/2/2025""}',
     'Timelineλ(""1/7/2026""), Timelineλ(""1/7/2026"",,,FALSE), {1;3}, {""1/7/2026""; ""15/8/2026""}', 1, 1),
    ('""22/3/2012"", ""10/4/2012""', '""23/7/2026"", ""11/8/2026""', 1, 1),
    ('EDATE(""1/1/2025"", SEQUENCE(,24, 0))', 'EDATE(""1/7/2026"", SEQUENCE(,24, 0))', 1, 1),
    ('"24%            →=oz.IRRλ', '"25%            →=oz.IRRλ', 1, 1),
    ('""15/2/2026"", EDATE(""1/1/2026""', '""15/8/2026"", EDATE(""1/7/2026""', 1, 1),
    ('""15/2/2025"", EDATE(""1/1/2026""', '""15/8/2025"", EDATE(""1/7/2026""', 1, 1),
    ('EDATE(""1/1/2025"", SEQUENCE( , 12, 0)), 12)', 'EDATE(""1/7/2026"", SEQUENCE( , 12, 0)), 12)', 1, 1),
    ('""2025-01-01""', '""2026-07-01""', 1, 1),
    ('""2026-01-01"", 3, ""2026-04-15""', '""2026-07-01"", 3, ""2026-10-15""', 1, 1),
    ('2026-04-01     →=oz.PeriodStartλ', '2026-10-01     →=oz.PeriodStartλ', 1, 1),
    ('VERSION:       →19 Aug 2026', 'VERSION:       →20 Aug 2026', 1, 1),
    ('VERSION:           →19 Aug 2026', 'VERSION:           →20 Aug 2026', 1, 1),
]


def apply_swaps(text: str, label: str, failures: list[str]) -> str:
    for old, new, wb_expected, src_expected in SWAPS:
        assert len(old) == len(new), f"length drift: {old[:40]!r}"
        hits = text.count(old)
        expected = wb_expected if label == "workbook" else src_expected
        if hits == 0:
            # Already applied (or never present): nothing to do. The replacement's
            # own count is not a valid check on a second run, because the first run
            # produced it at whatever total the store genuinely carries.
            continue
        if hits != expected:
            failures.append(f"{label}: {old[:50]!r} expected {expected} hits, got {hits}")
            continue
        text = text.replace(old, new)
    return text


def _read_text(path: Path) -> str:
    with open(path, encoding="utf-8", newline="") as handle:
        return handle.read()


def _write_text(path: Path, text: str) -> None:
    with open(path, "w", encoding="utf-8", newline="") as handle:
        handle.write(text)


def run(workbook: Path, src_dir: Path) -> list[str]:
    failures: list[str] = []

    src_texts = {}
    for module in MODULES:
        path = src_dir / f"{module}.txt"
        if not path.is_file():
            failures.append(f"src: missing {path}")
            continue
        src_texts[module] = apply_swaps(_read_text(path), "src", failures)

    with zipfile.ZipFile(workbook) as archive:
        parts = {n: archive.read(n) for n in archive.namelist()}
    book = parts["xl/workbook.xml"].decode("utf-8")
    parts["xl/workbook.xml"] = apply_swaps(book, "workbook", failures).encode("utf-8")

    if failures:
        raise ValueError("; ".join(failures))

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
        sys.exit("FAIL: usage: python tools/postbuild/fy27_help_text.py [workbook] [src dir]")
    workbook = Path(sys.argv[1] if len(sys.argv) > 1 else "ozzit.xlsx")
    src_dir = Path(sys.argv[2] if len(sys.argv) > 2 else "src")
    if not workbook.is_file():
        sys.exit(f"FAIL: no such workbook: {workbook}")
    try:
        changed = run(workbook, src_dir)
    except (OSError, ValueError, zipfile.BadZipFile) as exc:
        sys.exit(f"FAIL: {exc}")
    if changed:
        print(f"applied FY27 help text to: {', '.join(changed)}")
    else:
        print("FY27 help text already applied; no changes")
    print(f"OK: {workbook}")


if __name__ == "__main__":
    main()
