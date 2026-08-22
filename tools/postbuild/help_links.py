"""Help-link pass: repoint two WEBPAGE rows copied from a neighbouring function.

Usage: python tools/postbuild/help_links.py [workbook] [src dir]

CurrentRatioλ shipped CashRatioλ's cash-ratio article and ROIλ shipped ROEλ's
return-on-equity article. Each swap anchors on the preceding formula row of the
help so the neighbour that legitimately owns the old URL is untouched. Applies
to the defined names in xl/workbook.xml and to src/Ratios.txt; the ratio
functions have no dedicated worksheets, so there are no cached spills to
refresh.

Every swap carries an asserted hit count in both stores. A second run reports
"already applied" and writes nothing. Pure text surgery: no COM, no
recalculation.
"""

from __future__ import annotations

import sys
import zipfile
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from sanitise_workbook import write_deterministic

# (context prefix, old URL, new URL). The prefix is the row before WEBPAGE plus
# the label, which is unique to the owning function; the neighbour keeping the
# old URL never matches because its DESCRIPTION rows differ.
SWAPS = [
    (
        "(Current Assets/Liabilities)¶\" {AMP} \"WEBPAGE:       →",
        "https://www.investopedia.com/terms/c/cash-ratio.asp",
        "https://www.investopedia.com/terms/c/currentratio.asp",
    ),
    (
        "(Net Return on Investment/Cost of Investment)¶\" {AMP} \"WEBPAGE:               →",
        "https://www.investopedia.com/terms/r/returnonequity.asp",
        "https://www.investopedia.com/terms/r/returnoninvestment.asp",
    ),
]


def _forms(store: str) -> list[tuple[str, str]]:
    """Return (old, new) anchor pairs in this store's concatenation form."""
    amp = "&amp;" if store == "workbook" else "&"
    pairs = []
    for prefix, old_url, new_url in SWAPS:
        p = prefix.replace("{AMP}", amp)
        if store == "src":
            # src wraps the concatenation across lines after the ampersand.
            p = p.replace('¶" & "', '¶" & \n                        "')
        pairs.append((p + old_url, p + new_url))
    return pairs


def validate_store(text: str, store: str, failures: list[str]) -> None:
    for old, new in _forms(store):
        counts = {"wrong-link": text.count(old), "fixed-link": text.count(new)}
        if sum(counts.values()) != 1:
            failures.append(
                f"{store}: {old[-60:]!r} expected exactly 1 recognised anchor, "
                f"got wrong-link={counts['wrong-link']} fixed-link={counts['fixed-link']}"
            )


def apply_swaps(text: str, store: str) -> str:
    for old, new in _forms(store):
        text = text.replace(old, new)
    return text


def _read_text(path: Path) -> str:
    with open(path, encoding="utf-8", newline="") as handle:
        return handle.read()


def run(workbook: Path, src_dir: Path) -> list[str]:
    failures: list[str] = []
    ratios_path = src_dir / "Ratios.txt"
    if not ratios_path.is_file():
        raise ValueError(f"src: missing {ratios_path}")
    ratios = _read_text(ratios_path)

    with zipfile.ZipFile(workbook) as archive:
        parts = {n: archive.read(n) for n in archive.namelist()}
    book = parts["xl/workbook.xml"].decode("utf-8")

    validate_store(book, "workbook", failures)
    validate_store(ratios, "src", failures)
    if failures:
        raise ValueError("; ".join(failures))

    changed = []
    new_book = apply_swaps(book, "workbook")
    if new_book != book:
        parts["xl/workbook.xml"] = new_book.encode("utf-8")
        write_deterministic(workbook, parts)
        changed.append("workbook")
    new_ratios = apply_swaps(ratios, "src")
    if new_ratios != ratios:
        with open(ratios_path, "w", encoding="utf-8", newline="") as handle:
            handle.write(new_ratios)
        changed.append("src/Ratios.txt")
    return changed


def main() -> None:
    if len(sys.argv) > 3:
        sys.exit("FAIL: usage: python tools/postbuild/help_links.py [workbook] [src dir]")
    workbook = Path(sys.argv[1] if len(sys.argv) > 1 else "ozzit.xlsx")
    src_dir = Path(sys.argv[2] if len(sys.argv) > 2 else "src")
    if not workbook.is_file():
        sys.exit(f"FAIL: no such workbook: {workbook}")
    try:
        changed = run(workbook, src_dir)
    except (OSError, ValueError, zipfile.BadZipFile) as exc:
        sys.exit(f"FAIL: {exc}")
    if changed:
        print(f"applied help links to: {', '.join(changed)}")
    else:
        print("help links already applied; no changes")
    print(f"OK: {workbook}")


if __name__ == "__main__":
    main()
