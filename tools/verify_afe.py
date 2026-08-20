"""Check that the Advanced Formula Environment store matches what ships.

Usage: python tools/verify_afe.py [workbook] [src dir]

The AFE task pane is the documented editing surface, but the existing gates only
compare src/ with defined names and never compare the AFE store. Debt is
deliberately excluded: the module is recursive, and src/Debt.txt documents that
Excel Labs cannot hold it, so a compliant store must omit it.
"""

import base64
import html
import json
import re
import sys
import zipfile
from pathlib import Path

WORKBOOK = sys.argv[1] if len(sys.argv) > 1 else "ozzit.xlsx"
SRC = Path(sys.argv[2] if len(sys.argv) > 2 else "src")
SCHEMA = "http://schemas.advancedformulaenvironment.officeapps.live.com/afeprojects/0.2"
MODULES = ("Dates", "Essentials", "Financial", "Ratios", "Utilities")
DEBT_NAMES = {
    "oz.AmortiseBλ",
    "oz.DebtSculptFixedλ",
    "oz.DebtSculptVariableλ",
    "oz.DebtSculptVariableLRVλ",
    "oz.InterestLRVλ",
}
AFE_BLOB_RE = rb"[A-Za-z0-9+/]{100,}={0,2}"


def find_afe_blob(raw: bytes) -> bytes:
    """Return the longest base64 payload in the AFE store XML."""
    matches = re.findall(AFE_BLOB_RE, raw)
    if not matches:
        raise ValueError("AFE store has no base64 payload")
    return max(matches, key=len)


def fail(failures, message):
    failures.append(message)


def decode_store(workbook: Path):
    try:
        with zipfile.ZipFile(workbook) as archive:
            raw = archive.read("customXml/item1.xml")
    except (FileNotFoundError, KeyError, zipfile.BadZipFile) as exc:
        raise ValueError(f"cannot read AFE store: {exc}") from exc
    try:
        encoded = find_afe_blob(raw)
    except ValueError as exc:
        raise ValueError(str(exc)) from exc
    try:
        return json.loads(base64.b64decode(encoded).decode("utf-16-le"))
    except (ValueError, UnicodeDecodeError) as exc:
        raise ValueError(f"cannot decode AFE store: {exc}") from exc


def workbook_names(workbook: Path):
    with zipfile.ZipFile(workbook) as archive:
        text = archive.read("xl/workbook.xml").decode("utf-8")
    names = set()
    for raw in re.findall(r'<definedName\b[^>]*\bname="([^"]+)"[^>]*>', text):
        name = html.unescape(raw)
        if name.startswith("oz."):
            names.add(name)
    return names


def check(workbook: Path, src: Path) -> list[str]:
    failures = []
    try:
        store = decode_store(workbook)
    except ValueError as exc:
        return [str(exc)]

    if store.get("schema") != SCHEMA:
        fail(failures, f"AFE schema is {store.get('schema')!r}")

    files = {item.get("path"): item.get("text") for item in store.get("files", [])}
    for module in MODULES:
        path = f"/projects/{module}"
        expected_path = src / f"{module}.txt"
        if path not in files:
            fail(failures, f"AFE module {path} missing")
            continue
        try:
            expected = expected_path.read_text(encoding="utf-8")
        except FileNotFoundError:
            fail(failures, f"source module {expected_path} missing")
            continue
        if files[path] != expected:
            first = next(
                (
                    index
                    for index, pair in enumerate(zip(files[path], expected, strict=False))
                    if pair[0] != pair[1]
                ),
                min(len(files[path]), len(expected)),
            )
            actual = files[path][max(0, first - 30) : first + 50].replace("\n", " ")
            wanted = expected[max(0, first - 30) : first + 50].replace("\n", " ")
            fail(
                failures,
                f"AFE module {path} differs from src/{module}.txt at character {first}: "
                f"{actual!r} != {wanted!r}",
            )

    shipped = workbook_names(workbook)
    expected_names = shipped - DEBT_NAMES
    actual_names = set(store.get("projectNames", []))
    for name in sorted(actual_names & DEBT_NAMES):
        fail(failures, f"recursive Debt function {name} is present in the AFE store")
    for name in sorted(expected_names - actual_names):
        fail(failures, f"AFE projectNames is missing shipped non-Debt name {name}")
    for name in sorted(actual_names - expected_names - DEBT_NAMES):
        fail(failures, f"AFE projectNames has unknown name {name}")

    return failures


def main() -> None:
    failures = check(Path(WORKBOOK), SRC)
    if failures:
        print(f"FAIL: {len(failures)} AFE problem(s) in {WORKBOOK}")
        for failure in failures:
            print(f"  - {failure}")
        sys.exit(1)
    print(
        f"OK: AFE store in {WORKBOOK} matches {len(MODULES)} non-recursive modules "
        "and excludes recursive Debt"
    )


if __name__ == "__main__":
    main()
