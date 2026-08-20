"""Check that functions.csv is an index of the workbook and src/ that ship.

Usage: python tools/verify_index.py [workbook] [src dir] [functions.csv]

The previous-name gate proves migration history. This gate proves the other
published fields describe the current artefacts rather than a stale or
hand-edited index.
"""

import csv
import html
import re
import sys
import zipfile
from pathlib import Path

from verify_sources import NAME, statements

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")

WORKBOOK = Path(sys.argv[1] if len(sys.argv) > 1 else "ozzit.xlsx")
SRC = Path(sys.argv[2] if len(sys.argv) > 2 else "src")
INDEX = Path(sys.argv[3] if len(sys.argv) > 3 else "functions.csv")

NAME_RE = re.compile(
    r'<definedName name="(oz\.[^"]+)"([^>]*)>(.*?)</definedName>', re.DOTALL
)
SIG_RE = re.compile(r"FUNCTION:\s*→?\s*(.*?)¶")
DESC_RE = re.compile(r"DESCRIPTION:\s*→(.*?)¶")
ROW_RE = re.compile(r"→(.*?)¶")


def index_fields(attrs, body, bare):
    """Derive the signature and description exported by the build."""
    signature = SIG_RE.search(body)
    text, end = (
        (signature.group(1).strip(), signature.end()) if signature else (bare, 0)
    )
    while text.count("(") > text.count(")"):
        more = ROW_RE.search(body, end)
        if not more:
            break
        text = f"{text} {more.group(1).strip()}".strip()
        end = more.end()

    comment = re.search(r'comment="([^"]*)"', attrs)
    fallback = DESC_RE.search(body)
    blurb = (
        html.unescape(comment.group(1))
        if comment
        else (fallback.group(1).strip() if fallback else "")
    )
    blurb = (
        blurb.replace("_x000a_", " ")
        .replace("_x000D_", " ")
        .lstrip("→")
        .strip()
    )
    return text, re.sub(r"\s{2,}", " ", blurb)


def source_modules(src):
    """Map each qualified function to the source module declaring it."""
    modules = {}
    for path in sorted(src.glob("*.txt")):
        for statement in statements(path.read_text(encoding="utf-8")):
            match = NAME.match(statement)
            if match:
                modules[f"oz.{match.group(1)}"] = path.stem
    return modules


def main():
    failures = []
    try:
        with zipfile.ZipFile(WORKBOOK) as archive:
            workbook = archive.read("xl/workbook.xml").decode("utf-8")
        with INDEX.open(encoding="utf-8-sig", newline="") as source:
            rows = list(csv.DictReader(source))
    except (OSError, KeyError, UnicodeDecodeError, zipfile.BadZipFile) as exc:
        print(f"FAIL: cannot read index inputs: {exc}")
        return 1

    expected = {}
    modules = source_modules(SRC)
    for name, attrs, encoded in NAME_RE.findall(workbook):
        body = html.unescape(encoded)
        signature, description = index_fields(attrs, body, name.rsplit(".", 1)[1])
        expected[name] = {
            "module": modules.get(name, ""),
            "signature": signature,
            "description": description,
        }

    actual = {row.get("function", ""): row for row in rows}
    for name in sorted(set(expected) - set(actual)):
        failures.append(f"{name} is in the workbook but missing from functions.csv")
    for name in sorted(set(actual) - set(expected)):
        failures.append(f"{name} is in functions.csv but missing from the workbook")

    for name in sorted(set(expected) & set(actual)):
        for field in ("module", "signature", "description"):
            if actual[name].get(field, "") != expected[name][field]:
                source = "src/" if field == "module" else "workbook"
                failures.append(f"{name}: {field} differs from {source}")

    if failures:
        print(f"FAIL: {len(failures)} functions.csv problem(s)")
        for item in failures:
            print(f"  - {item}")
        return 1
    print(
        f"OK: all {len(expected)} functions.csv rows match the workbook and source modules"
    )
    return 0


if __name__ == "__main__":
    sys.exit(main())
