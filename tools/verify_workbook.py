"""Integrity checks for ozzit.xlsx.

Usage: python tools/verify_workbook.py [path/to/ozzit.xlsx]

Exits non-zero and prints every failure. Run by CI on each push.
"""
import base64
import html
import json
import re
import sys
import xml.etree.ElementTree as ET
import zipfile

# Every function name carries a λ, and a Windows console defaults to cp1252, which
# cannot encode it: without this the check dies printing its own result.
if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")


WORKBOOK = sys.argv[1] if len(sys.argv) > 1 else "ozzit.xlsx"

# Tokens that must never reappear: upstream branding, and any foreign tax content.
BANNED = re.compile(
    r"BX[DEFRLU]\.|BXLDebt|\bBXL\b|ozzit\.[a-z]+\.|beyondexcel|Eloquens|dropbox|Leonardo"
    r"|Starter Pack|Calibri|MACRS|Modified Accelerated|US GAAP|IRS Depreciation"
    # Excel stamps x15ac:absPath with the directory the file was last saved from, so
    # every release up to v2.2.0 published a path off the build machine. On a machine
    # whose account is not named "-" that path carries the account name with it.
    r"|x15ac:absPath|[A-Za-z]:\\{1,2}Users\\{1,2}|/Users/|/home/"
)
SHEET_RE = re.compile(r"xl/worksheets/sheet\d+\.xml$")
TOKEN_RE = re.compile(r"oz\.[A-Za-z0-9_]+λ?(?:DV)?")
FORBIDDEN_PARTS = (
    "xl/activeX/",
    "xl/ctrlProps/",
    "xl/externalLinks/",
    "xl/embeddings/",
    "xl/macrosheets/",
    "xl/dialogsheets/",
)

failures = []


def fail(msg):
    failures.append(msg)


def check_xml_part(name, data):
    """Reject declarations with entity expansion before parsing workbook XML."""
    if b"<!DOCTYPE" in data.upper():
        fail(f"DOCTYPE declaration in {name}")
        return
    try:
        ET.fromstring(data.decode("utf-8"))
    except (UnicodeDecodeError, ET.ParseError) as exc:
        fail(f"malformed XML in {name}: {exc}")


def main():
    z = zipfile.ZipFile(WORKBOOK)
    parts = z.namelist()

    for name in parts:
        if name == "xl/vbaProject.bin" or name.startswith(FORBIDDEN_PARTS):
            fail(f"forbidden workbook part {name}")

    for name in parts:
        if name.endswith((".xml", ".rels")):
            check_xml_part(name, z.read(name))

    for name in parts:
        if name.endswith((".xml", ".rels")) and name != "customXml/item1.xml":
            text = z.read(name).decode("utf-8")
            for hit in BANNED.finditer(text):
                fail(f"banned token {hit.group(0)!r} in {name}")
        elif name.endswith(".bin"):
            text = z.read(name).decode("utf-16-le", errors="ignore")
            for hit in BANNED.finditer(text):
                fail(f"banned token {hit.group(0)!r} in {name}")

    # The Advanced Formula Environment store holds the module sources as base64 UTF-16 JSON.
    afe = z.read("customXml/item1.xml").decode("utf-8")
    blob = re.search(r">([A-Za-z0-9+/=]{100,})<", afe)
    if not blob:
        fail("AFE project store missing")
        report(z)
        return
    store = json.loads(base64.b64decode(blob.group(1)).decode("utf-16-le"))
    store_text = json.dumps(store, ensure_ascii=False)
    for hit in BANNED.finditer(store_text):
        fail(f"banned token {hit.group(0)!r} in the AFE project store")

    workbook = z.read("xl/workbook.xml").decode("utf-8")
    defined = dict(re.findall(r'<definedName name="([^"]+)"[^>]*>([^<]*)</definedName>', workbook))
    sheets = set(re.findall(r'<sheet name="([^"]+)"', workbook))
    if not defined:
        fail("no defined names found")

    for name, body in defined.items():
        source = html.unescape(body)
        if source.count('"') % 2:
            fail(f"unbalanced quotes in {name}")
        # Help text is full of literal brackets, so only count parentheses outside string literals.
        code = re.sub(r'"(?:[^"]|"")*"', '""', source)
        if code.count("(") != code.count(")"):
            fail(f"unbalanced parentheses in {name}")
        for token in set(TOKEN_RE.findall(source)):
            if token not in defined and token not in sheets:
                fail(f"{name} references undefined {token}")

    for name in parts:
        if not SHEET_RE.match(name):
            continue
        text = z.read(name).decode("utf-8")
        if "#REF!" in text:
            fail(f"#REF! present in {name}")
        for formula in re.findall(r"<f[^>]*>([^<]*)</f>", text):
            for token in set(TOKEN_RE.findall(formula)):
                if token not in defined and token not in sheets:
                    fail(f"{name} uses undefined {token}")

    # Volatile formulas force the whole dependency chain to recalculate on every edit, which
    # is what makes a workbook this size feel slow on modest hardware. Only the sheet-name
    # titles, which use CELL(), are allowed to be volatile.
    # Only real formulas count: worksheet text quoting TODAY() as an example, and the
    # INDIRECT() that points a data-validation list at a table column, are not calculations.
    for part in parts:
        if not (SHEET_RE.match(part) or part.startswith("xl/tables/")):
            continue
        text = z.read(part).decode("utf-8")
        formulas = re.findall(r"<f[^>]*>(.*?)</f>", text, re.S)
        formulas += re.findall(r"<calculatedColumnFormula[^>]*>(.*?)</calculatedColumnFormula>", text, re.S)
        for formula in formulas:
            for volatile in ("RANDBETWEEN(", "RANDARRAY(", "RAND()", "NOW()", "TODAY()", "OFFSET(", "INDIRECT("):
                if volatile in formula:
                    fail(f"volatile {volatile} in a formula in {part}")
        for cell in re.findall(r"<c r=\"[A-Z]+\d+\"[^>]*>(?:(?!</c>).)*</c>", text, re.S):
            if ('ca="1"' in cell or 'aca="1"' in cell) and "CELL(" not in cell:
                fail(f"always-calculate flag on a non-volatile cell in {part}")
                break

    shared = ET.fromstring(z.read("xl/sharedStrings.xml").decode("utf-8"))
    declared = int(shared.get("uniqueCount"))
    if declared != len(list(shared)):
        fail(f"sharedStrings uniqueCount {declared} != {len(list(shared))} entries")

    # A reader has to see the right numbers, and there are two honest ways to get there.
    # Either the file forces a recalculation on load, which is what a workbook built purely
    # from XML needs because it has no formula engine to refresh what it edits, or its
    # cached values are already correct, which only Excel can produce and only
    # tools/verify_cache.py can confirm. Excel leaves a calculation chain behind when it
    # saves, so its absence together with no fullCalcOnLoad means neither is true: the file
    # would open showing whatever the build last left in it.
    if 'fullCalcOnLoad="1"' not in workbook and "xl/calcChain.xml" not in parts:
        fail("neither fullCalcOnLoad nor a calculation chain: this workbook would open "
             "showing cached values nothing has refreshed. Run tools/refresh_cache.py, or "
             "set fullCalcOnLoad")

    # docProps/app.xml repeats the sheet list, and nothing regenerates it: the build added
    # the Australian tax worksheet to five places and not to this one, so the part disagreed
    # with the workbook until Excel rewrote it. A reader's file properties dialogue and any
    # tool that trusts app.xml sees this list, not the real one.
    if "docProps/app.xml" in parts:
        app = z.read("docProps/app.xml").decode("utf-8")
        titles = re.search(r"<TitlesOfParts>(.*?)</TitlesOfParts>", app, re.S)
        listed = re.findall(r"<vt:lpstr>([^<]*)</vt:lpstr>", titles.group(1) if titles else "")
        sheets = re.findall(r'<sheet name="([^"]+)"', workbook)
        if listed[:len(sheets)] != sheets:
            fail("docProps/app.xml lists %d sheet titles that do not match the workbook's %d, "
                 "first difference at %r" % (len(listed), len(sheets),
                 next((a for a, b in zip(listed, sheets) if a != b), "the end")))

    # [MS-XLSX] requires each slicer cache to have a #N/A defined name of the same name.
    for part in parts:
        if not part.startswith("xl/slicerCaches/"):
            continue
        for cache in re.findall(r'<slicerCacheDefinition[^>]*name="([^"]+)"', z.read(part).decode("utf-8")):
            if f'<definedName name="{cache}">#N/A</definedName>' not in workbook:
                fail(f"slicer cache {cache} has no backing defined name")

    report(z, defined)


def report(z, defined=None):
    if failures:
        print(f"FAIL: {len(failures)} problem(s) in {WORKBOOK}")
        for item in failures:
            print(f"  - {item}")
        sys.exit(1)
    functions = sum(1 for name in defined if TOKEN_RE.fullmatch(name))
    print(f"OK: {WORKBOOK}, {functions} functions, {len(z.namelist())} parts")


if __name__ == "__main__":
    main()
