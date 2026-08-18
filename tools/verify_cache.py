"""Check that the workbook's cached values are the values its formulas produce.

Usage: python tools/verify_cache.py [path/to/nabla.xlsx]

An .xlsx stores two things for every calculated cell: the formula, and the answer Excel
last got from it. Nothing keeps them in step. The build edits values as XML with no formula
engine, so every cell downstream of an edit keeps the answer it had before: shifting the
sample dates forward two years left 3,193 cells across 43 sheets holding numbers their own
formulas no longer produce, and five cells shipped a saved #VALUE! from v1.2.0 to v2.2.0.

None of that was ever visible in Excel, which recalculates and quietly replaces the lot.
That is exactly what makes it worth a gate: the file can be wrong in a way only a second
tool can see, and everything that reads an .xlsx without a formula engine, from a diff to a
converter to a web preview, reads the cached answer.

Needs Excel, so this is a local gate like the arithmetic one, not a CI gate.
"""
import html
import os
import re
import subprocess
import sys
import tempfile
import zipfile

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")

WORKBOOK = sys.argv[1] if len(sys.argv) > 1 else "nabla.xlsx"
HERE = os.path.dirname(os.path.abspath(__file__))

# Below this the run proved nothing and the pass would be vacuous.
FLOOR = 15000

# Excel hands an error cell back over COM as its error code, not its text.
ERRORS = {-2146826288: "#NULL!", -2146826281: "#DIV/0!", -2146826273: "#VALUE!",
          -2146826265: "#REF!", -2146826259: "#NAME?", -2146826252: "#NUM!",
          -2146826246: "#N/A", -2146826245: "#GETTING_DATA"}

CELL = re.compile(r'<c r="([A-Z]+)(\d+)"([^>]*)>((?:(?!<c[ /]).)*?)</c>', re.S)


def column(letters):
    n = 0
    for ch in letters:
        n = n * 26 + ord(ch) - 64
    return n


def cached_values(path):
    """Every cached value in the workbook, keyed by (tab position, row, column).

    Keyed by position rather than by part name because Excel renumbers the sheet parts
    when it saves, so sheet4.xml is not the fourth sheet in a file it has written.
    """
    z = zipfile.ZipFile(path)
    book = z.read("xl/workbook.xml").decode("utf-8")
    rels = dict(re.findall(r'Id="([^"]+)"[^>]*Target="([^"]+)"',
                           z.read("xl/_rels/workbook.xml.rels").decode("utf-8")))
    order = [(m.group(1), rels[m.group(2)]) for m in
             re.finditer(r'<sheet name="([^"]+)"[^>]*r:id="([^"]+)"', book)]

    shared = []
    if "xl/sharedStrings.xml" in z.namelist():
        for si in re.findall(r"<si>(.*?)</si>", z.read("xl/sharedStrings.xml").decode("utf-8"), re.S):
            shared.append(html.unescape(re.sub(r"<.*?>", "", si)))

    out, names = {}, {}
    for pos, (name, target) in enumerate(order, start=1):
        names[pos] = name
        text = z.read("xl/" + target.lstrip("/")).decode("utf-8")
        for m in CELL.finditer(text):
            attrs, inner = m.group(3), m.group(4)
            kind = re.search(r't="([^"]+)"', attrs)
            kind = kind.group(1) if kind else "n"
            if kind == "inlineStr":
                got = re.search(r"<t[^>]*>(.*?)</t>", inner, re.S)
                value = html.unescape(got.group(1)) if got else None
            else:
                got = re.search(r"<v>(.*?)</v>", inner, re.S)
                if not got:
                    continue
                raw = html.unescape(got.group(1))
                if kind == "s":
                    value = shared[int(raw)]
                elif kind == "b":
                    value = "True" if raw == "1" else "False"
                else:
                    value = raw
            if value is not None:
                out[(pos, int(m.group(2)), column(m.group(1)))] = value
    return out, names


def fold(text):
    """One cell must be one line, the way the dumper writes it."""
    marker = chr(92) + "n"          # the two characters backslash and n
    return (text.replace(chr(13) + chr(10), marker)
                .replace(chr(10), marker)
                .replace(chr(13), marker))


def agree(cached, live):
    if fold(cached) == live:
        return True
    try:
        a, b = float(cached), float(live)
        return abs(a - b) <= max(1e-9, 1e-9 * max(abs(a), abs(b)))
    except ValueError:
        pass
    try:                                   # a saved error against the code Excel returns
        return ERRORS.get(int(float(live))) == cached
    except ValueError:
        return False


def main():
    if not os.path.exists(WORKBOOK):
        print("FAIL: no such workbook: %s" % WORKBOOK)
        return 1

    handle, dump = tempfile.mkstemp(suffix=".tsv", prefix="nabla-cache-")
    os.close(handle)
    try:
        proc = subprocess.run(
            ["powershell.exe", "-NoProfile", "-ExecutionPolicy", "Bypass",
             "-File", os.path.join(HERE, "dump_values.ps1"),
             "-Path", os.path.abspath(WORKBOOK), "-Out", dump],
            capture_output=True, text=True)
        if proc.returncode != 0:
            print("FAIL: could not read the workbook in Excel (exit %d)" % proc.returncode)
            print((proc.stdout or proc.stderr or "").strip()[:2000])
            return 1

        live = {}
        with open(dump, encoding="utf-8-sig") as fh:
            for line in fh:
                bits = line.rstrip("\n").split("\t", 3)
                if len(bits) == 4:
                    live[(int(bits[0]), int(bits[1]), int(bits[2]))] = bits[3]
    finally:
        try:
            os.unlink(dump)
        except OSError:
            pass

    cached, names = cached_values(WORKBOOK)
    both = sorted(set(cached) & set(live))
    stale = [(k, cached[k], live[k]) for k in both if not agree(cached[k], live[k])]

    if len(both) < FLOOR:
        print("FAIL: only %d cells carried both a cached and a calculated value, under the "
              "floor of %d. Something skipped most of the workbook rather than checking it."
              % (len(both), FLOOR))
        return 1

    if stale:
        print("FAIL: %d of %d cached values are not what the formula produces, on %d sheet(s)"
              % (len(stale), len(both), len({k[0] for k, _, _ in stale})))
        for (pos, row, col), was, now in stale[:20]:
            print("  - %s row %d col %d: cached %s, calculates to %s"
                  % (names[pos], row, col, str(was)[:40], str(now)[:40]))
        if len(stale) > 20:
            print("  ... and %d more" % (len(stale) - 20))
        print("Run: python tools/refresh_cache.py")
        return 1

    print("OK: all %d cached values in %s are the values their formulas produce"
          % (len(both), WORKBOOK))
    return 0


if __name__ == "__main__":
    sys.exit(main())
