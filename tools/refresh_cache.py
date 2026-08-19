"""Recalculate the workbook in Excel, then tidy what Excel adds on the way out.

Usage: python tools/refresh_cache.py [path/to/ozzit.xlsx]

The build is pure zip and XML surgery with no formula engine, so every value it edits
leaves the cells downstream of it holding an answer their formulas no longer produce.
Shifting the sample dates forward two years left 3,193 such cells across 43 sheets. Excel
replaced them all on open, so no reader ever saw one, but a file that disagrees with itself
can only be checked by opening it in Excel.

Two steps, because they need different tools:

    tools/refresh_cache.ps1   opens the workbook in Excel, forces a full rebuild and saves
    this file                 removes the two things that save puts back

Excel re-marks a couple of cells always-calculate, which the build clears because a
workbook that recalculates cells nothing changed is slower for no gain. It also drops
fullCalcOnLoad, which is correct and deliberate: that flag exists to force a recalculation
past stale values, and after this there are none. The workbook then opens showing the right
numbers without recalculating, which also stops Excel asking to save a file the reader
never edited.

Run tools/verify_cache.py afterwards. That is the gate that proves this worked.
"""
import os
import re
import shutil
import subprocess
import sys
import zipfile

# Every progress line here can name a function, and a Windows console defaults to cp1252,
# which cannot encode the lambda in the names.
if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")

WORKBOOK = sys.argv[1] if len(sys.argv) > 1 else "ozzit.xlsx"
HERE = os.path.dirname(os.path.abspath(__file__))

# A cell may legitimately be marked always-calculate when it is genuinely volatile. Only
# the sheet-name titles are, and they are the only formulas in the workbook using CELL().
CELL_TAG = re.compile(r"<c\b(?:(?!</c>|<c\b).)*?</c>|<c\b[^>]*/>", re.S)


ABS_PATH_BLOCK = re.compile(
    r"<mc:AlternateContent[^>]*>\s*<mc:Choice Requires=\"x15\">\s*"
    r"<x15ac:absPath[^>]*/>\s*</mc:Choice>\s*</mc:AlternateContent>", re.S)
ABS_PATH = re.compile(r"<x15ac:absPath[^>]*/>")


def strip_local_path(text):
    """Remove the directory Excel stamps on the workbook every time it saves.

    x15ac:absPath records where the file sat when it was last written, which is a fact
    about the build machine and nothing to do with the library. Every release up to v2.2.0
    published one, and on a machine whose account is not named "-" it carries the account
    name with it. The whole mc:AlternateContent wrapper goes when the path is all it holds.
    """
    out, n = ABS_PATH_BLOCK.subn("", text)
    if not n:
        out, n = ABS_PATH.subn("", text)
    return out, n


def tidy(text):
    """Clear always-calculate flags Excel put back on cells that are not volatile."""
    cleared = [0]

    def one(m):
        cell = m.group(0)
        if ("CELL(" in cell) or not re.search(r'\b(a?ca)="1"', cell):
            return cell
        cleared[0] += 1
        return re.sub(r'\s\b(a?ca)="1"', "", cell)

    return CELL_TAG.sub(one, text), cleared[0]


def main():
    if not os.path.exists(WORKBOOK):
        print("FAIL: no such workbook: %s" % WORKBOOK)
        return 1

    script = os.path.join(HERE, "refresh_cache.ps1")
    proc = subprocess.run(
        ["powershell.exe", "-NoProfile", "-ExecutionPolicy", "Bypass", "-File", script,
         "-Path", os.path.abspath(WORKBOOK)],
        capture_output=True, text=True)
    out = (proc.stdout or "").strip()
    if out:
        print(out)
    if proc.returncode != 0:
        print("FAIL: the Excel refresh did not complete (exit %d)" % proc.returncode)
        if proc.stderr.strip():
            print(proc.stderr.strip()[:2000])
        return 1

    # Rewrite the parts in place. openpyxl is never used here for the same reason the build
    # avoids it: it would drop the extensions and the formula-environment store.
    with zipfile.ZipFile(WORKBOOK) as zin:
        parts = {n: zin.read(n) for n in zin.namelist()}

    cleared = 0
    for name in list(parts):
        if re.match(r"xl/worksheets/sheet\d+\.xml$", name):
            text, n = tidy(parts[name].decode("utf-8"))
            if n:
                parts[name] = text.encode("utf-8")
                cleared += n

    book, stripped = strip_local_path(parts["xl/workbook.xml"].decode("utf-8"))
    parts["xl/workbook.xml"] = book.encode("utf-8")

    tmp = WORKBOOK + ".tmp"
    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for name, blob in parts.items():
            zout.writestr(name, blob)
    shutil.move(tmp, WORKBOOK)

    forced = 'fullCalcOnLoad="1"' in parts["xl/workbook.xml"].decode("utf-8")
    print("cleared %d always-calculate flag(s) Excel put back and %d local build path(s); "
          "fullCalcOnLoad is %s"
          % (cleared, stripped,
             "still set" if forced else "off, which is what a refreshed file wants"))
    print("OK: %s holds the values its own formulas produce" % WORKBOOK)
    return 0


if __name__ == "__main__":
    sys.exit(main())
