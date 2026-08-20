"""Recalculate the workbook in Excel, then sanitise what Excel adds on save.

Usage: python tools/refresh_cache.py [path/to/ozzit.xlsx]

The build is pure zip and XML surgery with no formula engine, so every value it edits
leaves the cells downstream of it holding an answer their formulas no longer produce.
Shifting the sample dates forward two years left 3,193 such cells across 43 sheets. Excel
replaced them all on open, so no reader ever saw one, but a file that disagrees with itself
can only be checked by opening it in Excel.

Two steps, because they need different tools:

    tools/refresh_cache.ps1      opens the workbook in Excel, forces a full rebuild and saves
    tools/sanitise_workbook.py  removes everything that save puts back

Excel re-marks a couple of cells always-calculate, which the build clears because a
workbook that recalculates cells nothing changed is slower for no gain. It also drops
fullCalcOnLoad, which is correct and deliberate: that flag exists to force a recalculation
past stale values, and after this there are none. The workbook then opens showing the right
numbers without recalculating, which also stops Excel asking to save a file the reader
never edited.

Run tools/verify_cache.py afterwards. That is the gate that proves this worked.
"""
import subprocess
import sys
from pathlib import Path

from sanitise_workbook import sanitise

# Every progress line here can name a function, and a Windows console defaults to cp1252,
# which cannot encode the lambda in the names.
if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")

WORKBOOK = sys.argv[1] if len(sys.argv) > 1 else "ozzit.xlsx"
HERE = Path(__file__).resolve().parent


def main():
    workbook = Path(WORKBOOK)
    if not workbook.exists():
        print(f"FAIL: no such workbook: {workbook}")
        return 1

    script = HERE / "refresh_cache.ps1"
    proc = subprocess.run(
        ["powershell.exe", "-NoProfile", "-ExecutionPolicy", "Bypass", "-File", str(script),
         "-Path", str(workbook.resolve())],
        capture_output=True, text=True, check=False)
    out = (proc.stdout or "").strip()
    if out:
        print(out)
    if proc.returncode != 0:
        print(f"FAIL: the Excel refresh did not complete (exit {proc.returncode})")
        if proc.stderr.strip():
            print(proc.stderr.strip()[:2000])
        return 1

    for line in sanitise(workbook):
        print(line)
    print(f"OK: {workbook} holds the values its own formulas produce")
    return 0


if __name__ == "__main__":
    sys.exit(main())
