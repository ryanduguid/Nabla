"""Synchronise the Advanced Formula Environment store with src/.

Usage: python tools/sync_afe_store.py [workbook] [src dir]

Only the five non-recursive modules that AFE can hold are rewritten. Debt stays
out by design, because src/Debt.txt documents that Excel Labs cannot hold its
recursive functions. The pass preserves the store's schema, locale, project
name order and JSON layout, re-encodes UTF-16LE/base64, and writes through the
canonical archive writer.
"""

from __future__ import annotations

import base64
import json
import re
import sys
import zipfile
from pathlib import Path

from sanitise_workbook import write_deterministic
from verify_afe import DEBT_NAMES

MODULES = ("Dates", "Essentials", "Financial", "Ratios", "Utilities")


def sync(workbook: Path, src: Path) -> list[str]:
    with zipfile.ZipFile(workbook) as archive:
        parts = {name: archive.read(name) for name in archive.namelist()}
    raw = parts["customXml/item1.xml"]
    matches = re.findall(rb"[A-Za-z0-9+/]{100,}={0,2}", raw)
    encoded = max(matches, key=len)
    store = json.loads(base64.b64decode(encoded).decode("utf-16-le"))

    files = {item.get("path"): item for item in store.get("files", [])}
    changes = []
    for module in MODULES:
        expected = (src / f"{module}.txt").read_text(encoding="utf-8")
        path = f"/projects/{module}"
        if path not in files:
            item = {"path": path, "text": expected}
            store["files"].append(item)
            files[path] = item
            changes.append(f"added AFE module {path}")
        elif files[path].get("text") != expected:
            files[path]["text"] = expected
            changes.append(f"synced AFE module {path}")

    if "/projects/Debt" in files or any(
        name in store.get("projectNames", []) for name in DEBT_NAMES
    ):
        raise ValueError("recursive Debt module must not be stored in AFE")

    replacement = base64.b64encode(
        json.dumps(store, ensure_ascii=False, separators=(",", ":")).encode("utf-16-le")
    )
    if replacement == encoded and not changes:
        return ["AFE store already matches src/"]
    parts["customXml/item1.xml"] = raw.replace(encoded, replacement)
    write_deterministic(workbook, parts)
    return changes


def main() -> None:
    if len(sys.argv) > 3:
        sys.exit("FAIL: usage: python tools/sync_afe_store.py [workbook] [src dir]")
    workbook = Path(sys.argv[1] if len(sys.argv) > 1 else "ozzit.xlsx")
    src = Path(sys.argv[2] if len(sys.argv) > 2 else "src")
    if not workbook.is_file():
        sys.exit(f"FAIL: no such workbook: {workbook}")
    try:
        changes = sync(workbook, src)
    except (OSError, KeyError, ValueError, json.JSONDecodeError, zipfile.BadZipFile) as exc:
        sys.exit(f"FAIL: cannot sync AFE store: {exc}")
    for change in changes:
        print(change)
    print(f"OK: {workbook} AFE store synced, {workbook.stat().st_size:,} bytes")


if __name__ == "__main__":
    main()
