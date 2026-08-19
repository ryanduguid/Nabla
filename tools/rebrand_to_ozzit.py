"""Rename the library from Nabla to Ozzit.

Nabla sits one edit from an acronym nobody wants beside their name, so the library
takes a new one. The rename is length-preserving on purpose: `Nabla` and `Ozzit` are
both five characters, `nb.` and `oz.` both three, bare `nb` and `oz` both two. Every
function carries a help block whose columns are aligned with padded arrows, so a
replacement of a different width would have to re-align 130 of them. This one does not,
and the script asserts the property on every string it touches rather than trusting it.

Usage: python tools/rebrand_to_ozzit.py [workbook]

Rewritten:

  the workbook      every text part: defined names, formulas, cached strings, help
                    text, drawings, the theme's colour scheme name, docProps
  src/*.txt         the Advanced Formula Environment module sources
  functions.csv     every column except `previous_name`
  the Markdown, the CI workflow, and the other tools

Deliberately left alone, because these record history rather than describe the library
as it stands:

  tools/released-names-v1.2.6.txt   the 130 names release v1.2.6 actually shipped
  functions.csv, previous_name      each function's name in that release
  CHANGELOG.md                      what past releases were called when they shipped

Rewriting either of the first two would break the migration path this repository
promises: somebody holding an older workbook needs to look up a name that existed, not
one that never did. The same reasoning keeps `nabla.*` in the changelog's older entries.
"""
import csv
import io
import os
import re
import sys
import zipfile

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
WORKBOOK = sys.argv[1] if len(sys.argv) > 1 else os.path.join(ROOT, "nabla.xlsx")

# Longest first: the capitalised and lowercase words cannot overlap the prefixes, but
# `nb.` must go before bare `nb` or the bare rule would eat the prefix's own letters.
WORDS = [("Nabla", "Ozzit"), ("nabla", "ozzit")]
PREFIX = [("nb.", "oz.")]
# Bare `nb` is the namespace as the help text names it in prose: "Suggested module
# name: nb", "Formula (nb is assumed to be the module name)". The lookaround keeps it
# off `nb_` in this repository's own Python and off the `nb` inside `onboarding` and
# `unbalanced`, while still reaching the `nb\.` written inside a regex literal.
BARE = re.compile(r"(?<![A-Za-z0-9_.])nb(?![A-Za-z0-9_.])")

# History, not description. See the module docstring.
SKIP = {
    "tools/released-names-v1.2.6.txt",
    "CHANGELOG.md",
    "tools/rebrand_to_ozzit.py",
}


def rebrand(text):
    """Apply every token rule, and prove the result is the same length as the input."""
    out = text
    for old, new in WORDS:
        out = out.replace(old, new)
    for old, new in PREFIX:
        out = out.replace(old, new)
    out = BARE.sub("oz", out)
    if len(out) != len(text):
        raise AssertionError("rebrand changed a string's length, which realigns help blocks")
    return out


def rebrand_workbook(path):
    """Rewrite every text part in place, leaving binary parts and zip metadata alone."""
    source = zipfile.ZipFile(path)
    entries = [(item, source.read(item)) for item in source.infolist()]
    source.close()
    touched = []
    buffer = io.BytesIO()
    with zipfile.ZipFile(buffer, "w", zipfile.ZIP_DEFLATED) as out:
        for item, data in entries:
            try:
                text = data.decode("utf-8")
            except UnicodeDecodeError:
                out.writestr(item, data)          # images, printer settings
                continue
            new = rebrand(text)
            if new != text:
                touched.append(item.filename)
            out.writestr(item, new.encode("utf-8"))
    with open(path, "wb") as fh:
        fh.write(buffer.getvalue())
    return touched


def rebrand_csv(path):
    """Rewrite functions.csv, holding `previous_name` at what those releases shipped."""
    with open(path, encoding="utf-8", newline="") as fh:
        rows = list(csv.DictReader(fh))
    fields = list(rows[0].keys())
    for row in rows:
        for key in fields:
            if key == "previous_name":
                continue
            row[key] = rebrand(row[key])
    with open(path, "w", encoding="utf-8", newline="") as fh:
        writer = csv.DictWriter(fh, fieldnames=fields, lineterminator="\n")
        writer.writeheader()
        writer.writerows(rows)


def text_targets():
    # Match whole path segments: a substring test would read `.github` as `.git` and
    # skip the workflow that runs the checks.
    ignored = {".git", ".ruff_cache", ".superpowers"}
    for folder, folders, names in os.walk(ROOT):
        folders[:] = [f for f in folders if f not in ignored]
        for name in names:
            full = os.path.join(folder, name)
            rel = os.path.relpath(full, ROOT).replace("\\", "/")
            if rel in SKIP:
                continue
            if name.endswith((".txt", ".md", ".py", ".ps1", ".yml", ".yaml", ".svg")) or name == "LICENCE":
                yield rel, full


def main():
    touched = rebrand_workbook(WORKBOOK)
    print("workbook: %d parts rewritten" % len(touched))

    rebrand_csv(os.path.join(ROOT, "functions.csv"))
    print("functions.csv: rewritten, previous_name held")

    count = 0
    for rel, full in text_targets():
        # newline="" both ways: read and write the file's own line endings rather than
        # normalising CRLF to LF and burying the real change in a whole-file diff.
        with open(full, encoding="utf-8", newline="") as fh:
            text = fh.read()
        new = rebrand(text)
        if new == text:
            continue
        with open(full, "w", encoding="utf-8", newline="") as fh:
            fh.write(new)
        count += 1
        print("  %s" % rel)
    print("text files: %d rewritten" % count)


if __name__ == "__main__":
    main()
