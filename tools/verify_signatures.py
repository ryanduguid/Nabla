"""Check that each function's help signature matches the function it belongs to.

Usage: python tools/verify_signatures.py [src dir]

Every function carries its own help, and the first line of that help is a signature:
the function's name, then its parameter list. Both halves are hand-written text inside
a string literal, so no structural check ever reads them, and both halves drift. A
function written by copying a neighbour keeps the neighbour's name, its parameter list,
or both. A parameter renamed in the LAMBDA is not renamed in the help. A capital lands
one key early and Flow1 becomes FLow1. Every one of those shipped at least once.

The LAMBDA's own declaration is the ground truth: it is what the function actually
takes. The comparison is character for character, because case is exactly the kind of
difference that goes unnoticed by eye.

Three conventions are honoured rather than reported:

  * a trailing DoNotUse parameter is an internal period counter, deliberately left out
    of the signature although the parameter table below still explains it
  * help is a two-column table, so a long signature wraps onto a row with an empty
    label, and the rows must be rejoined before the signature can be read
  * square brackets are ignored on both sides. Upstream declares every parameter
    optional so that a function called with no arguments can return its own help and
    police the omissions itself, so the declaration's brackets say nothing about which
    arguments a caller may leave out. Only the help distinguishes them, and there is
    nothing to check it against. Names are compared; requiredness is not.

Every declaration in every module is accounted for, and the counts are printed. A
function with no FUNCTION line in its help is listed, not failed: the debt module has
never carried one. Anything that is not a LAMBDA at all, such as the About tables, is
listed separately. The three totals must add up to the declarations read, and too few
declarations is itself a failure, since a checker that reads nothing passes everything.
"""
import glob
import os
import re
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from verify_sources import NAME, statements   # same statement splitter the src check uses

# Every function name carries a λ, and a Windows console defaults to cp1252, which
# cannot encode it: without this the check dies printing its own result.
if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")


SRC = sys.argv[1] if len(sys.argv) > 1 else "src"

SIGNATURE = re.compile(r"^([A-Za-z0-9_.]*λ[A-Za-z0-9_]*)\s*\((.*)\)\s*$", re.S)
INTERNAL = "DoNotUse"        # an internal counter, documented but kept out of the signature
FLOOR = 100                  # the library ships 130 functions; reading far fewer is a bug


def literals(text):
    """Concatenate every double-quoted literal in order, honouring the "" escape."""
    out, i, n = [], 0, len(text)
    while i < n:
        if text[i] != '"':
            i += 1
            continue
        i += 1
        while i < n:
            if text[i] == '"':
                if i + 1 < n and text[i + 1] == '"':
                    out.append('"')
                    i += 2
                    continue
                i += 1
                break
            out.append(text[i])
            i += 1
    return "".join(out)


def help_rows(body):
    """The help as (label, text) rows, with wrapped rows rejoined onto their label."""
    rows = []
    for raw in literals(body).split("¶"):
        label, sep, text = raw.partition("→")
        if not sep:
            continue
        if not label.strip() and rows:
            rows[-1][1] += text          # a blank label continues the row above
            continue
        rows.append([label.strip(), text])
    return rows


def names(text):
    """Bare parameter names, with the optional-marker brackets dropped."""
    return [p.strip().strip("[]").strip() for p in text.split(",") if p.strip()]


def declared(body):
    """The parameter names the LAMBDA takes."""
    head = body[body.index("LAMBDA(") + len("LAMBDA("):]
    head = re.sub(r"//[^\n]*", "", head)
    cut = head.find("LET(")
    return names(head[:cut] if cut != -1 else head)


def main():
    paths = sorted(glob.glob(os.path.join(SRC, "*.txt")))
    if not paths:
        sys.exit("no source modules found in %s/" % SRC)

    failures, no_signature, not_lambda, checked, read = [], [], [], 0, 0
    for path in paths:
        for st in statements(open(path, encoding="utf-8").read()):
            m = NAME.match(st)
            if not m:
                continue
            read += 1
            name, body = m.group(1), m.group(2).strip()
            if not body.startswith("LAMBDA("):
                not_lambda.append(name)
                continue

            row = next((t for label, t in help_rows(body) if label == "FUNCTION:"), None)
            if row is None:
                no_signature.append(name)
                continue
            checked += 1

            sig = SIGNATURE.match(row.strip())
            if not sig:
                failures.append("%s: the help's FUNCTION line is not a signature: %r"
                                % (name, row.strip()[:70]))
                continue
            if sig.group(1) != name:
                failures.append("%s: its help names %s" % (name, sig.group(1)))

            claimed = names(sig.group(2))
            real = [p for p in declared(body) if p != INTERNAL]
            if claimed != real:
                failures.append("%s: its signature says (%s) but it takes (%s)"
                                % (name, ", ".join(claimed), ", ".join(real)))

    print("%d declarations read: %d signatures checked, %d without a FUNCTION line (%s), "
          "%d not a LAMBDA (%s)"
          % (read, checked, len(no_signature), ", ".join(sorted(set(no_signature))) or "none",
             len(not_lambda), ", ".join(sorted(set(not_lambda))) or "none"))
    assert checked + len(no_signature) + len(not_lambda) == read

    if checked < FLOOR:
        sys.exit("only %d signatures parsed from %d module(s): the parser is not reading them"
                 % (checked, len(paths)))
    if failures:
        print("FAIL: %d signature problem(s)" % len(failures))
        for item in failures:
            print("  - %s" % item)
        return 1
    print("OK: all %d help signatures name their own function and its real parameters" % checked)
    return 0


if __name__ == "__main__":
    sys.exit(main())
