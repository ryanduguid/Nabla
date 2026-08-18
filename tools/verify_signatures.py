"""Check that each function's help describes the function it belongs to.

Usage: python tools/verify_signatures.py [src dir]

Every function carries its own help, and states its parameters twice: once on the
FUNCTION line as a signature, and again as a table below it, one row per parameter.
Both are hand-written text inside a string literal, so no structural check ever reads
them, and both drift, separately. A function written by copying a neighbour keeps the
neighbour's name, its signature, its table, or all three. A parameter renamed in the
LAMBDA is not renamed in the help. A capital lands one key early and Flow1 becomes
FLow1. Every one of those shipped at least once, and one function's table described a
different function's arguments for six releases, leaving its own three undocumented.

The LAMBDA's own declaration is the ground truth for both: it is what the function
actually takes. The comparisons are character for character, because case is exactly
the kind of difference that goes unnoticed by eye.

Four conventions are honoured rather than reported:

  * the internal DoNotUse parameter is a period counter, kept out of the signature; the
    table may explain it or leave it out, and both are accepted
  * help is a two-column table, so a long row wraps onto one with an empty label, and
    the rows must be rejoined before either can be read
  * a table row whose label ends in ! is an aside, not a parameter: NOTE!, NOTES!
  * square brackets are ignored. Upstream declares every parameter optional so that a
    function called with no arguments can return its own help and police the omissions
    itself, so the declaration's brackets say nothing about which arguments a caller may
    leave out. Only the help distinguishes them, and there is nothing to check that
    against. Names are compared; requiredness is not.

The two checks are independent, because the parts they read are independent: the debt
module has never carried a FUNCTION line, but it does carry parameter tables, and those
are checked. Every declaration in every module is accounted for and the counts are
printed. Too few is itself a failure, since a checker that reads nothing passes
everything.
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


def parameter_table(body):
    """The labels of the help's parameter table, or None if it has no such table.

    The table runs from the PARAMETERS row to the next section heading, which is any
    label ending in a colon. Asides are dropped: a row labelled NOTE! explains something
    about the function rather than naming one of its arguments.
    """
    labels, started = [], False
    for label, _text in help_rows(body):
        if not started:
            started = label.startswith("PARAMETERS")
            continue
        if label.endswith(":"):
            break
        if not label.endswith("!"):
            labels.append(label)
    return labels if started else None


def main():
    paths = sorted(glob.glob(os.path.join(SRC, "*.txt")))
    if not paths:
        sys.exit("no source modules found in %s/" % SRC)

    failures, no_signature, no_table, not_lambda = [], [], [], []
    signatures = tables = read = 0
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

            real = declared(body)
            without_internal = [p for p in real if p != INTERNAL]

            row = next((t for label, t in help_rows(body) if label == "FUNCTION:"), None)
            if row is None:
                no_signature.append(name)
            else:
                signatures += 1
                sig = SIGNATURE.match(row.strip())
                if not sig:
                    failures.append("%s: the help's FUNCTION line is not a signature: %r"
                                    % (name, row.strip()[:70]))
                else:
                    if sig.group(1) != name:
                        failures.append("%s: its help names %s" % (name, sig.group(1)))
                    claimed = names(sig.group(2))
                    if claimed != without_internal:
                        failures.append("%s: its signature says (%s) but it takes (%s)"
                                        % (name, ", ".join(claimed),
                                           ", ".join(without_internal)))

            listed = parameter_table(body)
            if listed is None:
                no_table.append(name)
            else:
                tables += 1
                # the internal counter may be documented or not, so accept either
                if listed not in (real, without_internal):
                    failures.append("%s: its parameter table describes (%s) but it takes (%s)"
                                    % (name, ", ".join(listed) or "nothing",
                                       ", ".join(real)))

    print("%d declarations read: %d signatures and %d parameter tables checked" % (read, signatures, tables))
    print("   no FUNCTION line (%d): %s" % (len(no_signature),
                                            ", ".join(sorted(set(no_signature))) or "none"))
    print("   no parameter table (%d): %s" % (len(no_table),
                                              ", ".join(sorted(set(no_table))) or "none"))
    print("   not a LAMBDA (%d): %s" % (len(not_lambda),
                                        ", ".join(sorted(set(not_lambda))) or "none"))
    assert signatures + len(no_signature) + len(not_lambda) == read
    assert tables + len(no_table) + len(not_lambda) == read

    if signatures < FLOOR or tables < FLOOR:
        sys.exit("only %d signatures and %d tables parsed from %d module(s): "
                 "the parser is not reading them" % (signatures, tables, len(paths)))
    if failures:
        print("FAIL: %d help problem(s)" % len(failures))
        for item in failures:
            print("  - %s" % item)
        return 1
    print("OK: all %d signatures and %d parameter tables describe their own function's "
          "real parameters" % (signatures, tables))
    return 0


if __name__ == "__main__":
    sys.exit(main())
