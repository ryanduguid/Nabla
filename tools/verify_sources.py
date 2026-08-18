"""Check that src/ really is the source of the functions in nabla.xlsx.

Usage: python tools/verify_sources.py [path/to/nabla.xlsx] [src dir]

src/ exists so the library can be read, diffed and imported back into Excel. That
only means anything if what is published matches what ships, and if it is written
in the form Excel accepts as input rather than the form the file format stores.

Four differences between the two are conventions, not divergences, and are mapped
rather than ignored:

    stored form            typed form         why
    _xlfn. _xlpm. _xlws.   (nothing)          markers for post-2007 functions
    _xlop.Name             [Name]             an OPTIONAL parameter
    [0]!Name               Name               "a name in this workbook"
    SINGLE(x)              @x                 implicit intersection

The parameter one matters most: stripping _xlop. instead of mapping it turns an
optional parameter into a required one, and any function whose body calls
ISOMITTED() on it is then rejected by Excel outright. That shipped once.

Module-local calls are qualified the way the Advanced Formula Environment does on
import, built-in names are upper-cased the way Excel does, and whitespace is
ignored, since help text is padded for alignment and TRIM() removes it.
"""
import html
import os
import re
import sys
import zipfile

# Every function name carries a λ, and a Windows console defaults to cp1252, which
# cannot encode it: without this the check dies printing its own result.
if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")


WORKBOOK = sys.argv[1] if len(sys.argv) > 1 else "nabla.xlsx"
SRC_DIR = sys.argv[2] if len(sys.argv) > 2 else "src"
MODULES = ["nabla.d", "nabla.e", "nabla.f", "nabla.r", "nabla.u", "nabla.debt"]

OPENERS, CLOSERS = "({[", ")}]"
NAME = re.compile(r"^\s*([A-Za-z_][A-Za-z0-9_.λ]*)\s*=\s*(.+)$", re.S)
OPTIONAL = re.compile(r"_xlop\.([A-Za-z_][A-Za-z0-9_]*\??)")   # ? is legal in a parameter name
PREFIX = re.compile(r"_xl[a-z]+\.", re.I)

failures = []


def fail(msg):
    failures.append(msg)


def read_string(text, i):
    """Copy a double-quoted literal verbatim, honouring the "" escape."""
    lit, n = ['"'], len(text)
    i += 1
    while i < n:
        if text[i] == '"':
            if i + 1 < n and text[i + 1] == '"':
                lit.append('""')
                i += 2
                continue
            lit.append('"')
            i += 1
            break
        lit.append(text[i])
        i += 1
    return "".join(lit), i


def statements(text):
    """Split a module into top-level statements, honouring strings and comments."""
    out, buf, depth = [], [], 0
    i, n = 0, len(text)
    while i < n:
        c = text[i]
        if c == '"':
            lit, i = read_string(text, i)
            buf.append(lit)
            continue
        if c == "/" and i + 1 < n and text[i + 1] == "/":
            while i < n and text[i] != "\n":
                i += 1
            continue
        if c == "/" and i + 1 < n and text[i + 1] == "*":
            j = text.find("*/", i + 2)
            i = j + 2 if j != -1 else n
            continue
        if c in OPENERS:
            depth += 1
        elif c in CLOSERS:
            depth -= 1
        elif c == ";" and depth == 0:
            out.append("".join(buf))
            buf = []
            i += 1
            continue
        buf.append(c)
        i += 1
    if "".join(buf).strip():
        out.append("".join(buf))
    return out


def split_literals(formula):
    """Alternating (is_string, chunk) pairs."""
    out, buf, i, n = [], [], 0, len(formula)
    while i < n:
        if formula[i] == '"':
            if buf:
                out.append((False, "".join(buf)))
                buf = []
            lit, i = read_string(formula, i)
            out.append((True, lit))
            continue
        buf.append(formula[i])
        i += 1
    if buf:
        out.append((False, "".join(buf)))
    return out


def unwrap_single(s):
    """SINGLE(expr) is how the file format stores the @ implicit-intersection operator."""
    while True:
        at = s.find("SINGLE(")
        if at < 0:
            return s
        depth, i = 0, at + len("SINGLE(") - 1
        while i < len(s):
            if s[i] == "(":
                depth += 1
            elif s[i] == ")":
                depth -= 1
                if depth == 0:
                    break
            i += 1
        if i >= len(s):
            return s
        s = s[:at] + "@" + s[at + len("SINGLE("):i] + s[i + 1:]


def canonical(formula):
    """Whitespace-free, upper-cased outside strings, in the typed form."""
    formula = OPTIONAL.sub(r"[\1]", formula)          # BEFORE the generic prefix strip
    formula = PREFIX.sub("", formula).replace("[0]!", "")
    parts, i, n = [], 0, len(formula)
    while i < n:
        c = formula[i]
        if c == '"':
            lit, i = read_string(formula, i)
            parts.append(re.sub(r"\s+", "", lit))
            continue
        if not c.isspace():
            parts.append(c.upper())
        i += 1
    return unwrap_single("".join(parts))


def qualify(formula, module, names):
    """Turn each module-local call into the qualified name AFE compiles it to."""
    chunks = []
    for is_string, chunk in split_literals(formula):
        if not is_string:
            for bare in sorted(names, key=len, reverse=True):
                chunk = re.sub(r"(?<![A-Za-z0-9_.!])" + re.escape(bare) + r"(?=\s*\()",
                               module + "." + bare, chunk)
        chunks.append(chunk)
    return "".join(chunks)


def main():
    parsed = {}
    for mod in MODULES:
        path = os.path.join(SRC_DIR, mod + ".txt")
        if not os.path.exists(path):
            fail("missing source module %s" % path)
            continue
        text = open(path, encoding="utf-8").read()
        # The published source must be in the form Excel accepts as input.
        for token in ("[0]!", "_xlfn.", "_xlpm.", "_xlop.", "_xlws."):
            if token in text:
                fail("%s contains the stored-form token %r, which Excel will not accept "
                     "as typed input" % (path, token))
        defs = {}
        for st in statements(text):
            m = NAME.match(st)
            if m:
                defs[m.group(1)] = m.group(2).strip()
            elif st.strip():
                fail("%s has an unparsable statement: %r" % (path, st.strip()[:60]))
        for bare, body in defs.items():
            parsed["%s.%s" % (mod, bare)] = qualify(body, mod, set(defs))

    z = zipfile.ZipFile(WORKBOOK)
    wbx = z.read("xl/workbook.xml").decode("utf-8")
    shipped = {n: html.unescape(b) for n, b in
               re.findall(r'<definedName name="(nabla\.[^"]+)"[^>]*>(.*?)</definedName>', wbx, re.S)}

    for name in sorted(set(shipped) - set(parsed)):
        fail("%s ships in the workbook but is not in src/" % name)
    for name in sorted(set(parsed) - set(shipped)):
        fail("%s is in src/ but does not ship in the workbook" % name)

    matched = 0
    for name in sorted(set(parsed) & set(shipped)):
        a, b = canonical(parsed[name]), canonical(shipped[name])
        if a == b:
            matched += 1
            continue
        cut = next((i for i in range(min(len(a), len(b))) if a[i] != b[i]), min(len(a), len(b)))
        fail("%s differs from its published source at character %d\n"
             "        src/ : %s\n        xlsx : %s"
             % (name, cut, a[max(0, cut - 30):cut + 50], b[max(0, cut - 30):cut + 50]))

    if failures:
        print("FAIL: %d problem(s) comparing %s with %s/" % (len(failures), WORKBOOK, SRC_DIR))
        for item in failures:
            print("  - %s" % item)
        return 1
    print("OK: %s/ reproduces all %d functions in %s" % (SRC_DIR, matched, WORKBOOK))
    return 0


if __name__ == "__main__":
    sys.exit(main())
