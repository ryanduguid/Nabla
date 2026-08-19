# Build nabla.xlsx from the upstream Financial Starter Pack workbook.
# Pure zip/XML surgery. Never resaves via openpyxl (preserves cached values, extensions, rich parts).
import zipfile, re, shutil, io, os, sys, datetime
import base64, json

# The progress lines name functions, and every function name carries a λ, which a Windows
# console's cp1252 cannot encode: without this the build dies partway through reporting
# what it just did, having already done it.
if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")

SRC = sys.argv[1] if len(sys.argv) > 1 else "2024-07-06.xlsx"
DST = sys.argv[2] if len(sys.argv) > 2 else "nabla.xlsx"
REPO_URL = "https://github.com/ryanduguid/Nabla"
TODAY_AU = "18 Aug 2026"
NOW_ISO = "2026-08-18T00:00:00Z"

zin = zipfile.ZipFile(SRC)
parts = {n: zin.read(n) for n in zin.namelist()}

REMOVE = {
    "xl/richData/richValueRel.xml",
    "xl/richData/rdrichvalue.xml",
    "xl/richData/rdrichvaluestructure.xml",
    "xl/richData/rdRichValueTypes.xml",
    "xl/richData/_rels/richValueRel.xml.rels",
    "xl/media/image1.png",
    "xl/media/image2.png",
    "xl/media/image3.png",
    # empty Power Query DataMashup store
    "customXml/item2.xml",
    "customXml/itemProps2.xml",
    "customXml/_rels/item2.xml.rels",
}
for r in REMOVE:
    assert r in parts, r
    del parts[r]

def get(n): return parts[n].decode("utf-8")
def put(n, s): parts[n] = s.encode("utf-8")

# ---------- 1. Content types: drop richData overrides + png default ----------
ct = get("[Content_Types].xml")
for pn in ["/xl/richData/richValueRel.xml", "/xl/richData/rdrichvalue.xml",
           "/xl/richData/rdrichvaluestructure.xml", "/xl/richData/rdRichValueTypes.xml"]:
    ct2 = re.sub(r'<Override PartName="%s"[^>]*/>' % re.escape(pn), "", ct)
    assert ct2 != ct, pn
    ct = ct2
ct2 = re.sub(r'<Default Extension="png"[^>]*/>', "", ct)
assert ct2 != ct
ct = ct2
ct2 = re.sub(r'<Override PartName="/customXml/itemProps2\.xml"[^>]*/>', "", ct)
assert ct2 != ct
ct = ct2
put("[Content_Types].xml", ct)

# ---------- 2. workbook rels: drop richData + empty DataMashup relationships ----------
rels = get("xl/_rels/workbook.xml.rels")
for rid in ["rId55", "rId56", "rId57", "rId58", "rId61"]:
    rels2 = re.sub(r'<Relationship Id="%s"[^>]*/>' % rid, "", rels)
    assert rels2 != rels, rid
    rels = rels2
put("xl/_rels/workbook.xml.rels", rels)

# ---------- 3. metadata.xml: drop XLRICHVALUE, keep XLDAPR ----------
md = get("xl/metadata.xml")
md = md.replace('<metadataTypes count="2">', '<metadataTypes count="1">')
md = re.sub(r'<metadataType name="XLRICHVALUE"[^>]*/>', "", md)
md = re.sub(r'<futureMetadata name="XLRICHVALUE".*?</futureMetadata>', "", md, flags=re.S)
md = re.sub(r'<valueMetadata.*?</valueMetadata>', "", md, flags=re.S)
assert "XLRICHVALUE" not in md and "XLDAPR" in md
put("xl/metadata.xml", md)

# ---------- 4. rich-value cells -> empty styled cells ----------
for sheet, cell in [("xl/worksheets/sheet1.xml", "B6"),
                    ("xl/worksheets/sheet16.xml", "K3"),
                    ("xl/worksheets/sheet21.xml", "K3")]:
    d = get(sheet)
    d2 = re.sub(r'<c r="%s" s="(\d+)" t="e" vm="\d+"><v>#VALUE!</v></c>' % cell,
                r'<c r="%s" s="\1"/>' % cell, d)
    assert d2 != d, (sheet, cell)
    put(sheet, d2)

# ---------- 5. Cover: drop Additional Content cells + Dropbox hyperlinks; sheet16 YouTube ----------
cov = get("xl/worksheets/sheet1.xml")
for cell in ["A29", "B30", "B32", "B33", "B35", "B36", "B38", "B39", "B7"]:
    cov2 = re.sub(r'<c r="%s"[^>]*(?:/>|>.*?</c>)' % cell, "", cov)
    assert cov2 != cov, cell
    cov = cov2
for ref in ["B39", "B36", "B33"]:
    cov2 = re.sub(r'<hyperlink ref="%s" r:id="rId\d+"[^>]*/>' % ref, "", cov)
    assert cov2 != cov, ref
    cov = cov2
put("xl/worksheets/sheet1.xml", cov)
r1 = get("xl/worksheets/_rels/sheet1.xml.rels")
n_before = r1.count("<Relationship")
r1 = re.sub(r'<Relationship Id="rId\d+" Type="[^"]*hyperlink" Target="https://www\.dropbox[^"]*"[^>]*/>', "", r1)
assert r1.count("<Relationship") == n_before - 3
put("xl/worksheets/_rels/sheet1.xml.rels", r1)

s16 = get("xl/worksheets/sheet16.xml")
s16b = re.sub(r'<hyperlink ref="K3:P15"[^>]*/>', "", s16)
assert s16b != s16
put("xl/worksheets/sheet16.xml", s16b)
r16 = get("xl/worksheets/_rels/sheet16.xml.rels")
r16b = re.sub(r'<Relationship Id="rId1" Type="[^"]*hyperlink" Target="https://www\.youtube[^"]*"[^>]*/>', "", r16)
assert r16b != r16
put("xl/worksheets/_rels/sheet16.xml.rels", r16b)

# ---------- 6. Targeted prose rewrites (pre token sweep) ----------
ss = get("xl/sharedStrings.xml")
rewrites = [
    ("This table is created by BXL TOC add-in available on Eloquens",
     "Click any name below to jump to its worksheet"),
    ("Financial Starter Pack Introduction", "nabla Introduction"),
    ("This workbook contains two 5g libraries: Dates (BXD) and Financial Functions (BXF). ",
     "This workbook contains the Nabla function library, covering dates, array essentials, "
     "financial functions, financial ratios, utilities and debt. Every function shares the nb. prefix. "),
    ("click and worksheet name", "click any worksheet name"),
    # matched before the brand sweep runs, so this is the upstream wording
    ("This is a library of 5g functions for simplifying financial model development, "
     "especially models using dynamic arrays, and especially for Excel novices.",
     "This is a library of nabla functions for simplifying financial model development, "
     "especially models using dynamic arrays, and especially for Excel novices. "
     "It needs Excel with LAMBDA and dynamic arrays: Microsoft 365, or Excel 2024 and later. "
     "Where Excel 365 has since gained a native equivalent, the function's inline help says so "
     "on a SEE ALSO line (checked August 2026)."),
]
for old, new in rewrites:
    assert old in ss, old[:40]
    ss = ss.replace(old, new)
# blank shared strings orphaned by the removed Additional Content section + image caption
blank_pats = [
    r'<si><t[^>]*>[^<]*dropbox\.com[^<]*</t></si>',
    r'<si><t[^>]*>5g makes the dream[^<]*Leonardo\.AI</t></si>',
    r'<si><t[^>]*>FAST 3 Statement Model</t></si>',
    r'<si><t[^>]*>FAST \+5G 3 Statement Model</t></si>',
    r'<si><t[^>]*>Tables \+5G 3 Statement Model</t></si>',
    r'<si><t[^>]*>Additional Content</t></si>',
    r'<si><t[^>]*>Additionally we include links[^<]*</t></si>',
]
for pat in blank_pats:
    ss2, cnt = re.subn(pat, '<si><t/></si>', ss)
    assert cnt >= 1, pat
    ss = ss2
put("xl/sharedStrings.xml", ss)

# ---------- 7. Global transforms over text parts + UTF-16 bins ----------
BRAND = [
    # upstream typo only; author names in revision histories are preserved (attribution, not branding)
    ("Craig Hatmaekr", "Craig Hatmaker"),
    ("GIST URL:", "REPOSITORY:"),
    ("Gist URL:", "Repository:"),
    ("Displays the URL to this module's Gist which includes documentation",
     "Displays this module's repository URL and function list"),
    ("BXLDebt.", "nabla.debt."),
    ("BXD.", "nabla.d."), ("BXE.", "nabla.e."), ("BXF.", "nabla.f."),
    ("BXR.", "nabla.r."), ("BXU.", "nabla.u."),
    ("BXL", "nabla"),
]
BRAND_BARE = [("BXD", "nabla.d"), ("BXE", "nabla.e"), ("BXF", "nabla.f"),
              ("BXR", "nabla.r"), ("BXU", "nabla.u"), ("5g", "nabla"), ("5G", "nabla")]
TYPOS = [("equally equally", "equally"),
         # worked-example results the +2 year date shift invalidated
         ('"→2¶" &amp; "→=nabla.d.CountDOWλ', '"→3¶" &amp; "→=nabla.d.CountDOWλ'),
         ('"→2¶" & "→=nabla.d.CountDOWλ', '"→3¶" & "→=nabla.d.CountDOWλ'),
         ("2023-Feb-26", "2025-Feb-26"), ("2023-Feb¶", "2025-Feb¶"),
         ("2023:Q01", "2025:Q1"), ("→2023¶", "→2025¶"), ("specifice text", "specific text"),
         ("dynamice", "dynamic"), ("a lable for", "a label for"),
         # double substitution artefact: upstream read "every BXL 5g Library"
         ("nabla nabla Library", "nabla library"),
         ("Every Workday (USA normal)", "Every Workday (Monday to Friday)"),
         ("randomly generated", "sample"), ("Randomly generated", "Sample"),
         ("\"FUNCTION:      FilterContains", "\"FUNCTION:      →FilterContains"),
         ("\"FUNCTION:      PeriodDiff", "\"FUNCTION:      →PeriodDiff"),
         ("\"FUNCTION:      RollingMin", "\"FUNCTION:      →RollingMin"),
         ("Liabilites", "Liabilities"),
         ('lang="en-US"', 'lang="en-AU"'),
         # unfulfilled upstream placeholder in 46 help blocks
         ("&lt;coming soon&gt;", REPO_URL), ("<coming soon>", REPO_URL),
         ("&lt;Coming soon&gt;", REPO_URL), ("<Coming soon>", REPO_URL),
         ("→Coming soon¶", "→" + REPO_URL + "¶"),
         # the +2 year shift left explanatory prose quoting the old years
         ("That loan starts in 2020.", "That loan starts in 2022."),
         ("Our model starts in 2024.", "Our model starts in 2026.")]

AMORT = [("Amoritization", "Amortisation"), ("amoritization", "amortisation"),
         ("Amoritize", "Amortise"), ("amoritize", "amortise"),
         ("Amortization", "Amortisation"), ("amortization", "amortisation"),
         ("Amortize", "Amortise"), ("amortize", "amortise"),
         ("Amortizing", "Amortising"), ("amortizing", "amortising"),
         ("Occurence", "Occurrence"), ("occurence", "occurrence")]
# Americanisms -> Australian equivalents (content, not just spelling)
AMERICAN = [
    # the MACRS row on the Data Validation sheet becomes the diminishing-balance row
    ("IRS Depreciation", "Diminishing balance 200%"),
    ("Some of Years", "Sum of Years"),
    ("Wal*Art", "Wool*Art"),
    ("Apt. ", "Unit "),
    ("Apartment", "Unit"),
    (">USD<", ">AUD<"),
]

# substring replacements (cover -s/-d/-ing and ALL-CAPS forms)
AU_SUB = [("modeling", "modelling"), ("Modeling", "Modelling"),
          ("modeler", "modeller"), ("Modeler", "Modeller"),
          ("ummarize", "ummarise"), ("apitaliz", "apitalis"),
          ("AMORTIZATION", "AMORTISATION"),
          ("preceeding", "preceding"), ("Preceeding", "Preceding"),
          ("dynamice", "dynamic")]
AU_WORD = [("gray", "grey"), ("Gray", "Grey")]
# US sample-data terms -> AU equivalents (exact cell-value tokens)
AU_DATA = [(">Paycheck<", ">Pay<"), (">Home Owners Insurance<", ">Home insurance<"),
           (">HOA Dues<", ">Strata levies<"), (">Gas<", ">Petrol<")]

URL_RE = re.compile(r'https://(?:sites\.google\.com/site/beyondexcel|gist\.github\.com/CHatmaker)[^\s"<>&¶]*')

# Sample/demo date refresh: +2 years, calendar-aware (29 Feb -> 28 Feb when target year is not a leap year)
YEAR_SHIFT = 2
EPOCH = datetime.datetime(1899, 12, 30)

def _shift_ymd(y, m, d):
    y2 = y + YEAR_SHIFT
    try:
        datetime.date(y2, m, d)
    except ValueError:
        d = 28
    return y2, m, d

def _shift_serial(v):
    frac = v - int(v)
    dt = EPOCH + datetime.timedelta(days=int(v))
    y2, m, d = _shift_ymd(dt.year, dt.month, dt.day)
    return (datetime.datetime(y2, m, d) - EPOCH).days + frac

def _mdy_to_au(mo):
    m_, d_, y_ = int(mo.group(1)), int(mo.group(2)), int(mo.group(3))
    if y_ < 100:  # upstream also wrote two-digit years, e.g. 02/26/23
        y_ += 2000
    y2, mm, dd = _shift_ymd(y_, m_, d_)
    return "%d/%d/%d" % (dd, mm, y2)

def _iso_shift(mo):
    y2, mm, dd = _shift_ymd(int(mo.group(1)), int(mo.group(2)), int(mo.group(3)))
    return "%04d-%02d-%02d" % (y2, mm, dd)

def _arr_shift(mo):
    return re.sub(r'\d{5}',
                  lambda n: str(int(_shift_serial(float(n.group(0)))))
                  if 40000 < int(n.group(0)) < 48000 else n.group(0),
                  mo.group(0))

def transform_text(s):
    s = URL_RE.sub(REPO_URL, s)
    for old, new in BRAND:
        s = s.replace(old, new)
    for old, new in BRAND_BARE:
        s = re.sub(r'\b%s\b' % old, new, s)
    for old, new in TYPOS:
        s = s.replace(old, new)
    for old, new in AMORT:
        s = s.replace(old, new)
    for old, new in AMERICAN:
        s = s.replace(old, new)
    for old, new in AU_SUB:
        s = s.replace(old, new)
    for old, new in AU_WORD:
        s = re.sub(r'\b%s\b' % old, new, s)
    for old, new in AU_DATA:
        s = s.replace(old, new)
    # CountDOWλ's worked example: shifting its dates +2 years changed the answer from 2 to 3
    s = re.sub(r'"2(\s+)→=nabla\.d\.CountDOWλ', r'"3\g<1>→=nabla.d.CountDOWλ', s)
    s = re.sub(r'(?i)(ersion:\s*→\s*)[A-Z][a-z]{2} \d{1,2} \d{4}', r'\g<1>' + TODAY_AU, s)
    # sample dates: quoted or cell-value m/d/yyyy -> +2y, AU day-first
    # (trailing char may be a backslash inside JSON-escaped AFE text)
    s = re.sub(r'(?<=["">])(\d{1,2})/(\d{1,2})/(20\d\d)(?=["<\\])', _mdy_to_au, s)
    s = re.sub(r'(?<=["">])(\d{1,2})/(\d{1,2})/(\d{2})(?=["<\\])', _mdy_to_au, s)
    # The About tables carry the same URL twice, as "Repository:" and "Website:". Drop the
    # second in the XML form only; the AFE module sources are de-duplicated after parsing,
    # where the quotes are real rather than JSON-escaped.
    s = re.sub(r'"Website:\s*→%s\s*¶" &amp; ' % re.escape(REPO_URL), "", s)
    s = re.sub(r'\b(20\d\d)-(\d\d)-(\d\d)\b', _iso_shift, s)
    s = re.sub(r'\{[\d;\s]+\}', _arr_shift, s)
    return s

# ---------- Australian additions: function specs ----------
# Each spec generates both the compiled defined-name body and the readable AFE module source
# from one description, so the two representations cannot drift apart.

def xesc(s):
    return s.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")

def help_lines(spec):
    out = ['FUNCTION:      →%s¶' % spec["sig"],
           'DESCRIPTION:   →%s¶' % spec["desc"],
           'WEBPAGE:       →%s¶' % REPO_URL,
           'VERSION:       →%s¶' % TODAY_AU,
           'PARAMETERS:    →¶']
    out += ['%-15s→%s¶' % (p, d) for p, d in spec["params"]]
    out += ['EXAMPLES:      →¶',
            "→Formula (%s is assumed to be the module's name)¶" % spec["module"],
            '→=%s¶' % spec["example"], '→Result¶', '→%s' % spec["result"]]
    return out

# The compiled name and the readable module source are generated from ONE expression each.
# They used to be two hand-kept lists, which is how the shipped FinancialYearλ and its
# published source silently drifted apart; de-prefixing makes that impossible.
DEPREFIX = re.compile(r"_xl[a-z]+\.")   # _xlfn. _xlpm. _xlop. _xlws. and friends

def build_xml(spec):
    hl = help_lines(spec)
    help_expr = ('TRIM(_xlfn.TEXTSPLIT(' + " &amp; ".join('"%s"' % xesc(l) for l in hl) + ', "→", "¶"))')
    lets = ['_xlpm.Help, ' + help_expr, '_xlpm.Help?, ' + xesc(spec["help_test"])]
    lets += ['_xlpm.%s, %s' % (n, xesc(e)) for _, n, e in spec["lets"]]
    return '_xlfn.LAMBDA(%s, _xlfn.LET(%s, CHOOSE(_xlpm.Help? + 1, _xlpm.Result, _xlpm.Help)))' % (
        spec["xml_decl"], ', '.join(lets))

def build_afe(spec):
    hl = help_lines(spec)
    body = "".join('                            "%s"%s\n' % (l, ' &' if i < len(hl) - 1 else ',')
                   for i, l in enumerate(hl))
    lets = "".join("    //  %s\n        %-16s%s,\n" % (c, n + ",", DEPREFIX.sub("", e))
                   for c, n, e in spec["lets"])
    return (
        "/*  FUNCTION NAME:  %s\n" % spec["name"]
        + "    DESCRIPTION:*//**%s*/\n" % spec["desc"].rstrip(".")
        + "/*  REVISIONS:      Date        Developer       Description  \n"
        + "                    %s nabla           Original development\n" % TODAY_AU
        + "*/\n\n"
        + "%s = LAMBDA(\n" % spec["name"]
        + "//  Parameter Declaration\n"
        + "".join("    [%s],\n" % p for p, _ in spec["params"])
        + "    LET(\n"
        + "    //  Help\n"
        + "        Help,           TRIM(TEXTSPLIT(\n" + body
        + '                            "→", "¶"\n'
        + "                        )),\n"
        + "    //  Check inputs - Omitted required arguments\n"
        + "        Help?,          %s,\n" % DEPREFIX.sub("", spec["help_test"])
        + lets
        + "    //  Return Result or Help\n"
        + "        CHOOSE( Help? + 1, Result, Help)\n"
        + "    )\n"
        + ");\n")

FUNCS = [
    {
        "module": "nabla.f", "name": "DiminishingValueλ",
        "sig": "DiminishingValueλ(Cost, Life)",
        "desc": "Diminishing-balance depreciation schedule at 200% of the straight-line rate, writing the residual off in the final period. A modelling schedule, not a tax calculation.",
        "params": [("Cost", "(Required) Asset's cost."),
                   ("Life", "(Required) Asset's effective life in years.")],
        "example": "nabla.f.DiminishingValueλ(1000, 5)",
        "result": "400.00,240.00,144.00,86.40,129.60",
        "xml_decl": "_xlop.Cost,_xlop.Life",
        "help_test": "OR(_xlfn.ISOMITTED(_xlpm.Cost), _xlfn.ISOMITTED(_xlpm.Life))",
        "lets": [
            ("Rate is capped: a life under two years would otherwise exceed 100%",
             "Rate", "MIN(2/_xlpm.Life, 1)"),
            ("Whole periods, so a part-year effective life still gets its final period",
             "Count", "MAX(1, ROUNDUP(_xlpm.Life, 0))"),
            ("Set Constants", "Periods", "_xlfn.SEQUENCE(, _xlpm.Count)"),
            ("The Periods=1 guard avoids 0^0, which Excel returns as #NUM! once Rate reaches 1",
             "Raw", "_xlpm.Cost * _xlpm.Rate * IF(_xlpm.Periods = 1, 1, (1-_xlpm.Rate)^(_xlpm.Periods-1))"),
            ("Write the undeducted residual off in the final period",
             "Result", "_xlpm.Raw + (_xlpm.Periods = _xlpm.Count) * (_xlpm.Cost - SUM(_xlpm.Raw))"),
        ],
    },
    {
        "module": "nabla.f", "name": "PrimeCostλ",
        "sig": "PrimeCostλ(Cost, Life)",
        "desc": "Straight-line depreciation schedule for one asset or asset class, in whole years. A modelling schedule, not a tax calculation.",
        "params": [("Cost", "(Required) Asset's cost."),
                   ("Life", "(Required) Asset's effective life in years.")],
        "example": "nabla.f.PrimeCostλ(1000, 5)",
        "result": "200.00,200.00,200.00,200.00,200.00",
        "xml_decl": "_xlop.Cost,_xlop.Life",
        "help_test": "OR(_xlfn.ISOMITTED(_xlpm.Cost), _xlfn.ISOMITTED(_xlpm.Life))",
        "lets": [
            ("Set Constants", "Annual", "_xlpm.Cost/_xlpm.Life"),
            ("Whole periods, so a part-year effective life still gets its final period",
             "Count", "MAX(1, ROUNDUP(_xlpm.Life, 0))"),
            ("The final period carries the part-year remainder, so the schedule sums to cost",
             "Result", "IF(_xlfn.SEQUENCE(, _xlpm.Count) = _xlpm.Count, "
                       "_xlpm.Cost - _xlpm.Annual * (_xlpm.Count - 1), _xlpm.Annual)"),
        ],
    },
    {
        "module": "nabla.f", "name": "GSTAddλ",
        "sig": "GSTAddλ(Amounts, [Rate])",
        "desc": "Adds GST to one or more GST-exclusive amounts.",
        "params": [("Amounts", "(Required) One or more GST-exclusive amounts."),
                   ("Rate", "(Optional: Default = 0.1) GST rate as a fraction.")],
        "example": "nabla.f.GSTAddλ(100)",
        "result": "110",
        "xml_decl": "_xlop.Amounts,_xlop.Rate",
        "help_test": "_xlfn.ISOMITTED(_xlpm.Amounts)",
        "lets": [
            ("A blank Rate cell is not an omitted argument, so test for both",
             "GSTRate", 'IF(OR(_xlfn.ISOMITTED(_xlpm.Rate), TRIM(_xlpm.Rate & "")=""), 0.1, _xlpm.Rate)'),
            ("Blanks stay blank, so a part-filled column of amounts does not fill with zeros",
             "Result", 'IF(TRIM(_xlpm.Amounts & "")="", "", _xlpm.Amounts * (1 + _xlpm.GSTRate))'),
        ],
    },
    {
        "module": "nabla.f", "name": "GSTExtractλ",
        "sig": "GSTExtractλ(Amounts, [Rate])",
        "desc": "Returns the GST contained in one or more GST-inclusive amounts.",
        "params": [("Amounts", "(Required) One or more GST-inclusive amounts."),
                   ("Rate", "(Optional: Default = 0.1) GST rate as a fraction.")],
        "example": "nabla.f.GSTExtractλ(110)",
        "result": "10",
        "xml_decl": "_xlop.Amounts,_xlop.Rate",
        "help_test": "_xlfn.ISOMITTED(_xlpm.Amounts)",
        "lets": [
            ("A blank Rate cell is not an omitted argument, so test for both",
             "GSTRate", 'IF(OR(_xlfn.ISOMITTED(_xlpm.Rate), TRIM(_xlpm.Rate & "")=""), 0.1, _xlpm.Rate)'),
            ("Blanks stay blank, so a part-filled column of amounts does not fill with zeros",
             "Result", 'IF(TRIM(_xlpm.Amounts & "")="", "", '
                       '_xlpm.Amounts * _xlpm.GSTRate / (1 + _xlpm.GSTRate))'),
        ],
    },
    {
        "module": "nabla.d", "name": "FinancialYearλ",
        "sig": "FinancialYearλ(Dates, [StartMonth])",
        "desc": "Labels dates with their Australian financial year, which starts on 1 July.",
        "params": [("Dates", "(Required) One or more dates."),
                   ("StartMonth", "(Optional: Default = 7) Month the financial year starts.")],
        "example": 'nabla.d.FinancialYearλ(DATE(2026,8,15))',
        "result": "FY2027",
        "xml_decl": "_xlop.Dates,_xlop.StartMonth",
        "help_test": "_xlfn.ISOMITTED(_xlpm.Dates)",
        "lets": [
            ("Set defaults", "FYStart", "IF(_xlfn.ISOMITTED(_xlpm.StartMonth), 7, _xlpm.StartMonth)"),
            ("Multiply rather than AND, which would collapse a range of dates to one answer",
             "Result", 'IF(N(_xlpm.Dates)=0, "", "FY" & TEXT(YEAR(_xlpm.Dates) '
                       '+ (_xlpm.FYStart>1) * (MONTH(_xlpm.Dates) >= _xlpm.FYStart), "0000"))'),
        ],
    },
]

# MACRS (US Modified Accelerated Cost Recovery System) is removed outright: the library is
# Australian-only. Depreciation method slot 6 becomes diminishing balance at 200% of straight
# line and a seventh slot adds straight line. Both are modelling schedules, not tax calculations:
# neither knows an income year, an acquisition date or days held.
AFE_MACRS = [
    ('{"SLN","SYD","DB","DDB","VDB", "MACRS"}', '{"SLN","SYD","DB","DDB","VDB","DV","PC"}'),
    ('@INDEX( LifeInYears, Asset) + N(Method = "MACRS")', '@INDEX( LifeInYears, Asset)'),
    ('IF( Method = "MACRS", 0, @INDEX( SalvageValues, Asset))',
     'IF( OR( Method = "DV", Method = "PC"), 0, @INDEX( SalvageValues, Asset))'),
    ('"SLN,SYD,DB,DDB,VDB,MACRS"', '"SLN,SYD,DB,DDB,VDB,DV,PC"'),
    ('Methods must be omitted or one of: SLN, SYD, DB, DDB, MACRS, or VDB.',
     'Methods must be omitted or one of: SLN, SYD, DB, DDB, VDB, DV, or PC.'),
    ('Must be one of these Excel function names: ', 'Must be one of these method codes: '),
    ('"→MACRS=Modified Accelerated Cost Recovery System. NOTE: Salvage value ignored¶" & ',
     '"→DV =Diminishing balance at 200% of straight line. Salvage value ignored¶" & \n'
     '                                           "→PC =Straight line, whole years. Salvage value ignored¶" & '),
]
AFE_MACRS_RE = [
    (r'DisposalDate,   IF\( Method = "MACRS", \s*\n\s*MAX\(EDATE\( InserviceDate, Years \* MpY\), '
     r'@INDEX\( DisposalDates, Asset\)\),\s*\n\s*@INDEX\( DisposalDates, Asset\)\), ',
     'DisposalDate,   @INDEX( DisposalDates, Asset), '),
    (r'//  6\. Modified accelerated cost recovery system \s*\n\s*MACRSλ\( InitialValue, Years - 1\),',
     '//  6. Diminishing balance at 200% of straight line\n'
     + ' ' * 56 + 'DiminishingValueλ( InitialValue, Years),\n'
     + ' ' * 52 + '//  7. Straight line, whole years\n'
     + ' ' * 56 + 'PrimeCostλ( InitialValue, Years),'),
]

SEE_ALSO = {
    "nabla.e.RangeToDAλ": "Excel 365 now has TRIMRANGE and trim references (.:.) for this.",
    "nabla.u.RangeToDAλ": "Excel 365 now has TRIMRANGE and trim references (.:.) for this.",
    "nabla.f.RangeToDAλ": "Excel 365 now has TRIMRANGE and trim references (.:.) for this.",
    "nabla.f.FilterContainsλ": "Excel 365 now has REGEXTEST and REGEXEXTRACT for pattern matching.",
    "nabla.f.SumPeriodsλ": "Excel 365 now has GROUPBY and PIVOTBY for formula-driven aggregation.",
    "nabla.f.SumContainsλ": "Excel 365 now has GROUPBY and PIVOTBY for formula-driven aggregation.",
}
# AFE (Excel Labs) project store: base64-wrapped UTF-16 JSON holding LAMBDA source
afe = get("customXml/item1.xml")
m = re.search(r'>([A-Za-z0-9+/=]{100,})<', afe)
assert m
j = base64.b64decode(m.group(1)).decode("utf-16-le")
obj_afe = json.loads(transform_text(j))
mods = {f["path"].rsplit("/", 1)[1]: f for f in obj_afe["files"]}
# drop the duplicated "Website:" About line now that the text is unescaped
dropped = 0
for f in obj_afe["files"]:
    f["text"], k = re.subn(r'\n\s*"Website:\s*→%s\s*¶" &' % re.escape(REPO_URL), "", f["text"])
    dropped += k
assert dropped >= 4, dropped

ftext = mods["nabla.f"]["text"]
for old, new in AFE_MACRS:
    assert ftext.count(old) == 1, old[:60]
    ftext = ftext.replace(old, new)
for pat, new in AFE_MACRS_RE:
    ftext, n = re.subn(pat, new, ftext)
    assert n == 1, pat[:60]
# excise the whole MACRSλ definition
i_mac = ftext.index("/*  FUNCTION NAME:  MACRSλ")
i_next = ftext.index("/*  FUNCTION NAME:", i_mac + 10)
ftext = ftext[:i_mac] + ftext[i_next:]
assert "MACRS" not in ftext
mods["nabla.f"]["text"] = ftext

# the SEE ALSO lines must live in the module source too, or an Excel Labs save drops them
see_afe = 0
for full, note in SEE_ALSO.items():
    mod, fname = full.split(".")[1], full.split(".")[2]
    text = mods["nabla." + mod]["text"]
    # anchor on the help signature line: not every function has a FUNCTION NAME header
    sig = re.search(r'"FUNCTION:\s*→?\s*%s\(' % re.escape(fname), text)
    assert sig, full
    start = sig.start()
    label = re.compile(r'\n(\s*)"(?:WEBPAGE|WEBSITE|VERSION|PARAMETERS):')
    lab = label.search(text, start)   # not `m`: that still holds the project-store blob match
    assert lab, full
    line = '\n%s"SEE ALSO:      →%s¶" &' % (lab.group(1), note)
    mods["nabla." + mod]["text"] = text[:lab.start()] + line + text[lab.start():]
    see_afe += 1
assert see_afe == len(SEE_ALSO), see_afe
print("SEE ALSO added to", see_afe, "module sources")

# Upstream's installed SumDepreciateλ is a later revision than its own module source:
# the name carries a blank help row and a different (behaviour-identical) way of
# testing for the omitted argument. Bring the source up to the version that ships, so
# what is published is what people actually get.
ftxt = mods["nabla.f"]["text"]
_s = ftxt.index("SumDepreciateλ = LAMBDA")
_e = ftxt.index("\n);", _s) + 3
block = ftxt[_s:_e]
for _old, _new in (
    ('"DepreciationSchedule→(Required) An array produced by Depreciateλ¶" &',
     '"DepreciationSchedule→(Required) An array produced by Depreciateλ¶" &\n'
     + " " * 28 + '"→¶" &'),
    ("        Help?,          ISOMITTED( DepreciationSchedule),\n",
     "        OmittedArgs,    VSTACK(ISOMITTED( DepreciationSchedule)),\n"
     "        Help?,          AND( OmittedArgs),\n"),
    ("        CHOOSE(Help? + 1, Result, Help)",
     "        Return,         IF( Help?, 2, 1),\n"
     "        CHOOSE(Return, Result, Help)"),
):
    assert block.count(_old) == 1, _old[:40]
    block = block.replace(_old, _new)
mods["nabla.f"]["text"] = ftxt[:_s] + block + ftxt[_e:]

# Four functions were written by copying a neighbour, and their help still announces the
# neighbour's name on the FUNCTION line. The description and the parameter list below it
# are their own, so only the name is wrong. AvgColsλ additionally defeated the collision
# tagging further down, because the module text renames the bare name it finds there and
# the defined name does not, leaving the shipped function and its published source
# disagreeing. Name the function each one documents, in both representations.
HELP_NAMES_ITSELF = [
    # function,               the neighbour its help names,  modules holding it
    ("AvgColsλ",            "SumColsλ",                ("nabla.e", "nabla.u")),
    ("CorkScrewReversalλ",  "Corkscrewλ",              ("nabla.f",)),
    ("DDBλ",                "DBλ",                     ("nabla.f",)),
    ("EquityRatioλ",        "InterestCoverageRatioλ",  ("nabla.r",)),
]

_wbx = get("xl/workbook.xml")
for _fn, _wrong, _in_mods in HELP_NAMES_ITSELF:
    _from, _to = "→%s(" % _wrong, "→%s(" % _fn
    for _mod in _in_mods:
        _t = mods[_mod]["text"]
        _d = re.search(r"(?m)^%s\s*=\s*LAMBDA" % re.escape(_fn), _t)   # not a longer name
        assert _d, (_fn, _mod)
        _a = _d.start()
        _b = _t.index("\n);", _a) + 3
        _blk = _t[_a:_b]
        assert _blk.count(_from) == 1, (_fn, _mod)
        mods[_mod]["text"] = _t[:_a] + _blk.replace(_from, _to) + _t[_b:]

    _hits = []
    def _name_itself(m, _from=_from, _to=_to, _hits=_hits):
        body = m.group(2)
        if _from in body:
            _hits.append(m.group(1))
            body = body.replace(_from, _to)
        return m.group(1) + body + m.group(3)
    # the module sources are already rebranded here, the workbook part is not: match either
    _wbx = re.sub(r'(<definedName name="[A-Za-z0-9.]*\.%s"[^>]*>)(.*?)(</definedName>)'
                  % re.escape(_fn), _name_itself, _wbx, flags=re.S)
    assert len(_hits) == len(_in_mods), (_fn, _hits)
put("xl/workbook.xml", _wbx)

# Excel stamps xl/workbook.xml with the directory the file was last saved from. That is a
# fact about the build machine, not about the library, and every release up to v2.2.0
# published one: on a machine whose account is not named "-" it carries the account name.
# tools/refresh_cache.py strips it again after Excel puts it back, and the banned-token
# gate fails on it, so it cannot come back quietly.
_wbx = get("xl/workbook.xml")
_wbx, _abs = re.subn(
    r'<mc:AlternateContent[^>]*>\s*<mc:Choice Requires="x15">\s*<x15ac:absPath[^>]*/>\s*'
    r"</mc:Choice>\s*</mc:AlternateContent>", "", _wbx, flags=re.S)
assert _abs == 1, ("absPath", _abs)
put("xl/workbook.xml", _wbx)
print("stripped the build machine's path from the workbook")
print("help now names its own function in", sum(len(m) for _, _, m in HELP_NAMES_ITSELF), "definitions")

# The Corkscrew pair's help signature spells its second parameter FLow1, with a capital L.
# The parameter table below it spells the same argument Flow1, and so does the LAMBDA, so
# the signature line is the one that is wrong. Nothing structural catches it, because it
# lives inside a string literal. Fix it in all three places it is stored: the module
# sources, the defined names, and the help output already cached on the demonstration
# sheet, which would otherwise go on showing the typo until something forced a recalc.
HELP_PARAM_TYPOS = [("FLow1", "Flow1")]
for _wrong, _right in HELP_PARAM_TYPOS:
    _in_src = sum(m["text"].count(_wrong) for m in mods.values())
    assert _in_src, (_wrong, "not present in any module")
    for _mod in mods:
        mods[_mod]["text"] = mods[_mod]["text"].replace(_wrong, _right)

    _wbx = get("xl/workbook.xml")
    assert _wbx.count(_wrong) == _in_src, (_wrong, _in_src, _wbx.count(_wrong))
    put("xl/workbook.xml", _wbx.replace(_wrong, _right))

    _cached = [n for n in list(parts) if n.startswith("xl/worksheets/") and _wrong in get(n)]
    for _sheet in _cached:
        put(_sheet, get(_sheet).replace(_wrong, _right))
    print("fixed %s -> %s in %d sources, %d defined names, %d cached help output(s)"
          % (_wrong, _right, _in_src, _in_src, len(_cached)))


# ---------- help signatures that describe a different function ----------
# Each function's help opens with a signature, its name and parameter list, and repeats
# the parameters as a table three rows below. Where the two disagree the table has been
# right every time: it is what the LAMBDA declares. The signatures drifted by being
# copied wholesale from a neighbour (four of the ratios document the neighbour's
# arguments outright), by being left behind when a parameter was renamed, or by a plain
# typing slip. None of it is reachable by any structural check, because every one of them
# sits inside a string literal. tools/verify_signatures.py reads all 117 signatures back
# out of src/ and fails the build on any that no longer matches its own function.
#
# Both stores are patched from this one table: the module source that src/ is exported
# from, here, and the defined name Excel installs, further down. They are patched at
# different points because they are swept for spelling and rebranded at different points,
# and a fragment written in Australian English matches only after that sweep.
HELP_SIGNATURES = [
    # module,     function,          what the signature said,          what the function takes
    ("nabla.f", "CorkScrewReversalλ", "( Opening, Flow1,", "( Opening, ReversalFlags, Flow1,"),
    ("nabla.f", "Movementλ", "( [BeginningValue], Values)", "( [BeginningValues], Values)"),
    ("nabla.f", "LabelAmortiseλ", "λ([LoanNames])",
     "λ(LoanNames, [LoanAmounts], [LoanAPRs], [LoanTerms])"),
    ("nabla.f", "Depreciateλ", "[Methods], [Factor])", "[Methods], [Factors])"),
    ("nabla.f", "DBλ", "(Cost, Salvage, Life, [Month])", "(Cost, Salvage, Life, [Months])"),
    ("nabla.f", "TimelineOffsetλ", "λ(ArrayStart, Timeline)", "λ(Date, Timeline)"),
    ("nabla.f", "FilterContainsλ", "FilterByArray, Text,", "FilterByArray, FilterByText,"),
    ("nabla.r", "QuickRatioλ", "λ( LiquidAssets, Liabilities)", "λ( QuickAssets, Liabilities)"),
    ("nabla.r", "WorkingCapitalTurnoverRatioλ", "λ( CostOfGoodsSold, AverageInventory)",
     "λ( NetAnnualSales, WorkingCapital)"),
    ("nabla.r", "DSCRλ", "λ( NetOperatingIncome, Totaldebtservice)",
     "λ( NetOperatingIncome, TotalDebtService)"),
    ("nabla.r", "CashFlowMarginλ", "λ( NetIncome, Revenue)",
     "λ( CashFlowFromOperatingActivities, Revenue)"),
    ("nabla.r", "PriceToBookRatioλ", "BookValuePerShareBvps)", "BookValuePerShare)"),
    ("nabla.r", "PriceToCashRatioλ", "λ( MarketPricePerShare, SalesPerShare)",
     "λ( MarketPricePerShare, OperatingCashFlowPerShare)"),
    # Two parameter tables spell an argument differently again, in the label column. The
    # labels are padded to a fixed width, so the replacement keeps the column aligned.
    ("nabla.f", "LabelAmortiseλ", '"LoanAPR       →', '"LoanAPRs      →'),
    ("nabla.f", "LabelAmortiseλ", '"LoanTerm      →', '"LoanTerms     →'),
    ("nabla.r", "DSCRλ", '"TotaldebtService   →', '"TotalDebtService   →'),
    # Four labels put the colon after the padding instead of before it, so the help
    # reads "EXAMPLES :" and its arrow sits one column right of every other row's.
    # These four are clones of one another; every other EXAMPLES label is already right.
    ("nabla.e", "IsBetweenλ", '"EXAMPLES       :→', '"EXAMPLES:      →'),
    ("nabla.e", "IsInListλ", '"EXAMPLES       :→', '"EXAMPLES:      →'),
    ("nabla.u", "IsBetweenλ", '"EXAMPLES       :→', '"EXAMPLES:      →'),
    ("nabla.u", "IsInListλ", '"EXAMPLES       :→', '"EXAMPLES:      →'),
    # IsBetweenλ's parameter table calls its second argument Lo, which is not what the
    # LAMBDA declares or what the signature above it says, and describes the upper limit
    # as the lower one, copied from the row above. The Dates module's own IsBetweenλ has
    # both right and reads "The higher limit that the value(s) must be less than".
    ("nabla.e", "IsBetweenλ", '"Lo             →', '"Low            →'),
    ("nabla.e", "IsBetweenλ", "(Required) The lower limit that the value must be less than",
     "(Required) The higher limit that the value must be less than"),
    ("nabla.e", "IsBetweenλ", "equal to Lo and/or Hi", "equal to Low and/or Hi"),
    ("nabla.u", "IsBetweenλ", '"Lo             →', '"Low            →'),
    ("nabla.u", "IsBetweenλ", "(Required) The lower limit that the value must be less than",
     "(Required) The higher limit that the value must be less than"),
    ("nabla.u", "IsBetweenλ", "equal to Lo and/or Hi", "equal to Low and/or Hi"),
    # EquityMultiplierλ's table drops the Total from its second parameter. The LAMBDA and
    # the signature agree on TotalShareholdersEquity; only the table shortens it, which is
    # the name the neighbouring DebtToEquityRatioλ genuinely uses.
    ("nabla.r", "EquityMultiplierλ", '"ShareholdersEquity →(Required)',
     '"TotalShareholdersEquity →(Required)'),
    # EquityRatioλ's table is not its own at all: it documents InterestCoverageRatioλ's two
    # arguments, which is the same copy that gave this function the wrong name until v1.2.3.
    # Its three real parameters were never described. The wording follows the neighbours
    # that already document the same quantities.
    # Three worked examples call a neighbour instead of the function they document, so
    # copying the example as printed runs the wrong function. The results they claim are
    # right for the function they belong to, which is how it went unnoticed: only the
    # call is wrong. RollingAvgλ is the same copy going the other way, its call correct
    # and its result lifted from RollingSumλ. Excel gives 1, 1.5, 2, 3, 4.
    ("nabla.f", "RollingMinλ", "RollingMaxλ({5,3,4,6,2,7}, 3)", "RollingMinλ({5,3,4,6,2,7}, 3)"),
    ("nabla.f", "RollingSumλ", "RollingMinλ(SEQUENCE(, 5), 3)", "RollingSumλ(SEQUENCE(, 5), 3)"),
    ("nabla.f", "RollingAvgλ", '"1,3,6,9,12     →', '"1,1.5,2,3,4    →'),
    ("nabla.d", "ScheduleValuesByItemsλ", "ScheduleRatesByItemsλ(", "ScheduleValuesByItemsλ("),
    # and one help points the reader at a function that does not exist, by transposition
    ("nabla.f", "Amortiseλ", "LableAmortiseλ", "LabelAmortiseλ"),
    # MaxColsλ's description was copied from MinColsλ and never changed, so both copies of
    # the function tell the reader they return the minimum. The Utilities clone says it too.
    ("nabla.e", "MaxColsλ", "→Get the minimum for each Column", "→Get the maximum for each Column"),
    ("nabla.u", "MaxColsλ", "→Get the minimum for each Column", "→Get the maximum for each Column"),
    # and CountColsλ's signature carries a trailing comma, so the printed call has an empty
    # second argument in it
    ("nabla.e", "CountColsλ", "CountColsλ( Array,)", "CountColsλ( Array)"),
    ("nabla.u", "CountColsλ", "CountColsλ( Array,)", "CountColsλ( Array)"),
    ("nabla.r", "EquityRatioλ",
     '"OperatingIncome    →(Required) Operating Income (EBIT) ¶"',
     '"ShareholdersEquity →(Required) Shareholders\' Equity (total assets - total liabilities) ¶"'),
    ("nabla.r", "EquityRatioλ",
     '"InterestExpenses   →(Required) Interest Expense ¶"',
     '"TotalAssets        →(Required) Both current and long-term assets¶" & \n'
     '                        "IntangibleAssets   →(Required) Deducted from total assets before '
     'the comparison¶"',
     # the stored formula is one line, and it is XML, so the joining ampersand is escaped
     '"InterestExpenses   →(Required) Interest Expense ¶"',
     '"TotalAssets        →(Required) Both current and long-term assets¶" &amp; '
     '"IntangibleAssets   →(Required) Deducted from total assets before the comparison¶"'),
    # TimelineOffsetλ's worked example is the one line in the library a reader cannot copy:
    # the call is missing the two closing brackets that finish EDATE and the function call
    # itself, so pasting it gets a syntax error rather than an answer, and the Result column
    # beside it is empty where every other example prints what it returns. Upstream wrote it
    # against 2/15/2022 and a timeline starting 1/1/2023; the date sweep moved both forward
    # two years with everything else, which still left it two years behind the 1 Jan 2026
    # timeline the demonstration sheet builds. Rewritten as two rows against that same
    # timeline, one date inside it and one before it, since a date before the timeline is
    # what the function's own DISCUSSION comment is for. Both results are checked in a real
    # Excel by tools/excel_selftest.ps1. The name is still the module-qualified one here:
    # the flat nb. namespace is applied much further down.
    ("nabla.f", "TimelineOffsetλ",
     '                                           "→=nabla.f.TimelineOffsetλ(""15/2/2024"", '
     'EDATE(""1/1/2025"", SEQUENCE( , 12, 0)"',
     '                            "1              →=nabla.f.TimelineOffsetλ(""15/2/2026"", '
     'EDATE(""1/1/2026"", SEQUENCE( , 12, 0)))¶" &\n'
     '                            "-11            →=nabla.f.TimelineOffsetλ(""15/2/2025"", '
     'EDATE(""1/1/2026"", SEQUENCE( , 12, 0)))"',
     '"→=nabla.f.TimelineOffsetλ(""15/2/2024"", EDATE(""1/1/2025"", SEQUENCE( , 12, 0)"',
     '"1              →=nabla.f.TimelineOffsetλ(""15/2/2026"", EDATE(""1/1/2026"", '
     'SEQUENCE( , 12, 0)))¶" &amp; "-11            →=nabla.f.TimelineOffsetλ(""15/2/2025"", '
     'EDATE(""1/1/2026"", SEQUENCE( , 12, 0)))"'),
    # and its parameter table misspells the word the function is named after
    ("nabla.f", "TimelineOffsetλ", "A model's timline (Row", "A model's timeline (Row"),
    # TimelinePositionλ's parameter table misspells the word the function is named after.
    # TimelineOffsetλ's copy of the same row was corrected in v2.5.0; this is the last one.
    ("nabla.f", "TimelinePositionλ", "A model's timline", "A model's timeline"),
    # and LabelDepreciateλ's table transposes "the"
    ("nabla.f", "LabelDepreciateλ", "each asset in teh depreciation", "each asset in the depreciation"),
    # Amortiseλ now answers on a timeline shorter than a month, so its description says what
    # it does there rather than leaving a reader to infer it from a spilled block of noughts.
    # This rewrites the line rather than adding one. A function's help spills down its own
    # demonstration sheet, and this one has a single free row beneath it before the sample
    # data starts; two added rows blocked the spill and the whole help block came back an
    # error. Depreciateλ's sheet below has seven free rows and can take the two it gains.
    ("nabla.f", "Amortiseλ",
     '"→It assumes all payments are made monthly.¶"',
     '"→Payments are assumed monthly. On periods shorter than a month the figures for '
     'each month land in the period holding its start, the rest nil.¶"'),
    # Depreciateλ's timeline row named the three intervals its dead SWITCH lookups listed.
    # It has always accepted any interval and now collects the final period on the short ones.
    ("nabla.f", "Depreciateλ",
     "→Timeline can be in Months, Quarters, or Years¶",
     "→Timeline can be any interval: days, weeks, months, quarters or years¶"),
    # and its life row states the bound the function now enforces
    ("nabla.f", "Depreciateλ",
     '                            "LifeInYears    →(Required) The number of years with which to depreciate each asset¶" & \n',
     '                            "LifeInYears    →(Required) The number of years with which to depreciate each asset.¶" & \n'
     '                            "               →More than 0 and no more than 100. A date here is not a life:¶" & \n'
     '                            "               →the arguments read InitialValues, InServiceDates, LifeInYears, Timeline.¶" & \n',
     '"LifeInYears    →(Required) The number of years with which to depreciate each asset¶" &amp; ',
     '"LifeInYears    →(Required) The number of years with which to depreciate each asset.¶" &amp; '
     '"               →More than 0 and no more than 100. A date here is not a life:¶" &amp; '
     '"               →the arguments read InitialValues, InServiceDates, LifeInYears, Timeline.¶" &amp; '),
]


def stores(entry):
    """One correction as (module, function, source pair, defined-name pair).

    The two forms differ only when a replacement spans more than one row: the module
    source wraps its help across lines and indents the continuation, while the stored
    formula is a single line joining the same literals with a plain ampersand.
    """
    module, fn, old, new = entry[:4]
    return module, fn, (old, new), (entry[4], entry[5]) if len(entry) == 6 else (old, new)

# IsInListλ is the one case where the declaration is the odd one out: it shouts LIST,
# while its own signature, its parameter table and one of its two references all write
# List. Excel resolves identifiers case-insensitively, so this renames the parameter to
# match everything else rather than making the help shout back.
ISINLIST = ("nabla.e", "nabla.u")


def function_block(module, fn):
    """The span of one function's definition in its module source."""
    text = mods[module]["text"]
    start = re.search(r"(?m)^%s\s*=\s*LAMBDA" % re.escape(fn), text)
    assert start, (fn, module, "no declaration")
    a = start.start()
    return a, text.index("\n);", a) + 3


# ---------- the comment banners above each function ----------
# Each function is introduced by a comment naming it and describing it. Those comments are
# stripped before anything is stored in the workbook, so no part of the build and no gate
# has ever read them, and they are the one piece of the published source that drifted with
# nothing watching. Three name a function that does not exist and three describe a
# different one. Each is anchored to the name above it, because the wording on its own is
# also correct somewhere else: "Count numbers in each row" is right above CountRowsλ.
BANNER_FIXES = [
    ("nabla.e", "CountColsλ", "Count numbers in each row", "Count numbers in each column"),
    ("nabla.u", "CountColsλ", "Count numbers in each row", "Count numbers in each column"),
    ("nabla.e", "CountAColsλ", "Count everything in each row", "Count everything in each column"),
    ("nabla.u", "CountAColsλ", "Count everything in each row", "Count everything in each column"),
    ("nabla.e", "IsInListλ", "Determine if a value is between a lower and upper limit",
     "Determine whether a value is one of a list of items"),
    ("nabla.u", "IsInListλ", "Determine if a value is between a lower and upper limit",
     "Determine whether a value is one of a list of items"),
]

# and three banners misspell the name they announce
BANNER_NAMES = [
    ("nabla.f", "FUNCTION NAME:  →SumContainsλ", "FUNCTION NAME:  SumContainsλ"),
    ("nabla.r", "FUNCTION NAME:  InterestCoverateRatioλ", "FUNCTION NAME:  InterestCoverageRatioλ"),
    ("nabla.r", "FUNCTION NAME:  PriceToCashsRatioλ", "FUNCTION NAME:  PriceToCashRatioλ"),
]

for _mod, _fn, _old, _new in BANNER_FIXES:
    _pat = re.compile(r"(FUNCTION NAME:\s+%s\s+DESCRIPTION:\*//\*\*)%s(\*/)"
                      % (re.escape(_fn), re.escape(_old)))
    mods[_mod]["text"], _n = _pat.subn(lambda m, w=_new: m.group(1) + w + m.group(2),
                                       mods[_mod]["text"])
    assert _n == 1, (_mod, _fn, _n)

for _mod, _old, _new in BANNER_NAMES:
    assert mods[_mod]["text"].count(_old) == 1, (_mod, _old, mods[_mod]["text"].count(_old))
    mods[_mod]["text"] = mods[_mod]["text"].replace(_old, _new)

print("corrected %d comment banners that described the wrong function and %d that misnamed it"
      % (len(BANNER_FIXES), len(BANNER_NAMES)))

for _entry in HELP_SIGNATURES:
    _mod, _fn, (_old, _new), _ = stores(_entry)
    _a, _b = function_block(_mod, _fn)
    _text = mods[_mod]["text"]
    _blk = _text[_a:_b]
    assert _blk.count(_old) == 1, (_fn, _mod, "source", _blk.count(_old))
    mods[_mod]["text"] = _text[:_a] + _blk.replace(_old, _new) + _text[_b:]

# IsInListλ also builds its help wrong. TEXTSPLIT takes the text, then a column
# delimiter, then an optional row delimiter, and this call supplies only one: the
# arrow that should separate the two columns was left concatenated onto the end of
# the text, and the pilcrow that should end each row became the column delimiter. The
# help therefore returns a single 11-column row instead of an 11-row table, and spills
# sideways across the sheet. The other 125 functions supply both. Restore the column
# delimiter; the trailing arrow stays, as the empty last row it was meant to be.
TEXTSPLIT_ARGS = ('"→", "¶"', '"→", "→", "¶"')

for _mod in ISINLIST:
    _a, _b = function_block(_mod, "IsInListλ")
    _text = mods[_mod]["text"]
    _blk, _n = re.subn(r"\bLIST\b", "List", _text[_a:_b])
    assert _n == 2, (_mod, "source", _n)      # the declaration and one of two references
    _old, _new = TEXTSPLIT_ARGS
    assert _blk.count(_old) == 1, (_mod, "source", "TEXTSPLIT", _blk.count(_old))
    mods[_mod]["text"] = _text[:_a] + _blk.replace(_old, _new) + _text[_b:]

print("corrected %d help signatures in the module sources, renamed IsInListλ's LIST "
      "parameter and restored its help's column delimiter" % len(HELP_SIGNATURES))

# ---------- defects in the functions themselves ----------
# Everything above corrects what a function says about itself. These change what a function
# computes. All were reported by an outside review of v1.2.6 and confirmed in Excel, and in
# several the workbook's own worked example is the evidence: the printed answer is the
# answer the bug produces.
#
# They take the same shape as the table above, so they run through the same two passes:
# the module source here, the defined name Excel installs further down. The entries that
# carry six values give the stored form separately, because Excel prefixes every parameter
# with _xlpm. and writes < and > as XML entities.
LOGIC_FIXES = [
    # OverLapDaysλ converts all four of its dates, so that a date written as text becomes
    # a serial number, and then compares the raw arguments anyway. The four conversions
    # are never read. Two text dates therefore compare as text, which orders "17/1/2025"
    # before "7/1/2025", and the help's own third example claims 12 shared days for two
    # January 2025 periods that share 2. Compare the converted values, which is what they
    # are for. Numbers and real dates are unaffected: converting one returns it unchanged.
    ("nabla.d", "OverLapDaysλ",
     "IF(Period2End <= Period1End, Period2End, Period1End)",
     "IF(CvtP2End <= CvtP1End, CvtP2End, CvtP1End)",
     "IF(_xlpm.Period2End &lt;= _xlpm.Period1End, _xlpm.Period2End, _xlpm.Period1End)",
     "IF(_xlpm.CvtP2End &lt;= _xlpm.CvtP1End, _xlpm.CvtP2End, _xlpm.CvtP1End)"),
    ("nabla.d", "OverLapDaysλ",
     "IF(Period2Start >= Period1Start, Period2Start, Period1Start)",
     "IF(CvtP2Start >= CvtP1Start, CvtP2Start, CvtP1Start)",
     "IF(_xlpm.Period2Start &gt;= _xlpm.Period1Start, _xlpm.Period2Start, _xlpm.Period1Start)",
     "IF(_xlpm.CvtP2Start &gt;= _xlpm.CvtP1Start, _xlpm.CvtP2Start, _xlpm.CvtP1Start)"),
    ("nabla.d", "OverLapDaysλ",
     '"12             →=nabla.d.OverLapDaysλ', '"2              →=nabla.d.OverLapDaysλ'),
    # Periodsλ promises "Returns negative values if Date1 is after Date2" and cannot: the
    # difference is floored at 1 before SIGN sees it, so the sign is always +1 and the two
    # negative examples in its own help are unreachable. Take the sign of the difference
    # itself. Equal dates now give SIGN 0 rather than 1, which changes nothing, because
    # DATEDIF of a date with itself is 0 either way.
    ("nabla.d", "Periodsλ",
     "SIGN(Max(DateTwo - DateOne, 1))", "SIGN(DateTwo - DateOne)",
     "SIGN(MAX(_xlpm.DateTwo - _xlpm.DateOne, 1))", "SIGN(_xlpm.DateTwo - _xlpm.DateOne)"),
    # VDBλ declares No_Switch, defaults it to FALSE, and then calls VDB without it, so the
    # argument does nothing. Its help demonstrates the function with No_Switch TRUE and
    # prints the FALSE answer, which is how it went unnoticed. Excel gives 300.00, 210.00,
    # 147.00, 102.90, 72.03 once the argument is passed, against the 121.50, 121.50 tail
    # that switching to straight line produces.
    ("nabla.f", "VDBλ",
     "VDB( Cost, Salvage, Life, SEQUENCE( , Life, 0), SEQUENCE( , Life), Factor)",
     "VDB( Cost, Salvage, Life, SEQUENCE( , Life, 0), SEQUENCE( , Life), Factor, No_Switch)",
     "VDB(_xlpm.Cost, _xlpm.Salvage, _xlpm.Life, _xlfn.SEQUENCE(, _xlpm.Life, 0), "
     "_xlfn.SEQUENCE(, _xlpm.Life), _xlpm.Factor)",
     "VDB(_xlpm.Cost, _xlpm.Salvage, _xlpm.Life, _xlfn.SEQUENCE(, _xlpm.Life, 0), "
     "_xlfn.SEQUENCE(, _xlpm.Life), _xlpm.Factor, _xlpm.No_Switch)"),
    ("nabla.f", "VDBλ",
     "→300.00,210.00,147.00,121.50,121.50", "→300.00,210.00,147.00,102.90,72.03"),
    # Periodsλ counted complete intervals where every one of its four worked examples counts
    # the period starts crossed between the two dates. It says so itself: the description
    # lists "End Date is inclusive" as a difference from DATEDIF, and then the procedure
    # calls DATEDIF. From 31 March to 15 May is one complete month and two month starts, and
    # the help says 2. Count ordinals instead, the same way for all five intervals, so a part
    # period at the end counts and a whole one does not count twice. "D" is unchanged, since
    # a day ordinal is the serial number itself. The week ordinal follows PeriodLabelλ's own
    # week numbering, which restarts each 1 January and therefore labels every year with 53
    # weeks, the last of them one or two days long. That is what makes the help's fourth
    # example 53 rather than the 52 whole weeks the two dates are apart.
    ("nabla.d", "Periodsλ",
     "        DPW, 7,     //Days Per Week\n",
     "        DPW, 7,     //Days Per Week\n"
     "        WPY, 53,    //Week labels Per Year: the last one runs 1 or 2 days\n"
     "        QPY, 4,     //Quarters Per Year\n",
     "_xlpm.DPW, 7, ",
     "_xlpm.DPW, 7, _xlpm.WPY, 53, _xlpm.QPY, 4, "),
    ("nabla.d", "Periodsλ",
     '                                Latest,     MAX(DateOne, DateTwo) , //+ 1,\n'
     '                                Sign,       SIGN(DateTwo - DateOne),\n'
     '                                Periods,    Switch(Interval,   \n'
     '                                                "D", DATEDIF(Earliest, Latest, "D"),\n'
     '                                                "W", INT(DATEDIF(Earliest, Latest, "D")/DPW),\n'
     '                                                "M", DATEDIF(Earliest, Latest, "M"),\n'
     '                                                "Q", INT(DATEDIF(Earliest, Latest, "M")/MPQ),\n'
     '                                                "Y", DATEDIF(Earliest, Latest, "Y")\n'
     '                                            ),\n',
     '                                Latest,     MAX(DateOne, DateTwo),\n'
     '                                Sign,       SIGN(DateTwo - DateOne),\n'
     '                                Periods,    Switch(Interval,   \n'
     '                                                "D", Latest - Earliest,\n'
     '                                                "W", WPY * (YEAR(Latest) - YEAR(Earliest))\n'
     '                                                     + QUOTIENT(Latest   - DATE(YEAR(Latest),   1, 0) - 1, DPW)\n'
     '                                                     - QUOTIENT(Earliest - DATE(YEAR(Earliest), 1, 0) - 1, DPW),\n'
     '                                                "M", MPY * (YEAR(Latest) - YEAR(Earliest))\n'
     '                                                     + MONTH(Latest) - MONTH(Earliest),\n'
     '                                                "Q", QPY * (YEAR(Latest) - YEAR(Earliest))\n'
     '                                                     + QUOTIENT(MONTH(Latest)   - 1, MPQ)\n'
     '                                                     - QUOTIENT(MONTH(Earliest) - 1, MPQ),\n'
     '                                                "Y", YEAR(Latest) - YEAR(Earliest)\n'
     '                                            ),\n',
     '_xlpm.Latest, MAX(_xlpm.DateOne, _xlpm.DateTwo), _xlpm.Sign, '
     'SIGN(_xlpm.DateTwo - _xlpm.DateOne), _xlpm.Periods, _xlfn.SWITCH(_xlpm.Interval, '
     '"D", DATEDIF(_xlpm.Earliest, _xlpm.Latest, "D"), '
     '"W", INT(DATEDIF(_xlpm.Earliest, _xlpm.Latest, "D") / _xlpm.DPW), '
     '"M", DATEDIF(_xlpm.Earliest, _xlpm.Latest, "M"), '
     '"Q", INT(DATEDIF(_xlpm.Earliest, _xlpm.Latest, "M") / _xlpm.MPQ), '
     '"Y", DATEDIF(_xlpm.Earliest, _xlpm.Latest, "Y")), ',
     '_xlpm.Latest, MAX(_xlpm.DateOne, _xlpm.DateTwo), _xlpm.Sign, '
     'SIGN(_xlpm.DateTwo - _xlpm.DateOne), _xlpm.Periods, _xlfn.SWITCH(_xlpm.Interval, '
     '"D", _xlpm.Latest - _xlpm.Earliest, '
     '"W", _xlpm.WPY * (YEAR(_xlpm.Latest) - YEAR(_xlpm.Earliest)) '
     '+ QUOTIENT(_xlpm.Latest - DATE(YEAR(_xlpm.Latest), 1, 0) - 1, _xlpm.DPW) '
     '- QUOTIENT(_xlpm.Earliest - DATE(YEAR(_xlpm.Earliest), 1, 0) - 1, _xlpm.DPW), '
     '"M", _xlpm.MPY * (YEAR(_xlpm.Latest) - YEAR(_xlpm.Earliest)) '
     '+ MONTH(_xlpm.Latest) - MONTH(_xlpm.Earliest), '
     '"Q", _xlpm.QPY * (YEAR(_xlpm.Latest) - YEAR(_xlpm.Earliest)) '
     '+ QUOTIENT(MONTH(_xlpm.Latest) - 1, _xlpm.MPQ) '
     '- QUOTIENT(MONTH(_xlpm.Earliest) - 1, _xlpm.MPQ), '
     '"Y", YEAR(_xlpm.Latest) - YEAR(_xlpm.Earliest)), '),
    # and the third example passes four arguments to a function that takes three, so it is
    # the one line in this help a reader cannot copy. The -12 it claims is right without it.
    ("nabla.d", "Periodsλ", '", , FALSE)', '")'),
    # the description's shorthand for what changed, in the row that already carried it
    ("nabla.d", "Periodsλ", "* End Date is inclusive",
     "* Counts the period starts crossed, so a part period at the end counts"),
    # Four more functions convert a date argument and then use the raw one, the same defect
    # OverLapDaysλ had. None is reachable through a demonstration sheet, because every one of
    # them is called there with real dates, and converting a date returns it unchanged. Pass
    # any of them a date written as text and the conversion is still thrown away.
    ("nabla.d", "PeriodLabelλ",
     '        Result,     SWITCH(Interval,\n'
     '                        "D", TEXT(Date, "yyyy-mmm-dd"),\n'
     '                        "W", YEAR(Date) & ":W" & TEXT(QUOTIENT(Date - DATE(YEAR(Date), 1, 0) - 1, 7) + 1, "00"),\n'
     '                        "I", YEAR(Date) & ":W" & TEXT(ISOWEEKNUM(Date), "00"),\n'
     '                        "M", TEXT(Date, "YYYY-MMM"),\n'
     '                        "Q", YEAR(Date) & ":Q" & QUOTIENT(MONTH(Date) - 1, 3) + 1,\n'
     '                        "S", YEAR(Date) & ":S" & QUOTIENT(MONTH(Date) - 1, 6) + 1,\n'
     '                        "A", TEXT(Date, "YYYY"),\n'
     '                        "Y", TEXT(Date, "YYYY"),',
     '        Result,     SWITCH(Interval,\n'
     '                        "D", TEXT(CvtDate, "yyyy-mmm-dd"),\n'
     '                        "W", YEAR(CvtDate) & ":W" & TEXT(QUOTIENT(CvtDate - DATE(YEAR(CvtDate), 1, 0) - 1, 7) + 1, "00"),\n'
     '                        "I", YEAR(CvtDate) & ":W" & TEXT(ISOWEEKNUM(CvtDate), "00"),\n'
     '                        "M", TEXT(CvtDate, "YYYY-MMM"),\n'
     '                        "Q", YEAR(CvtDate) & ":Q" & QUOTIENT(MONTH(CvtDate) - 1, 3) + 1,\n'
     '                        "S", YEAR(CvtDate) & ":S" & QUOTIENT(MONTH(CvtDate) - 1, 6) + 1,\n'
     '                        "A", TEXT(CvtDate, "YYYY"),\n'
     '                        "Y", TEXT(CvtDate, "YYYY"),',
     '_xlfn.SWITCH(_xlpm.Interval, "D", TEXT(_xlpm.Date, "yyyy-mmm-dd"), '
     '"W", YEAR(_xlpm.Date) &amp; ":W" &amp; TEXT(QUOTIENT(_xlpm.Date - DATE(YEAR(_xlpm.Date), 1, 0) - 1, 7) + 1, "00"), '
     '"I", YEAR(_xlpm.Date) &amp; ":W" &amp; TEXT(_xlfn.ISOWEEKNUM(_xlpm.Date), "00"), '
     '"M", TEXT(_xlpm.Date, "YYYY-MMM"), '
     '"Q", YEAR(_xlpm.Date) &amp; ":Q" &amp; QUOTIENT(MONTH(_xlpm.Date) - 1, 3) + 1, '
     '"S", YEAR(_xlpm.Date) &amp; ":S" &amp; QUOTIENT(MONTH(_xlpm.Date) - 1, 6) + 1, '
     '"A", TEXT(_xlpm.Date, "YYYY"), "Y", TEXT(_xlpm.Date, "YYYY"), #VALUE!',
     '_xlfn.SWITCH(_xlpm.Interval, "D", TEXT(_xlpm.CvtDate, "yyyy-mmm-dd"), '
     '"W", YEAR(_xlpm.CvtDate) &amp; ":W" &amp; TEXT(QUOTIENT(_xlpm.CvtDate - DATE(YEAR(_xlpm.CvtDate), 1, 0) - 1, 7) + 1, "00"), '
     '"I", YEAR(_xlpm.CvtDate) &amp; ":W" &amp; TEXT(_xlfn.ISOWEEKNUM(_xlpm.CvtDate), "00"), '
     '"M", TEXT(_xlpm.CvtDate, "YYYY-MMM"), '
     '"Q", YEAR(_xlpm.CvtDate) &amp; ":Q" &amp; QUOTIENT(MONTH(_xlpm.CvtDate) - 1, 3) + 1, '
     '"S", YEAR(_xlpm.CvtDate) &amp; ":S" &amp; QUOTIENT(MONTH(_xlpm.CvtDate) - 1, 6) + 1, '
     '"A", TEXT(_xlpm.CvtDate, "YYYY"), "Y", TEXT(_xlpm.CvtDate, "YYYY"), #VALUE!'),
    ("nabla.d", "ScheduleRatesλ",
     'XLOOKUP(PeriodEnds, RateStarts, Rates, "", -1)',
     'XLOOKUP(CvtEnds, CvtStarts, Rates, "", -1)',
     '_xlfn.XLOOKUP(_xlpm.PeriodEnds, _xlpm.RateStarts, _xlpm.Rates, "", -1)',
     '_xlfn.XLOOKUP(_xlpm.CvtEnds, _xlpm.CvtStarts, _xlpm.Rates, "", -1)'),
    ("nabla.d", "ScheduleValuesλ",
     '                            SEQUENCE(, COLUMNS(PeriodStarts)), \n'
     '                            LAMBDA(Period, \n'
     '                                LET(StartDate,  INDEX(PeriodStarts, Period), \n'
     '                                    EndDate,    INDEX(PeriodEnds, Period),',
     '                            SEQUENCE(, COLUMNS(CvtPrdStarts)), \n'
     '                            LAMBDA(Period, \n'
     '                                LET(StartDate,  INDEX(CvtPrdStarts, Period), \n'
     '                                    EndDate,    INDEX(CvtPrdEnds, Period),',
     '_xlfn.SEQUENCE(, COLUMNS(_xlpm.PeriodStarts)), _xlfn.LAMBDA(_xlpm.Period, '
     '_xlfn.LET(_xlpm.StartDate, INDEX(_xlpm.PeriodStarts, _xlpm.Period), '
     '_xlpm.EndDate, INDEX(_xlpm.PeriodEnds, _xlpm.Period),',
     '_xlfn.SEQUENCE(, COLUMNS(_xlpm.CvtPrdStarts)), _xlfn.LAMBDA(_xlpm.Period, '
     '_xlfn.LET(_xlpm.StartDate, INDEX(_xlpm.CvtPrdStarts, _xlpm.Period), '
     '_xlpm.EndDate, INDEX(_xlpm.CvtPrdEnds, _xlpm.Period),'),
    ("nabla.d", "Timelineλ",
     '                            "Y", EDATE(StartDate, SEQUENCE(1, Periods, EndDates * MPY, MPY)) - EndDates,\n'
     '                            "Q", EDATE(StartDate, SEQUENCE(1, Periods, EndDates * MPQ, MPQ)) - EndDates,\n'
     '                            "M", EDATE(StartDate, SEQUENCE(1, Periods, EndDates, 1)) - EndDates,\n'
     '                            "W", SEQUENCE(1, Periods, StartDate + IF(PeriodStarts?, 0, DPW - 1), DPW),\n'
     '                            "D", SEQUENCE(1, Periods, StartDate, 1),',
     '                            "Y", EDATE(CvtStart, SEQUENCE(1, Periods, EndDates * MPY, MPY)) - EndDates,\n'
     '                            "Q", EDATE(CvtStart, SEQUENCE(1, Periods, EndDates * MPQ, MPQ)) - EndDates,\n'
     '                            "M", EDATE(CvtStart, SEQUENCE(1, Periods, EndDates, 1)) - EndDates,\n'
     '                            "W", SEQUENCE(1, Periods, CvtStart + IF(PeriodStarts?, 0, DPW - 1), DPW),\n'
     '                            "D", SEQUENCE(1, Periods, CvtStart, 1),',
     '"Y", EDATE(_xlpm.StartDate, _xlfn.SEQUENCE(1, _xlpm.Periods, _xlpm.EndDates * _xlpm.MPY, _xlpm.MPY)) - _xlpm.EndDates, '
     '"Q", EDATE(_xlpm.StartDate, _xlfn.SEQUENCE(1, _xlpm.Periods, _xlpm.EndDates * _xlpm.MPQ, _xlpm.MPQ)) - _xlpm.EndDates, '
     '"M", EDATE(_xlpm.StartDate, _xlfn.SEQUENCE(1, _xlpm.Periods, _xlpm.EndDates, 1)) - _xlpm.EndDates, '
     '"W", _xlfn.SEQUENCE(1, _xlpm.Periods, _xlpm.StartDate + IF(_xlpm.PeriodStarts?, 0, _xlpm.DPW - 1), _xlpm.DPW), '
     '"D", _xlfn.SEQUENCE(1, _xlpm.Periods, _xlpm.StartDate, 1),',
     '"Y", EDATE(_xlpm.CvtStart, _xlfn.SEQUENCE(1, _xlpm.Periods, _xlpm.EndDates * _xlpm.MPY, _xlpm.MPY)) - _xlpm.EndDates, '
     '"Q", EDATE(_xlpm.CvtStart, _xlfn.SEQUENCE(1, _xlpm.Periods, _xlpm.EndDates * _xlpm.MPQ, _xlpm.MPQ)) - _xlpm.EndDates, '
     '"M", EDATE(_xlpm.CvtStart, _xlfn.SEQUENCE(1, _xlpm.Periods, _xlpm.EndDates, 1)) - _xlpm.EndDates, '
     '"W", _xlfn.SEQUENCE(1, _xlpm.Periods, _xlpm.CvtStart + IF(_xlpm.PeriodStarts?, 0, _xlpm.DPW - 1), _xlpm.DPW), '
     '"D", _xlfn.SEQUENCE(1, _xlpm.Periods, _xlpm.CvtStart, 1),'),
    # The two ByItems functions carry the same defect their siblings had until v2.2.0, and
    # the Cvt check added then does not see it: each one does read its conversion, but only
    # to hand it to the recursive call, so the first row of the result is computed from the
    # raw argument and every row below it from the converted one, off the same call.
    ("nabla.d", "ScheduleValuesByItemsλ",
     "(ItemDates >= PeriodStarts) * (ItemDates <= PeriodEnds)",
     "(ItemDates >= CvtPrdStarts) * (ItemDates <= CvtPrdEnds)",
     "(_xlpm.ItemDates &gt;= _xlpm.PeriodStarts) * (_xlpm.ItemDates &lt;= _xlpm.PeriodEnds)",
     "(_xlpm.ItemDates &gt;= _xlpm.CvtPrdStarts) * (_xlpm.ItemDates &lt;= _xlpm.CvtPrdEnds)"),
    ("nabla.d", "ScheduleRatesByItemsλ",
     "XLOOKUP( PeriodEnds, ItemDates, ItemRates, 0, -1)",
     "XLOOKUP( CvtPrdEnds, ItemDates, ItemRates, 0, -1)",
     "_xlfn.XLOOKUP(_xlpm.PeriodEnds, _xlpm.ItemDates, _xlpm.ItemRates, 0, -1)",
     "_xlfn.XLOOKUP(_xlpm.CvtPrdEnds, _xlpm.ItemDates, _xlpm.ItemRates, 0, -1)"),
    # PeriodStartλ walks the calendar a month at a time and then repairs what that walk
    # got wrong, and the repair is a single day: where the anchor is the 31st and the
    # month it lands in is shorter, DATE() rolls the date into the next month and the
    # correction takes one day back off it. Anchored on 31 January, monthly, the period
    # containing 5 March came back as 2 March, which is neither a period start nor even
    # in the right month, and 28 February, which is, was unreachable. Anchors on the 29th,
    # 30th and 31st are all affected, and 29 February with them; every other anchor was
    # already right. Step the anchor by whole periods with EDATE instead, which is what
    # defines a monthly schedule off a month end: it holds the day where the target month
    # has one and clamps to the month's end where it does not. Checked against every date
    # from 2024 to 2030 at 13-day steps, for eleven anchors and five period lengths: 169
    # of 9,900 came back wrong, all of them month-end anchors.
    ("nabla.f", "PeriodStartλ",
     '    //  Procedure\n'
     '        DateDay,            DAY(DateOfInterest), \n'
     '        PeriodDay,          Day(AnyPeriodStart),\n'
     '        DateMonth,          MONTH(DateOfInterest), \n'
     '        PeriodMonth,        MONTH(AnyPeriodStart),\n'
     '        MonthModDifference, MOD(DateMonth - PeriodMonth, MpP),\n'
     '        NeedPriorPeriod?,   AND( MonthModDifference = 0, DateDay < PeriodDay) * - MpP,\n'
     '        NewPeriodMonth,     MOD(DateMonth - MonthModDifference + NeedPriorPeriod? -1, MpY ) + 1,\n'
     '        YearCorrection,     -((DateMonth * 100 + DateDay) < (NewPeriodMonth * 100 + PeriodDay )),\n'
     '        PeriodDate,         DATE(YEAR(DateOfInterest) + YearCorrection,  NewPeriodMonth, PeriodDay),\n'
     '        LastDayOfMonthAdj,  PeriodDate - (MONTH(PeriodDate) <> NewPeriodMonth),\n'
     '        Result,             LastDayOfMonthAdj,\n',
     '    //  Check inputs - If a date is in text, convert to value\n'
     '        CvtStart,           IF( ISTEXT( AnyPeriodStart), DATEVALUE( AnyPeriodStart), AnyPeriodStart),\n'
     '        CvtDate,            IF( ISTEXT( DateOfInterest), DATEVALUE( DateOfInterest), DateOfInterest),\n'
     '    //  Procedure\n'
     '    //  Count the whole periods from the anchor to the date of interest and step the\n'
     '    //  anchor on by that many.\n'
     '        MonthDiff,          (YEAR( CvtDate) - YEAR( CvtStart)) * MpY + MONTH( CvtDate) - MONTH( CvtStart),\n'
     '        WholePeriods,       QUOTIENT( MonthDiff, MpP),\n'
     '        Candidate,          EDATE( CvtStart, WholePeriods * MpP),\n'
     '    //  QUOTIENT truncates towards zero, and the day of the month can leave the candidate\n'
     '    //  past the date of interest, so step back a period where it does. One step is always\n'
     '    //  enough: a truncated quotient is never more than one period out.\n'
     '        Result,             IF( Candidate > CvtDate, EDATE( CvtStart, (WholePeriods - 1) * MpP), Candidate),\n',
     '_xlpm.DateDay, DAY(_xlpm.DateOfInterest), _xlpm.PeriodDay, DAY(_xlpm.AnyPeriodStart), '
     '_xlpm.DateMonth, MONTH(_xlpm.DateOfInterest), _xlpm.PeriodMonth, MONTH(_xlpm.AnyPeriodStart), '
     '_xlpm.MonthModDifference, MOD(_xlpm.DateMonth - _xlpm.PeriodMonth, _xlpm.MpP), '
     '_xlpm.NeedPriorPeriod?, AND(_xlpm.MonthModDifference = 0, _xlpm.DateDay &lt; _xlpm.PeriodDay) '
     '* -_xlpm.MpP, _xlpm.NewPeriodMonth, MOD(_xlpm.DateMonth - _xlpm.MonthModDifference + '
     '_xlpm.NeedPriorPeriod? - 1, _xlpm.MpY) + 1, _xlpm.YearCorrection, '
     '-((_xlpm.DateMonth * 100 + _xlpm.DateDay) &lt; (_xlpm.NewPeriodMonth * 100 + _xlpm.PeriodDay)), '
     '_xlpm.PeriodDate, DATE(YEAR(_xlpm.DateOfInterest) + _xlpm.YearCorrection, _xlpm.NewPeriodMonth, '
     '_xlpm.PeriodDay), _xlpm.LastDayOfMonthAdj, _xlpm.PeriodDate - (MONTH(_xlpm.PeriodDate) '
     '&lt;&gt; _xlpm.NewPeriodMonth), _xlpm.Result, _xlpm.LastDayOfMonthAdj, ',
     '_xlpm.CvtStart, IF(ISTEXT(_xlpm.AnyPeriodStart), DATEVALUE(_xlpm.AnyPeriodStart), '
     '_xlpm.AnyPeriodStart), _xlpm.CvtDate, IF(ISTEXT(_xlpm.DateOfInterest), '
     'DATEVALUE(_xlpm.DateOfInterest), _xlpm.DateOfInterest), _xlpm.MonthDiff, '
     '(YEAR(_xlpm.CvtDate) - YEAR(_xlpm.CvtStart)) * _xlpm.MpY + MONTH(_xlpm.CvtDate) '
     '- MONTH(_xlpm.CvtStart), _xlpm.WholePeriods, QUOTIENT(_xlpm.MonthDiff, _xlpm.MpP), '
     '_xlpm.Candidate, EDATE(_xlpm.CvtStart, _xlpm.WholePeriods * _xlpm.MpP), _xlpm.Result, '
     'IF(_xlpm.Candidate &gt; _xlpm.CvtDate, EDATE(_xlpm.CvtStart, (_xlpm.WholePeriods - 1) '
     '* _xlpm.MpP), _xlpm.Candidate), '),
    # TimelineOffsetλ reads a timeline's interval off its first two dates and converts it
    # to whole months. A daily, weekly or fortnightly timeline is no whole number of months
    # long, so that rounds to zero and the next line divides by it: every such call returned
    # #DIV/0!, and nb.Amortiseλ calls this on every timeline it is given. Those intervals are
    # a fixed number of days, which is exactly what makes them easy: count the days instead.
    # The month path is untouched, and gives the same answer it always did on monthly,
    # quarterly and yearly timelines, including those anchored on a month end.
    ("nabla.f", "TimelineOffsetλ",
     '        MpP,            ROUND( (Period2 - Period1)/30.5, 0), //Months Per Period\n',
     '        DpP,            Period2 - Period1,                    //Days Per Period\n'
     '        MpP,            MAX( ROUND( DpP/30.5, 0), 1),         //Months Per Period, never zero\n'
     '        SubMonthly?,    ROUND( DpP/30.5, 0) = 0,\n',
     '_xlpm.MpP, ROUND((_xlpm.Period2 - _xlpm.Period1) / 30.5, 0), ',
     '_xlpm.DpP, _xlpm.Period2 - _xlpm.Period1, _xlpm.MpP, MAX(ROUND(_xlpm.DpP / 30.5, 0), 1), '
     '_xlpm.SubMonthly?, ROUND(_xlpm.DpP / 30.5, 0) = 0, '),
    ("nabla.f", "TimelineOffsetλ",
     '        Result,         MATCH( Date, SearchTimeline, 1) - IF( Direction = 1, 1, PeriodDiff + 2),        \n',
     '        ByMonth,        MATCH( Date, SearchTimeline, 1) - IF( Direction = 1, 1, PeriodDiff + 2),\n'
     '    //  A sub-monthly timeline is a fixed number of days per period, so the offset is the\n'
     '    //  whole number of periods between the timeline start and the date, floored, which\n'
     '    //  puts a date before the timeline in the period that ends where the timeline begins.\n'
     '        Result,         IF( SubMonthly?, INT( (Date - Period1) / DpP), ByMonth),\n',
     '_xlpm.Result, MATCH(_xlpm.Date, _xlpm.SearchTimeline, 1) - IF(_xlpm.Direction = 1, 1, '
     '_xlpm.PeriodDiff + 2), ',
     '_xlpm.ByMonth, MATCH(_xlpm.Date, _xlpm.SearchTimeline, 1) - IF(_xlpm.Direction = 1, 1, '
     '_xlpm.PeriodDiff + 2), _xlpm.Result, IF(_xlpm.SubMonthly?, INT((_xlpm.Date - _xlpm.Period1) '
     '/ _xlpm.DpP), _xlpm.ByMonth), '),
    # Amortiseλ infers its timeline's period length from the first two dates and rounds it
    # to whole months, which is nought for anything shorter than about a fortnight, and the
    # next two lines divide by it. Every daily, weekly and fortnightly call returned
    # #DIV/0!. Whole-month intervals other than 1, 3 and 12 were never affected: 2 and 6
    # measure correctly and the schedule arithmetic is generic in the count, confirmed in
    # Excel on a six-month timeline. PpY divides twelve by that count and nothing reads it,
    # so it goes rather than gaining a guard it does not need.
    ("nabla.f", "Amortiseλ",
     '        MpP,                @ROUND(( SecondPeriod - FirstPeriod) / ADpM, 0),\n'
     '        PpY,                MpY / MpP,\n',
     '        DpP,                SecondPeriod - FirstPeriod,                 //Days Per Period\n'
     '        LastGap,            MAX( Timeline) - INDEX( Timeline, COUNTA( Timeline) - 1),\n'
     '        SubMonthly?,        DpP < 28,                                   //No month is shorter\n'
     '        MpP,                MAX( @ROUND( DpP / ADpM, 0), 1),            //Months Per Period, never nought\n',
     '_xlpm.MpP, _xlfn.SINGLE(ROUND((_xlpm.SecondPeriod - _xlpm.FirstPeriod) / _xlpm.ADpM, 0)), '
     '_xlpm.PpY, _xlpm.MpY / _xlpm.MpP, ',
     '_xlpm.DpP, _xlpm.SecondPeriod - _xlpm.FirstPeriod, '
     '_xlpm.LastGap, MAX(_xlpm.Timeline) - INDEX(_xlpm.Timeline, COUNTA(_xlpm.Timeline) - 1), '
     '_xlpm.SubMonthly?, _xlpm.DpP &lt; 28, '
     '_xlpm.MpP, MAX(_xlfn.SINGLE(ROUND(_xlpm.DpP / _xlpm.ADpM, 0)), 1), '),
    # Flooring the count at one month is not on its own enough. The schedule is always
    # solved monthly and only then folded into the timeline's periods, and TimelinePositionλ
    # lays the folded block down one period at a time from an offset. On a weekly timeline
    # that dates month two a week after month one and reports twelve months of interest
    # inside a quarter. A sub-monthly period is a fixed number of days, so put each month's
    # figures in the period that contains that month's start and leave the periods between
    # them nil, which is what Depreciateλ has always done on the same timelines. The month
    # path is untouched.
    ("nabla.f", "Amortiseλ",
     '                    NewBlock,           TimelinePositionλ(AmortisationSched, Timeline, Offset),\n',
     '                //  A sub-monthly period is a fixed number of days, so each month lands in\n'
     '                //  the one period that contains its start and the periods between are nil.\n'
     '                //  Rows 2 and 5 are balances rather than flows; a balance is carried into\n'
     '                //  the period it is dated in and nowhere else, which is the same treatment\n'
     '                //  Depreciateλ gives its own opening balance row.\n'
     '                    MonthStarts,        EDATE( StartDate, SEQUENCE( , Periods, 0)),\n'
     '                    ByDate,             MAKEARRAY( ROWS( AmortisationSched), COUNTA( Timeline),\n'
     '                                            LAMBDA( R, C,\n'
     '                                                LET( PeriodOpens,   INDEX( Timeline, C),\n'
     '                                                //  A period ends where the next one opens. Only the last\n'
     '                                                //  has no successor to ask, so it runs on as far as the\n'
     '                                                //  period before it did.\n'
     '                                                     PeriodShut,    IF( C = COUNTA( Timeline),\n'
     '                                                                        PeriodOpens + LastGap,\n'
     '                                                                        INDEX( Timeline, C + 1)),\n'
     '                                                    SUM( CHOOSEROWS( AmortisationSched, R)\n'
     '                                                         * ( MonthStarts >= PeriodOpens)\n'
     '                                                         * ( MonthStarts < PeriodShut))))),\n'
     '                    NewBlock,           IF( SubMonthly?,\n'
     '                                            ByDate,\n'
     '                                            TimelinePositionλ(AmortisationSched, Timeline, Offset)),\n',
     '_xlpm.NewBlock, nabla.f.TimelinePositionλ(_xlpm.AmortisationSched, _xlpm.Timeline, _xlpm.Offset), ',
     '_xlpm.MonthStarts, EDATE(_xlpm.StartDate, _xlfn.SEQUENCE(, _xlpm.Periods, 0)), '
     '_xlpm.ByDate, _xlfn.MAKEARRAY(ROWS(_xlpm.AmortisationSched), COUNTA(_xlpm.Timeline), '
     '_xlfn.LAMBDA(_xlpm.R,_xlpm.C, _xlfn.LET(_xlpm.PeriodOpens, INDEX(_xlpm.Timeline, _xlpm.C), '
     '_xlpm.PeriodShut, IF(_xlpm.C = COUNTA(_xlpm.Timeline), _xlpm.PeriodOpens + _xlpm.LastGap, '
     'INDEX(_xlpm.Timeline, _xlpm.C + 1)), '
     'SUM(_xlfn.CHOOSEROWS(_xlpm.AmortisationSched, _xlpm.R) '
     '* (_xlpm.MonthStarts &gt;= _xlpm.PeriodOpens) '
     '* (_xlpm.MonthStarts &lt; _xlpm.PeriodShut))))), '
     '_xlpm.NewBlock, IF(_xlpm.SubMonthly?, _xlpm.ByDate, '
     'nabla.f.TimelinePositionλ(_xlpm.AmortisationSched, _xlpm.Timeline, _xlpm.Offset)), '),
    # Depreciateλ reads the same interval and, unlike Amortiseλ, survives a sub-monthly one:
    # measured in Excel, a weekly timeline already returns each month's depreciation in the
    # week holding that month's start. One thing does not survive. Every period but the last
    # takes its end from the next period's start; the last has no successor and is given
    # EDATE( its own start, MpP) - 1, which at nought months is the day BEFORE it begins, so
    # the final period of a daily, weekly or fortnightly timeline collects nothing. On 48
    # weekly periods from 1 January 2026 that dropped December outright: 1,833.37 of a
    # 2,000.00 year, against 2,000.00 over 49 weeks and over 12 months.
    #
    # Interval and PpY are the SWITCH lookups that made this function look as though it
    # only understood monthly, quarterly and yearly timelines. They never did anything: each
    # is read by exactly one binding, LastPeriod and LifeInPeriods, and nothing reads those.
    # Excel never evaluates them, which is why two-, four- and six-month timelines have
    # always returned correct schedules rather than the #N/A the source implies. Removed
    # with their readers below rather than generalised.
    ("nabla.f", "Depreciateλ",
     '        MpP,            @ROUND(( SecondPeriod - FirstPeriod) / 30.5, 0), //Months Per Period\n'
     '        Interval,       SWITCH( MpP, 1, "M", 3, "Q", 12, "Y"),\n'
     '        PpY,            SWITCH( MpP, 1, 12, 3, 4, 12, 1),                //Periods Per Year\n'
     '        EndDates,       HSTACK( MAP(  SEQUENCE(, TimelineCols - 1), LAMBDA( n, INDEX( Timeline, n + 1) - 1)), EDATE( MAX( Timeline), MpP) -1 ),\n',
     '        DpP,            SecondPeriod - FirstPeriod,                      //Days Per Period\n'
     '        SubMonthly?,    DpP < 28,                                        //No month is shorter\n'
     '        MpP,            MAX( @ROUND( DpP / 30.5, 0), 1),                 //Months Per Period, never nought\n'
     '    //  The last period has no successor to take its end date from, so it ends one period\n'
     '    //  on from its own start: a whole number of months where the timeline is monthly or\n'
     '    //  longer, the same number of days as every other period where it is shorter.\n'
     '        LastGap,        MAX( Timeline) - INDEX( Timeline, TimelineCols - 1),\n'
     '        LastEnd,        IF( SubMonthly?, MAX( Timeline) + LastGap, EDATE( MAX( Timeline), MpP)) - 1,\n'
     '        EndDates,       HSTACK( MAP(  SEQUENCE(, TimelineCols - 1), LAMBDA( n, INDEX( Timeline, n + 1) - 1)), LastEnd),\n',
     '_xlpm.MpP, _xlfn.SINGLE(ROUND((_xlpm.SecondPeriod - _xlpm.FirstPeriod) / 30.5, 0)), '
     '_xlpm.Interval, _xlfn.SWITCH(_xlpm.MpP, 1, "M", 3, "Q", 12, "Y"), '
     '_xlpm.PpY, _xlfn.SWITCH(_xlpm.MpP, 1, 12, 3, 4, 12, 1), '
     '_xlpm.EndDates, _xlfn.HSTACK(_xlfn.MAP(_xlfn.SEQUENCE(, _xlpm.TimelineCols - 1), '
     '_xlfn.LAMBDA(_xlpm.n, INDEX(_xlpm.Timeline, _xlpm.n + 1) - 1)), '
     'EDATE(MAX(_xlpm.Timeline), _xlpm.MpP) - 1), ',
     '_xlpm.DpP, _xlpm.SecondPeriod - _xlpm.FirstPeriod, '
     '_xlpm.SubMonthly?, _xlpm.DpP &lt; 28, '
     '_xlpm.MpP, MAX(_xlfn.SINGLE(ROUND(_xlpm.DpP / 30.5, 0)), 1), '
     '_xlpm.LastGap, MAX(_xlpm.Timeline) - INDEX(_xlpm.Timeline, _xlpm.TimelineCols - 1), '
     '_xlpm.LastEnd, IF(_xlpm.SubMonthly?, MAX(_xlpm.Timeline) + _xlpm.LastGap, '
     'EDATE(MAX(_xlpm.Timeline), _xlpm.MpP)) - 1, '
     '_xlpm.EndDates, _xlfn.HSTACK(_xlfn.MAP(_xlfn.SEQUENCE(, _xlpm.TimelineCols - 1), '
     '_xlfn.LAMBDA(_xlpm.n, INDEX(_xlpm.Timeline, _xlpm.n + 1) - 1)), _xlpm.LastEnd), '),
    # and the two bindings that read them, which nothing else reads
    ("nabla.f", "Depreciateλ",
     '                                    LifeInPeriods,  Years * PpY,\n'
     '                                    LifeInMonths,   PeriodDiffλ( InserviceDate, DisposalDate, "M"),\n'
     '                                    LastPeriod,     PeriodDiffλ( InserviceDate, DisposalDate, Interval),\n',
     '                                    LifeInMonths,   PeriodDiffλ( InserviceDate, DisposalDate, "M"),\n',
     '_xlpm.LifeInPeriods, _xlpm.Years * _xlpm.PpY, '
     '_xlpm.LifeInMonths, nabla.f.PeriodDiffλ(_xlpm.InServiceDate, _xlpm.DisposalDate, "M"), '
     '_xlpm.LastPeriod, nabla.f.PeriodDiffλ(_xlpm.InServiceDate, _xlpm.DisposalDate, _xlpm.Interval), ',
     '_xlpm.LifeInMonths, nabla.f.PeriodDiffλ(_xlpm.InServiceDate, _xlpm.DisposalDate, "M"), '),
    # Depreciateλ takes InitialValues, InServiceDates, LifeInYears, Timeline. Transposing the
    # middle two puts a date serial where the life belongs, and 1 January 2026 is 46,023, so
    # the function is asked for a schedule 552,276 months long and stops responding rather
    # than answering. No asset has a life over a hundred years, so say so and return a
    # message naming the argument order. The life is clamped where the check fails as well as
    # reported, because a message is no use if the arrays are built before anything reads it.
    ("nabla.f", "Depreciateλ",
     '        Mpy,            12, //Months Per Year\n',
     '        Mpy,            12, //Months Per Year\n'
     '    //  Check inputs - a life in years must be a life, not a date\n'
     '        LifeGiven,      IF( ISOMITTED( LifeInYears), 1, LifeInYears),\n'
     '        LifeNums,       IFERROR( VALUE( LifeGiven), -1),\n'
     '        BadLife?,       OR( MIN( LifeNums) <= 0, MAX( LifeNums) > 100),\n'
     '        LifeMessage,    "LifeInYears must be a number of years greater than 0 and no more than 100. "\n'
     '                        & "Check the argument order: InitialValues, InServiceDates, LifeInYears, Timeline.",\n',
     '_xlpm.Mpy, 12, ',
     '_xlpm.Mpy, 12, '
     '_xlpm.LifeGiven, IF(_xlfn.ISOMITTED(_xlpm.LifeInYears), 1, _xlpm.LifeInYears), '
     '_xlpm.LifeNums, IFERROR(VALUE(_xlpm.LifeGiven), -1), '
     '_xlpm.BadLife?, OR(MIN(_xlpm.LifeNums) &lt;= 0, MAX(_xlpm.LifeNums) &gt; 100), '
     '_xlpm.LifeMessage, "LifeInYears must be a number of years greater than 0 and no more than 100. " '
     '&amp; "Check the argument order: InitialValues, InServiceDates, LifeInYears, Timeline.", '),
    ("nabla.f", "Depreciateλ",
     '                                    Years,          @INDEX( LifeInYears, Asset),\n',
     '                                    Years,          IF( BadLife?, 1, @INDEX( LifeInYears, Asset)),\n',
     '_xlpm.Years, _xlfn.SINGLE(INDEX(_xlpm.LifeInYears, _xlpm.Asset))',
     '_xlpm.Years, IF(_xlpm.BadLife?, 1, _xlfn.SINGLE(INDEX(_xlpm.LifeInYears, _xlpm.Asset)))'),
    ("nabla.f", "Depreciateλ",
     '    //  Return Result\n'
     '        CHOOSE(Help? + 1, Result, Help)\n',
     '    //  Return Result, Help, or the message\n'
     '        Return,         IF( Help?, 2, IF( BadLife?, 3, 1)),\n'
     '        CHOOSE( Return, Result, Help, LifeMessage)\n',
     'CHOOSE(_xlpm.Help? + 1, _xlpm.Result, _xlpm.Help)',
     '_xlpm.Return, IF(_xlpm.Help?, 2, IF(_xlpm.BadLife?, 3, 1)), '
     'CHOOSE(_xlpm.Return, _xlpm.Result, _xlpm.Help, _xlpm.LifeMessage)'),
    # Allocateλ grew its answer by HSTACKing each new group onto everything already built,
    # inside a REDUCE, so the whole accumulator is copied on every pass and the work is
    # quadratic in the number of amounts. It is quick at any sane input and it is what turned
    # the transposed Depreciateλ call above from an error into a workbook that stops
    # responding: 46,023 amounts allocated to 552,276 months is of the order of 10^10 element
    # copies. The answer is a closed form, so size the row once and compute each column from
    # its own index. Every published result is unchanged, the last-column adjustment included:
    # it is still the amount less the SUM of the same array of equal parts, in the same order,
    # so the floating-point residue is identical to the bit.
    ("nabla.f", "Allocateλ",
     '        Result,         REDUCE( 0, SEQUENCE( , FromCount),\n'
     '                            LAMBDA( Acc, n,\n'
     '                                LET( \n'
     '                                    FromAmount,     INDEX( Amounts, n), \n'
     '                                    ToAmount,       ROUND( FromAmount * From / To, 2),\n'
     '                                    BaseArray,      EXPAND( ToAmount, 1, ToCount, ToAmount),\n'
     '                                    ToArray,        IF( ToCount = 1, \n'
     '                                                        BaseArray, \n'
     '                                                        HSTACK( \n'
     '                                                            TAKE(BaseArray, 1, ToCount - 1), \n'
     '                                                            FromAmount - SUM( TAKE(BaseArray, 1, ToCount - 1))\n'
     '                                                        )\n'
     '                                                    ),\n'
     '                                    Result,         IF( n = 1, ToArray, HSTACK( Acc, ToArray )),\n'
     '                                    Result\n'
     '                                )\n'
     '                            )\n'
     '                        ),\n',
     '        Result,         MAKEARRAY( 1, FromCount * ToCount,\n'
     '                            LAMBDA( R, C,\n'
     '                                LET(\n'
     '                                //  Column C holds part C of amount QUOTIENT( C - 1, ToCount) + 1\n'
     '                                    FromAmount,     INDEX( Amounts, QUOTIENT( C - 1, ToCount) + 1),\n'
     '                                    ToAmount,       ROUND( FromAmount * From / To, 2),\n'
     '                                //  The last part of each amount carries the rounding difference\n'
     '                                    Last?,          MOD( C - 1, ToCount) + 1 = ToCount,\n'
     '                                    IF( AND( Last?, ToCount > 1),\n'
     '                                        FromAmount - SUM( EXPAND( ToAmount, 1, MAX( ToCount - 1, 1), ToAmount)),\n'
     '                                        ToAmount)\n'
     '                                )\n'
     '                            )\n'
     '                        ),\n',
     '_xlpm.Result, _xlfn.REDUCE(0, _xlfn.SEQUENCE(, _xlpm.FromCount), _xlfn.LAMBDA(_xlpm.Acc,_xlpm.n, '
     '_xlfn.LET(_xlpm.FromAmount, INDEX(_xlpm.Amounts, _xlpm.n), '
     '_xlpm.ToAmount, ROUND(_xlpm.FromAmount * _xlpm.From / _xlpm.To, 2), '
     '_xlpm.BaseArray, _xlfn.EXPAND(_xlpm.ToAmount, 1, _xlpm.ToCount, _xlpm.ToAmount), '
     '_xlpm.ToArray, IF(_xlpm.ToCount = 1, _xlpm.BaseArray, '
     '_xlfn.HSTACK(_xlfn.TAKE(_xlpm.BaseArray, 1, _xlpm.ToCount - 1), '
     '_xlpm.FromAmount - SUM(_xlfn.TAKE(_xlpm.BaseArray, 1, _xlpm.ToCount - 1)))), '
     '_xlpm.Result, IF(_xlpm.n = 1, _xlpm.ToArray, _xlfn.HSTACK(_xlpm.Acc, _xlpm.ToArray)), '
     '_xlpm.Result))), ',
     '_xlpm.Result, _xlfn.MAKEARRAY(1, _xlpm.FromCount * _xlpm.ToCount, _xlfn.LAMBDA(_xlpm.R,_xlpm.C, '
     '_xlfn.LET(_xlpm.FromAmount, INDEX(_xlpm.Amounts, QUOTIENT(_xlpm.C - 1, _xlpm.ToCount) + 1), '
     '_xlpm.ToAmount, ROUND(_xlpm.FromAmount * _xlpm.From / _xlpm.To, 2), '
     '_xlpm.Last?, MOD(_xlpm.C - 1, _xlpm.ToCount) + 1 = _xlpm.ToCount, '
     'IF(AND(_xlpm.Last?, _xlpm.ToCount &gt; 1), '
     '_xlpm.FromAmount - SUM(_xlfn.EXPAND(_xlpm.ToAmount, 1, MAX(_xlpm.ToCount - 1, 1), _xlpm.ToAmount)), '
     '_xlpm.ToAmount)))), '),
]

for _entry in LOGIC_FIXES:
    _mod, _fn, (_old, _new), _ = stores(_entry)
    _a, _b = function_block(_mod, _fn)
    _text = mods[_mod]["text"]
    _blk = _text[_a:_b]
    assert _blk.count(_old) == 1, (_fn, _mod, "source", _old, _blk.count(_old))
    mods[_mod]["text"] = _text[:_a] + _blk.replace(_old, _new) + _text[_b:]

print("fixed %d defects in the module sources across %d functions"
      % (len(LOGIC_FIXES), len({(e[0], e[1]) for e in LOGIC_FIXES})))

# ---------- the debt module, which has no module source ----------
# The five debt functions are recursive: each calls itself by name. The Advanced Formula
# Environment takes a function's prefix from the container it is imported into, so a
# recursive definition held in an AFE module would call a name that does not exist there.
# Upstream leaves all five out of its project store for that reason and ships them as
# defined names only, and so does this build. They therefore get the defined-name pass
# alone: there is no module source to correct first, and no demonstration sheet has ever
# called one, so nothing is cached either.
DEBT_FIXES = [
    # DebtSculptVariableLRVλ subtracted the interest when working out the payment and then
    # added it back into the closing balance. Subtracting it is only right if the interest
    # is paid out of the period's cash, which is what a debt service coverage ratio means;
    # adding it back is only right if it is not. Both were there, so every closing balance
    # carried one period's interest too much and the error compounded into the next opening
    # balance. On 1,000 of debt at 6% with CFADS 300 and DSCR 1.2 over five periods it
    # repaid 1,078.60 of a 1,000 loan and still showed 92.80 outstanding.
    #
    # The payment is the principal repayment, so cap it at the principal itself rather than
    # at the principal less the interest, and take it off the balance. A negative repayment
    # needs no special case: when the cash cannot cover the interest, subtracting a negative
    # capitalises the shortfall, which is what should happen.
    ("nabla.debt", "DebtSculptVariableLRVλ",
     "MIN(_xlpm.PeriodCFADS / _xlpm.PeriodDSCR - _xlpm.Interest, _xlpm.Principal - _xlpm.Interest), "
     "_xlpm.ClosingDebt, _xlpm.Principal + _xlpm.Interest - _xlpm.Payment",
     "MIN(_xlpm.PeriodCFADS / _xlpm.PeriodDSCR - _xlpm.Interest, _xlpm.Principal), "
     "_xlpm.ClosingDebt, _xlpm.Principal - _xlpm.Payment"),
    # The other two sculpting functions compute the balance correctly, but their third row
    # holds the whole debt service and both call it principal repayments. On the same
    # figures that row reads 250 where the principal repaid is 190. Only the label is wrong.
    # DebtSculptVariableLRVλ keeps the label, because with the fix above its third row is
    # the principal repayment.
    ("nabla.debt", "DebtSculptFixedλ",
     "→Principal repayments¶", "→Debt service (interest and principal)¶"),
    ("nabla.debt", "DebtSculptVariableλ",
     "→Principal repayments¶", "→Debt service (interest and principal)¶"),
    # and the LRV function's one worked example calls its sibling rather than itself, with
    # the closing bracket missing as well, so the line a reader copies runs the schedule
    # this release has just made behave differently.
    ("nabla.debt", "DebtSculptVariableLRVλ",
     "→=DebtSculptVariableλ(, Debt, CFADS, DSCR, APR\"",
     "→=DebtSculptVariableLRVλ(, Debt, CFADS, DSCR, APR)\""),
    # InterestLRVλ solves for the interest on the average balance over the period, taking the
    # principal repayment as the cash available for debt service less that interest. Nobody
    # repays more than they owe, and its caller says so: DebtSculptVariableLRVλ caps the
    # payment at the principal. This function did not, so wherever the cap binds it solved a
    # repayment larger than the debt, put the average balance below half the opening balance,
    # and reported interest that is too small. Interest on 1,000 at 5% a period with 1,200 of
    # cash came back as 20.51 where the balance runs from 1,000 to nil and the interest is
    # 25.00. Cap the repayment inside the iteration, both where it is assumed and where it is
    # solved, and the two agree. It still converges, and faster: once the cap binds the
    # assumed and solved repayments are both the principal, so the first pass is the last.
    #
    # The one period of every sculpted schedule where the debt is retired is affected, and no
    # other. The published example is not: 6,666.37 over 3.50 is 1,904.68 against 90,000 of
    # principal, nowhere near the cap, and still prints 222.90. No balance moves anywhere,
    # because the closing balance is the principal less the payment and never read the
    # interest, which is why the balance identities in the self-test could not see this.
    ("nabla.debt", "InterestLRVλ",
     "_xlpm.Payment, IF(_xlpm.DoNotUse = 0, _xlpm.CFADS / _xlpm.DSCR, _xlpm.DoNotUse), "
     "_xlpm.Interest, (_xlpm.Principal - _xlpm.Payment / 2) * _xlpm.InterestRate, "
     "_xlpm.Repayment, -_xlpm.CFADS / _xlpm.DSCR + _xlpm.Interest, ",
     "_xlpm.Payment, MIN(IF(_xlpm.DoNotUse = 0, _xlpm.CFADS / _xlpm.DSCR, _xlpm.DoNotUse), _xlpm.Principal), "
     "_xlpm.Interest, (_xlpm.Principal - _xlpm.Payment / 2) * _xlpm.InterestRate, "
     "_xlpm.Repayment, -MIN(_xlpm.CFADS / _xlpm.DSCR - _xlpm.Interest, _xlpm.Principal), "),
    # and its description says which balance the interest is charged on
    ("nabla.debt", "InterestLRVλ",
     "→Calculates debt sculpting interest using method presented by Lance Rubin¶",
     "→Calculates debt sculpting interest using method presented by Lance Rubin.¶"
     "→Charged on the average balance over the period, with the principal repaid¶"
     "→capped at the principal outstanding.¶"),
]

# fix upstream copy-paste bug: the u module's About suggested "nabla.e" (was BXE) as its own name
assert "Suggested module name: nabla.e" in mods["nabla.u"]["text"]
mods["nabla.u"]["text"] = mods["nabla.u"]["text"].replace(
    "Suggested module name: nabla.e", "Suggested module name: nabla.u")

# append the Australian additions to their modules and list them in the f module's About table
for spec in FUNCS:
    mod = spec["module"].split(".", 1)[1]
    mods["nabla." + mod]["text"] = mods["nabla." + mod]["text"].rstrip() + "\n\n\n\n" + build_afe(spec)
_dep = [f for f in FUNCS if f["module"] == "nabla.f" and "GST" not in f["name"]]
_gst = [f for f in FUNCS if f["module"] == "nabla.f" and "GST" in f["name"]]
about_add = "".join('"%-19s→%s¶" & \n%s' % (f["name"], f["desc"], " " * 43) for f in _dep)
about_add += '"→¶" & \n%s"AUSTRALIAN TAX     →¶" & \n%s' % (" " * 43, " " * 43)
about_add += "".join('"%-19s→%s¶" & \n%s' % (f["name"], f["desc"], " " * 43) for f in _gst)
anchor = '"VDBλ               →Variable declining balance depreciation method for one asset or asset class.¶" & '
assert mods["nabla.f"]["text"].count(anchor) == 1
mods["nabla.f"]["text"] = mods["nabla.f"]["text"].replace(anchor, anchor + "\n" + " " * 43 + about_add)

# list the new dates function in its module's About table
d_anchor = '"Timelineλ              →Creates a horizontal list of start or end dates for a timeline¶" & '
assert mods["nabla.d"]["text"].count(d_anchor) == 1
mods["nabla.d"]["text"] = mods["nabla.d"]["text"].replace(
    d_anchor,
    d_anchor + '\n        "%-23s→%s¶" &' % ("FinancialYearλ", "Labels dates with their Australian financial year, starting 1 July"))

# the store still declared the upstream authoring locale
obj_afe["locale"]["localeName"] = "en-au"
obj_afe["locale"]["dateOrder"] = "DMY"

names = obj_afe["projectNames"]
assert "nabla.f.MACRSλ" in names
names.remove("nabla.f.MACRSλ")
if "nabla.f.SumDepreciateλ" not in names:  # upstream omitted it from the index
    names.append("nabla.f.SumDepreciateλ")
for spec in FUNCS:
    full = spec["module"] + "." + spec["name"]
    if full not in names:
        names.append(full)

j2 = json.dumps(obj_afe, ensure_ascii=False, separators=(",", ":"))
afe = afe.replace(m.group(1), base64.b64encode(j2.encode("utf-16-le")).decode("ascii"))
put("customXml/item1.xml", afe)

for n in list(parts):
    if n in ("customXml/item1.xml", "docProps/core.xml"):
        continue  # AFE handled above; core.xml provenance dates must not shift
    if n.endswith((".xml", ".rels")):
        put(n, transform_text(get(n)))
    elif re.match(r'xl/customProperty\d+\.bin$', n):
        txt = parts[n].decode("utf-16-le")
        parts[n] = transform_text(txt).encode("utf-16-le")

# ---------- help signatures, second store ----------
# The module sources were corrected before the AFE store was written. The defined names
# could not be, because until the sweep just above they still carried the upstream brand
# and the upstream spelling: QuickRatioλ's help said Liabilites, and LabelAmortiseλ was
# still LabelAmortizeλ. Now that both stores read the same, apply the same table.
_wbx = get("xl/workbook.xml")
for _entry in HELP_SIGNATURES + LOGIC_FIXES + DEBT_FIXES:
    _mod, _fn, _, (_old, _new) = stores(_entry)
    _hit = []

    # the full name, not the base: several functions exist in more than one module, and
    # a correction belongs to the copy the table names
    def _fix(m, _old=_old, _new=_new, _hit=_hit, _want=_mod + "." + _fn):
        if m.group(1) != _want or _old not in m.group(2):
            return m.group(0)
        _hit.append(m.group(1))
        return m.group(0).replace(_old, _new)

    _wbx = re.sub(r'<definedName name="([^"]+)"[^>]*>(.*?)</definedName>', _fix, _wbx, flags=re.S)
    assert len(_hit) == 1, (_fn, _mod, _old, _hit)

# and the same parameter rename, which Excel stores prefixed: _xlop. on the declaration,
# _xlpm. on the references, both already normalised to the declared spelling
_renamed = 0
for _mod in ISINLIST:
    _hit = []

    def _rename(m, _hit=_hit):
        if m.group(1) != _mod + ".IsInListλ":
            return m.group(0)
        body, n = re.subn(r"(_xl(?:op|pm)\.)LIST\b", lambda k: k.group(1) + "List", m.group(2))
        _hit.append(n)
        # and its missing column delimiter, stored the same way it is written
        _old, _new = TEXTSPLIT_ARGS
        assert body.count(_old) == 1, (_mod, "defined name", "TEXTSPLIT", body.count(_old))
        return m.group(0).replace(m.group(2), body.replace(_old, _new))

    _wbx = re.sub(r'<definedName name="([^"]+)"[^>]*>(.*?)</definedName>', _rename, _wbx, flags=re.S)
    assert _hit == [3], (_mod, "defined name", _hit)   # one declaration, two references
    _renamed += _hit[0]
put("xl/workbook.xml", _wbx)
print("corrected %d help signatures and %d defects in the defined names, %d of them in the "
      "debt module, renamed %d LIST tokens, restored 2 column delimiters"
      % (len(HELP_SIGNATURES), len(LOGIC_FIXES) + len(DEBT_FIXES), len(DEBT_FIXES), _renamed))

# Several of these functions are demonstrated on a sheet that calls them with no
# arguments, so their help is spilled there and the old text sits in the file as a
# cached value. It would go on being displayed until something forced a recalculation.
# A signature is cached whole; a parameter label is cached on its own, trimmed, because
# the help is a two-column table and TRIM() has already run on it.


def cached_forms(fragment):
    """The fragment as a cached spill holds it, and whether it is a whole cell.

    Returns (is_whole_cell, value). A parameter label spills into a cell of its own, so
    it is matched whole: the help is built with TRIM(), which strips the padding and
    collapses any run of spaces inside the label, so a label whose colon sits after its
    padding caches with one space before the colon rather than seven. Anything else is a
    piece of a longer cell and is matched as a substring.
    """
    if fragment.startswith('"') and fragment.rstrip().endswith("→"):
        return True, " ".join(fragment[1:].rstrip()[:-1].split())
    return False, fragment


CACHED_VALUE = re.compile(r"<v>(.*?)</v>", re.S)


def refresh_cached(text, whole, old, new):
    """Replace inside cached values only, never inside a formula.

    A worksheet holds both: <f> is what Excel will recalculate, <v> is what it last
    produced. Correcting a spilled help block means rewriting the stale <v>. Rewriting an
    <f> would change which function a demonstration sheet calls, which is exactly what a
    blanket replace did once: a correction aimed at one function's help text renamed a
    live call on a neighbouring sheet and spilled #SPILL! across it.
    """
    hits = [0]

    def one(m):
        value = m.group(1)
        if whole:
            if value.strip() != old:
                return m.group(0)
            hits[0] += 1
            return "<v>%s</v>" % value.replace(old, new)
        if old not in value:
            return m.group(0)
        hits[0] += 1
        return "<v>%s</v>" % value.replace(old, new)

    return CACHED_VALUE.sub(one, text), hits[0]


_refreshed = {}
for _entry in HELP_SIGNATURES:
    _mod, _fn, (_old, _new), _ = stores(_entry)
    _whole, _co = cached_forms(_old)
    _cn = cached_forms(_new)[1]
    for _sheet in [n for n in list(parts) if re.match(r"xl/worksheets/sheet\d+\.xml$", n)]:
        _text, _hits = refresh_cached(get(_sheet), _whole, _co, _cn)
        if not _hits:
            continue
        put(_sheet, _text)
        _refreshed.setdefault(_sheet, []).append(_fn)
    _left = [v for n in parts if re.match(r"xl/worksheets/sheet\d+\.xml$", n)
             for v in CACHED_VALUE.findall(get(n))
             if (v.strip() == _co if _whole else _co in v)]
    assert not _left, (_fn, _co, _left[:2])
print("refreshed cached help on %d sheets: %s"
      % (len(_refreshed), ", ".join("%s (%s)" % (s.rsplit("/", 1)[1], ", ".join(f))
                                    for s, f in sorted(_refreshed.items()))))


def left_of(ref):
    """The cell one column to the left of an A1-style reference."""
    col, row = re.match(r"([A-Z]+)(\d+)$", ref).groups()
    n = 0
    for ch in col:
        n = n * 26 + ord(ch) - 64
    n -= 1
    assert n >= 1, ref
    out = ""
    while n:
        n, rem = divmod(n - 1, 26)
        out = chr(65 + rem) + out
    return out + row


# One corrected example result is also cached: OverLapDaysλ has a demonstration sheet, so
# its help is spilled there and the old answer would go on being displayed until something
# forced a recalculation. It cannot be matched the way the help corrections above are. The
# help is a two-column table, the result spills into a cell of its own, and a bare 12 is
# not distinctive: forty-odd cells across the workbook cache that number. Find the row by
# the formula beside it, which names the function and its arguments, and rewrite only the
# cell to its left. The marker leaves off the module prefix, because this runs before the
# flat rename and the cell still reads nabla.d. rather than the nb. it ships as.
CACHED_EXAMPLES = [('OverLapDaysλ("17/1/2025"', "12", "2")]

CELL = re.compile(r'<c r="([A-Z]+\d+)"[^>]*>((?:(?!<c[ /]).)*?)</c>', re.S)

for _marker, _was, _now in CACHED_EXAMPLES:
    _found = []
    for _sheet in [n for n in list(parts) if re.match(r"xl/worksheets/sheet\d+\.xml$", n)]:
        _text = get(_sheet)
        if _marker not in _text:
            continue
        _cells = {m.group(1): m.group(2) for m in CELL.finditer(_text)}
        _rows = [r for r, body in _cells.items() if _marker in body]
        assert len(_rows) == 1, (_sheet, _marker, _rows)
        _label = left_of(_rows[0])
        assert _cells[_label] == "<v>%s</v>" % _was, (_sheet, _label, _cells[_label])
        _open = re.search(r'<c r="%s"[^>]*>' % _label, _text).group(0)
        _text = _text.replace(_open + _cells[_label], _open + "<v>%s</v>" % _now, 1)
        put(_sheet, _text)
        _found.append("%s!%s" % (_sheet.rsplit("/", 1)[1], _label))
    assert len(_found) == 1, (_marker, _found)
    print("refreshed the cached example result at %s: %s -> %s" % (_found[0], _was, _now))

# Periodsλ counts differently now, so the sheet that demonstrates it caches answers that no
# longer match its own formula, and two of the five have moved. The same sheet is where the
# workbook's cached #VALUE! cells live: the demonstration spills through a LAMBDA that Excel
# evaluates correctly on open but that something once saved in an error state, so the file
# has been shipping five errors that the reader never sees and a recalculation never
# reproduces. Correcting them here means the file agrees with itself before Excel is opened.
# Cells are named outright rather than searched for, because a cached number is not
# distinctive enough to find, and each one carries the value it must be replacing.
CACHED_SHEET_CELLS = [
    ("Periodsλ(A25:A29,B25:B29,C25:C29)", [
        ("B8", "* End Date is inclusive",
         "* Counts the period starts crossed, so a part period at the end counts"),
        ("B19", '=nabla.d.Periodsλ("15/1/2026", "16/1/2025", , FALSE)',
         '=nabla.d.Periodsλ("15/1/2026", "16/1/2025")'),
        ("D25", "#VALUE!", "9"),      # 28 Feb 2026 to 30 Nov 2026, months
        ("D26", "#VALUE!", "5"),      # 30 Jun 2026 to 1 Aug 2027, quarters
        ("D27", "#VALUE!", "3"),      # 1 Jan 2026 to 1 Jan 2029, years
        ("D28", "#VALUE!", "53"),     # 30 Apr 2026 to 1 May 2027, weeks
        ("D29", "#VALUE!", "365"),    # 31 Aug 2026 to 31 Aug 2027, days
    ]),
]

for _marker, _cells in CACHED_SHEET_CELLS:
    _sheets = [n for n in list(parts) if re.match(r"xl/worksheets/sheet\d+\.xml$", n)
               and _marker in get(n)]
    assert len(_sheets) == 1, (_marker, _sheets)
    _sheet = _sheets[0]
    _text = get(_sheet)
    for _ref, _was, _now in _cells:
        _m = re.search(r'<c r="%s"([^>]*)>((?:(?!<c[ /]).)*?)</c>' % _ref, _text, re.S)
        assert _m, (_sheet, _ref, "no such cell")
        _attrs, _inner = _m.group(1), _m.group(2)
        assert "<v>%s</v>" % _was in _inner, (_sheet, _ref, _inner[:80])
        # an error cell carries t="e"; a number carries no type at all
        if _was.startswith("#") and not _now.startswith("#"):
            assert ' t="e"' in _attrs, (_sheet, _ref, _attrs)
            _attrs = _attrs.replace(' t="e"', "")
        _text = (_text[:_m.start()] + '<c r="%s"%s>%s</c>' % (_ref, _attrs,
                 _inner.replace("<v>%s</v>" % _was, "<v>%s</v>" % _now))
                 + _text[_m.end():])
    put(_sheet, _text)
    print("refreshed %d cached cells on %s"
          % (len(_cells), _sheet.rsplit("/", 1)[1]))

# cached spill copies of help version lines in worksheets -> today
for n in list(parts):
    if re.match(r'xl/worksheets/sheet\d+\.xml$', n):
        put(n, re.sub(r'(?<=>)(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec) \d{1,2} 20\d\d(?=<)',
                      TODAY_AU, get(n)))

# demo input cells stored as date serials: +2 years, calendar-aware
st_now = get("xl/styles.xml")
cellxfs = re.search(r'<cellXfs count="\d+">(.*?)</cellXfs>', st_now, re.S).group(1)
xfs = re.findall(r'<xf [^>]*numFmtId="(\d+)"[^>]*/?>', cellxfs)
DATE_BUILTIN = set(range(14, 23)) | set(range(27, 37)) | {45, 46, 47} | set(range(50, 59))
custom_date = set()
for m in re.finditer(r'<numFmt numFmtId="(\d+)" formatCode="([^"]*)"/>', st_now):
    code = re.sub(r'\[[^\]]*\]|"[^"]*"|\\.', "", m.group(2))
    if re.search(r'y|mmm', code):
        custom_date.add(int(m.group(1)))
date_styles = {i for i, nf in enumerate(xfs) if int(nf) in DATE_BUILTIN or int(nf) in custom_date}
shifted_cells = 0
for n in list(parts):
    if re.match(r'xl/worksheets/sheet\d+\.xml$', n):
        d = get(n)
        def _cell(mo):
            global shifted_cells
            sid = int(mo.group("s") or 0)
            if sid in date_styles and 20000 <= float(mo.group("v")) <= 80000:
                v = float(mo.group("v"))
                nv = _shift_serial(v)
                out = str(int(nv)) if nv == int(nv) else repr(nv)
                shifted_cells += 1
                return mo.group(0).replace("<v>%s</v>" % mo.group("v"), "<v>%s</v>" % out)
            return mo.group(0)
        d = re.sub(r'<c r="[A-Z]+\d+"(?: s="(?P<s>\d+)")?><v>(?P<v>[0-9.]+)</v></c>', _cell, d)
        put(n, d)
assert shifted_cells >= 60, shifted_cells
print("shifted", shifted_cells, "serial date cells by +%d years" % YEAR_SHIFT)

# demo-consistency anchors named by review: assert each landed on its shifted value
ANCHOR_CHECKS = [
    ("xl/worksheets/sheet8.xml", "C23", 45337), ("xl/worksheets/sheet8.xml", "D23", 45542),
    ("xl/worksheets/sheet8.xml", "C24", 45306), ("xl/worksheets/sheet8.xml", "D24", 45372),
    ("xl/worksheets/sheet8.xml", "C25", 45292), ("xl/worksheets/sheet8.xml", "D25", 45322),
    ("xl/worksheets/sheet13.xml", "B20", 45292),
    ("xl/worksheets/sheet11.xml", "B21", 45292), ("xl/worksheets/sheet11.xml", "B26", 45292),
    ("xl/worksheets/sheet11.xml", "B27", 45444),
    ("xl/worksheets/sheet16.xml", "E24", 44927), ("xl/worksheets/sheet16.xml", "E25", 45292),
    ("xl/worksheets/sheet16.xml", "E26", 45658),
    ("xl/worksheets/sheet17.xml", "E24", 44927), ("xl/worksheets/sheet18.xml", "E24", 44927),
    ("xl/worksheets/sheet12.xml", "E25", 44950), ("xl/worksheets/sheet12.xml", "E38", 45385),
    ("xl/worksheets/sheet32.xml", "E23", 43831),
]
for part, cell, old in ANCHOR_CHECKS:
    d = get(part)
    want = str(int(_shift_serial(float(old))))
    assert re.search(r'<c r="%s"[^>]*><v>%s</v></c>' % (cell, want), d), (part, cell, old, want)
print("anchor checks pass:", len(ANCHOR_CHECKS))

# ---------- 8a4. Amortiseλ demo: start the first loan inside the timeline ----------
# Upstream started it 12 months before the model timeline with a 10-month term, so the loan
# was fully repaid before the first period and its six rows rendered as zeros. Starting it
# 1 March 2026 puts a partial schedule on screen. The caption is restated to match.
OLD_START, NEW_START = "45658", "46082"      # 1 Jan 2025 -> 1 Mar 2026
assert (EPOCH + datetime.timedelta(days=int(NEW_START))) == datetime.datetime(2026, 3, 1)
for sheet in ("sheet16", "sheet17", "sheet18"):
    p = "xl/worksheets/%s.xml" % sheet
    d = get(p)
    d2, n = re.subn(r'(<c r="E24"[^>]*>)<v>%s</v>' % OLD_START, r'\g<1><v>%s</v>' % NEW_START, d)
    assert n == 1, sheet
    put(p, d2)
d13 = get("xl/drawings/drawing13.xml")
d13b = d13.replace("The first one predates our model's timeline.",
                   "The first one starts partway through our model's first period.")
assert d13b != d13
put("xl/drawings/drawing13.xml", d13b)
print("Amortiseλ demo: first loan moved to 1 Mar 2026, caption restated")

# ---------- 8a5. sheet32: repair inherited #REF! Timeline argument ----------
s32 = get("xl/worksheets/sheet32.xml")
s32b = s32.replace('nabla.d.Timelineλ( E23, D23, "Y",#REF!)', 'nabla.d.Timelineλ( E23, D23, "Y")')
assert s32b != s32
put("xl/worksheets/sheet32.xml", s32b)

# ---------- 8a6. drawing18: drop the US GAAP framing ----------
d18 = get("xl/drawings/drawing18.xml")
d18b = d18.replace(
    "US GAAP allows companies to, in the last period, depreciate all remaining book value down to "
    "the salvage value. Excel's DB() and DDB() functions do not, thus, they deprive companies of "
    "some of the tax benefits of offsetting income with depreciation.",
    "In the final period it writes the remaining book value down to the salvage value. "
    "Excel's DB() and DDB() functions stop short of that, understating the deduction in the last period.")
assert d18b != d18
put("xl/drawings/drawing18.xml", d18b)

# ---------- 8. Fonts: Calibri -> Aptos ----------
for n in ["xl/styles.xml", "xl/theme/theme1.xml", "xl/sharedStrings.xml"]:
    d = get(n)
    d = d.replace("Calibri Light", "Aptos Display").replace("Calibri", "Aptos")
    put(n, d)

# ---------- 8x. Date number formats: US order -> AU day-first ----------
st2 = get("xl/styles.xml")
for old, new in [('formatCode="mm/dd/yyyy\\ h:mm"', 'formatCode="dd/mm/yyyy\\ h:mm"'),
                 ('formatCode="mm/dd/yyyy"', 'formatCode="dd/mm/yyyy"'),
                 ('formatCode="m/d/yyyy"', 'formatCode="d/m/yyyy"')]:
    assert old in st2, old
    st2 = st2.replace(old, new)
put("xl/styles.xml", st2)

# ---------- 8a2. Locale-safe RANDBETWEEN demo inputs (US text dates fail on AU-locale Excel) ----------
for path, old, new, cnt_want in [
    ("xl/worksheets/sheet9.xml", 'RANDBETWEEN("1/1/2026", "31/12/2028")',
     'RANDBETWEEN(DATE(2026,1,1), DATE(2028,12,31))', 3),
    ("xl/worksheets/sheet13.xml", 'RANDBETWEEN("1/1/2026", "1/6/2026")',
     'RANDBETWEEN(DATE(2026,1,1), DATE(2026,6,1))', 1),
    # the table definition carries its own copy of the column formula
    ("xl/tables/table10.xml", 'RANDBETWEEN("1/1/2026", "1/6/2026")',
     'RANDBETWEEN(DATE(2026,1,1), DATE(2026,6,1))', 1),
]:
    d = get(path)
    assert d.count(old) == cnt_want, (path, d.count(old))
    put(path, d.replace(old, new))
for path in parts:
    if path.startswith("xl/tables/"):
        assert 'RANDBETWEEN("' not in get(path), path

# ---------- 8a3. Remove dead TOC hyperlink (row points at a worksheet that never existed) ----------
toc = get("xl/worksheets/sheet2.xml")
toc2 = re.sub(r'<hyperlink ref="A28" location="\'nabla\.f\.Aboutλ\'![^"]*"[^>]*/>', "", toc)
assert toc2 != toc
put("xl/worksheets/sheet2.xml", toc2)

# ---------- 8b. Fix undefined Sheetλ title formulas (original defect: cached #NAME? in every A1) ----------
wbnow = get("xl/workbook.xml")
sheetmap = {}  # part path -> transformed sheet name
relsnow = get("xl/_rels/workbook.xml.rels")
relmap = dict(re.findall(r'Id="(rId\d+)"[^>]*Target="([^"]+)"', relsnow))
for nm, rid in re.findall(r'<sheet name="([^"]+)"[^>]*r:id="(rId\d+)"', wbnow):
    sheetmap["xl/" + relmap[rid]] = nm
fixed_titles = 0
for path, nm in sheetmap.items():
    d = get(path)
    new_c = ('<c r="A1" s="\\1" t="str">'
             '<f>_xlfn.TEXTAFTER(CELL("filename",A1),"]")</f>'
             '<v>%s</v></c>' % nm)
    d2, c1 = re.subn(r'<c r="A1" s="(\d+)" t="e" cm="1"><f t="array" ref="A1">nabla\.u\.Sheetλ</f><v>#NAME\?</v></c>', new_c, d)
    d2, c2 = re.subn(r'<c r="A1" s="(\d+)" t="e"><f>nabla\.u\.Sheetλ</f><v>#NAME\?</v></c>', new_c, d2)
    if c1 + c2:
        fixed_titles += c1 + c2
        put(path, d2)
leftover = [p for p in sheetmap if "nabla.u.Sheetλ" in get(p)]
assert fixed_titles >= 40 and not leftover, (fixed_titles, leftover)
print("fixed", fixed_titles, "Sheetλ title cells")

# ---------- 9. workbook.xml: drop Slicer_Type, force full recalc ----------
wb = get("xl/workbook.xml")
# Slicer_Type is NOT stale: [MS-XLSX] requires a #N/A defined name for each slicer cache,
# and the TOC sheet's Type slicer is live. It must survive the transform.
assert '<definedName name="Slicer_Type">#N/A</definedName>' in wb
m = re.search(r'<calcPr[^>]*/>', wb)
assert m, "calcPr not found"
if "fullCalcOnLoad" not in m.group(0):
    wb = wb.replace(m.group(0), m.group(0)[:-2] + ' fullCalcOnLoad="1"/>')
put("xl/workbook.xml", wb)

# ---------- 9b. Define nabla.e.Aboutλ (original defect: called on its sheet, never defined; source in AFE) ----------
store = json.loads(j2)
wb = get("xl/workbook.xml")
for mod in ("e", "d"):  # both modules ship an Aboutλ source that was never installed
    text = next(f["text"] for f in store["files"] if f["path"] == "/projects/nabla." + mod)
    i = text.index("Aboutλ = ")
    expr = text[i + len("Aboutλ = "):]
    depth = 0; in_str = False; end = None
    for k, ch in enumerate(expr):
        if ch == '"': in_str = not in_str
        elif not in_str and ch == '(': depth += 1
        elif not in_str and ch == ')':
            depth -= 1
            if depth == 0: end = k + 1; break
    assert end, "Aboutλ body not parsed for " + mod
    body = " ".join(expr[:end].split())
    assert re.match(r'TRIM\(\s*TEXTSPLIT\(', body) and body.count('"') % 2 == 0
    body = body.replace("TEXTSPLIT(", "_xlfn.TEXTSPLIT(")
    full = "nabla.%s.Aboutλ" % mod
    assert '<definedName name="%s"' % full not in wb
    wb = wb.replace("</definedNames>",
                    '<definedName name="%s" comment="Displays this module\'s repository URL and function list">%s</definedName></definedNames>'
                    % (full, xesc(body)))
    print("%s defined, %d chars" % (full, len(body)))
# same upstream copy-paste bug in the installed u-module About
wb = re.sub(r'(<definedName name="nabla\.u\.Aboutλ"[^>]*>[^<]*?)Suggested module name: nabla\.e',
            r'\g<1>Suggested module name: nabla.u', wb)
put("xl/workbook.xml", wb)

# ---------- 9c. Remove MACRS from the compiled workbook and register the new functions ----------
WB_MACRS = [
    ('{"SLN","SYD","DB","DDB","VDB","MACRS"}', '{"SLN","SYD","DB","DDB","VDB","DV","PC"}'),
    (' + N(_xlpm.Method = "MACRS")', ''),
    ('IF(_xlpm.Method = "MACRS", 0, ', 'IF(OR(_xlpm.Method = "DV", _xlpm.Method = "PC"), 0, '),
    ('_xlpm.DisposalDate, IF(_xlpm.Method = "MACRS", MAX(EDATE(_xlpm.InServiceDate, '
     '_xlpm.Years * _xlpm.Mpy), _xlfn.SINGLE(INDEX(_xlpm.DisposalDates, _xlpm.Asset))), '
     '_xlfn.SINGLE(INDEX(_xlpm.DisposalDates, _xlpm.Asset)))',
     '_xlpm.DisposalDate, _xlfn.SINGLE(INDEX(_xlpm.DisposalDates, _xlpm.Asset))'),
    ('nabla.f.MACRSλ(_xlpm.InitialValue, _xlpm.Years - 1)',
     'nabla.f.DiminishingValueλ(_xlpm.InitialValue, _xlpm.Years), '
     'nabla.f.PrimeCostλ(_xlpm.InitialValue, _xlpm.Years)'),
    ('"SLN,SYD,DB,DDB,VDB,MACRS"', '"SLN,SYD,DB,DDB,VDB,DV,PC"'),
    ('Methods must be omitted or one of: SLN, SYD, DB, DDB, MACRS, or VDB.',
     'Methods must be omitted or one of: SLN, SYD, DB, DDB, VDB, DV, or PC.'),
    ('Must be one of these Excel function names: ', 'Must be one of these method codes: '),
    ('"→MACRS=Modified Accelerated Cost Recovery System. NOTE: Salvage value ignored¶" &amp; ',
     '"→DV =Diminishing balance at 200% of straight line. Salvage value ignored¶" &amp; '
     '"→PC =Straight line, whole years. Salvage value ignored¶" &amp; '),
]
wb = get("xl/workbook.xml")
for old, new in WB_MACRS:
    assert wb.count(old) == 1, old[:70]
    wb = wb.replace(old, new)
wb, n = re.subn(r'<definedName name="nabla\.f\.MACRSλ"[^>]*>[^<]*</definedName>', "", wb)
assert n == 1
# list the new f-module functions in its About table
anchor = ('"VDBλ               →Variable declining balance depreciation method for one asset '
          'or asset class.¶" &amp; ')
assert wb.count(anchor) == 1
wb = wb.replace(anchor, anchor
    + "".join('"%-19s→%s¶" &amp; ' % (f["name"], xesc(f["desc"])) for f in _dep)
    + '"→¶" &amp; "AUSTRALIAN TAX     →¶" &amp; '
    + "".join('"%-19s→%s¶" &amp; ' % (f["name"], xesc(f["desc"])) for f in _gst))
# register every new function as a defined name
for spec in FUNCS:
    full = spec["module"] + "." + spec["name"]
    assert '<definedName name="%s"' % full not in wb, full
    wb = wb.replace("</definedNames>", '<definedName name="%s" comment="%s">%s</definedName></definedNames>'
                    % (full, xesc(spec["desc"]), build_xml(spec)))
assert "MACRS" not in wb
put("xl/workbook.xml", wb)
print("MACRS removed;", len(FUNCS), "Australian functions registered")

# ---------- 9d. Clear MACRS out of cached help spills on worksheets ----------
cache_fixes = 0
for n in list(parts):
    if re.match(r'xl/worksheets/sheet\d+\.xml$', n):
        d = get(n)
        d2 = d.replace("MACRS=Modified Accelerated Cost Recovery System. NOTE: Salvage value ignored",
                       "DV =Diminishing balance at 200% of straight line. Salvage value ignored")
        d2 = d2.replace("Must be one of these Excel function names: ",
                        "Must be one of these method codes: ")
        d2 = d2.replace("<v>MACRS</v>", "<v>DV</v>")
        if d2 != d:
            cache_fixes += 1
            put(n, d2)
ss0 = get("xl/sharedStrings.xml")
ss1 = re.sub(r'<si><t[^>]*>MACRS</t></si>', '<si><t>DV</t></si>', ss0)
assert ss1 != ss0
put("xl/sharedStrings.xml", ss1)
for n in list(parts):
    if n.endswith(".xml"):
        assert "MACRS" not in get(n), n
print("cleared cached MACRS text on", cache_fixes, "sheets")

# ---------- 9e. Data Validation sheet: add the prime cost method row ----------
ss = get("xl/sharedStrings.xml")
cnt = int(re.search(r'count="(\d+)"', ss).group(1))
uniq = int(re.search(r'uniqueCount="(\d+)"', ss).group(1))
new_si = ["PC", "Straight line"]
ss = ss.replace("</sst>", "".join("<si><t>%s</t></si>" % v for v in new_si) + "</sst>")
ss = ss.replace('count="%d" uniqueCount="%d"' % (cnt, uniq),
                'count="%d" uniqueCount="%d"' % (cnt + 2, uniq + 2), 1)
put("xl/sharedStrings.xml", ss)
s3 = get("xl/worksheets/sheet3.xml")
row12 = re.search(r'<row r="12"[^>]*>.*?</row>', s3, re.S)
assert row12
style = re.search(r'<c r="A12" s="(\d+)"', row12.group(0))
style_b = re.search(r'<c r="B12" s="(\d+)"', row12.group(0))
sa = ' s="%s"' % style.group(1) if style else ""
sb = ' s="%s"' % style_b.group(1) if style_b else ""
row13 = ('<row r="13" spans="1:13">'
         '<c r="A13"%s t="s"><v>%d</v></c><c r="B13"%s t="s"><v>%d</v></c></row>'
         % (sa, uniq, sb, uniq + 1))
s3 = s3.replace(row12.group(0), row12.group(0) + row13)
s3 = s3.replace('<dimension ref="A1:M12"/>', '<dimension ref="A1:M13"/>')
put("xl/worksheets/sheet3.xml", s3)
t2 = get("xl/tables/table2.xml")  # tblMethods must cover the new row
assert 'ref="A6:B12"' in t2
put("xl/tables/table2.xml", t2.replace('ref="A6:B12"', 'ref="A6:B13"'))
print("Data Validation sheet: PC row added, tblMethods extended")

# ---------- 9e1. Current-Excel guidance: point at the natives that now overlap ----------
# Excel 365 has gained functions since the upstream release that do natively what a few of
# these helpers were written to work around. Checked against Microsoft's documentation on
# 18 August 2026.
wb = get("xl/workbook.xml")
added = 0
for fname, note in SEE_ALSO.items():
    # anchor on the label after the description, which may run over several lines
    pat = (r'(<definedName name="%s"[^>]*>.*?"DESCRIPTION:.*?)("(?:VERSION|WEBPAGE|WEBSITE|PARAMETERS):)'
           % re.escape(fname))
    wb, k = re.subn(pat, r'\g<1>"SEE ALSO:      →%s¶" &amp; \g<2>' % note, wb, count=1)
    assert k == 1, fname
    added += k
put("xl/workbook.xml", wb)
print("current-Excel notes added to", added, "functions")

# ---------- 9e2. TOC: the Aboutλ row is a function, not a worksheet ----------
s2 = get("xl/worksheets/sheet2.xml")
assert '<c r="B28" s="108" t="s"><v>134</v></c>' in s2   # 134 = "Worksheet"
put("xl/worksheets/sheet2.xml",
    s2.replace('<c r="B28" s="108" t="s"><v>134</v></c>',
               '<c r="B28" s="108" t="s"><v>459</v></c>'))  # 459 = "Function"

# ---------- 9e3. Cover: drop the merge left behind by the removed section ----------
cov = get("xl/worksheets/sheet1.xml")
m_cnt = re.search(r'<mergeCells count="(\d+)">', cov)
cov2 = cov.replace('<mergeCell ref="A29:B29"/>', "")
assert cov2 != cov
cov2 = cov2.replace(m_cnt.group(0), '<mergeCells count="%d">' % (int(m_cnt.group(1)) - 1))
put("xl/worksheets/sheet1.xml", cov2)

# ---------- 9e4. Drop calcChain: a pure calculation-order cache Excel rebuilds ----------
del parts["xl/calcChain.xml"]
rels = get("xl/_rels/workbook.xml.rels")
rels2 = re.sub(r'<Relationship Id="rId\d+"[^>]*calcChain[^>]*/>', "", rels)
assert rels2 != rels
put("xl/_rels/workbook.xml.rels", rels2)
ct = get("[Content_Types].xml")
ct2 = re.sub(r'<Override PartName="/xl/calcChain\.xml"[^>]*/>', "", ct)
assert ct2 != ct
put("[Content_Types].xml", ct2)

# ---------- 9f. Drop the upstream printer configuration; print on A4 ----------
# The printerSettings parts embed the original author's printer name and US Letter paper.
for i in (1, 2, 3):
    p = "xl/printerSettings/printerSettings%d.bin" % i
    assert p in parts
    del parts[p]
ct = get("[Content_Types].xml")
ct2 = re.sub(r'<Default Extension="bin"[^>]*printerSettings[^>]*/>', "", ct)
assert ct2 != ct
assert len(re.findall(r'<Override PartName="/xl/customProperty\d+\.bin"', ct2)) == 38
put("[Content_Types].xml", ct2)
for sheet, rid in [("sheet1", "rId4"), ("sheet2", "rId1"), ("sheet6", "rId1")]:
    rp = "xl/worksheets/_rels/%s.xml.rels" % sheet
    d = get(rp)
    d2 = re.sub(r'<Relationship Id="%s"[^>]*printerSettings[^>]*/>' % rid, "", d)
    assert d2 != d, rp
    put(rp, d2)
    sp = "xl/worksheets/%s.xml" % sheet
    d = get(sp)
    d2 = re.sub(r'<pageSetup([^>]*) r:id="%s"/>' % rid, r'<pageSetup paperSize="9"\1/>', d)
    assert d2 != d, sp
    put(sp, d2)
assert not [n for n in parts if "printerSettings" in n]
print("printer configuration removed; A4 set on 3 sheets")

# ---------- 9f2. A worksheet demonstrating the Australian functions ----------
# The five additions had no demonstration sheet, unlike every other headline function.
st_now = get("xl/styles.xml")
cellxfs_m = re.search(r'<cellXfs count="(\d+)">', st_now)
base_xf = int(cellxfs_m.group(1))
NEW_XFS = ('<xf numFmtId="0" fontId="4" fillId="0" borderId="0" xfId="0" applyFont="1"/>'      # heading
           '<xf numFmtId="164" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>'  # currency
           '<xf numFmtId="171" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>')  # date
st_now = st_now.replace(cellxfs_m.group(0), '<cellXfs count="%d">' % (base_xf + 3), 1)
st_now = st_now.replace("</cellXfs>", NEW_XFS + "</cellXfs>", 1)
put("xl/styles.xml", st_now)
H, MONEY, DATE = base_xf, base_xf + 1, base_xf + 2

def txt(ref, s, value):
    return '<c r="%s" s="%d" t="inlineStr"><is><t>%s</t></is></c>' % (ref, s, value)
def num(ref, s, value):
    return '<c r="%s" s="%d"><v>%s</v></c>' % (ref, s, value)
def fml(ref, s, formula, spill=None):
    f = ('<f t="array" ref="%s">%s</f>' % (spill, xesc(formula))) if spill else '<f>%s</f>' % xesc(formula)
    cm = ' cm="1"' if spill else ''
    return '<c r="%s" s="%d"%s>%s</c>' % (ref, s, cm, f)

AU_ROWS = {
    1:  '<c r="A1" s="270" t="str"><f>_xlfn.TEXTAFTER(CELL("filename",A1),"]")</f><v>Australian tax</v></c>',
    2:  txt("A2", 0, "The Australian additions to this library. Change the input cells and the results follow."),
    4:  txt("A4", H, "DEPRECIATION"),
    5:  txt("A5", H, "Cost") + txt("B5", H, "Effective life (years)"),
    6:  num("A6", MONEY, 50000) + num("B6", 0, 5),
    8:  txt("A8", 0, "Diminishing balance at 200% of straight line")
        + fml("B8", MONEY, "nabla.f.DiminishingValueλ($A$6,$B$6)", spill="B8:F8"),
    9:  txt("A9", 0, "Straight line, whole years")
        + fml("B9", MONEY, "nabla.f.PrimeCostλ($A$6,$B$6)", spill="B9:F9"),
    # Both schedules must write off the whole cost, whatever the effective life. Shown on
    # the sheet because it is the property that a part-year or sub-two-year life breaks.
    10: txt("A10", 0, "Total written off, diminishing balance")
        + fml("B10", MONEY, "SUM(_xlfn.ANCHORARRAY(B8))"),
    11: txt("A11", 0, "Total written off, straight line")
        + fml("B11", MONEY, "SUM(_xlfn.ANCHORARRAY(B9))"),
    12: txt("A12", 0, "Both totals equal the cost above for any effective life, "
                      "including part years such as 6 2/3 and lives under 2 years, "
                      "because the diminishing-balance schedule writes its residual off "
                      "in the final period."),
    # its own row: a cell belongs to the row its address names, and Excel refuses to open
    # a file where A13 sits inside <row r="12">
    13: txt("A13", 0, "Modelling schedules, not tax calculations: no acquisition date, "
                      "no income year, no days held, no disposal. Do not use them to "
                      "prepare a return."),
    14: txt("A14", H, "GST"),
    15: txt("A15", H, "GST-exclusive amount") + txt("B15", H, "Plus GST"),
    # marked as dynamic-array cells so Excel does not store them as legacy formulas and
    # display an implicit-intersection @ in front of a function documented to take a range
    16: num("A16", MONEY, 1000)
        + fml("B16", MONEY, "nabla.f.GSTAddλ(A16)", spill="B16:B16"),
    18: txt("A18", H, "GST-inclusive amount") + txt("B18", H, "GST included"),
    19: num("A19", MONEY, 1100)
        + fml("B19", MONEY, "nabla.f.GSTExtractλ(A19)", spill="B19:B19"),
    21: txt("A21", H, "FINANCIAL YEAR"),
    # One spilled call over the whole column, which is how the function is meant to be
    # used. Labelling dates one cell at a time would not exercise the array path.
    22: txt("A22", H, "Date") + txt("B22", H, "Financial year"),
    23: num("A23", DATE, 46203)                                                  # 30 Jun 2026
        + fml("B23", 0, "nabla.d.FinancialYearλ(A23:A26)", spill="B23:B26"),
    24: num("A24", DATE, 46204),                                                 # 1 Jul 2026
    25: num("A25", DATE, 46249),                                                 # 15 Aug 2026
    26: num("A26", DATE, 46387),                                                 # 31 Dec 2026
}
for _serial, _date in ((46203, (2026, 6, 30)), (46204, (2026, 7, 1)),
                       (46249, (2026, 8, 15)), (46387, (2026, 12, 31))):
    assert (EPOCH + datetime.timedelta(days=_serial)) == datetime.datetime(*_date), _serial
AU_SHEET = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
    '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
    'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
    '<dimension ref="A1:F26"/>'
    '<sheetViews><sheetView showGridLines="0" workbookViewId="0">'
    '<selection activeCell="A1" sqref="A1"/></sheetView></sheetViews>'
    '<sheetFormatPr defaultRowHeight="15"/>'
    '<cols><col min="1" max="1" width="34" customWidth="1"/>'
    '<col min="2" max="6" width="16" customWidth="1"/></cols><sheetData>'
    + "".join('<row r="%d">%s</row>' % (r, AU_ROWS[r]) for r in sorted(AU_ROWS))
    + '</sheetData>'
    '<hyperlinks><hyperlink ref="A1" location="\'TOC\'!$A$1" display="\'TOC\'!$A$1"/></hyperlinks>'
    '<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>'
    '<pageSetup paperSize="9" orientation="portrait"/></worksheet>')
AU_PART, AU_RID, AU_SHEETID = "xl/worksheets/sheet50.xml", "rId62", 132
assert AU_PART not in parts
parts[AU_PART] = AU_SHEET.encode("utf-8")
ct = get("[Content_Types].xml")
put("[Content_Types].xml", ct.replace("</Types>",
    '<Override PartName="/xl/worksheets/sheet50.xml" '
    'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>'))
rels = get("xl/_rels/workbook.xml.rels")
assert AU_RID not in rels
put("xl/_rels/workbook.xml.rels", rels.replace("</Relationships>",
    '<Relationship Id="%s" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" '
    'Target="worksheets/sheet50.xml"/></Relationships>' % AU_RID))
wbx = get("xl/workbook.xml")
anchor_sheet = re.search(r'<sheet name="Data Validation"[^>]*/>', wbx)
assert anchor_sheet
put("xl/workbook.xml", wbx.replace(anchor_sheet.group(0), anchor_sheet.group(0)
    + '<sheet name="Australian tax" sheetId="%d" r:id="%s"/>' % (AU_SHEETID, AU_RID)))
# list it in the table of contents
toc = get("xl/worksheets/sheet2.xml")
assert '<row r="70"' not in toc
row70 = ('<row r="70" spans="1:4">'
         + txt("A70", 110, "Australian tax") + txt("B70", 108, "Worksheet")
         + txt("C70", 111, "Australian tax") + txt("D70", 109,
             "depreciation, GST and financial-year helpers") + '</row>')
toc = toc.replace("</sheetData>", row70 + "</sheetData>")
toc = toc.replace('<dimension ref="A1:F69"/>', '<dimension ref="A1:F70"/>')
toc = toc.replace('<hyperlinks>', '<hyperlinks><hyperlink ref="A70" location="\'Australian tax\'!$A$1" '
                  'display="\'Australian tax\'!$A$1"/>', 1)
put("xl/worksheets/sheet2.xml", toc)
t1 = get("xl/tables/table1.xml")
assert 'ref="A4:D69"' in t1
t1 = t1.replace('ref="A4:D69"', 'ref="A4:D70"')

# Upstream ships the contents filtered to Type = Worksheet, so the workbook opens with
# every tblBudget row hidden: 16 of 66 entries invisible until someone notices the slicer.
# Drop the criteria and unhide the rows. The <autoFilter> element itself stays, because
# the Type slicer binds to it; with no filterColumn it simply has nothing applied.
fcol = re.search(r'<filterColumn colId="1">.*?</filterColumn>', t1, re.S)
assert fcol and 'val="Worksheet"' in fcol.group(0), "TOC filter not found where expected"
t1 = t1.replace(fcol.group(0), "")
assert "<filterColumn" not in t1
put("xl/tables/table1.xml", t1)

toc = get("xl/worksheets/sheet2.xml")
unhidden = 0
def _show(m):
    global unhidden
    row = int(m.group(1))
    if not (5 <= row <= 70):        # only the table's own data rows
        return m.group(0)
    unhidden += 1
    return m.group(0).replace(' hidden="1"', "")
toc = re.sub(r'<row r="(\d+)"[^>]*hidden="1"[^>]*>', _show, toc)
assert unhidden == 16, unhidden
assert 'hidden="1"' not in toc
put("xl/worksheets/sheet2.xml", toc)
print("added the Australian tax worksheet; cleared the TOC filter and unhid", unhidden, "rows")

# docProps/app.xml repeats the sheet list for the file properties dialogue and for anything
# that reads a workbook's shape without opening it. The new worksheet was registered in the
# content types, the relationships, the workbook, the contents table and its table range,
# and not here, so the part described a 49-sheet workbook that has 50. Excel rewrote it on
# the first save, which is why it looked right in the shipped file and wrong in a fresh
# build. Insert the title where the sheet sits, and bump both counts that describe it.
_app = get("docProps/app.xml")
assert "<vt:lpstr>Australian tax</vt:lpstr>" not in _app
_needle = "<vt:lpstr>FMTs</vt:lpstr>"
assert _app.count(_needle) == 1, "no FMTs entry to insert before"
_app = _app.replace(_needle, "<vt:lpstr>Australian tax</vt:lpstr>" + _needle, 1)
_app, _n = re.subn(r'(<TitlesOfParts><vt:vector size=")(\d+)("\s+baseType="lpstr")',
                   lambda m: m.group(1) + str(int(m.group(2)) + 1) + m.group(3), _app, count=1)
assert _n == 1, "no TitlesOfParts vector size"
_app, _n = re.subn(r"(<vt:lpstr>Worksheets</vt:lpstr></vt:variant><vt:variant><vt:i4>)(\d+)(</vt:i4>)",
                   lambda m: m.group(1) + str(int(m.group(2)) + 1) + m.group(3), _app, count=1)
assert _n == 1, "no worksheet count in HeadingPairs"
put("docProps/app.xml", _app)
print("listed the new worksheet in docProps/app.xml")

# ---------- 9g. Performance: freeze the volatile demo data ----------
# 38 RANDBETWEEN formulas made 93% of the workbook's formula cells volatile, so every edit
# recalculated almost everything. The sample data does not need to be random: fixed values
# make the demonstrations reproducible and cost nothing to recalculate.
def static_cells(part, values, numeric=True):
    """Replace whole cells with literal values, dropping their formulas."""
    d = get(part)
    for cell, val in values.items():
        pat = r'<c r="%s"( [^>]*?)?(?:/>|>.*?</c>)' % cell
        m = re.search(pat, d, re.S)
        assert m, (part, cell)
        attrs = (m.group(1) or "")
        attrs = re.sub(r' (?:t|cm|ca|aca)="[^"]*"', "", attrs)
        body = '<v>%s</v>' % val if numeric else '<is><t>%s</t></is>' % val
        if not numeric:
            attrs += ' t="inlineStr"'
        d = d[:m.start()] + '<c r="%s"%s>%s</c>' % (cell, attrs, body) + d[m.end():]
    put(part, d)

# Read the customer names straight out of tblCT (sheet13 B28:B32) so the frozen invoice rows
# can never name a customer the lookup cannot find.
_ss = get("xl/sharedStrings.xml")
_si = re.findall(r'<si>(.*?)</si>', _ss, re.S)
def _shared(i):
    return "".join(re.findall(r'<t[^>]*>([^<]*)</t>', _si[i]))
_s13 = get("xl/worksheets/sheet13.xml")
CUSTOMERS = []
for _r in range(28, 33):
    _m = re.search(r'<c r="B%d"[^>]*t="s"[^>]*><v>(\d+)</v></c>' % _r, _s13)
    assert _m, "tblCT customer row %d not found" % _r
    CUSTOMERS.append(_shared(int(_m.group(1))))
assert len(CUSTOMERS) == 5 and all(CUSTOMERS), CUSTOMERS
print("tblCT customers:", CUSTOMERS)
# tblCO: 21 invoices, issued across the first half of 2026, amounts 1,000 to 5,000
issued = {"G%d" % r: 46023 + ((r - 20) * 7 + 3) for r in range(20, 41)}
amounts = {"I%d" % r: 1000 + ((r - 20) * 197 % 4001) for r in range(20, 41)}
custs = {"F%d" % r: CUSTOMERS[(r - 20) % len(CUSTOMERS)] for r in range(20, 41)}
static_cells("xl/worksheets/sheet13.xml", issued)
static_cells("xl/worksheets/sheet13.xml", amounts)
static_cells("xl/worksheets/sheet13.xml", custs, numeric=False)
# tblCO must not re-inject the formulas when the table is edited
t10 = get("xl/tables/table10.xml")
t10b, n = re.subn(r'<calculatedColumnFormula[^>]*>[^<]*</calculatedColumnFormula>', "", t10)
assert n >= 3, n
put("xl/tables/table10.xml", t10b)

# Periodsλ: fixed start/end pairs and interval codes
# the "Y" row must span more than a year or the demo returns zero
static_cells("xl/worksheets/sheet9.xml", {"A25": 46081, "B25": 46356, "A26": 46203, "B26": 46600,
                                          "A27": 46023, "B27": 47119, "A28": 46142, "B28": 46508,
                                          "A29": 46265, "B29": 46630})
static_cells("xl/worksheets/sheet9.xml", {"C25": "M", "C26": "Q", "C27": "Y", "C28": "W", "C29": "D"},
             numeric=False)
# TimelineOffsetλ: fixed event dates through 2026
static_cells("xl/worksheets/sheet31.xml", {"A%d" % r: 46023 + (r - 19) * 47 for r in range(19, 26)})
# IsBetweenλ: onboarding dates spaced a month apart instead of randomly scattered
RAND_ONBOARD = ('RANDBETWEEN(1,$B$22*_xlfn.SWITCH($B$23, "Y", 365, "M", 30, "Q", 90, "W", 7, "D", 1))')
SPACED = '365 * (ROW() - ROW(tblOnBoarding[#Headers]))'
for part in ("xl/worksheets/sheet47.xml", "xl/tables/table17.xml"):
    d = get(part)
    assert RAND_ONBOARD in d, part
    put(part, d.replace(RAND_ONBOARD, SPACED))
# IRRλ: a fixed investment and a rising series of returns instead of random ones
s26 = get("xl/worksheets/sheet26.xml")
for old, new in [("_xlpm.Pad,RANDBETWEEN(1,3)", "_xlpm.Pad,2"),
                 ("RANDBETWEEN(1000,2000)", "1000 + _xlpm.C * 100"),
                 ("-RANDBETWEEN(1000,9000)", "-5000")]:
    assert old in s26, old
    s26 = s26.replace(old, new)
put("xl/worksheets/sheet26.xml", s26)

# RANDARRAY is volatile too. Replace each with a deterministic MAKEARRAY over the same
# shape and range, so the grids stay varied but stop recalculating on every edit.
def _split_args(text, start):
    """Split the argument list of a call whose '(' is at `start`; returns (args, end)."""
    depth, args, cur, in_str, i = 0, [], "", False, start
    while i < len(text):
        ch = text[i]
        if ch == '"':
            in_str = not in_str
        if not in_str:
            if ch == "(":
                depth += 1
                if depth == 1:
                    i += 1
                    continue
            elif ch == ")":
                depth -= 1
                if depth == 0:
                    args.append(cur)
                    return args, i + 1
            elif ch == "," and depth == 1:
                args.append(cur); cur = ""; i += 1; continue
        cur += ch
        i += 1
    raise ValueError("unbalanced call")

randarray = 0
for n in list(parts):
    if not re.match(r'xl/worksheets/sheet\d+\.xml$', n):
        continue
    d = get(n)
    while "_xlfn.RANDARRAY(" in d:
        at = d.index("_xlfn.RANDARRAY(")
        args, end = _split_args(d, at + len("_xlfn.RANDARRAY"))
        assert len(args) in (2, 3, 4, 5), (n, args)   # min/max and the whole-number flag are optional
        rows, cols = (a.strip() for a in args[:2])
        if len(args) >= 4:
            lo, hi = args[2].strip(), args[3].strip()
            body = "%s + MOD(_xlpm.r * 7 + _xlpm.c * 13, %s - %s + 1)" % (lo, hi, lo)
        else:   # RANDARRAY's own default range is 0 to 1, returned as decimals
            body = "MOD(_xlpm.r * 7 + _xlpm.c * 13, 100) / 100"
        repl = '_xlfn.MAKEARRAY(%s, %s, _xlfn.LAMBDA(_xlpm.r,_xlpm.c, %s))' % (rows, cols, body)
        d = d[:at] + repl + d[end:]
        randarray += 1
    put(n, d)
print("replaced", randarray, "volatile RANDARRAY grids")

for n in list(parts):
    if re.match(r'xl/(?:worksheets/sheet\d+|tables/table\d+)\.xml$', n):
        d = get(n)
        formulas = re.findall(r'<f[^>]*>(.*?)</f>', d, re.S)
        formulas += re.findall(r'<calculatedColumnFormula[^>]*>(.*?)</calculatedColumnFormula>', d, re.S)
        for f in formulas:   # worksheet prose may still mention these by name
            assert "RANDBETWEEN(" not in f and "RANDARRAY(" not in f, (n, f[:60])

# Drop the always-calculate flags now that nothing but the sheet-name titles is volatile.
cleared = 0
for n in list(parts):
    if not re.match(r'xl/worksheets/sheet\d+\.xml$', n):
        continue
    d = get(n)
    def strip_flags(mo):
        global cleared
        cell = mo.group(0)
        if "CELL(" in cell:       # the A1 title formula is legitimately volatile
            return cell
        new = cell.replace(' ca="1"', "").replace(' aca="1"', "")
        cleared += cell.count('ca="1"')
        return new
    d = re.sub(r'<c r="[A-Z]+\d+"[^>]*>(?:.*?</c>)?', strip_flags, d, flags=re.S)
    put(n, d)
print("volatile demo data frozen;", cleared, "always-calculate flags cleared")

# ---------- 9h. Presentation ----------
TAB = {"Australian tax": "FF157A5F", "nabla.d": "FF1F4E79", "nabla.e": "FF2E75B6", "nabla.f": "FF157A5F", "nabla.r": "FF375623"}
wbx = get("xl/workbook.xml")
relmap = dict(re.findall(r'Id="(rId\d+)"[^>]*Target="([^"]+)"', get("xl/_rels/workbook.xml.rels")))
styled = 0
for nm, rid in re.findall(r'<sheet name="([^"]+)"[^>]*r:id="(rId\d+)"', wbx):
    path = "xl/" + relmap[rid]
    d = get(path)
    colour = TAB.get(".".join(nm.split(".")[:2]), "FF595959")
    tab = '<tabColor rgb="%s"/>' % colour
    if "<sheetPr" in d:
        d = re.sub(r'<sheetPr([^>]*?)/>', r'<sheetPr\g<1>>%s</sheetPr>' % tab, d, count=1)
        d = re.sub(r'(<sheetPr[^>]*[^/]>)(?!<tabColor)', r'\g<1>%s' % tab, d, count=1)
    else:
        d = re.sub(r'(<worksheet[^>]*>)', r'\g<1><sheetPr>%s</sheetPr>' % tab, d, count=1)
    # hide gridlines and open every sheet at the top left.
    # the lookahead for a space or > keeps this off the <sheetViews> container.
    d = re.sub(r'<sheetView(?=[ >])(?![^>]*showGridLines)', '<sheetView showGridLines="0"', d)
    d = d.replace('showGridLines="1"', 'showGridLines="0"')
    d = re.sub(r'(<sheetView[^>]*?)\s*topLeftCell="[^"]*"', r'\g<1>', d)
    # only the cover should open selected, matching the removed activeTab
    d = re.sub(r'(<sheetView[^>]*?)\s*tabSelected="1"', r'\g<1>', d)
    if nm == "Cover":
        d = d.replace("<sheetView ", '<sheetView tabSelected="1" ', 1)
    if "<pane " not in d:   # leave split/frozen sheets to their own pane selections
        d = re.sub(r'<selection[^>]*/>', '<selection activeCell="A1" sqref="A1"/>', d)
    put(path, d)
    styled += 1
assert styled == 50, styled
# open on the cover, not on whichever tab was last active
wbx2 = re.sub(r'(<workbookView[^>]*?)\s*activeTab="\d+"', r'\g<1>', wbx)
assert wbx2 != wbx
put("xl/workbook.xml", wbx2)
toc_cols = get("xl/worksheets/sheet2.xml")
for col, width in (("1", "31.5"), ("3", "30.5")):
    toc_cols = re.sub(r'(<col min="%s" max="%s" width=")[\d.]+' % (col, col), r'\g<1>' + width, toc_cols)
put("xl/worksheets/sheet2.xml", toc_cols)
print("presentation: tab colours, gridlines and opening view set on", styled, "sheets")

# ---------- 10. core.xml: title + modified ----------
core = get("docProps/core.xml")
core = re.sub(r'(<dcterms:modified[^>]*>)[^<]*(</dcterms:modified>)',
              r'\g<1>' + NOW_ISO + r'\g<2>', core)
core = core.replace("<cp:lastModifiedBy>",
                    "<dc:title>nabla</dc:title><dc:subject>LAMBDA function library for dynamic-array financial models</dc:subject>"
                    "<dc:creator>Craig Hatmaker (original library); nabla project (derivative)</dc:creator><cp:lastModifiedBy>")
put("docProps/core.xml", core)

# ---------- 11. Flat nb. namespace ----------
# Every function lives in one namespace, nb., so callers type three characters
# instead of eight. Where two modules shipped the same base name, the fuller
# implementation keeps the plain name and the other takes a module tag
# (B borrowings, E essentials, U utilities). The five About tables take words,
# because nb.AboutRλ would tell a reader nothing.
FLAT_EXPLICIT = [
    ("nabla.debt.Amortiseλ", "nb.AmortiseBλ"),
    ("nabla.u.CountAColsλ", "nb.CountAColsUλ"),
    ("nabla.u.CountARowsλ", "nb.CountARowsUλ"),
    ("nabla.e.IsBetweenλ", "nb.IsBetweenEλ"),
    ("nabla.e.RangeToDAλ", "nb.RangeToDAEλ"),
    ("nabla.u.CountColsλ", "nb.CountColsUλ"),
    ("nabla.u.CountRowsλ", "nb.CountRowsUλ"),
    ("nabla.u.IsBetweenλ", "nb.IsBetweenUλ"),
    ("nabla.u.RangeToDAλ", "nb.RangeToDAUλ"),
    ("nabla.u.IsInListλ", "nb.IsInListUλ"),
    ("nabla.u.AvgColsλ", "nb.AvgColsUλ"),
    ("nabla.u.AvgRowsλ", "nb.AvgRowsUλ"),
    ("nabla.u.MaxColsλ", "nb.MaxColsUλ"),
    ("nabla.u.MaxRowsλ", "nb.MaxRowsUλ"),
    ("nabla.u.MinColsλ", "nb.MinColsUλ"),
    ("nabla.u.MinRowsλ", "nb.MinRowsUλ"),
    ("nabla.u.SumColsλ", "nb.SumColsUλ"),
    ("nabla.u.SumRowsλ", "nb.SumRowsUλ"),
    ("nabla.u.CountCλ", "nb.CountCUλ"),
    ("nabla.d.Aboutλ", "nb.AboutDatesλ"),
    ("nabla.e.Aboutλ", "nb.AboutEssentialsλ"),
    ("nabla.f.Aboutλ", "nb.AboutFinancialλ"),
    ("nabla.r.Aboutλ", "nb.AboutRatiosλ"),
    ("nabla.u.Aboutλ", "nb.AboutUtilitiesλ"),
]
FLAT_PREFIX = [("nabla.debt.", "nb."), ("nabla.d.", "nb."), ("nabla.e.", "nb."),
               ("nabla.f.", "nb."), ("nabla.r.", "nb."), ("nabla.u.", "nb.")]
FLAT_BARE = [("nabla.debt", "nb"), ("nabla.d", "nb"), ("nabla.e", "nb"),
             ("nabla.f", "nb"), ("nabla.r", "nb"), ("nabla.u", "nb")]
# The About tables list their functions by bare name, so the tagged ones need it here too.
ABOUT_ENTRY_FIXES = {
    "nb.AboutUtilitiesλ": [(b + "λ", b + "Uλ") for b in (
        "CountC", "SumCols", "SumRows", "AvgCols",
        "AvgRows", "MinCols", "MinRows", "MaxCols",
        "MaxRows", "CountCols", "CountRows", "CountACols",
        "CountARows", "IsBetween", "IsInList", "RangeToDA",
    )],
    "nb.AboutEssentialsλ": [("IsBetweenλ", "IsBetweenEλ"), ("RangeToDAλ", "RangeToDAEλ")],
    "nb.AboutFinancialλ": [("Aboutλ", "AboutFinancialλ")],
    "nb.AboutRatiosλ": [("Aboutλ", "AboutRatiosλ")],
}
# The AFE module containers cannot all collapse to "nb" without colliding, so they take words.
AFE_MODULES = [("nabla.debt", "Debt"), ("nabla.d", "Dates"), ("nabla.e", "Essentials"),
               ("nabla.f", "Financial"), ("nabla.r", "Ratios"), ("nabla.u", "Utilities")]

def flatten_names(s):
    for a, b in FLAT_EXPLICIT:
        s = s.replace(a, b)
    for a, b in FLAT_PREFIX:
        s = s.replace(a, b)
    for a, b in FLAT_BARE:
        s = s.replace(a, b)
    return s

def capitalise_brand(s):
    # the chart template filename and the theme colour scheme are asset ids, not prose
    s = s.replace("nabla Combo Area", "@@CRTX@@").replace("nabla TnC", "@@TNC@@")
    s = re.sub(r"(?<![A-Za-z/])nabla(?![A-Za-z])", "Nabla", s)
    return s.replace("@@CRTX@@", "nabla Combo Area").replace("@@TNC@@", "nabla TnC")

def fix_about_entries(s):
    for nm, subs in ABOUT_ENTRY_FIXES.items():
        m = re.search(r'(<definedName name="%s"[^>]*>)(.*?)(</definedName>)' % re.escape(nm), s, re.S)
        if not m:
            continue
        body = m.group(2)
        for a, b in subs:
            body = re.sub(r"(?<![A-Za-z])" + re.escape(a), b, body)
        s = s[:m.start()] + m.group(1) + body + m.group(3) + s[m.end():]
    return s

# A renamed function still prints its old bare name in its own inline help.
SELF_NAME_FIXES = [
    ('nb.AmortiseBλ', 'Amortiseλ', 'AmortiseBλ'),
    ('nb.CountAColsUλ', 'CountAColsλ', 'CountAColsUλ'),
    ('nb.CountARowsUλ', 'CountARowsλ', 'CountARowsUλ'),
    ('nb.IsBetweenEλ', 'IsBetweenλ', 'IsBetweenEλ'),
    ('nb.RangeToDAEλ', 'RangeToDAλ', 'RangeToDAEλ'),
    ('nb.CountColsUλ', 'CountColsλ', 'CountColsUλ'),
    ('nb.CountRowsUλ', 'CountRowsλ', 'CountRowsUλ'),
    ('nb.IsBetweenUλ', 'IsBetweenλ', 'IsBetweenUλ'),
    ('nb.RangeToDAUλ', 'RangeToDAλ', 'RangeToDAUλ'),
    ('nb.IsInListUλ', 'IsInListλ', 'IsInListUλ'),
    ('nb.AvgColsUλ', 'AvgColsλ', 'AvgColsUλ'),
    ('nb.AvgRowsUλ', 'AvgRowsλ', 'AvgRowsUλ'),
    ('nb.MaxColsUλ', 'MaxColsλ', 'MaxColsUλ'),
    ('nb.MaxRowsUλ', 'MaxRowsλ', 'MaxRowsUλ'),
    ('nb.MinColsUλ', 'MinColsλ', 'MinColsUλ'),
    ('nb.MinRowsUλ', 'MinRowsλ', 'MinRowsUλ'),
    ('nb.SumColsUλ', 'SumColsλ', 'SumColsUλ'),
    ('nb.SumRowsUλ', 'SumRowsλ', 'SumRowsUλ'),
    ('nb.CountCUλ', 'CountCλ', 'CountCUλ'),
    ('nb.AboutDatesλ', 'Aboutλ', 'AboutDatesλ'),
    ('nb.AboutEssentialsλ', 'Aboutλ', 'AboutEssentialsλ'),
    ('nb.AboutFinancialλ', 'Aboutλ', 'AboutFinancialλ'),
    ('nb.AboutRatiosλ', 'Aboutλ', 'AboutRatiosλ'),
    ('nb.AboutUtilitiesλ', 'Aboutλ', 'AboutUtilitiesλ'),
]
# Same correction inside the Excel Labs module source.
AFE_BARE_FIXES = {
    'Dates': [
        ('Aboutλ', 'AboutDatesλ'),
    ],
    'Debt': [
        ('Amortiseλ', 'AmortiseBλ'),
    ],
    'Essentials': [
        ('Aboutλ', 'AboutEssentialsλ'),
        ('IsBetweenλ', 'IsBetweenEλ'),
        ('RangeToDAλ', 'RangeToDAEλ'),
    ],
    'Financial': [
        ('Aboutλ', 'AboutFinancialλ'),
    ],
    'Ratios': [
        ('Aboutλ', 'AboutRatiosλ'),
    ],
    'Utilities': [
        ('Aboutλ', 'AboutUtilitiesλ'),
        ('AvgColsλ', 'AvgColsUλ'),
        ('AvgRowsλ', 'AvgRowsUλ'),
        ('CountAColsλ', 'CountAColsUλ'),
        ('CountARowsλ', 'CountARowsUλ'),
        ('CountColsλ', 'CountColsUλ'),
        ('CountCλ', 'CountCUλ'),
        ('CountRowsλ', 'CountRowsUλ'),
        ('IsBetweenλ', 'IsBetweenUλ'),
        ('IsInListλ', 'IsInListUλ'),
        ('MaxColsλ', 'MaxColsUλ'),
        ('MaxRowsλ', 'MaxRowsUλ'),
        ('MinColsλ', 'MinColsUλ'),
        ('MinRowsλ', 'MinRowsUλ'),
        ('RangeToDAλ', 'RangeToDAUλ'),
        ('SumColsλ', 'SumColsUλ'),
        ('SumRowsλ', 'SumRowsUλ'),
    ],
}

def fix_self_names(s):
    for full, ob, nb_ in SELF_NAME_FIXES:
        m = re.search(r'(<definedName name="%s"[^>]*>)(.*?)(</definedName>)' % re.escape(full), s, re.S)
        if not m:
            continue
        body = re.sub(r"(?<![A-Za-z])" + re.escape(ob), nb_, m.group(2))
        s = s[:m.start()] + m.group(1) + body + m.group(3) + s[m.end():]
    return s
def flat_text(s):
    return capitalise_brand(flatten_names(s))

# Which module each function came from, captured before the prefixes disappear.
# The exporters below need it: one namespace leaves nothing in the name to group by.
FLAT_MODULE_OF = {}
# What each function was called before the prefixes collapsed. functions.csv publishes
# it, because renaming every name in the library breaks every formula written against
# the old ones and a reader needs somewhere to look the replacement up.
# The names in the loop below are the build's own intermediate ones, not the names any
# release carried, and today those happen to coincide. They stop coinciding the moment a
# function is added: a new nabla.f.PayrollTaxλ would flatten to nb.PayrollTaxλ and publish
# a predecessor no release ever shipped, which is worse than publishing nothing. So the
# predecessor is only recorded when the released baseline confirms it existed.
RELEASED = "released-names-v1.2.6.txt"
with open(os.path.join(os.path.dirname(os.path.abspath(__file__)), RELEASED),
          encoding="utf-8") as _fh:
    BASELINE = {ln.strip() for ln in _fh if ln.strip() and not ln.startswith("#")}
assert len(BASELINE) == 130, (RELEASED, len(BASELINE))

FLAT_PREVIOUS_OF = {}
_pre_wb = get("xl/workbook.xml")
_mod_word = dict(AFE_MODULES)
for _m in re.finditer(r'<definedName name="(nabla\.(?:debt|d|e|f|r|u))\.([^"]+)"', _pre_wb):
    _was = _m.group(1) + "." + _m.group(2)
    _now = flatten_names(_was)
    FLAT_MODULE_OF[_now] = _mod_word[_m.group(1)]
    # two old names collapsing onto one new one would silently lose a function
    assert FLAT_PREVIOUS_OF.setdefault(_now, _was) == _was, (_now, _was, FLAT_PREVIOUS_OF[_now])
    if _was not in BASELINE:
        FLAT_PREVIOUS_OF[_now] = ""          # added since the baseline, so it replaced nothing

# A baseline name nobody claims is a function that has silently disappeared, which is the
# one thing the column exists to prevent. Say which, rather than reporting a count.
_orphans = sorted(BASELINE - set(FLAT_PREVIOUS_OF.values()))
assert not _orphans, ("no function claims these %s names: %s"
                      % (RELEASED, ", ".join(_orphans)))
print("rename map: %d functions, %d of them keeping their bare name, %d new since %s"
      % (len(FLAT_PREVIOUS_OF),
         sum(1 for _n, _o in FLAT_PREVIOUS_OF.items()
             if _o and _n.split(".", 1)[1] == _o.rsplit(".", 1)[1]),
         sum(1 for _o in FLAT_PREVIOUS_OF.values() if not _o), RELEASED))

# AFE store: rename the module containers, then flatten the source inside them
afe_flat = get("customXml/item1.xml")
m_flat = re.search(r'>([A-Za-z0-9+/=]{100,})<', afe_flat)
assert m_flat
j_flat = base64.b64decode(m_flat.group(1)).decode("utf-16-le")
obj_flat = json.loads(j_flat)
# rename the containers first: flattening them all to "nb" would collide
for f in obj_flat["files"]:
    for a, b in AFE_MODULES:
        if f["path"].endswith("/" + a) or f["path"] == a:
            f["path"] = f["path"][: -len(a)] + b
            break
for f in obj_flat["files"]:
    for ob, nb_ in AFE_BARE_FIXES.get(f["path"].rsplit("/", 1)[-1], []):
        f["text"] = re.sub(r"(?<![A-Za-z])" + re.escape(ob), nb_, f["text"])
# then flatten the whole store, which also covers the projectNames function index
j2_flat = flat_text(json.dumps(obj_flat, ensure_ascii=False, separators=(",", ":")))
assert "nabla." not in j2_flat
afe_flat = afe_flat.replace(m_flat.group(1), base64.b64encode(j2_flat.encode("utf-16-le")).decode("ascii"))
put("customXml/item1.xml", afe_flat)

flat_parts = 0
for n in list(parts):
    if n == "customXml/item1.xml":
        continue  # handled above
    if n.endswith((".xml", ".rels")):
        s = flat_text(get(n))
        if n == "xl/workbook.xml":
            s = fix_self_names(fix_about_entries(s))
        put(n, s)
        flat_parts += 1
    elif re.match(r"xl/customProperty\d+\.bin$", n):
        parts[n] = flat_text(parts[n].decode("utf-16-le")).encode("utf-16-le")
        flat_parts += 1
for n in parts:
    if n.endswith((".xml", ".rels")):
        assert "nabla." not in get(n), n
print("flattened to the nb. namespace across", flat_parts, "parts")

# ---------- write ----------
os.makedirs(os.path.dirname(DST) or ".", exist_ok=True)
order = [n for n in zin.namelist() if n in parts]
order += [n for n in parts if n not in set(zin.namelist())]   # parts this build adds
with zipfile.ZipFile(DST, "w", zipfile.ZIP_DEFLATED, compresslevel=9) as zout:
    for n in order:
        zout.writestr(n, parts[n])
zin.close()
print("written", DST, os.path.getsize(DST), "bytes,", len(order), "parts")

# ---------- 11. export src/ and functions.csv FROM the built workbook ----------
# Both used to be maintained by hand, which is how the shipped FinancialYearλ and its
# published source drifted apart. Generating them from the artefact removes that class of bug.
import csv, html

repo = os.path.dirname(os.path.abspath(DST))
src_dir = os.path.join(repo, "src")
os.makedirs(src_dir, exist_ok=True)

store = json.loads(base64.b64decode(
    re.search(r">([A-Za-z0-9+/=]{100,})<",
              parts["customXml/item1.xml"].decode("utf-8")).group(1)).decode("utf-16-le"))
exported = []
for f in store["files"]:
    mod = f["path"].rsplit("/", 1)[-1]
    if mod == "Workbook":
        continue                      # a stub, not a module
    with open(os.path.join(src_dir, mod + ".txt"), "w", encoding="utf-8", newline="\n") as fh:
        fh.write(f["text"])
    exported.append(mod)

wbx_final = parts["xl/workbook.xml"].decode("utf-8")
NAME_RE = re.compile(r'<definedName name="(nb\.[^"]+)"([^>]*)>(.*?)</definedName>', re.S)
entries = [(n, a, html.unescape(b)) for n, a, b in NAME_RE.findall(wbx_final)]

# The Debt functions are recursive, so Excel Labs cannot hold them; export from the names.
# Two tokens in the stored form are NOT valid to type back in, and both have to be
# translated or the exported module cannot be imported at all:
#   _xlop.Name   marks an OPTIONAL parameter. Stripping the prefix leaves a required
#                one, and the body's ISOMITTED() calls then make Excel reject the
#                whole definition. The typed form is [Name].
#   [0]!         is the internal token for "a name in this workbook". The typed form
#                is a bare reference.
OPTIONAL = re.compile(r"_xlop\.([A-Za-z_][A-Za-z0-9_]*\??)")

def as_typed(body):
    body = OPTIONAL.sub(r"[\1]", body)
    return DEPREFIX.sub("", body).replace("[0]!", "").lstrip("=")

debt = [(n, b) for n, _, b in entries if FLAT_MODULE_OF.get(n) == "Debt"]
with open(os.path.join(src_dir, "Debt.txt"), "w", encoding="utf-8", newline="\n") as fh:
    fh.write("//  Debt module - debt sculpting and amortisation functions\n"
             "//  Exported from the workbook's defined names. These are recursive and are not part\n"
             "//  of the Advanced Formula Environment project store, so import them by pasting each\n"
             "//  definition into Name Manager rather than through the AFE module importer.\n\n")
    for n, body in sorted(debt):
        typed = as_typed(body)
        assert "_xl" not in typed and "[0]!" not in typed, n
        # trailing ; so each block can be pasted straight into Name Manager
        fh.write("%s =\n%s;\n\n" % (n.rsplit(".", 1)[1], typed))
exported.append("Debt")
print("exported", len(exported), "module sources to", src_dir)

SIG_RE = re.compile(r"FUNCTION:\s*→?\s*(.*?)¶")   # arrow optional: some upstream help omits it
DESC_RE = re.compile(r"DESCRIPTION:\s*→(.*?)¶")
ROW_RE = re.compile(r"→(.*?)¶")           # the next help row, label or not
# every shipped name came from somewhere, so a missing entry is a build error rather
# than an empty cell: index directly instead of .get()
assert len(entries) == len(FLAT_PREVIOUS_OF), (len(entries), len(FLAT_PREVIOUS_OF))
with open(os.path.join(repo, "functions.csv"), "w", encoding="utf-8", newline="") as fh:
    out = csv.writer(fh, lineterminator="\n")
    out.writerow(["function", "module", "previous_name", "signature", "description"])
    for name, attrs, body in sorted(entries):
        bare = name.rsplit(".", 1)[1]
        sig = SIG_RE.search(body)
        # A long signature wraps onto the next help row, which carries no label. Reading
        # only the first row published nb.Depreciateλ with nine of its parameters missing
        # and a bracket left open. Keep taking rows until the brackets balance.
        text, end = (sig.group(1).strip(), sig.end()) if sig else (bare, 0)
        while text.count("(") > text.count(")"):
            more = ROW_RE.search(body, end)
            if not more:
                break
            text = (text + " " + more.group(1).strip()).strip()
            end = more.end()
        assert text.count("(") == text.count(")"), (name, text)

        desc = re.search(r'comment="([^"]*)"', attrs)
        fallback = DESC_RE.search(body)
        blurb = (html.unescape(desc.group(1)) if desc else
                 (fallback.group(1).strip() if fallback else ""))
        # A Name Manager comment stores a line break as the OOXML escape _x000a_, which is
        # not an XML entity, so unescaping the attribute leaves the escape itself behind.
        # And a description lifted from the help table can still carry its column delimiter.
        blurb = blurb.replace("_x000a_", " ").replace("_x000D_", " ").lstrip("→").strip()
        blurb = re.sub(r"\s{2,}", " ", blurb)
        out.writerow([name, FLAT_MODULE_OF.get(name, ""), FLAT_PREVIOUS_OF[name], text, blurb])
print("exported functions.csv,", len(entries), "functions")
