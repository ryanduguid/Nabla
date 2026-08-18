# Build nabla.xlsx from the upstream Financial Starter Pack workbook.
# Pure zip/XML surgery. Never resaves via openpyxl (preserves cached values, extensions, rich parts).
import zipfile, re, shutil, io, sys, datetime
import base64, json

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
     "This workbook contains six nabla modules: dates (nabla.d), array essentials (nabla.e), "
     "financial functions (nabla.f), financial ratios (nabla.r), utilities (nabla.u) and debt (nabla.debt). "),
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
    # the MACRS row on the Data Validation sheet becomes the ATO diminishing value row
    ("IRS Depreciation", "Diminishing value (ATO)"),
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
        "desc": "Diminishing value depreciation (ATO 200% method), writing the residual off in the final period.",
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
        "desc": "Prime cost (straight line) depreciation, ATO method, for one asset or asset class.",
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
# Australian-only. Depreciation method slot 6 becomes the ATO diminishing value method and a
# seventh slot adds ATO prime cost.
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
     '"→DV =Diminishing value (ATO 200% method). Salvage value ignored¶" & \n'
     '                                           "→PC =Prime cost (ATO straight line). Salvage value ignored¶" & '),
]
AFE_MACRS_RE = [
    (r'DisposalDate,   IF\( Method = "MACRS", \s*\n\s*MAX\(EDATE\( InserviceDate, Years \* MpY\), '
     r'@INDEX\( DisposalDates, Asset\)\),\s*\n\s*@INDEX\( DisposalDates, Asset\)\), ',
     'DisposalDate,   @INDEX( DisposalDates, Asset), '),
    (r'//  6\. Modified accelerated cost recovery system \s*\n\s*MACRSλ\( InitialValue, Years - 1\),',
     '//  6. Diminishing value (ATO 200% method)\n'
     + ' ' * 56 + 'DiminishingValueλ( InitialValue, Years),\n'
     + ' ' * 52 + '//  7. Prime cost (ATO straight line)\n'
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
     '"→DV =Diminishing value (ATO 200% method). Salvage value ignored¶" &amp; '
     '"→PC =Prime cost (ATO straight line). Salvage value ignored¶" &amp; '),
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
                       "DV =Diminishing value (ATO 200% method). Salvage value ignored")
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
new_si = ["PC", "Prime cost (ATO)"]
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
    8:  txt("A8", 0, "Diminishing value (ATO 200% method)")
        + fml("B8", MONEY, "nabla.f.DiminishingValueλ($A$6,$B$6)", spill="B8:F8"),
    9:  txt("A9", 0, "Prime cost (ATO straight line)")
        + fml("B9", MONEY, "nabla.f.PrimeCostλ($A$6,$B$6)", spill="B9:F9"),
    # Both schedules must write off the whole cost, whatever the effective life. Shown on
    # the sheet because it is the property that a part-year or sub-two-year life breaks.
    10: txt("A10", 0, "Total written off, diminishing value")
        + fml("B10", MONEY, "SUM(_xlfn.ANCHORARRAY(B8))"),
    11: txt("A11", 0, "Total written off, prime cost")
        + fml("B11", MONEY, "SUM(_xlfn.ANCHORARRAY(B9))"),
    12: txt("A12", 0, "Both totals equal the cost above for any effective life, "
                      "including part years such as 6 2/3 and lives under 2 years."),
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
             "ATO depreciation, GST and financial-year helpers") + '</row>')
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

# ---------- write ----------
import os
os.makedirs(os.path.dirname(DST), exist_ok=True)
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
    if not mod.startswith("nabla."):
        continue                      # /projects/Workbook is a stub, not a module
    with open(os.path.join(src_dir, mod + ".txt"), "w", encoding="utf-8", newline="\n") as fh:
        fh.write(f["text"])
    exported.append(mod)

wbx_final = parts["xl/workbook.xml"].decode("utf-8")
NAME_RE = re.compile(r'<definedName name="(nabla\.[^"]+)"([^>]*)>(.*?)</definedName>', re.S)
entries = [(n, a, html.unescape(b)) for n, a, b in NAME_RE.findall(wbx_final)]

# nabla.debt is recursive, so Excel Labs cannot hold it; export it from the names instead.
# Two tokens in the stored form are NOT valid to type back in, and both have to be
# translated or the exported module cannot be imported at all:
#   _xlop.Name   marks an OPTIONAL parameter. Stripping the prefix leaves a required
#                one, and the body's ISOMITTED() calls then make Excel reject the
#                whole definition. The typed form is [Name].
#   [0]!         is the internal token for "a name in this workbook". The typed form
#                is a bare reference.
OPTIONAL = re.compile(r"_xlop\.([A-Za-z_][A-Za-z0-9_]*)")

def as_typed(body):
    body = OPTIONAL.sub(r"[\1]", body)
    return DEPREFIX.sub("", body).replace("[0]!", "").lstrip("=")

debt = [(n, b) for n, _, b in entries if n.startswith("nabla.debt.")]
with open(os.path.join(src_dir, "nabla.debt.txt"), "w", encoding="utf-8", newline="\n") as fh:
    fh.write("//  nabla.debt module - debt sculpting and amortisation functions\n"
             "//  Exported from the workbook's defined names. These are recursive and are not part\n"
             "//  of the Advanced Formula Environment project store, so import them by pasting each\n"
             "//  definition into Name Manager rather than through the AFE module importer.\n\n")
    for n, body in sorted(debt):
        typed = as_typed(body)
        assert "_xl" not in typed and "[0]!" not in typed, n
        # trailing ; so each block can be pasted straight into Name Manager
        fh.write("%s =\n%s;\n\n" % (n.rsplit(".", 1)[1], typed))
exported.append("nabla.debt")
print("exported", len(exported), "module sources to", src_dir)

SIG_RE = re.compile(r"FUNCTION:\s*→?\s*(.*?)¶")   # arrow optional: some upstream help omits it
DESC_RE = re.compile(r"DESCRIPTION:\s*→(.*?)¶")
with open(os.path.join(repo, "functions.csv"), "w", encoding="utf-8", newline="") as fh:
    out = csv.writer(fh, lineterminator="\n")
    out.writerow(["function", "module", "signature", "description"])
    for name, attrs, body in sorted(entries):
        bare = name.rsplit(".", 1)[1]
        sig = SIG_RE.search(body)
        desc = re.search(r'comment="([^"]*)"', attrs)
        fallback = DESC_RE.search(body)
        out.writerow([name, name.rsplit(".", 1)[0],
                      sig.group(1).strip() if sig else bare,
                      html.unescape(desc.group(1)) if desc else
                      (fallback.group(1).strip() if fallback else "")])
print("exported functions.csv,", len(entries), "functions")
