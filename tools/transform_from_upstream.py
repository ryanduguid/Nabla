# Transform Craig Hatmaker's 5g Financial Starter Pack (2024-07-06.xlsx) into nabla.xlsx
# Pure zip/XML surgery. Never resaves via openpyxl (preserves cached values, extensions, rich parts).
import zipfile, re, shutil, io, sys, datetime
import base64, json

SRC = r"C:\Users\-\Downloads\2024-07-06.xlsx"
DST = r"C:\Users\-\AppData\Local\Temp\claude\C--\fea021f8-8627-4730-bcc8-53de478d7f07\scratchpad\nabla-build\nabla.xlsx"
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
AMORT = [("Amoritization", "Amortisation"), ("amoritization", "amortisation"),
         ("Amoritize", "Amortise"), ("amoritize", "amortise"),
         ("Amortization", "Amortisation"), ("amortization", "amortisation"),
         ("Amortize", "Amortise"), ("amortize", "amortise"),
         ("Amortizing", "Amortising"), ("amortizing", "amortising"),
         ("Occurence", "Occurrence"), ("occurence", "occurrence")]
# Americanisms -> Australian equivalents (content, not just spelling)
AMERICAN = [
    ("MACRS=Modified Accelerated Cost Recovery System. NOTE: Salvage value ignored",
     "MACRS=Modified Accelerated Cost Recovery System (US legacy; for ATO diminishing value use DDB with a factor of 2). Salvage value ignored"),
    ("IRS Depreciation", "US IRS depreciation (legacy)"),
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
    s = re.sub(r'(?i)(ersion:\s*→\s*)[A-Z][a-z]{2} \d{1,2} \d{4}', r'\g<1>' + TODAY_AU, s)
    # sample dates: quoted or cell-value m/d/yyyy -> +2y, AU day-first
    # (trailing char may be a backslash inside JSON-escaped AFE text)
    s = re.sub(r'(?<=["">])(\d{1,2})/(\d{1,2})/(20\d\d)(?=["<\\])', _mdy_to_au, s)
    s = re.sub(r'\b(20\d\d)-(\d\d)-(\d\d)\b', _iso_shift, s)
    s = re.sub(r'\{[\d;\s]+\}', _arr_shift, s)
    return s

# New AU function: nabla.f.DiminishingValueλ (ATO 200% diminishing value)
DV_HELP_LINES = [
    'FUNCTION:      →DiminishingValueλ(Cost, Life)¶',
    'DESCRIPTION:   →Diminishing value depreciation (ATO 200% method) for one asset or asset class.¶',
    'WEBPAGE:       →%s¶' % REPO_URL,
    'VERSION:       →%s¶' % TODAY_AU,
    'PARAMETERS:    →¶',
    "Cost           →(Required) Asset's initial cost.¶",
    "Life           →(Required) Asset's effective life in years.¶",
    'EXAMPLES:      →¶',
    "→Formula (nabla.f is assumed to be the module's name)¶",
    '→=nabla.f.DiminishingValueλ(1000, 5)¶',
    '→Result¶',
    '→400.00,240.00,144.00,86.40,51.84',
]
DV_XML = (
    '_xlfn.LAMBDA(_xlop.Cost,_xlop.Life, _xlfn.LET(_xlpm.Help, TRIM(_xlfn.TEXTSPLIT('
    + " &amp; ".join('"%s"' % l for l in DV_HELP_LINES)
    + ', "→", "¶")), _xlpm.Help?, OR(_xlfn.ISOMITTED(_xlpm.Cost), _xlfn.ISOMITTED(_xlpm.Life)), '
    '_xlpm.Rate, 2/_xlpm.Life, '
    '_xlpm.Result, _xlpm.Cost * _xlpm.Rate * (1-_xlpm.Rate)^(_xlfn.SEQUENCE(, _xlpm.Life)-1), '
    'CHOOSE(_xlpm.Help? + 1, _xlpm.Result, _xlpm.Help)))'
)
DV_AFE = (
    "/*  FUNCTION NAME:  DiminishingValueλ\n"
    "    DESCRIPTION:*//**Diminishing value depreciation (ATO 200% declining balance) for one asset or asset class*/\n"
    "/*  REVISIONS:      Date        Developer       Description  \n"
    "                    18 Aug 2026 nabla           Added Australian diminishing value method\n"
    "*/\n\n"
    "DiminishingValueλ = LAMBDA(\n"
    "//  Parameter Declaration\n"
    "    [Cost],\n"
    "    [Life], \n"
    "    LET(\n"
    "    //  Help\n"
    "        Help,           TRIM(TEXTSPLIT(\n"
    + "".join('                            "%s"%s\n' % (l, ' &' if i < len(DV_HELP_LINES) - 1 else ',')
              for i, l in enumerate(DV_HELP_LINES))
    + '                            "→", "¶"\n'
    "                        )),\n"
    "    //  Check inputs - Omitted required arguments\n"
    "        Help?,          OR( ISOMITTED( Cost), ISOMITTED( Life)),\n"
    "    //  Set Constants\n"
    "        Rate,           2 / Life,\n"
    "    //  Procedure\n"
    "        Result,         Cost * Rate * (1 - Rate)^(SEQUENCE( , Life) - 1), \n"
    "    //  Return Result or Help\n"
    "        CHOOSE( Help? + 1, Result, Help)\n"
    "    )\n"
    ");\n"
)

# AFE (Excel Labs) project store: base64-wrapped UTF-16 JSON holding LAMBDA source
afe = get("customXml/item1.xml")
m = re.search(r'>([A-Za-z0-9+/=]{100,})<', afe)
assert m
j = base64.b64decode(m.group(1)).decode("utf-16-le")
j2 = transform_text(j)
json.loads(j2)  # must stay valid JSON
# insert DV source into the nabla.f module, after the MACRSλ block
esc = json.dumps(DV_AFE, ensure_ascii=False)[1:-1]
i_mac = j2.find("MACRSλ = LAMBDA")
assert i_mac > 0
anchor = ");\\n\\n\\n\\n\\n//  Diagnostic Routines"
i_end = j2.find(anchor, i_mac)
assert i_end > 0
j2 = j2[:i_end + 4] + "\\n\\n" + esc + j2[i_end + 4:]
# register in AFE projectNames
assert '"nabla.f.MACRSλ",' in j2
j2 = j2.replace('"nabla.f.MACRSλ",', '"nabla.f.MACRSλ","nabla.f.DiminishingValueλ",', 1)
# fix upstream copy-paste bug: the u module's About suggested "nabla.e" (was BXE) as its own name
obj_afe = json.loads(j2)
for f in obj_afe["files"]:
    if f["path"] == "/projects/nabla.u":
        assert "Suggested module name: nabla.e" in f["text"]
        f["text"] = f["text"].replace("Suggested module name: nabla.e", "Suggested module name: nabla.u")
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
            if sid in date_styles and 20000 <= float(mo.group("v")) <= 60000:
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

# ---------- 8a5. sheet32: repair inherited #REF! Timeline argument ----------
s32 = get("xl/worksheets/sheet32.xml")
s32b = s32.replace('nabla.d.Timelineλ( E23, D23, "Y",#REF!)', 'nabla.d.Timelineλ( E23, D23, "Y")')
assert s32b != s32
put("xl/worksheets/sheet32.xml", s32b)

# ---------- 8a6. drawing18: label the US GAAP note as legacy ----------
d18 = get("xl/drawings/drawing18.xml")
d18b = d18.replace(
    "US GAAP allows companies",
    "US GAAP (a US-legacy convention; for the ATO method see nabla.f.DiminishingValueλ) allows companies")
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
]:
    d = get(path)
    assert d.count(old) == cnt_want, (path, d.count(old))
    put(path, d.replace(old, new))

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
wb2 = re.sub(r'<definedName name="Slicer_Type"[^>]*>[^<]*</definedName>', "", wb)
assert wb2 != wb
wb = wb2
m = re.search(r'<calcPr[^>]*/>', wb)
assert m, "calcPr not found"
if "fullCalcOnLoad" not in m.group(0):
    wb = wb.replace(m.group(0), m.group(0)[:-2] + ' fullCalcOnLoad="1"/>')
put("xl/workbook.xml", wb)

# ---------- 9b. Define nabla.e.Aboutλ (original defect: called on its sheet, never defined; source in AFE) ----------
etext = next(f["text"] for f in json.loads(j2)["files"] if f["path"] == "/projects/nabla.e")
i = etext.index("Aboutλ = ")
expr = etext[i + len("Aboutλ = "):]
depth = 0; in_str = False; end = None
for k, ch in enumerate(expr):
    if ch == '"': in_str = not in_str
    elif not in_str and ch == '(': depth += 1
    elif not in_str and ch == ')':
        depth -= 1
        if depth == 0: end = k + 1; break
assert end, "Aboutλ body not parsed"
body = " ".join(expr[:end].split())
assert body.startswith("TRIM(TEXTSPLIT(") and body.count('"') % 2 == 0
body = body.replace("TEXTSPLIT(", "_xlfn.TEXTSPLIT(")
body_xml = body.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
wb = get("xl/workbook.xml")
assert '<definedName name="nabla.e.Aboutλ"' not in wb
wb = wb.replace("</definedNames>",
                '<definedName name="nabla.e.Aboutλ" comment="Displays this module\'s repository URL and function list">%s</definedName></definedNames>' % body_xml)
# same upstream copy-paste bug in the installed u-module About
wb = re.sub(r'(<definedName name="nabla\.u\.Aboutλ"[^>]*>[^<]*?)Suggested module name: nabla\.e',
            r'\g<1>Suggested module name: nabla.u', wb)
put("xl/workbook.xml", wb)
print("nabla.e.Aboutλ defined,", len(body), "chars")

# ---------- 9c. Register nabla.f.DiminishingValueλ defined name ----------
wb = get("xl/workbook.xml")
assert '<definedName name="nabla.f.DiminishingValueλ"' not in wb
wb = wb.replace("</definedNames>",
                '<definedName name="nabla.f.DiminishingValueλ" comment="Diminishing value depreciation (ATO 200%% method) for one asset or asset class.">%s</definedName></definedNames>' % DV_XML)
put("xl/workbook.xml", wb)
print("nabla.f.DiminishingValueλ defined")

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
with zipfile.ZipFile(DST, "w", zipfile.ZIP_DEFLATED) as zout:
    for n in order:
        zout.writestr(n, parts[n])
zin.close()
print("written", DST, os.path.getsize(DST), "bytes,", len(order), "parts")
