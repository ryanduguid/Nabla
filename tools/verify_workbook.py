import zipfile, re, base64, json
import xml.etree.ElementTree as ET
DST = r"nabla-build\nabla.xlsx"
z = zipfile.ZipFile(DST)
errs = []
for n in z.namelist():
    if n.endswith((".xml", ".rels")):
        try: ET.fromstring(z.read(n).decode("utf-8"))
        except Exception as e: errs.append(f"XML parse {n}: {e}")
BAD = re.compile(r'BX[DEFRLU]|Calibri|beyondexcel|sites\.google|dropbox|Eloquens|Starter Pack|Hatmaker|5g|5G|Leonardo|CHatmaker')
YT = re.compile(r'youtube')
for n in z.namelist():
    if n.endswith((".xml", ".rels")) and n != "customXml/item1.xml":
        d = z.read(n).decode("utf-8")
        for m in list(BAD.finditer(d))[:3]: errs.append(f"brand {n}: ...{d[max(0,m.start()-50):m.start()+50]}...")
        for m in list(YT.finditer(d))[:2]: errs.append(f"youtube {n}")
    elif re.match(r'xl/customProperty\d+\.bin$', n):
        if BAD.search(z.read(n).decode("utf-16-le")): errs.append(f"brand in {n}")
afe = z.read("customXml/item1.xml").decode("utf-8")
j = base64.b64decode(re.search(r'>([A-Za-z0-9+/=]{100,})<', afe).group(1)).decode("utf-16-le")
json.loads(j)
for m in list(BAD.finditer(j))[:5]: errs.append(f"brand AFE: ...{j[max(0,m.start()-50):m.start()+55]}...")
wb = z.read("xl/workbook.xml").decode("utf-8")
names = re.findall(r'<definedName name="([^"]+)"', wb)
nameset = set(names); sheets = re.findall(r'<sheet name="([^"]+)"', wb)
used = set(); err_cells = []
for n in z.namelist():
    if re.match(r'xl/worksheets/sheet\d+\.xml$', n):
        d = z.read(n).decode("utf-8")
        for m in re.finditer(r'<c r="([A-Z]+\d+)"[^>]*t="e"[^>]*>(?:<f[^>]*>[^<]*</f>)?<v>([^<]*)</v>', d):
            err_cells.append((n, m.group(1), m.group(2)))
        for f in re.findall(r'<f[^>]*>([^<]*)</f>', d):
            used.update(re.findall(r'nabla\.[a-z]+\.[A-Za-z0-9_]+λ?(?:DV)?', f))
        for m in re.finditer(r'<f[^>]*>[^<]*"(\d{1,2}/\d{1,2}/\d{4})"[^<]*</f>', d):
            errs.append(f"text-date literal in formula {n}: {m.group(1)}")
sheetset = set(sheets)
really_missing = {u for u in used if u not in nameset and u not in sheetset}
if really_missing: errs.append(f"missing fn tokens: {sorted(really_missing)}")
for nm, body in re.findall(r'<definedName name="([^"]+)"[^>]*>([^<]*)</definedName>', wb):
    for tok in set(re.findall(r'nabla\.[a-z]+\.[A-Za-z0-9_]+λ?(?:DV)?', body)):
        if tok not in nameset and tok not in sheetset: errs.append(f"name {nm} refs missing {tok}")
print("names:", len(names), "| sheets:", len(sheets), "| err cells:", err_cells)
print("nabla.e.Aboutλ defined:", "nabla.e.Aboutλ" in nameset)
from collections import Counter
allx = " ".join(z.read(n).decode("utf-8") for n in z.namelist() if n.endswith(".xml") and n != "customXml/item1.xml")
print("stale Version dates (non-AFE):", Counter(re.findall(r'(?i)ersion:\s*→\s*([A-Za-z]{3} \d{1,2} \d{4})', allx)).most_common(3))
print("stale Version dates (AFE):", Counter(re.findall(r'(?i)ersion:\s*→\s*([A-Za-z]{3} \d{1,2} \d{4})', j)).most_common(3))
import warnings, openpyxl
with warnings.catch_warnings():
    warnings.simplefilter("ignore")
    wb2 = openpyxl.load_workbook(DST)
print("openpyxl OK:", len(wb2.sheetnames), len(wb2.defined_names))
print()
print("ERRORS", len(errs)) if errs else print("ALL CHECKS PASS")
for e in errs[:15]: print(" -", e[:180])
