import shutil
import subprocess
import sys
import tempfile
import unittest
import xml.etree.ElementTree as ET
import zipfile
from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]
WORKBOOK = ROOT / "ozzit.xlsx"
NS = {"m": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}

EXPECTED_STRAPLINES = {
    "oz.CountCλ": "Count the number of times one or more characters appear in a string",
    "oz.CountRowsλ": "Count the number of numbers in each row of an array",
    "oz.CountColsλ": "Count the number of numbers in each column of an array",
    "oz.CountARowsλ": "Count all non-empty cells in each row of an array",
    "oz.CountAColsλ": "Count all non-empty cells in each column of an array",
    "oz.IsBetweenEλ": "Determine if a value is between a lower and upper limit",
    "oz.RangeToDAEλ": "Convert a static range into a dynamic array",
    "oz.FinancialRatios": "Three dozen financial ratios",
}


def workbook_parts():
    with zipfile.ZipFile(WORKBOOK) as archive:
        return {name: archive.read(name) for name in archive.namelist()}


def sheet_map(parts):
    workbook = ET.fromstring(parts["xl/workbook.xml"])
    rels = ET.fromstring(parts["xl/_rels/workbook.xml.rels"])
    targets = {
        rel.attrib["Id"]: rel.attrib["Target"]
        for rel in rels
    }
    rid = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"
    return {
        sheet.attrib["name"]: "xl/" + targets[sheet.attrib[rid]].lstrip("/")
        for sheet in workbook.find("m:sheets", NS)
    }


def shared_strings(parts):
    root = ET.fromstring(parts["xl/sharedStrings.xml"])
    return ["".join(node.itertext()) for node in root.findall("m:si", NS)]


def cell_text(root, address, strings):
    cell = root.find(f".//m:c[@r='{address}']", NS)
    if cell is None:
        return ""
    value = cell.find("m:v", NS)
    if value is None:
        return ""
    return strings[int(value.text)] if cell.attrib.get("t") == "s" else value.text


class WorkbookPresentationTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.parts = workbook_parts()
        cls.sheets = sheet_map(cls.parts)
        cls.strings = shared_strings(cls.parts)
        cls.styles = ET.fromstring(cls.parts["xl/styles.xml"])
        cls.cell_xfs = cls.styles.find("m:cellXfs", NS)
        cls.fills = cls.styles.find("m:fills", NS)
        cls.fonts = cls.styles.find("m:fonts", NS)

    def test_no_legacy_mint_fill_is_used(self):
        mint_fill_ids = {
            index
            for index, fill in enumerate(self.fills)
            if any(colour.attrib.get("indexed") == "42" for colour in fill.iter())
        }
        used = []
        for name, path in self.sheets.items():
            root = ET.fromstring(self.parts[path])
            for cell in root.findall(".//m:c[@s]", NS):
                style = self.cell_xfs[int(cell.attrib["s"])]
                if int(style.attrib.get("fillId", "0")) in mint_fill_ids:
                    used.append(f"{name}!{cell.attrib['r']}")
        self.assertEqual(used, [])

    def test_populated_dates_do_not_use_locale_dependent_builtin_14(self):
        bad = []
        for name, path in self.sheets.items():
            root = ET.fromstring(self.parts[path])
            for cell in root.findall(".//m:c[@s]", NS):
                value = cell.find("m:v", NS)
                if value is None or value.text in (None, ""):
                    continue
                if cell.attrib.get("t") in {"s", "str", "inlineStr"}:
                    continue
                try:
                    serial = float(value.text)
                except ValueError:
                    continue
                style = self.cell_xfs[int(cell.attrib["s"])]
                if serial > 20_000 and style.attrib.get("numFmtId", "0") == "14":
                    bad.append(f"{name}!{cell.attrib['r']}")
        self.assertEqual(bad, [])

    def test_essentials_and_ratio_straplines_match_toc_descriptions(self):
        for sheet, expected in EXPECTED_STRAPLINES.items():
            root = ET.fromstring(self.parts[self.sheets[sheet]])
            self.assertEqual(cell_text(root, "A2", self.strings), expected, sheet)

    def test_function_strapline_font_meets_house_neutral(self):
        bad = []
        for name, path in self.sheets.items():
            if not name.startswith("oz."):
                continue
            root = ET.fromstring(self.parts[path])
            cell = root.find(".//m:c[@r='A2'][@s]", NS)
            if cell is None or not cell_text(root, "A2", self.strings):
                continue
            style = self.cell_xfs[int(cell.attrib["s"])]
            font = self.fonts[int(style.attrib.get("fontId", "0"))]
            rgb = next(
                (node.attrib.get("rgb") for node in font.iter() if node.tag.endswith("color")),
                None,
            )
            if rgb not in {"FF6E6862", "FF04001F", "FF2B2733", "FF5C2D91"}:
                bad.append((name, rgb))
        self.assertEqual(bad, [])

    def test_cover_describes_brand_neutral_formula_highlights(self):
        cover = ET.fromstring(self.parts[self.sheets["Cover"]])
        text = " ".join(
            cell_text(cover, address, self.strings)
            for address in ("A25", "B25", "A26", "B26", "A27", "B27")
        )
        self.assertNotRegex(text.lower(), r"\bgreen\s+shaded\b")
        self.assertIn("purple", text.lower())

    def test_polish_tool_is_idempotent(self):
        directory = Path(tempfile.mkdtemp(prefix="ozzit-polish-"))
        try:
            target = directory / "ozzit.xlsx"
            shutil.copy2(WORKBOOK, target)
            command = [sys.executable, str(ROOT / "tools" / "polish_workbook.py"), str(target)]
            first = subprocess.run(command, capture_output=True, text=True, check=False)
            self.assertEqual(first.returncode, 0, first.stdout + first.stderr)
            before = target.read_bytes()
            second = subprocess.run(command, capture_output=True, text=True, check=False)
            self.assertEqual(second.returncode, 0, second.stdout + second.stderr)
            self.assertEqual(target.read_bytes(), before)
        finally:
            shutil.rmtree(directory, ignore_errors=True)


if __name__ == "__main__":
    unittest.main()
