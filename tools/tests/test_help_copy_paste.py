"""Guards on help text that was copied from a neighbour.

verify_signatures.py already reads FUNCTION lines, parameter tables and
worked examples. These checks cover the leftovers that live outside those
three: Name Manager comments in functions.csv, the About-table prose, and
source comments that never make it into a LAMBDA body.
"""

import csv
import re
import unittest
import zipfile
from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]
SRC = ROOT / "src"
INDEX = ROOT / "functions.csv"
FINANCIAL = SRC / "Financial.txt"


class HelpCopyPasteTests(unittest.TestCase):
    def test_corkscrew_dv_describes_corkscrew_not_depreciate(self):
        text = FINANCIAL.read_text(encoding="utf-8")
        block = re.search(
            r"FUNCTION NAME:\s+CorkscrewλDV\s+DESCRIPTION:\*//\*\*(.*?)\*/",
            text,
            re.S,
        )
        self.assertIsNotNone(block, "CorkscrewλDV description comment missing")
        description = block.group(1)
        self.assertIn("Corkscrewλ", description)
        self.assertNotIn("Depreciateλ", description)

        with INDEX.open(encoding="utf-8-sig", newline="") as source:
            row = next(
                r for r in csv.DictReader(source) if r["function"] == "oz.CorkscrewλDV"
            )
        self.assertIn("Corkscrewλ", row["description"])
        self.assertNotIn("Depreciateλ", row["description"])

    def test_about_financial_does_not_describe_rolling_avg_as_a_maximum(self):
        text = FINANCIAL.read_text(encoding="utf-8")
        match = re.search(r'"RollingAvgλ\s+→([^"]+)¶"', text)
        self.assertIsNotNone(match, "AboutFinancialλ RollingAvg row missing")
        self.assertNotIn("maximum", match.group(1).lower())
        self.assertIn("average", match.group(1).lower())

    def test_column_functions_are_not_indexed_as_row_functions(self):
        with INDEX.open(encoding="utf-8-sig", newline="") as source:
            rows = {r["function"]: r["description"] for r in csv.DictReader(source)}
        for name in (
            "oz.CountColsλ",
            "oz.CountColsUλ",
            "oz.CountAColsλ",
            "oz.CountAColsUλ",
        ):
            description = rows[name].lower()
            self.assertIn("column", description, name)
            self.assertNotIn("each row", description, name)

    def test_source_has_no_coming_soon_placeholder_or_year_2924(self):
        for path in SRC.glob("*.txt"):
            text = path.read_text(encoding="utf-8")
            self.assertNotRegex(text, r"(?i)coming soon", msg=path.name)
            self.assertNotIn("2924", text, path.name)

    def test_workbook_no_longer_ships_the_copied_help(self):
        with zipfile.ZipFile(ROOT / "ozzit.xlsx") as archive:
            book = archive.read("xl/workbook.xml").decode("utf-8")
            sheets = "".join(
                archive.read(name).decode("utf-8")
                for name in archive.namelist()
                if name.startswith("xl/worksheets/")
            )
        self.assertNotIn("Coming Soon", book)
        self.assertNotIn("Coming Soon", sheets)
        self.assertNotIn(
            'name="oz.CorkscrewλDV" comment="Provides extended data validation of arguments for Depreciateλ."',
            book,
        )
        self.assertIn(
            'name="oz.CorkscrewλDV" comment="Provides extended data validation of arguments for Corkscrewλ."',
            book,
        )
        self.assertNotIn(
            'RollingAvgλ        →Finds the maximum value the n preceding values',
            book,
        )
        for name in (
            "oz.CountColsλ",
            "oz.CountColsUλ",
            "oz.CountAColsλ",
            "oz.CountAColsUλ",
        ):
            self.assertNotRegex(
                book,
                rf'name="{re.escape(name)}" comment="[^"]*each row',
                msg=name,
            )


if __name__ == "__main__":
    unittest.main()
