import csv
import shutil
import subprocess
import sys
import tempfile
import unittest
import zipfile
from pathlib import Path

TOOLS = Path(__file__).resolve().parents[1]
ROOT = TOOLS.parent
WORKBOOK = ROOT / "ozzit.xlsx"
INDEX = ROOT / "functions.csv"


class IndexAndInventoryTests(unittest.TestCase):
    def setUp(self):
        self.directory = Path(tempfile.mkdtemp(prefix="ozzit-index-"))

    def tearDown(self):
        shutil.rmtree(self.directory, ignore_errors=True)

    def _run(self, tool, *args):
        return subprocess.run(
            [sys.executable, str(TOOLS / tool), *map(str, args)],
            capture_output=True,
            text=True,
            check=False,
        )

    def _mutated_index(self, field, value):
        target = self.directory / "functions.csv"
        with INDEX.open(encoding="utf-8-sig", newline="") as source:
            rows = list(csv.DictReader(source))
            fields = list(rows[0])
        rows[0][field] = value
        with target.open("w", encoding="utf-8", newline="") as output:
            writer = csv.DictWriter(output, fieldnames=fields, lineterminator="\n")
            writer.writeheader()
            writer.writerows(rows)
        return target

    def test_index_gate_rejects_fabricated_description(self):
        index = self._mutated_index("description", "Fabricated description")
        result = self._run("verify_index.py", WORKBOOK, ROOT / "src", index)
        self.assertNotEqual(result.returncode, 0)
        self.assertIn("description differs from workbook", result.stdout)

    def test_index_gate_rejects_fabricated_module(self):
        index = self._mutated_index("module", "FabricatedModule")
        result = self._run("verify_index.py", WORKBOOK, ROOT / "src", index)
        self.assertNotEqual(result.returncode, 0)
        self.assertIn("module differs from src/", result.stdout)

    def assert_forbidden_part_rejected(self, part):
        target = self.directory / "forbidden.xlsx"
        with zipfile.ZipFile(WORKBOOK) as source, zipfile.ZipFile(
            target, "w", zipfile.ZIP_DEFLATED
        ) as output:
            for info in source.infolist():
                output.writestr(info, source.read(info.filename))
            output.writestr(part, b"forbidden executable or linked content")
        result = self._run("verify_workbook.py", target)
        self.assertNotEqual(result.returncode, 0)
        self.assertIn(f"forbidden workbook part {part}", result.stdout)

    def test_workbook_gate_rejects_executable_and_external_parts(self):
        for part in (
            "xl/vbaProject.bin",
            "xl/activeX/activeX1.xml",
            "xl/ctrlProps/ctrlProp1.xml",
            "xl/externalLinks/externalLink1.xml",
            "xl/embeddings/oleObject1.bin",
            "xl/macrosheets/sheet1.xml",
            "xl/dialogsheets/sheet1.xml",
        ):
            with self.subTest(part=part):
                self.assert_forbidden_part_rejected(part)


if __name__ == "__main__":
    unittest.main()
