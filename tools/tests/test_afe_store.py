import base64
import json
import shutil
import subprocess
import sys
import tempfile
import unittest
import zipfile
from pathlib import Path

TOOLS = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(TOOLS))
ROOT = TOOLS.parent
WORKBOOK = ROOT / "ozzit.xlsx"

from verify_afe import find_afe_blob


def afe_parts(workbook: Path) -> tuple[dict, bytes]:
    with zipfile.ZipFile(workbook) as archive:
        raw = archive.read("customXml/item1.xml")
    encoded = find_afe_blob(raw)
    return json.loads(base64.b64decode(encoded).decode("utf-16-le")), encoded


class AfeGateTests(unittest.TestCase):
    def setUp(self):
        self.directory = Path(tempfile.mkdtemp(prefix="ozzit-afe-"))

    def tearDown(self):
        shutil.rmtree(self.directory, ignore_errors=True)

    def _run(self, path):
        return subprocess.run(
            [sys.executable, str(TOOLS / "verify_afe.py"), str(path), str(ROOT / "src")],
            capture_output=True,
            text=True,
            check=False,
        )

    def test_gate_rejects_stale_version_lines(self):
        target = self.directory / "stale-dates.xlsx"
        store, encoded = afe_parts(WORKBOOK)
        for item in store["files"]:
            if item.get("path", "").startswith("/projects/"):
                item["text"] = item["text"].replace("20 Aug 2026", "19 Aug 2026", 1)
        replacement = base64.b64encode(
            json.dumps(store, ensure_ascii=False, separators=(",", ":")).encode("utf-16-le")
        )
        with zipfile.ZipFile(WORKBOOK) as source, zipfile.ZipFile(
            target, "w", zipfile.ZIP_DEFLATED
        ) as output:
            for info in source.infolist():
                data = source.read(info.filename)
                if info.filename == "customXml/item1.xml":
                    data = data.replace(encoded, replacement)
                output.writestr(info, data)
        result = self._run(target)
        self.assertNotEqual(result.returncode, 0)
        self.assertIn("19 Aug 2026", result.stdout)
        self.assertIn("20 Aug 2026", result.stdout)

    def test_gate_rejects_recursive_debt_if_someone_adds_it(self):
        target = self.directory / "debt-added.xlsx"
        store, encoded = afe_parts(WORKBOOK)
        store["files"].append(
            {"path": "/projects/Debt", "text": (ROOT / "src" / "Debt.txt").read_text()}
        )
        store["projectNames"].extend(
            [
                "oz.AmortiseBλ",
                "oz.DebtSculptFixedλ",
                "oz.DebtSculptVariableλ",
                "oz.DebtSculptVariableLRVλ",
                "oz.InterestLRVλ",
            ]
        )
        replacement = base64.b64encode(
            json.dumps(store, ensure_ascii=False, separators=(",", ":")).encode("utf-16-le")
        )
        with zipfile.ZipFile(WORKBOOK) as source, zipfile.ZipFile(
            target, "w", zipfile.ZIP_DEFLATED
        ) as output:
            for info in source.infolist():
                data = source.read(info.filename)
                if info.filename == "customXml/item1.xml":
                    data = data.replace(encoded, replacement)
                output.writestr(info, data)
        result = self._run(target)
        self.assertNotEqual(result.returncode, 0)
        self.assertIn("recursive Debt", result.stdout)

    def test_gate_rejects_missing_non_debt_name(self):
        target = self.directory / "missing-name.xlsx"
        store, encoded = afe_parts(WORKBOOK)
        store["projectNames"].remove("oz.CountDOWλ")
        replacement = base64.b64encode(
            json.dumps(store, ensure_ascii=False, separators=(",", ":")).encode("utf-16-le")
        )
        with zipfile.ZipFile(WORKBOOK) as source, zipfile.ZipFile(
            target, "w", zipfile.ZIP_DEFLATED
        ) as output:
            for info in source.infolist():
                data = source.read(info.filename)
                if info.filename == "customXml/item1.xml":
                    data = data.replace(encoded, replacement)
                output.writestr(info, data)
        result = self._run(target)
        self.assertNotEqual(result.returncode, 0)
        self.assertIn("oz.CountDOWλ", result.stdout)


if __name__ == "__main__":
    unittest.main()
