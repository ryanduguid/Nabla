import base64
import json
import os
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

    def _run(self, path, encoding):
        env = os.environ.copy()
        env["PYTHONIOENCODING"] = encoding
        return subprocess.run(
            [sys.executable, str(TOOLS / "verify_afe.py"), str(path), str(ROOT / "src")],
            capture_output=True,
            text=True,
            encoding=encoding,
            errors="strict",
            env=env,
            check=False,
        )

    def _assert_controlled_failure(self, result, expected_count):
        self.assertEqual(result.returncode, 1)
        self.assertTrue(
            result.stdout.startswith(f"FAIL: {expected_count} AFE problem(s)"),
            result.stdout,
        )
        self.assertEqual(result.stderr, "")
        combined = result.stdout + result.stderr
        self.assertNotIn("Traceback", combined)
        self.assertNotIn("UnicodeEncodeError", combined)

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
        modules = ("Dates", "Essentials", "Financial", "Ratios", "Utilities")
        cp1252 = self._run(target, "cp1252")
        self._assert_controlled_failure(cp1252, len(modules))
        for module in modules:
            self.assertIn(f"/projects/{module}", cp1252.stdout)
        self.assertEqual(cp1252.stdout.count("at character "), len(modules))
        self.assertEqual(cp1252.stdout.count("19 Aug 2026"), len(modules))
        self.assertEqual(cp1252.stdout.count("20 Aug 2026"), len(modules))
        cp1252_details = [
            line for line in cp1252.stdout.splitlines() if line.startswith("  - AFE module ")
        ]
        self.assertEqual(len(cp1252_details), len(modules))
        for detail in cp1252_details:
            self.assertIn(r"\u2192", detail)

        utf8 = self._run(target, "utf-8")
        self._assert_controlled_failure(utf8, len(modules))
        for module in modules:
            self.assertIn(f"/projects/{module}", utf8.stdout)
        utf8_details = [
            line for line in utf8.stdout.splitlines() if line.startswith("  - AFE module ")
        ]
        self.assertEqual(len(utf8_details), len(modules))
        for detail in utf8_details:
            self.assertIn("→", detail)
        self.assertNotIn(r"\u2192", utf8.stdout)

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
        debt_names = [
            "oz.AmortiseBλ",
            "oz.DebtSculptFixedλ",
            "oz.DebtSculptVariableλ",
            "oz.DebtSculptVariableLRVλ",
            "oz.InterestLRVλ",
        ]
        cp1252 = self._run(target, "cp1252")
        self._assert_controlled_failure(cp1252, len(debt_names))
        for name in debt_names:
            self.assertIn(name.replace("λ", r"\u03bb"), cp1252.stdout)

        utf8 = self._run(target, "utf-8")
        self._assert_controlled_failure(utf8, len(debt_names))
        for name in debt_names:
            self.assertIn(name, utf8.stdout)
        self.assertNotIn(r"\u03bb", utf8.stdout)

    def test_gate_rejects_missing_non_debt_name(self):
        target = self.directory / "missing-name-λ→.xlsx"
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
        cp1252 = self._run(target, "cp1252")
        self._assert_controlled_failure(cp1252, 1)
        self.assertIn(r"missing-name-\u03bb\u2192.xlsx", cp1252.stdout)
        self.assertIn(r"oz.CountDOW\u03bb", cp1252.stdout)

        utf8 = self._run(target, "utf-8")
        self._assert_controlled_failure(utf8, 1)
        self.assertIn("missing-name-λ→.xlsx", utf8.stdout)
        self.assertIn("oz.CountDOWλ", utf8.stdout)
        self.assertNotIn(r"\u03bb", utf8.stdout)

    def test_gate_reports_success_in_supported_encodings(self):
        target = self.directory / "valid-λ→.xlsx"
        shutil.copyfile(WORKBOOK, target)
        for encoding in ("cp1252", "utf-8"):
            with self.subTest(encoding=encoding):
                result = self._run(target, encoding)
                self.assertEqual(result.returncode, 0)
                self.assertEqual(result.stderr, "")
                self.assertIn("OK: AFE store", result.stdout)
                self.assertIn("matches 5 non-recursive modules", result.stdout)
                self.assertIn("excludes recursive Debt", result.stdout)
                expected_path = (
                    r"valid-\u03bb\u2192.xlsx"
                    if encoding == "cp1252"
                    else "valid-λ→.xlsx"
                )
                self.assertIn(expected_path, result.stdout)


if __name__ == "__main__":
    unittest.main()
