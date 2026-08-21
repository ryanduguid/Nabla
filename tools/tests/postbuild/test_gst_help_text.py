"""Contract for tools/postbuild/gst_help_text.py.

The pass inserts a NOTES! legislative scope block into oz.GSTAddλ and
oz.GSTExtractλ inline help. On the current workbook it must be a byte no-op;
on a workbook reverted to the pre-note text it must apply each swap exactly
once; AboutFinancialλ and comment= must keep the one-line descriptions.
"""

import importlib.util
import re
import shutil
import subprocess
import sys
import tempfile
import unittest
import zipfile
from pathlib import Path

ROOT = Path(__file__).resolve().parents[3]
TOOLS = ROOT / "tools"
WORKBOOK = ROOT / "ozzit.xlsx"
PASS_SCRIPT = TOOLS / "postbuild" / "gst_help_text.py"

_spec = importlib.util.spec_from_file_location("gst_help_text", PASS_SCRIPT)
_mod = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(_mod)
SWAPS = _mod.SWAPS
NOTES = _mod.NOTES
GSTADD_OLD, GSTADD_NEW = SWAPS[0][0], SWAPS[0][1]


def defined_name(book: str, name: str) -> str:
    match = re.search(
        rf'<definedName name="{re.escape(name)}"[^>]*>(.*?)</definedName>',
        book,
        re.DOTALL,
    )
    if match is None:
        raise AssertionError(f"{name} missing from workbook.xml")
    return match.group(0)


class GstHelpTextTests(unittest.TestCase):
    def setUp(self):
        self.directory = Path(tempfile.mkdtemp(prefix="ozzit-gst-help-"))

    def tearDown(self):
        shutil.rmtree(self.directory, ignore_errors=True)

    def _run(self, workbook, src_dir):
        return subprocess.run(
            [sys.executable, str(PASS_SCRIPT), str(workbook), str(src_dir)],
            capture_output=True,
            text=True,
            check=False,
        )

    def _copy_tree(self):
        workbook = self.directory / "ozzit.xlsx"
        shutil.copy2(WORKBOOK, workbook)
        src = self.directory / "src"
        shutil.copytree(ROOT / "src", src)
        return workbook, src

    def test_notes_label_is_padded_to_15_and_has_no_raw_ampersand(self):
        self.assertTrue(NOTES.startswith("NOTES!         →"))
        self.assertEqual(len("NOTES!         "), 15)
        self.assertEqual(len("               "), 15)
        self.assertNotIn("&", NOTES)
        self.assertIn("ss 9-70 and 9-75", NOTES)

    def test_pass_is_byte_noop_on_current_workbook(self):
        workbook, src = self._copy_tree()
        before = workbook.read_bytes()
        src_before = {p.name: p.read_bytes() for p in src.glob("*.txt")}
        result = self._run(workbook, src)
        self.assertEqual(result.returncode, 0, result.stdout + result.stderr)
        self.assertIn("already", result.stdout.lower())
        self.assertEqual(workbook.read_bytes(), before)
        self.assertEqual({p.name: p.read_bytes() for p in src.glob("*.txt")}, src_before)

    def test_pass_does_not_rewrite_about_or_comment(self):
        workbook, src = self._copy_tree()
        with zipfile.ZipFile(workbook) as archive:
            book = archive.read("xl/workbook.xml").decode("utf-8")
        financial = (src / "Financial.txt").read_text(encoding="utf-8")
        about = defined_name(book, "oz.AboutFinancialλ")
        gst_add = defined_name(book, "oz.GSTAddλ")
        gst_extract = defined_name(book, "oz.GSTExtractλ")

        self.assertIn(
            'comment="Adds GST to one or more GST-exclusive amounts."', gst_add
        )
        self.assertIn(
            'comment="Returns the GST contained in one or more GST-inclusive amounts."',
            gst_extract,
        )
        self.assertNotIn("ss 9-70", about)
        self.assertNotIn("NOTES!", about)
        self.assertIn("ss 9-70 and 9-75", gst_add)
        self.assertIn("ss 9-70 and 9-75", gst_extract)
        self.assertEqual(book.count("ss 9-70"), 2)
        self.assertIn(
            "GSTAddλ            →Adds GST to one or more GST-exclusive amounts.¶",
            financial,
        )
        self.assertIn(
            "GSTExtractλ        →Returns the GST contained in one or more GST-inclusive amounts.¶",
            financial,
        )
        self.assertEqual(financial.count("ss 9-70"), 2)

    def test_pass_transforms_pre_note_text(self):
        workbook, src = self._copy_tree()
        with zipfile.ZipFile(workbook) as archive:
            parts = {n: archive.read(n) for n in archive.namelist()}
        book = parts["xl/workbook.xml"].decode("utf-8")
        self.assertIn(GSTADD_NEW, book, "precondition: current workbook carries the GST note")
        parts["xl/workbook.xml"] = book.replace(GSTADD_NEW, GSTADD_OLD).encode("utf-8")
        with zipfile.ZipFile(workbook, "w", zipfile.ZIP_DEFLATED) as archive:
            for name, data in parts.items():
                archive.writestr(name, data)
        target = src / "Financial.txt"
        text = target.read_text(encoding="utf-8")
        self.assertIn(GSTADD_NEW, text, "precondition: src carries the GST note")
        target.write_text(text.replace(GSTADD_NEW, GSTADD_OLD), encoding="utf-8")

        result = self._run(workbook, src)
        self.assertEqual(result.returncode, 0, result.stdout + result.stderr)
        with zipfile.ZipFile(workbook) as archive:
            book_after = archive.read("xl/workbook.xml").decode("utf-8")
        self.assertIn(GSTADD_NEW, book_after)
        self.assertNotIn(GSTADD_OLD, book_after)
        self.assertIn(GSTADD_NEW, target.read_text(encoding="utf-8"))
        about = defined_name(book_after, "oz.AboutFinancialλ")
        self.assertNotIn("ss 9-70", about)
        self.assertIn(
            'comment="Adds GST to one or more GST-exclusive amounts."',
            defined_name(book_after, "oz.GSTAddλ"),
        )

    def test_pass_fails_loudly_when_expected_anchors_are_missing(self):
        self.assertTrue(PASS_SCRIPT.exists(), "pass script must exist for this test")
        workbook, src = self._copy_tree()
        with zipfile.ZipFile(workbook) as archive:
            parts = {n: archive.read(n) for n in archive.namelist()}
        book = parts["xl/workbook.xml"].decode("utf-8")
        book = book.replace(GSTADD_NEW, GSTADD_OLD)
        self.assertEqual(book.count(GSTADD_OLD), 1, "precondition: one pre-note anchor")
        book = book.replace(GSTADD_OLD, GSTADD_OLD + GSTADD_OLD, 1)
        parts["xl/workbook.xml"] = book.encode("utf-8")
        with zipfile.ZipFile(workbook, "w", zipfile.ZIP_DEFLATED) as archive:
            for name, data in parts.items():
                archive.writestr(name, data)
        result = self._run(workbook, src)
        self.assertNotEqual(result.returncode, 0, result.stdout + result.stderr)
        self.assertIn("FAIL", result.stdout + result.stderr)

    def test_pass_fails_when_an_anchor_and_its_replacement_are_both_absent(self):
        self.assertTrue(PASS_SCRIPT.exists(), "pass script must exist for this test")
        workbook, src = self._copy_tree()
        dummy = 'DESCRIPTION:   →Adds GST to one or more GST-exclusive amounts (removed).¶"'
        with zipfile.ZipFile(workbook) as archive:
            parts = {n: archive.read(n) for n in archive.namelist()}
        book = parts["xl/workbook.xml"].decode("utf-8")
        self.assertIn(GSTADD_NEW, book, "precondition: replacement present in workbook")
        parts["xl/workbook.xml"] = book.replace(GSTADD_NEW, dummy).encode("utf-8")
        with zipfile.ZipFile(workbook, "w", zipfile.ZIP_DEFLATED) as archive:
            for name, data in parts.items():
                archive.writestr(name, data)
        for module in ("Dates", "Essentials", "Financial", "Ratios", "Utilities", "Debt"):
            target = src / f"{module}.txt"
            text = target.read_text(encoding="utf-8")
            target.write_text(text.replace(GSTADD_NEW, dummy), encoding="utf-8")
        result = self._run(workbook, src)
        self.assertNotEqual(result.returncode, 0, result.stdout + result.stderr)
        self.assertIn("FAIL", result.stdout + result.stderr)


if __name__ == "__main__":
    unittest.main()
