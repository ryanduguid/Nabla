"""Contract for tools/postbuild/fy27_help_text.py.

The pass rewrites dated help-text examples from the v3.0.0 state to the v3.1.0
(FY27) state, length-preserving, with asserted hit counts. On the current
workbook it must be a byte no-op; on a workbook reverted to the v3.0.0 text it
must apply each swap exactly the recorded number of times.
"""

import shutil
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[3]
TOOLS = ROOT / "tools"
WORKBOOK = ROOT / "ozzit.xlsx"
PASS_SCRIPT = TOOLS / "postbuild" / "fy27_help_text.py"

# A swap recorded in the v3.1.0 commit: old -> new, hits in workbook.xml + src.
EXAMPLE_SWAPS = [
    ('""22/3/2012"", ""10/4/2012""', '""23/7/2026"", ""11/8/2026""'),
    ('EDATE(""1/1/2025"", SEQUENCE(,24, 0))', 'EDATE(""1/7/2026"", SEQUENCE(,24, 0))'),
]


class Fy27HelpTextTests(unittest.TestCase):
    def setUp(self):
        self.directory = Path(tempfile.mkdtemp(prefix="ozzit-postbuild-"))

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

    def test_pass_is_byte_noop_on_current_workbook(self):
        workbook, src = self._copy_tree()
        before = workbook.read_bytes()
        src_before = {p.name: p.read_bytes() for p in src.glob("*.txt")}
        result = self._run(workbook, src)
        self.assertEqual(result.returncode, 0, result.stdout + result.stderr)
        self.assertIn("already", result.stdout.lower())
        self.assertEqual(workbook.read_bytes(), before)
        self.assertEqual({p.name: p.read_bytes() for p in src.glob("*.txt")}, src_before)

    def test_pass_transforms_v30_text_and_is_length_preserving(self):
        workbook, src = self._copy_tree()
        # Revert one swap inside the workbook and src to simulate v3.0.0 state.
        import zipfile

        old, new = EXAMPLE_SWAPS[1]
        with zipfile.ZipFile(workbook) as archive:
            parts = {n: archive.read(n) for n in archive.namelist()}
        book = parts["xl/workbook.xml"].decode("utf-8")
        self.assertIn(new, book, "precondition: current workbook carries the FY27 text")
        parts["xl/workbook.xml"] = book.replace(new, old).encode("utf-8")
        with zipfile.ZipFile(workbook, "w", zipfile.ZIP_DEFLATED) as archive:
            for name, data in parts.items():
                archive.writestr(name, data)
        target = src / "Financial.txt"
        text = target.read_text(encoding="utf-8")
        self.assertIn(new, text, "precondition: src carries the FY27 text")
        target.write_text(text.replace(new, old), encoding="utf-8")

        result = self._run(workbook, src)
        self.assertEqual(result.returncode, 0, result.stdout + result.stderr)
        with zipfile.ZipFile(workbook) as archive:
            book_after = archive.read("xl/workbook.xml").decode("utf-8")
        self.assertIn(new, book_after)
        self.assertNotIn(old, book_after)
        self.assertIn(new, target.read_text(encoding="utf-8"))

    def test_pass_fails_loudly_when_expected_anchors_are_missing(self):
        self.assertTrue(PASS_SCRIPT.exists(), "pass script must exist for this test")
        workbook, src = self._copy_tree()
        # Force an anchor to appear at the wrong multiplicity: the CountDOW swap
        # expects exactly 1 hit in the workbook. Revert the replacement and then
        # duplicate the old text, so the pass sees 2 hits and must refuse.
        import zipfile

        old, new = EXAMPLE_SWAPS[0]
        with zipfile.ZipFile(workbook) as archive:
            parts = {n: archive.read(n) for n in archive.namelist()}
        book = parts["xl/workbook.xml"].decode("utf-8")
        book = book.replace(new, old)  # back to v3.0.0 text (1 hit)
        self.assertEqual(book.count(old), 1, "precondition: one v3.0.0 anchor")
        book = book.replace(old, old + old, 1)  # now 2 hits where 1 is expected
        parts["xl/workbook.xml"] = book.encode("utf-8")
        with zipfile.ZipFile(workbook, "w", zipfile.ZIP_DEFLATED) as archive:
            for name, data in parts.items():
                archive.writestr(name, data)
        result = self._run(workbook, src)
        self.assertNotEqual(result.returncode, 0, result.stdout + result.stderr)
        self.assertIn("FAIL", result.stdout + result.stderr)


if __name__ == "__main__":
    unittest.main()
