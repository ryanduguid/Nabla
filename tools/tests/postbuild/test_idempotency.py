"""Every XML-only postbuild pass must be a byte no-op on the committed workbook.

This is the reproducibility property testable without the upstream workbook: a
second run either reports "already applied" and changes nothing, or fails on an
assertion. A silent mutation means the pass is not idempotent and must not ship.
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

XML_PASSES = [
    (TOOLS / "postbuild" / "fy27_help_text.py", True),
    (TOOLS / "postbuild" / "luma_palette.py", False),
    (TOOLS / "postbuild" / "gst_help_text.py", True),
    (TOOLS / "postbuild" / "help_links.py", True),
    (TOOLS / "postbuild" / "sheet_names.py", False),
]


class PostbuildIdempotencyTests(unittest.TestCase):
    def setUp(self):
        self.directory = Path(tempfile.mkdtemp(prefix="ozzit-idem-"))
        self.workbook = self.directory / "ozzit.xlsx"
        shutil.copy2(WORKBOOK, self.workbook)
        self.src = self.directory / "src"
        shutil.copytree(ROOT / "src", self.src)

    def tearDown(self):
        shutil.rmtree(self.directory, ignore_errors=True)

    def test_each_xml_pass_is_a_byte_noop(self):
        for script, takes_src in XML_PASSES:
            with self.subTest(script=script.name):
                command = [sys.executable, str(script), str(self.workbook)]
                if takes_src:
                    command.append(str(self.src))
                first = subprocess.run(command, capture_output=True, text=True, check=False)
                self.assertEqual(first.returncode, 0, first.stdout + first.stderr)
                before = self.workbook.read_bytes()
                src_before = {p.name: p.read_bytes() for p in self.src.glob("*.txt")}

                second = subprocess.run(command, capture_output=True, text=True, check=False)
                self.assertEqual(second.returncode, 0, second.stdout + second.stderr)
                self.assertIn("already", second.stdout.lower())
                self.assertEqual(
                    self.workbook.read_bytes(), before, f"{script.name} mutated the workbook"
                )
                self.assertEqual(
                    {p.name: p.read_bytes() for p in self.src.glob("*.txt")},
                    src_before,
                    f"{script.name} mutated src/",
                )


if __name__ == "__main__":
    unittest.main()
