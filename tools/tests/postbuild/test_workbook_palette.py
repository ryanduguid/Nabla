"""Contract for tools/postbuild/workbook_palette.py.

The pass applies the house palette: explicit colour remaps in theme,
styles, sheets and drawings, plus the font-family consolidation, the help-label
greens folded to brand purple, and the mint help-block fill folded to pale
lavender. It must be a byte no-op on the current workbook and must refuse to
touch the deliberate TOC warning red (dxf FFCCCC/C00000).
"""

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
PASS_SCRIPT = TOOLS / "postbuild" / "workbook_palette.py"


class WorkbookPaletteTests(unittest.TestCase):
    def setUp(self):
        self.directory = Path(tempfile.mkdtemp(prefix="ozzit-palette-"))
        self.workbook = self.directory / "ozzit.xlsx"
        shutil.copy2(WORKBOOK, self.workbook)

    def tearDown(self):
        shutil.rmtree(self.directory, ignore_errors=True)

    def _run(self):
        return subprocess.run(
            [sys.executable, str(PASS_SCRIPT), str(self.workbook)],
            capture_output=True,
            text=True,
            check=False,
        )

    def _styles(self):
        with zipfile.ZipFile(self.workbook) as archive:
            return archive.read("xl/styles.xml").decode("utf-8")

    def test_pass_is_byte_noop_on_current_workbook(self):
        before = self.workbook.read_bytes()
        result = self._run()
        self.assertEqual(result.returncode, 0, result.stdout + result.stderr)
        self.assertIn("already", result.stdout.lower())
        self.assertEqual(self.workbook.read_bytes(), before)

    def test_pass_preserves_the_deliberate_toc_warning_red(self):
        # Even after a run, the TOC length-guard dxf keeps its warning colours.
        result = self._run()
        self.assertEqual(result.returncode, 0, result.stdout + result.stderr)
        styles = self._styles()
        dxfs = styles[styles.index("<dxfs") : styles.index("</dxfs>")]
        self.assertIn("FFFFCCCC", dxfs, "TOC warning fill must survive the palette pass")
        self.assertIn("FFC00000", dxfs, "TOC warning font must survive the palette pass")

    def test_pass_refolds_a_reverted_green_font_and_mint_fill(self):
        # Simulate pre-v3.1.0 state: put one legacy green font and one mint fill
        # back into styles.xml, then require the pass to fold them again.
        with zipfile.ZipFile(self.workbook) as archive:
            parts = {n: archive.read(n) for n in archive.namelist()}
        styles = parts["xl/styles.xml"].decode("utf-8")
        fonts = styles[styles.index("<fonts") : styles.index("</fonts>")]
        styles = styles.replace(
            fonts,
            fonts.replace('rgb="FF5C2D91"', 'rgb="FF006600"', 1),
            1,
        )
        fills = styles[styles.index("<fills") : styles.index("</fills>")]
        self.assertIn('rgb="FFDED9E8"', fills, "precondition: lavender fill present")
        styles = styles.replace(
            fills,
            fills.replace('rgb="FFDED9E8"', 'indexed="42"', 1),
            1,
        )
        parts["xl/styles.xml"] = styles.encode("utf-8")
        with zipfile.ZipFile(self.workbook, "w", zipfile.ZIP_DEFLATED) as archive:
            for name, data in parts.items():
                archive.writestr(name, data)

        result = self._run()
        self.assertEqual(result.returncode, 0, result.stdout + result.stderr)
        styles = self._styles()
        fonts = styles[styles.index("<fonts") : styles.index("</fonts>")]
        self.assertNotIn('rgb="FF006600"', fonts, "green font must be folded to purple")
        fills = styles[styles.index("<fills") : styles.index("</fills>")]
        self.assertNotIn('indexed="42"', fills, "mint fill must be folded to lavender")

    def test_pass_folds_rich_text_font_and_colour_in_shared_strings(self):
        # The TOC "Totals ↓ Dates→" string carries a bold run in Aptos Narrow /
        # FF2E4950. The pass must fold rich-text runs the same way as styles fonts.
        result = self._run()
        self.assertEqual(result.returncode, 0, result.stdout + result.stderr)
        with zipfile.ZipFile(self.workbook) as archive:
            shared = archive.read("xl/sharedStrings.xml").decode("utf-8")
        self.assertNotIn('rFont val="Aptos Narrow"', shared)
        self.assertNotIn('rgb="FF2E4950"', shared)

    def test_pass_refolds_a_reverted_rich_text_run_in_shared_strings(self):
        # Exercise the folding path itself: put the pre-v3.1.0 run back into
        # sharedStrings.xml, then require the pass to fold it again. The end-state
        # test above cannot catch a broken sharedStrings path on a clean workbook.
        with zipfile.ZipFile(self.workbook) as archive:
            parts = {n: archive.read(n) for n in archive.namelist()}
        shared = parts["xl/sharedStrings.xml"].decode("utf-8")
        self.assertIn('rgb="FF2B2733"', shared, "precondition: folded colour present")
        self.assertIn('rFont val="Aptos"', shared, "precondition: folded font present")
        shared = shared.replace('rgb="FF2B2733"', 'rgb="FF2E4950"', 1)
        shared = shared.replace('rFont val="Aptos"', 'rFont val="Aptos Narrow"', 1)
        parts["xl/sharedStrings.xml"] = shared.encode("utf-8")
        with zipfile.ZipFile(self.workbook, "w", zipfile.ZIP_DEFLATED) as archive:
            for name, data in parts.items():
                archive.writestr(name, data)

        result = self._run()
        self.assertEqual(result.returncode, 0, result.stdout + result.stderr)
        with zipfile.ZipFile(self.workbook) as archive:
            shared = archive.read("xl/sharedStrings.xml").decode("utf-8")
        self.assertNotIn('rFont val="Aptos Narrow"', shared)
        self.assertNotIn('rgb="FF2E4950"', shared)
        self.assertIn('rgb="FF2B2733"', shared)


if __name__ == "__main__":
    unittest.main()
