import io
import shutil
import sys
import tempfile
import unittest
from contextlib import redirect_stdout
from pathlib import Path
from unittest import mock

TOOLS = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(TOOLS))

import verify_signatures
import verify_sources

ROOT = TOOLS.parent


class VerifyGateTests(unittest.TestCase):
    def test_declared_handles_lambda_without_let(self):
        self.assertEqual(
            verify_signatures.declared("LAMBDA(Cost, Life, Cost/Life)"),
            ["Cost", "Life"],
        )

    def test_declared_ignores_nested_let_in_lambda_body(self):
        self.assertEqual(
            verify_signatures.declared("LAMBDA(x, IF(x>0, LET(a, x, a), 0))"),
            ["x"],
        )

    def test_declared_keeps_current_top_level_let_form(self):
        self.assertEqual(
            verify_signatures.declared(
                "LAMBDA(Cost, Life, LET(Result, Cost/Life, Result))"
            ),
            ["Cost", "Life"],
        )

    def test_verify_sources_reads_new_txt_modules(self):
        directory = Path(tempfile.mkdtemp(prefix="ozzit-src-"))
        try:
            for source in (ROOT / "src").glob("*.txt"):
                shutil.copy2(source, directory / source.name)
            (directory / "Payroll.txt").write_text(
                "PayrollTaxλ = LAMBDA(Income, LET(Result, Income, Result));\n",
                encoding="utf-8",
            )

            verify_sources.failures.clear()
            output = io.StringIO()
            with (
                mock.patch.object(verify_sources, "SRC_DIR", str(directory)),
                mock.patch.object(verify_sources, "WORKBOOK", str(ROOT / "ozzit.xlsx")),
                redirect_stdout(output),
            ):
                result = verify_sources.main()

            self.assertEqual(result, 1)
            self.assertIn(
                "oz.PayrollTaxλ is in src/ but does not ship in the workbook",
                output.getvalue(),
            )
        finally:
            shutil.rmtree(directory, ignore_errors=True)

    def _copy_src(self):
        directory = Path(tempfile.mkdtemp(prefix="ozzit-src-"))
        for source in (ROOT / "src").glob("*.txt"):
            shutil.copy2(source, directory / source.name)
        return directory

    def test_verify_signatures_rejects_coming_soon_webpage(self):
        directory = self._copy_src()
        try:
            path = directory / "Financial.txt"
            path.write_text(
                path.read_text(encoding="utf-8").replace(
                    '"WEBPAGE:       →https://github.com/ryanduguid/Ozzit¶"',
                    '"WEBPAGE:       →<Coming Soon>¶"',
                    1,
                ),
                encoding="utf-8",
            )
            output = io.StringIO()
            with (
                mock.patch.object(verify_signatures, "SRC", str(directory)),
                redirect_stdout(output),
            ):
                result = verify_signatures.main()
            self.assertEqual(result, 1)
            self.assertIn("Coming Soon", output.getvalue())
        finally:
            shutil.rmtree(directory, ignore_errors=True)

    def test_verify_signatures_rejects_copied_about_descriptions(self):
        directory = self._copy_src()
        try:
            path = directory / "Financial.txt"
            path.write_text(
                path.read_text(encoding="utf-8").replace(
                    '"RollingAvgλ        →Creates averages for preceding values of a set size moving from beginning to end over a row of values.¶" &',
                    '"RollingAvgλ        →Finds the maximum value the n preceding values in a set moving from beginning to end over a row of values.¶" &',
                    1,
                ),
                encoding="utf-8",
            )
            output = io.StringIO()
            with (
                mock.patch.object(verify_signatures, "SRC", str(directory)),
                redirect_stdout(output),
            ):
                result = verify_signatures.main()
            self.assertEqual(result, 1)
            self.assertIn("share the same About-table description", output.getvalue())
        finally:
            shutil.rmtree(directory, ignore_errors=True)


if __name__ == "__main__":
    unittest.main()
