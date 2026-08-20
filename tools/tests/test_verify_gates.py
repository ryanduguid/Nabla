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


if __name__ == "__main__":
    unittest.main()
