import shutil
import subprocess
import sys
import tempfile
import unittest
import zipfile
from pathlib import Path
from types import SimpleNamespace
from unittest import mock

TOOLS = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(TOOLS))

import refresh_cache
import sanitise_workbook
import verify_workbook

ROOT = TOOLS.parent
WORKBOOK = ROOT / "ozzit.xlsx"


class WorkbookToolTests(unittest.TestCase):
    def setUp(self):
        self.directory = Path(tempfile.mkdtemp(prefix="ozzit-sanitise-"))
        self.workbook = self.directory / "dirty.xlsx"
        shutil.copy2(WORKBOOK, self.workbook)

    def tearDown(self):
        shutil.rmtree(self.directory, ignore_errors=True)

    def _parts(self):
        with zipfile.ZipFile(self.workbook) as z:
            return {name: z.read(name) for name in z.namelist()}

    def _rewrite(self, parts):
        with zipfile.ZipFile(self.workbook, "w", zipfile.ZIP_DEFLATED) as z:
            for name, data in parts.items():
                z.writestr(name, data)

    def test_canonical_clean_workbook_is_byte_identical(self):
        sanitise_workbook.write_deterministic(self.workbook, self._parts())
        before = self.workbook.read_bytes()
        self.assertEqual(sanitise_workbook.sanitise(self.workbook), ["already clean"])
        self.assertEqual(self.workbook.read_bytes(), before)

    def test_clean_noncanonical_workbook_is_repacked(self):
        self._rewrite(self._parts())
        before = self.workbook.read_bytes()
        log = sanitise_workbook.sanitise(self.workbook)
        self.assertIn("canonicalised archive metadata and compression", log)
        self.assertNotEqual(self.workbook.read_bytes(), before)
        with zipfile.ZipFile(self.workbook) as archive:
            self.assertEqual(archive.namelist(), sorted(archive.namelist()))
            self.assertTrue(
                all(
                    info.date_time == sanitise_workbook.FIXED_DATE
                    and info.create_system == 3
                    and (info.external_attr >> 16) == 0o644
                    for info in archive.infolist()
                )
            )

    def test_privacy_gate_matches_real_windows_user_path(self):
        self.assertIsNotNone(
            verify_workbook.BANNED.search(r"C:\Users\Example\Documents\Ozzit")
        )

    def test_integrity_gate_rejects_xml_doctype_before_parsing(self):
        verify_workbook.failures.clear()
        verify_workbook.check_xml_part(
            "xl/worksheets/sheet1.xml",
            b'<!DOCTYPE x [<!ENTITY a "boom">]><x>&a;</x>',
        )
        self.assertIn(
            "DOCTYPE declaration in xl/worksheets/sheet1.xml",
            verify_workbook.failures,
        )

    def test_cell_match_covers_shared_formula_follower_to_end_tag(self):
        cell = (
            '<c ca="1" r="A31" s="121" t="str">'
            '<f t="shared" si="0"/><v>Town Hall</v></c>'
        )
        self.assertEqual(sanitise_workbook.CELL_RE.search(cell).group(0), cell)

    def test_strips_printer_settings_and_writes_deterministically(self):
        parts = self._parts()
        printer = "xl/printerSettings/printerSettings99.bin"
        parts[printer] = b"printer-specific bytes"

        rels = "xl/worksheets/_rels/sheet2.xml.rels"
        parts[rels] = (
            b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            b'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            b'<Relationship Id="rId99" '
            b'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/printerSettings" '
            b'Target="../printerSettings/printerSettings99.bin"/>'
            b"</Relationships>"
        )

        sheet = "xl/worksheets/sheet2.xml"
        text = parts[sheet].decode()
        text = text.replace("<pageSetup ", '<pageSetup r:id="rId99" ', 1)
        parts[sheet] = text.encode()
        self._rewrite(parts)

        log = sanitise_workbook.sanitise(self.workbook)
        self.assertIn("removed 1 printerSettings parts", log)
        self.assertIn("cleared pageSetup r:id on 1 sheets", log)
        self.assertIn("dropped 1 empty worksheet rels", log)

        with zipfile.ZipFile(self.workbook) as z:
            self.assertEqual(z.namelist(), sorted(z.namelist()))
            self.assertTrue(
                all(
                    info.date_time == sanitise_workbook.FIXED_DATE
                    for info in z.infolist()
                )
            )
            self.assertNotIn(printer, z.namelist())
            self.assertNotIn(rels, z.namelist())
            self.assertNotIn('r:id="rId99"', z.read(sheet).decode())

        first = self.workbook.read_bytes()
        self.assertEqual(sanitise_workbook.sanitise(self.workbook), ["already clean"])
        self.assertEqual(self.workbook.read_bytes(), first)

    def test_atomic_writer_preserves_original_and_cleans_temp_if_replace_fails(self):
        original = self.workbook.read_bytes()
        parts = self._parts()
        parts["docProps/core.xml"] += b" "

        writer = sanitise_workbook.write_deterministic
        with (
            mock.patch("sanitise_workbook.os.replace", side_effect=OSError("blocked")),
            self.assertRaisesRegex(OSError, "blocked"),
        ):
            writer(self.workbook, parts)

        self.assertEqual(self.workbook.read_bytes(), original)
        self.assertFalse(self.workbook.with_name(self.workbook.name + ".tmp").exists())

    def test_deterministic_writer_matches_explicit_level_9_reference(self):
        parts = self._parts()
        actual = self.directory / "actual.xlsx"
        expected = self.directory / "expected.xlsx"

        sanitise_workbook.write_deterministic(actual, parts)
        with zipfile.ZipFile(
            expected, "w", zipfile.ZIP_DEFLATED, compresslevel=9
        ) as archive:
            for name in sorted(parts):
                info = zipfile.ZipInfo(name, date_time=sanitise_workbook.FIXED_DATE)
                info.compress_type = zipfile.ZIP_DEFLATED
                info.create_system = 3
                info.external_attr = 0o644 << 16
                archive.writestr(info, parts[name], compresslevel=9)

        self.assertEqual(actual.read_bytes(), expected.read_bytes())

    def test_deterministic_writer_uses_portable_file_metadata(self):
        parts = self._parts()
        output = self.directory / "metadata.xlsx"
        sanitise_workbook.write_deterministic(output, parts)
        with zipfile.ZipFile(output) as archive:
            self.assertTrue(all(info.create_system == 3 for info in archive.infolist()))
            self.assertTrue(
                all((info.external_attr >> 16) == 0o644 for info in archive.infolist())
            )

    def test_cli_reports_non_zip_input_without_traceback(self):
        broken = self.directory / "broken.xlsx"
        broken.write_text("not a zip", encoding="utf-8")
        result = subprocess.run(
            [sys.executable, str(TOOLS / "sanitise_workbook.py"), str(broken)],
            capture_output=True,
            text=True,
            check=False,
        )
        self.assertNotEqual(result.returncode, 0)
        self.assertTrue(result.stderr.startswith("FAIL:"), result.stderr)
        self.assertNotIn("Traceback", result.stderr)

    def test_cli_reports_missing_input_without_traceback(self):
        missing = self.directory / "missing.xlsx"
        result = subprocess.run(
            [sys.executable, str(TOOLS / "sanitise_workbook.py"), str(missing)],
            capture_output=True,
            text=True,
            check=False,
        )
        self.assertNotEqual(result.returncode, 0)
        self.assertTrue(result.stderr.startswith("FAIL:"), result.stderr)
        self.assertNotIn("Traceback", result.stderr)

    def test_refresh_cache_delegates_all_cleanup_to_shared_sanitiser(self):
        refresh_cache.WORKBOOK = str(self.workbook)
        completed = SimpleNamespace(returncode=0, stdout="refreshed", stderr="")
        with (
            mock.patch.object(refresh_cache.subprocess, "run", return_value=completed),
            mock.patch.object(
                refresh_cache,
                "sanitise",
                return_value=["removed 1 x15ac:absPath"],
            ) as shared,
        ):
            self.assertEqual(refresh_cache.main(), 0)

        shared.assert_called_once_with(Path(self.workbook))


if __name__ == "__main__":
    unittest.main()
