"""Regression checks for Office discovery, export errors and work-PC diagnostics."""

import subprocess
import sys
import tempfile
import time
import unittest
from pathlib import Path
from unittest.mock import MagicMock, patch

WF = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(WF))
from core import office_convert as office, pdf_setup


class OfficeTests(unittest.TestCase):
    def setUp(self):
        self.temporary = tempfile.TemporaryDirectory(prefix="rca_office_test_")
        self.addCleanup(self.temporary.cleanup)
        self.folder = Path(self.temporary.name)
        self.docx = self.folder / "routing.docx"
        self.docx.write_bytes(b"test input; converter is mocked")

    def test_word_is_attempted_when_registry_detection_misses_office(self):
        expected = self.folder / "routing.pdf"
        with patch.object(office, "engines", return_value=["libreoffice", "msoffice"]), \
                patch.object(office, "find_soffice", return_value=None), \
                patch.object(office, "word_available", return_value=False) as hint, \
                patch.object(office, "_msoffice_to_pdf", return_value=expected) as export:
            self.assertEqual(office.office_to_pdf(self.docx, self.folder), expected)
            export.assert_called_once_with(self.docx, self.folder)
            hint.assert_not_called()

    def test_excel_export_does_not_require_word(self):
        source = self.folder / "analysis.xlsx"
        source.write_bytes(b"test input")
        expected = self.folder / "analysis.pdf"
        with patch.object(office, "engines", return_value=["msoffice"]), \
                patch.object(office, "word_available", return_value=False) as word, \
                patch.object(office, "excel_available", return_value=False) as excel, \
                patch.object(office, "_msoffice_to_pdf", return_value=expected):
            self.assertEqual(office.office_to_pdf(source, self.folder), expected)
            word.assert_not_called()
            excel.assert_not_called()

    def test_actual_office_error_is_preserved(self):
        with patch.object(office, "engines", return_value=["msoffice"]), \
                patch.object(office, "_msoffice_to_pdf", side_effect=RuntimeError("0x80040154 Class not registered")):
            with self.assertRaises(RuntimeError) as raised:
                office.office_to_pdf(self.docx, self.folder)
            message = str(raised.exception)
            self.assertIn("0x80040154 Class not registered", message)
            self.assertIn("Microsoft Word", message)
            self.assertIn("Check PDF setup", message)
            self.assertNotIn("Could not find a program", message)

    def test_missing_automation_component_is_not_reported_as_missing_word(self):
        with patch.object(office, "engines", return_value=["msoffice"]), \
                patch.object(office, "automation_support_error", return_value="Office automation component could not load: No module named win32com"):
            with self.assertRaisesRegex(RuntimeError, "No module named win32com"):
                office.office_to_pdf(self.docx, self.folder)

    def test_forced_libreoffice_is_respected_and_explained(self):
        with patch.dict(office.os.environ, {"RCA_OFFICE_ENGINE": "libreoffice"}), \
                patch.object(office, "find_soffice", return_value=None), \
                patch.object(office, "_msoffice_to_pdf") as export:
            with self.assertRaisesRegex(RuntimeError, "Office engine preference: libreoffice"):
                office.office_to_pdf(self.docx, self.folder)
            export.assert_not_called()

    def test_registry_hint_checks_other_view_and_com_registration(self):
        import winreg
        def open_key(root, key, reserved=0, access=0):
            if root == winreg.HKEY_CLASSES_ROOT and key == r"Word.Application\CLSID" and access & winreg.KEY_WOW64_32KEY:
                return MagicMock()
            raise FileNotFoundError(key)
        with patch.object(winreg, "OpenKey", side_effect=open_key), \
                patch.object(winreg, "QueryValueEx", return_value=("{test-clsid}", winreg.REG_SZ)):
            self.assertTrue(office._office_app_registered("Winword.exe"))

    def test_excel_recalculation_attempts_com_without_registry_hint(self):
        import pythoncom
        import win32com.client
        app = MagicMock()
        workbook = app.Workbooks.Open.return_value
        destination = self.folder / "recalculated.xlsx"
        workbook.SaveAs.side_effect = lambda path, **kwargs: Path(path).write_bytes(b"recalculated")
        with patch.object(office, "excel_available", return_value=False) as hint, \
                patch.object(win32com.client, "DispatchEx", return_value=app) as start, \
                patch.object(pythoncom, "CoInitialize"), patch.object(pythoncom, "CoUninitialize"):
            self.assertEqual(office.recalc_with_excel(self.docx, destination), destination)
            start.assert_called_once_with("Excel.Application")
            hint.assert_not_called()
            app.CalculateFullRebuild.assert_called_once()
            workbook.Close.assert_called_once_with(False)
            app.Quit.assert_called_once()

    def test_packaged_diagnostic_uses_companion_exe(self):
        with patch.object(sys, "frozen", True, create=True), \
                patch.object(sys, "executable", str(self.folder / "RCA Wizard Suite.exe")):
            command = pdf_setup._check_command("word", self.folder)
            self.assertEqual(command[:2], [str(self.folder / "RCA Wizard Tools.exe"), "pdfcheck"])

    def test_timeout_is_reported_and_excel_is_still_checked(self):
        completed = subprocess.CompletedProcess([], 0, "PASS: Excel created a readable PDF (1 page(s)).", "")
        with patch.object(pdf_setup.subprocess, "run", side_effect=[subprocess.TimeoutExpired("word", 60), completed]) as run:
            report = pdf_setup.run_report()
            self.assertIn("Word did not finish within 60 seconds", report)
            self.assertIn("PASS: Excel", report)
            self.assertNotIn("Both desktop Office PDF exports passed", report)
            self.assertEqual(run.call_count, 2)

    def test_both_diagnostic_samples_are_readable_office_files(self):
        from pypdf import PdfWriter
        from docx import Document
        from openpyxl import load_workbook
        def export(source, folder):
            if source.suffix == ".docx":
                self.assertIn("Word PDF conversion test", Document(source).paragraphs[0].text)
            else:
                wb = load_workbook(source)
                self.assertIn("Excel PDF conversion test", wb.active["A1"].value)
                wb.close()
            result = folder / (source.stem + ".pdf")
            writer = PdfWriter()
            writer.add_blank_page(width=612, height=792)
            writer.write(result)
            writer.close()
            return result
        with patch.object(office, "_msoffice_to_pdf", side_effect=export):
            for app_name in ("word", "excel"):
                self.assertEqual(pdf_setup.check_export(app_name, self.folder), 1)

    def test_home_button_opens_copyable_report_without_case_files(self):
        import tkinter as tk
        import app
        from core import session
        def descendants(widget):
            for child in widget.winfo_children():
                yield child
                yield from descendants(child)
        with patch.object(session, "SETTINGS_PATH", self.folder / "session.json"), \
                patch.object(pdf_setup, "run_report", return_value="TEST REPORT: Word export failure details"):
            root = app.SuiteApp()
            root.attributes("-alpha", 0)
            try:
                root.update()
                button = next(w for w in descendants(root) if isinstance(w, tk.Button)
                              and w.cget("text") == "Check PDF setup")
                button.invoke()
                windows = [w for w in root.winfo_children() if isinstance(w, tk.Toplevel)]
                self.assertEqual(len(windows), 1)
                window = windows[0]
                window.attributes("-alpha", 0)
                deadline = time.monotonic() + 3
                report_widget = next(w for w in descendants(window) if isinstance(w, tk.Text))
                while "TEST REPORT" not in report_widget.get("1.0", "end") and time.monotonic() < deadline:
                    root.update()
                    time.sleep(0.02)
                self.assertIn("Word export failure details", report_widget.get("1.0", "end"))
                for widget in descendants(window):
                    if isinstance(widget, tk.Button) and widget.cget("text") in ("Copy report", "Save report"):
                        self.assertEqual(str(widget.cget("state")), "normal")
            finally:
                root.destroy()


if __name__ == "__main__":
    unittest.main(verbosity=2)
