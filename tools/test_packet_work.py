"""Packet retries must not depend on deleting an earlier attempt's folder."""

import importlib
from pathlib import Path
import shutil
import sys
import tempfile
import unittest
import warnings
from unittest.mock import patch

WF = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(WF))
from core import packet_work, office_convert
from core.paths import sample_input_dir, templates_dir
from core.workbooks import CALCULATION_TEMPLATES

TEMPLATES = {
    "mls": ("MLS_Routing Form.docx", "MLS_Director_Approval.docx", "PC_Approval_NEW.docx"),
    "mbl": ("MBL_Routing_Form.docx", "MBL_Directors_Approval.docx", "MBL_PC_Approval.docx"),
    "mls_renewal": ("MLS_Renewal_Routing_Form.docx", "MLS_REN_Director_Approval.docx", "MLS_REN_PC_Approval.docx"),
    "mbl_renewal": ("MBL_REN_Routing_Form.docx", "MBL_REN_Directors_Approval.docx", "MBL_REN_PC_Approval.docx"),
}


def sample_packet(key, output, attachment):
    module_name = "wizards.mls.templates_2026" if key == "mls" else f"wizards.{key}.wizard"
    module = importlib.import_module(module_name)
    workbook = sample_input_dir(key) / (Path(CALCULATION_TEMPLATES[key]).stem + "_FILLED_SAMPLE.xlsx")
    templates = dict(zip(("routing", "director", "pc"),
                        (templates_dir(key) / name for name in TEMPLATES[key])))
    if key == "mls":
        return module.generate_rca_packet_pdf(workbook, templates, attachment, attachment, output)
    args = [workbook, templates, attachment, "", attachment, ""]
    if key.endswith("renewal"):
        args.append(attachment)
    return module.generate_rca_packet_pdf(*args, output, include_staff_analysis=False)


class PacketWorkTests(unittest.TestCase):
    def setUp(self):
        self.temporary = tempfile.TemporaryDirectory(prefix="rca_packet_tests_")
        self.addCleanup(self.temporary.cleanup)
        self.root = Path(self.temporary.name).resolve()
        self.temp_root = self.root / "WindowsTemp"
        self.temp_root.mkdir()
        self.case = self.root / "OneDrive - City of Houston" / "Case" / "RCA"
        self.case.mkdir(parents=True)
        self.temp_patch = patch.object(tempfile, "gettempdir", return_value=str(self.temp_root))
        self.temp_patch.start()
        self.addCleanup(self.temp_patch.stop)
        self.real_rmtree = shutil.rmtree
        self.attachment = self.root / "attachment.pdf"
        self.write_pdf(self.attachment)

    def write_pdf(self, path):
        import fitz
        with fitz.open() as document:
            page = document.new_page()
            page.insert_text((72, 72), "RCA packet test fixture: " + path.name)
            document.save(path)

    def test_simultaneous_workspaces_are_unique_and_cleaned(self):
        with packet_work.packet_workspace() as first:
            with packet_work.packet_workspace() as second:
                self.assertNotEqual(first, second)
                self.assertEqual(first.parent, self.temp_root)
                self.assertEqual(second.parent, self.temp_root)
                self.assertFalse(first.is_relative_to(self.case))
                (first / "source.txt").write_text("first")
                (second / "source.txt").write_text("second")
            self.assertFalse(second.exists())
            self.assertEqual((first / "source.txt").read_text(), "first")
        self.assertFalse(first.exists())

    def test_cleanup_access_denied_does_not_fail_success_or_next_attempt(self):
        with patch.object(shutil, "rmtree", side_effect=PermissionError("WinError 5: Access is denied")):
            with packet_work.packet_workspace() as locked:
                (locked / "output.txt").write_text("completed")
        self.assertTrue(locked.exists())
        with packet_work.packet_workspace() as next_attempt:
            self.assertNotEqual(locked, next_attempt)
        self.assertFalse(next_attempt.exists())

    def test_cleanup_error_does_not_hide_original_conversion_error(self):
        with patch.object(shutil, "rmtree", side_effect=PermissionError("cleanup denied")):
            with self.assertRaisesRegex(RuntimeError, "actual Word export failure"):
                with packet_work.packet_workspace():
                    raise RuntimeError("actual Word export failure")

    def test_all_four_packets_retry_without_touching_locked_old_folders(self):
        from pypdf import PdfReader
        for key in TEMPLATES:
            with self.subTest(wizard=key):
                output_dir = self.case / key
                output_dir.mkdir()
                blocked = []
                for name in ("_packet_work", "_mbl_packet_work", "_renewal_packet_work"):
                    old = output_dir / name
                    old.mkdir()
                    (old / "keep.txt").write_text("previous attempt, do not delete")
                    blocked.append(old.resolve())
                touched = []
                stages = []
                def guarded_rmtree(path, *args, **kwargs):
                    resolved = Path(path).resolve()
                    if any(resolved == old or resolved.is_relative_to(old) for old in blocked):
                        touched.append(resolved)
                        raise PermissionError("WinError 5: old packet folder is locked")
                    return self.real_rmtree(path, *args, **kwargs)
                def export(source, folder):
                    source, folder = Path(source), Path(folder)
                    if source.name.startswith("01_"):
                        stages.append(folder)
                    result = folder / (source.stem + ".pdf")
                    self.write_pdf(result)
                    return result
                output = output_dir / "RCA_Packet.pdf"
                with patch.object(shutil, "rmtree", side_effect=guarded_rmtree), \
                        patch.object(office_convert, "office_to_pdf", side_effect=export):
                    for attempt in range(2):
                        count = sample_packet(key, output, self.attachment)
                        with output.open("rb") as stream:
                            self.assertEqual(len(PdfReader(stream).pages), count)
                self.assertEqual(touched, [])
                self.assertEqual(len(stages), 2)
                self.assertNotEqual(*stages)
                for stage in stages:
                    self.assertEqual(stage.parent, self.temp_root)
                    self.assertFalse(stage.exists())
                for old in blocked:
                    self.assertEqual((old / "keep.txt").read_text(), "previous attempt, do not delete")

    def test_completed_renewal_packet_survives_cleanup_denial(self):
        def deny_packet_cleanup(path, *args, **kwargs):
            if Path(path).name.startswith("rca_packet_"):
                raise PermissionError("Windows is holding the temporary directory")
            return self.real_rmtree(path, *args, **kwargs)
        def export(source, folder):
            result = Path(folder) / (Path(source).stem + ".pdf")
            self.write_pdf(result)
            return result
        output = self.case / "RCA_Packet.pdf"
        with patch.object(shutil, "rmtree", side_effect=deny_packet_cleanup), \
                patch.object(office_convert, "office_to_pdf", side_effect=export):
            self.assertEqual(sample_packet("mbl_renewal", output, self.attachment), 5)
        self.assertTrue(output.is_file())


if __name__ == "__main__":
    warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")
    unittest.main(verbosity=2)
