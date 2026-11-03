"""Regression checks for calculation downloads, defaults and saved edits.

Run: python tools/test_workflow.py
Uses temporary files and invisible Tk windows; does not touch user settings.
"""

from __future__ import annotations

import importlib
import json
import os
from pathlib import Path
import shutil
import sys
import tempfile
import unittest
from unittest.mock import patch

WF = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(WF))

from core import session, recalc
from core.paths import templates_dir
from core.workbooks import (
    CALCULATION_TEMPLATES, WorkbookReadCache,
    calculation_template_path, save_calculation_sheet,
)

WIZARDS = {
    "mls": "MailoutWizard", "mbl": "MBLWizard",
    "mls_renewal": "MLSRenewalWizard", "mbl_renewal": "MBLRenewalWizard",
}


class WorkflowTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory(prefix="rca_workflow_test_")
        self.addCleanup(self.temp.cleanup)
        self.folder = Path(self.temp.name)
        self.settings = patch.object(session, "SETTINGS_PATH", self.folder / "settings.json")
        self.settings.start()
        self.addCleanup(self.settings.stop)

    def test_downloads_preserve_every_template_byte(self):
        for key in WIZARDS:
            with self.subTest(wizard=key):
                source = calculation_template_path(key)
                before = source.read_bytes()
                saved = save_calculation_sheet(key, self.folder / (key + ".xlsx"))
                self.assertEqual(saved.read_bytes(), before)
                self.assertEqual(source.read_bytes(), before)
                with self.assertRaises(ValueError):
                    save_calculation_sheet(key, source)
                with self.assertRaises(ValueError):
                    save_calculation_sheet(key, self.folder / (key + ".xls"))

    def test_session_retains_cases_but_not_old_templates_or_samples(self):
        case = self.folder / "my_case.xlsx"
        case.write_bytes(b"case")
        old_template = self.folder / "old.docx"
        old_template.write_bytes(b"old")
        for key in WIZARDS:
            saved = {"workbook": str(case), "out": str(self.folder / "output"),
                     "template": str(old_template), "director_template": str(old_template)}
            session.SETTINGS_PATH.write_text(json.dumps({key: saved}), encoding="utf-8")
            self.assertEqual(session.load(key), {"workbook": str(case), "out": saved["out"]})
            session.save(key, saved)
            self.assertNotIn("director_template", json.loads(session.SETTINGS_PATH.read_text())[key])
            for filename in [CALCULATION_TEMPLATES[key], "old_FILLED_SAMPLE.xlsx"]:
                old_sample = self.folder / "old_suite" / "sample_input" / key / filename
                old_sample.parent.mkdir(parents=True, exist_ok=True)
                old_sample.write_bytes(b"old example")
                session.SETTINGS_PATH.write_text(json.dumps({key: {"excel": str(old_sample)}}))
                self.assertEqual(session.load(key), {})

    def test_all_wizards_refresh_after_same_path_same_size_same_timestamp_save(self):
        for key, classname in WIZARDS.items():
            with self.subTest(wizard=key):
                module = importlib.import_module(f"wizards.{key}.wizard")
                cls = getattr(module, classname)
                source = self.folder / (key + ".xlsx")
                source.write_bytes(b"before")
                holder = type("Holder", (), {})()
                holder.files = {"workbook": str(source)}
                holder.vars = {"excel": type("Value", (), {"get": lambda _: str(source)})()}
                copies = []

                def recalculate(path):
                    dest = self.folder / f"{key}_{len(copies)}.xlsx"
                    shutil.copyfile(path, dest)
                    copies.append(dest)
                    return dest

                with patch.object(recalc, "needs_recalc", return_value=True), \
                        patch.object(recalc, "recalculated_copy", side_effect=recalculate):
                    first = Path(cls._wb(holder))
                    self.assertEqual(first.read_bytes(), b"before")
                    self.assertEqual(cls._wb(holder), str(first))
                    self.assertEqual(len(copies), 1)
                    old_stat = source.stat()
                    source.write_bytes(b"edited")
                    os.utime(source, ns=(old_stat.st_atime_ns, old_stat.st_mtime_ns))
                    second = Path(cls._wb(holder))
                    self.assertEqual(second.read_bytes(), b"edited")
                    self.assertNotEqual(first, second)
                    second.unlink()
                    self.assertEqual(Path(cls._wb(holder)).read_bytes(), b"edited")
                    self.assertEqual(len(copies), 3)
                    source.unlink()
                    with self.assertRaises(FileNotFoundError):
                        cls._wb(holder)

    def test_save_during_recalculation_does_not_cache_old_values(self):
        source = self.folder / "case.xlsx"
        source.write_bytes(b"before")
        cache = WorkbookReadCache()
        copies = []

        def recalculate(path):
            dest = self.folder / f"copy{len(copies)}.xlsx"
            shutil.copyfile(path, dest)
            copies.append(dest)
            if len(copies) == 1:
                source.write_bytes(b"edited")
            return dest

        with patch.object(recalc, "needs_recalc", return_value=True), \
                patch.object(recalc, "recalculated_copy", side_effect=recalculate):
            self.assertEqual(Path(cache.resolve(source)).read_bytes(), b"edited")
            self.assertEqual(len(copies), 2)

    def test_excel_calculated_save_stops_using_temporary_copy(self):
        source = self.folder / "case.xlsx"
        source.write_bytes(b"uncalculated")
        previous = self.folder / "previous.xlsx"
        previous.write_bytes(b"previous result")
        cache = WorkbookReadCache()
        with patch.object(recalc, "needs_recalc", side_effect=[True, False]), \
                patch.object(recalc, "recalculated_copy", return_value=previous):
            self.assertEqual(cache.resolve(source), str(previous))
            source.write_bytes(b"saved and calculated by Excel")
            self.assertEqual(cache.resolve(source), str(source))

    def test_home_download_buttons_and_wizard_defaults(self):
        import app
        # Prevent contact with the user's saved settings or Office installations.
        with patch.object(session, "load", return_value={}), \
                patch.object(app.SuiteApp, "_setup_summary", return_value="Test", create=True), \
                patch.object(app.messagebox, "showerror") as error:
            root = app.SuiteApp()
            root.attributes("-alpha", 0)
            self.addCleanup(root.destroy)

            def descendants(widget):
                for child in widget.winfo_children():
                    yield child
                    yield from descendants(child)

            for geometry in ("1280x860", "1100x760"):
                root.geometry(geometry)
                root.update()
                buttons = [w for w in descendants(root) if w.winfo_class() == "Button"
                           and w.cget("text") == "Download calculation sheet"]
                self.assertEqual(len(buttons), 4)
                for button in buttons:
                    self.assertTrue(button.winfo_ismapped())
                    self.assertGreaterEqual(button.winfo_height(), button.winfo_reqheight())
                    self.assertLessEqual(button.winfo_rooty() + button.winfo_height(),
                                         root.winfo_rooty() + root.winfo_height())
                    # Both actions must fit fully inside their card.
                    self.assertGreaterEqual(button.master.winfo_height(), button.master.winfo_reqheight())
                    self.assertGreaterEqual(button.winfo_width(), button.winfo_reqwidth())
                    self.assertGreaterEqual(button.master.winfo_width(), button.master.winfo_reqwidth())

            buttons = [w for w in descendants(root) if w.winfo_class() == "Button"
                       and w.cget("text") == "Download calculation sheet"]
            for button, key in zip(buttons, WIZARDS):
                dest = self.folder / f"home_{key}.xlsx"
                with patch.object(app.filedialog, "asksaveasfilename", return_value=str(dest)), \
                        patch.object(app.os, "startfile", create=True) as opened:
                    button.invoke()
                    self.assertEqual(dest.read_bytes(), calculation_template_path(key).read_bytes())
                    opened.assert_called_once_with(str(dest.resolve()))
            with patch.object(app.filedialog, "asksaveasfilename", return_value=""), \
                    patch.object(app, "save_calculation_sheet") as copy:
                buttons[0].invoke()
                copy.assert_not_called()

            for key, classname in WIZARDS.items():
                root.open_wizard(key)
                root.update()
                wizard = next(w for w in descendants(root) if type(w).__name__ == classname)
                workbook = wizard.vars["excel"].get() if key == "mls" else wizard.files["workbook"]
                self.assertEqual(workbook, str(calculation_template_path(key)))
                selections = ({k: v.get() for k, v in wizard.vars.items()}
                              if key == "mls" else wizard.files)
                document_templates = {k: Path(v) for k, v in selections.items()
                                      if k == "template" or k.endswith("_template")}
                self.assertTrue(document_templates)
                for actual in document_templates.values():
                    self.assertEqual(actual.parent, templates_dir(key))
                    self.assertTrue(actual.is_file())
                manifest = WF / "templates" / "manifest.json"
                if manifest.exists():
                    for role in json.loads(manifest.read_text())["wizards"][key]["roles"]:
                        self.assertEqual(document_templates[role["key"]],
                                         templates_dir(key) / role["file"])
                root.go_home()
            error.assert_not_called()


if __name__ == "__main__":
    unittest.main(verbosity=2)
