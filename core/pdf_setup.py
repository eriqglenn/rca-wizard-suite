"""Test actual desktop Office PDF exports without using a planner's case."""

from __future__ import annotations

import argparse
from datetime import datetime
from pathlib import Path
import queue
import subprocess
import sys
import tempfile
import threading

from .embedding import SUITE_VERSION
from . import office_convert


def _check_command(app_name: str, folder: Path) -> list[str]:
    if getattr(sys, "frozen", False):
        command = [str(Path(sys.executable).parent / "RCA Wizard Tools.exe"), "pdfcheck"]
    else:
        command = [sys.executable, str(Path(__file__).resolve().parents[1]
                                      / "tools" / "check_pdf_setup.py")]
    return command + ["--app", app_name, "--out", str(folder)]


def check_export(app_name: str, folder: Path) -> int:
    """Create one harmless sample, export through Office, and read the PDF."""
    from pypdf import PdfReader

    folder.mkdir(parents=True, exist_ok=True)
    if app_name == "word":
        from docx import Document
        source = folder / "Word_PDF_test.docx"
        document = Document()
        document.add_paragraph("RCA Wizard Suite - Word PDF conversion test")
        document.save(source)
    elif app_name == "excel":
        from openpyxl import Workbook
        source = folder / "Excel_PDF_test.xlsx"
        workbook = Workbook()
        sheet = workbook.active
        sheet["A1"] = "RCA Wizard Suite - Excel PDF conversion test"
        sheet.column_dimensions["A"].width = 60
        sheet.print_area = "A1:A1"
        workbook.save(source)
        workbook.close()
    else:
        raise ValueError(f"Unknown Office application: {app_name}")
    # Exercise the Office path even on a computer that also has LibreOffice.
    result = office_convert._msoffice_to_pdf(source, folder)
    with result.open("rb") as stream:
        pages = len(PdfReader(stream).pages)
    if pages < 1:
        raise RuntimeError("Office created a PDF with no pages.")
    return pages


def run_report() -> str:
    """Use isolated workers so an Office startup problem does not freeze Tk."""
    mode = "packaged EXE" if getattr(sys, "frozen", False) else "Python source launcher"
    lines = [f"RCA Wizard Suite {SUITE_VERSION} - PDF setup report",
             datetime.now().isoformat(timespec="seconds"),
             f"Launch mode: {mode}", f"Runtime: {sys.executable}",
             f"Office engine preference: {office_convert._preference()}",
             f"LibreOffice: {office_convert.find_soffice() or 'not found'}",
             f"Office automation component: {office_convert.automation_support_error() or 'loaded'}",
             "", "Tests below use desktop Word and Excel directly.",
             "They create temporary samples; no case files are opened.", ""]
    passed = 0
    with tempfile.TemporaryDirectory(prefix="rca_pdf_setup_", ignore_cleanup_errors=True) as temporary:
        for app_name in ("word", "excel"):
            try:
                result = subprocess.run(
                    _check_command(app_name, Path(temporary) / app_name),
                    capture_output=True, text=True, encoding="utf-8", errors="replace",
                    timeout=60,
                    creationflags=subprocess.CREATE_NO_WINDOW if sys.platform == "win32" else 0,
                )
                output = (result.stdout + "\n" + result.stderr).strip()
                lines.append(output or f"FAIL: {app_name.title()} worker returned {result.returncode} without a report.")
                if result.returncode == 0 and output.startswith("PASS:"):
                    passed += 1
            except subprocess.TimeoutExpired:
                lines.append(f"FAIL: {app_name.title()} did not finish within 60 seconds. "
                             "Open the desktop app manually and check for a setup or sign-in prompt.")
            except Exception as exc:
                lines.append(f"FAIL: {app_name.title()} check could not run: {type(exc).__name__}: {exc}")
            lines.append("")
    if passed == 2:
        lines.append("Both desktop Office PDF exports passed. Try generating the RCA again. "
                     "A remaining error may involve the case's particular document or attachment.")
    else:
        lines.append("Copy this report when asking for help. Word handles document PDFs; "
                     "Excel handles spreadsheet PDFs. Browser-only Office cannot perform these "
                     "desktop exports. This report distinguishes a missing automation component "
                     "from an Office startup/export error.")
    return "\n".join(lines)


def show_dialog(parent):
    import tkinter as tk
    from tkinter import filedialog, messagebox, scrolledtext

    window = tk.Toplevel(parent)
    window.title("Check PDF setup")
    window.geometry("820x580")
    window.minsize(620, 420)
    tk.Label(window, text="Testing desktop Word and Excel PDF export",
             font=("Segoe UI", 13, "bold"), anchor="w").pack(fill="x", padx=16, pady=(16, 8))
    output = scrolledtext.ScrolledText(window, wrap="word", font=("Consolas", 10))
    output.pack(fill="both", expand=True, padx=16, pady=8)
    output.insert("1.0", "Checking with temporary sample files. This can take up to two minutes.\n")
    output.configure(state="disabled")
    actions = tk.Frame(window)
    actions.pack(fill="x", padx=16, pady=(0, 16))

    def copy_report():
        window.clipboard_clear()
        window.clipboard_append(output.get("1.0", "end-1c"))

    def save_report():
        destination = filedialog.asksaveasfilename(parent=window, title="Save PDF setup report",
                    initialfile="RCA_PDF_setup_report.txt", defaultextension=".txt",
                    filetypes=[("Text report", "*.txt")])
        if destination:
            try:
                Path(destination).write_text(output.get("1.0", "end-1c"), encoding="utf-8")
            except OSError as exc:
                messagebox.showerror("Save report", str(exc), parent=window)

    copy_button = tk.Button(actions, text="Copy report", command=copy_report, state="disabled")
    copy_button.pack(side="left", padx=(0, 8))
    save_button = tk.Button(actions, text="Save report", command=save_report, state="disabled")
    save_button.pack(side="left")
    tk.Button(actions, text="Close", command=window.destroy).pack(side="right")
    results = queue.Queue()

    def worker():
        try:
            results.put(run_report())
        except Exception as exc:
            results.put(f"PDF setup check failed: {type(exc).__name__}: {exc}")

    def poll():
        if not window.winfo_exists():
            return
        try:
            report = results.get_nowait()
        except queue.Empty:
            window.after(100, poll)
            return
        output.configure(state="normal")
        output.delete("1.0", "end")
        output.insert("1.0", report)
        output.configure(state="disabled")
        copy_button.configure(state="normal")
        save_button.configure(state="normal")

    threading.Thread(target=worker, daemon=True).start()
    window.after(100, poll)
    return window


def main() -> int:
    if sys.stdout is not None and hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")
    parser = argparse.ArgumentParser(description="Test desktop Office PDF export")
    parser.add_argument("--app", required=True, choices=("word", "excel"))
    parser.add_argument("--out", required=True, type=Path)
    args = parser.parse_args()
    try:
        pages = check_export(args.app, args.out)
        print(f"PASS: {args.app.title()} created a readable PDF ({pages} page(s)).", flush=True)
        return 0
    except Exception as exc:
        print(f"FAIL: {args.app.title()}: {type(exc).__name__}: {exc}", flush=True)
        return 1
