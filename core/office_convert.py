"""Office-to-PDF conversion and workbook recalculation with two engines.

The suite historically required LibreOffice for two jobs:

  * converting generated .docx / .xlsx files to PDF for the RCA packet
  * recalculating a workbook whose formula cache is empty ("poisoned")

City of Houston work computers ship with Microsoft Office but usually not
LibreOffice, and installing extra software there needs IT.  This module keeps
LibreOffice as the preferred engine when it is present and falls back to
Microsoft Word / Excel through COM automation (pywin32) when it is not, so the
suite works out of the box on a standard City PC.

Engine selection (environment variable ``RCA_OFFICE_ENGINE``):

    auto        (default) LibreOffice if installed, otherwise Microsoft Office
    libreoffice force LibreOffice
    msoffice    force Microsoft Word / Excel

Every public function is safe to call from a worker thread: the COM path
initialises and releases COM per call.
"""

from __future__ import annotations

import os
import subprocess
import sys
from pathlib import Path

from .soffice import find_soffice

ENGINE_ENV = "RCA_OFFICE_ENGINE"

_WORD_EXTS = {".docx", ".doc", ".docm", ".rtf"}
_EXCEL_EXTS = {".xlsx", ".xls", ".xlsm"}


# ---------------------------------------------------------------- discovery
def _preference() -> str:
    return (os.environ.get(ENGINE_ENV) or "auto").strip().lower()


def automation_support_error() -> str | None:
    """Explain a runtime problem separately from Office installation status."""
    if sys.platform != "win32":
        return "Microsoft Office automation requires Windows."
    try:
        import win32com.client  # noqa: F401
        import pythoncom  # noqa: F401
        return None
    except Exception as exc:
        return ("The wizard's Microsoft Office automation component could not load: "
                f"{type(exc).__name__}: {exc}. "
                "Use the complete Windows package and keep its _internal folder "
                "beside RCA Wizard Suite.exe. This does not mean Word is missing.")


def _pywin32_available() -> bool:
    return automation_support_error() is None


def _office_app_registered(exe_name: str) -> bool:
    """Read installation hints in both registry views; never gate conversion.

    Office can be activated through COM even when an App Paths entry is absent.
    These hints only describe the setup on the home screen.
    """
    if sys.platform != "win32":
        return False
    try:
        import winreg
    except Exception:
        return False
    progid = {"winword.exe": "Word.Application", "excel.exe": "Excel.Application"}[exe_name.lower()]
    views = (0, winreg.KEY_WOW64_64KEY, winreg.KEY_WOW64_32KEY)
    for view in views:
        try:
            with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, progid + r"\CLSID",
                                0, winreg.KEY_READ | view) as k:
                clsid, _ = winreg.QueryValueEx(k, None)
                if clsid:
                    return True
        except OSError:
            pass
        key = r"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths" + "\\" + exe_name
        for root in (winreg.HKEY_LOCAL_MACHINE, winreg.HKEY_CURRENT_USER):
            try:
                with winreg.OpenKey(root, key, 0, winreg.KEY_READ | view) as k:
                    path, _ = winreg.QueryValueEx(k, None)
                    if path and Path(os.path.expandvars(str(path).strip('"'))).is_file():
                        return True
            except OSError:
                continue
    return False


def word_available() -> bool:
    return _pywin32_available() and _office_app_registered("Winword.exe")


def excel_available() -> bool:
    return _pywin32_available() and _office_app_registered("Excel.exe")


def engines() -> list[str]:
    """Engines to try, in order, honouring ``RCA_OFFICE_ENGINE``."""
    pref = _preference()
    if pref in ("libreoffice", "soffice", "lo"):
        return ["libreoffice"]
    if pref in ("msoffice", "word", "office", "com"):
        return ["msoffice"]
    return ["libreoffice", "msoffice"]


def available() -> bool:
    """True if at least one engine can convert Word documents to PDF."""
    for eng in engines():
        if eng == "libreoffice" and find_soffice():
            return True
        if eng == "msoffice" and word_available():
            return True
    return False


def describe_engine() -> str:
    """Human-readable description of the engine that will be used."""
    for eng in engines():
        if eng == "libreoffice":
            so = find_soffice()
            if so:
                return f"LibreOffice ({so})"
        elif eng == "msoffice":
            if not _pywin32_available():
                return "Office support unavailable — use Check PDF setup"
            if word_available() and excel_available():
                return "Microsoft Word / Excel (detected)"
            return "Microsoft Office — use Check PDF setup"
    return "LibreOffice not found — use Check PDF setup"


def missing_message() -> str:
    return (
        "Use Check PDF setup on the home screen to test Word and Excel and "
        "copy the diagnostic report. Automatic PDF creation uses desktop Word "
        "for documents and desktop Excel for spreadsheets, or LibreOffice. "
        "Browser-only Office cannot perform this desktop conversion."
    )


# ---------------------------------------------------------------- engines
def _libreoffice_to_pdf(soffice: str, input_path: Path, output_dir: Path) -> Path:
    subprocess.run(
        [soffice, "--headless", "--convert-to", "pdf", "--outdir",
         str(output_dir), str(input_path)],
        check=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE,
    )
    pdf = output_dir / (input_path.stem + ".pdf")
    if not pdf.exists():
        raise RuntimeError(f"PDF conversion failed for {input_path}")
    return pdf


def _msoffice_to_pdf(input_path: Path, output_dir: Path) -> Path:
    support_error = automation_support_error()
    if support_error:
        raise RuntimeError(support_error)
    import pythoncom
    import win32com.client

    pdf = output_dir / (input_path.stem + ".pdf")
    if pdf.exists():
        try:
            pdf.unlink()
        except OSError:
            pass
    src = str(input_path.resolve())
    dst = str(pdf.resolve())
    suffix = input_path.suffix.lower()

    pythoncom.CoInitialize()
    try:
        if suffix in _WORD_EXTS:
            app = win32com.client.DispatchEx("Word.Application")
            try:
                app.Visible = False
                app.DisplayAlerts = 0
                doc = app.Documents.Open(src, ReadOnly=True, AddToRecentFiles=False,
                                         Visible=False)
                try:
                    # 17 = wdExportFormatPDF
                    doc.ExportAsFixedFormat(OutputFileName=dst, ExportFormat=17)
                finally:
                    doc.Close(False)
            finally:
                app.Quit()
        elif suffix in _EXCEL_EXTS:
            app = win32com.client.DispatchEx("Excel.Application")
            try:
                app.Visible = False
                app.DisplayAlerts = False
                wb = app.Workbooks.Open(src, UpdateLinks=0, ReadOnly=True)
                try:
                    # 0 = xlTypePDF
                    wb.ExportAsFixedFormat(0, dst)
                finally:
                    wb.Close(False)
            finally:
                app.Quit()
        else:
            raise ValueError(f"Unsupported Office file type: {input_path}")
    finally:
        pythoncom.CoUninitialize()

    if not pdf.exists():
        raise RuntimeError(f"PDF conversion failed for {input_path}")
    return pdf


# ---------------------------------------------------------------- public API
def office_to_pdf(input_path: str | Path, output_dir: str | Path) -> Path:
    """Convert a Word or Excel file to PDF in ``output_dir`` and return the
    PDF path.  Tries the engines from :func:`engines` in order and raises a
    single actionable error if none is usable."""
    input_path = Path(input_path)
    output_dir = Path(output_dir)
    if not input_path.is_file():
        raise FileNotFoundError(f"File to convert was not found: {input_path}")
    if input_path.suffix.lower() not in _WORD_EXTS | _EXCEL_EXTS:
        raise ValueError(f"Unsupported Office file type: {input_path}")
    output_dir.mkdir(parents=True, exist_ok=True)

    errors: list[str] = []
    for eng in engines():
        if eng == "libreoffice":
            soffice = find_soffice()
            if not soffice:
                errors.append("LibreOffice: not found")
                continue
            try:
                return _libreoffice_to_pdf(soffice, input_path, output_dir)
            except Exception as e:  # fall through to the next engine
                errors.append(f"LibreOffice: {e}")
        elif eng == "msoffice":
            # The only definitive check is activation/export. Registry hints
            # may be missing, and an Excel export must not depend on Word.
            try:
                return _msoffice_to_pdf(input_path, output_dir)
            except Exception as e:
                app_name = "Excel" if input_path.suffix.lower() in _EXCEL_EXTS else "Word"
                errors.append(f"Microsoft {app_name}: {type(e).__name__}: {e}")

    if errors:
        raise RuntimeError(
            f"PDF conversion failed for {input_path.name}:\n  " + "\n  ".join(errors)
            + f"\n\nOffice engine preference: {_preference()}\n\n" + missing_message())
    raise RuntimeError(missing_message())


def recalc_with_excel(src: str | Path, dest: str | Path) -> Path | None:
    """Open ``src`` in Excel, force a full recalculation, and save the result
    as ``dest`` (.xlsx).  Returns ``dest`` on success or None if Excel is not
    available or the recalc fails.  Never raises."""
    # Do not reject a working Excel COM server because a registry hint is absent.
    src = Path(src).resolve()
    dest = Path(dest).resolve()
    try:
        import pythoncom
        import win32com.client

        pythoncom.CoInitialize()
        try:
            app = win32com.client.DispatchEx("Excel.Application")
            try:
                app.Visible = False
                app.DisplayAlerts = False
                wb = app.Workbooks.Open(str(src), UpdateLinks=0)
                try:
                    app.CalculateFullRebuild()
                    if dest.exists():
                        dest.unlink()
                    # 51 = xlOpenXMLWorkbook (.xlsx)
                    wb.SaveAs(str(dest), FileFormat=51)
                finally:
                    wb.Close(False)
            finally:
                app.Quit()
        finally:
            pythoncom.CoUninitialize()
    except Exception:
        return None
    return dest if dest.exists() else None
