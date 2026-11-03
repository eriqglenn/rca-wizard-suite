"""Build the self-contained Windows package of the RCA Wizard Suite.

Result: a zip file that planners extract anywhere (Documents, Desktop, a
shared drive) and run by double-clicking "RCA Wizard Suite.exe".  Nothing
else is required on the target computer:

    * no Python install, no pip, no internet, no administrator rights
    * PDF packets use Microsoft Word/Excel when LibreOffice is absent
    * the HCAD downloader drives the computer's own Google Chrome

Run from the suite folder with the Python that has the suite's packages and
PyInstaller installed (see "Build Package.bat"):

    py -B build\\build_package.py            # full build + zip
    py -B build\\build_package.py --no-zip   # folder only

Build work stays centrally under Maintenance/Builds/<suite>/package_current. Cleanup checks
resolved paths and cannot remove another version's build outputs.
"""

from __future__ import annotations

import argparse
import hashlib
import zipfile
import os
import stat
import re
import shutil
import subprocess
import sys
import time
from pathlib import Path

BUILD_DIR = Path(__file__).resolve().parent
SUITE_DIR = BUILD_DIR.parent
WF = SUITE_DIR / "Wizard_Files"
SPEC = BUILD_DIR / "rca_wizard_suite.spec"
ICON = BUILD_DIR / "rca.ico"

MAINTENANCE_ROOT = SUITE_DIR.parent / "Maintenance"
WORK_ROOT = MAINTENANCE_ROOT / "Builds" / SUITE_DIR.name / "package_current"
# Candidates stay separate from tested portable releases.
OUT_DIR = MAINTENANCE_ROOT / "Builds" / SUITE_DIR.name / "Candidate ZIP"

#: Folders copied next to the executables as the editable Wizard_Files tree.
EXTERNAL_DIRS = ("templates", "sample_input", "docs", "assets")
EXTERNAL_FILES = ("README.md",)
SKIP_DIR_NAMES = {"__pycache__"}


def remove_tree(path: Path) -> None:
    """Delete a folder, clearing read-only bits and retrying briefly.  Fails
    loudly if a file is still locked (an exe from the previous package still
    running is the usual cause) instead of leaving a half-deleted folder."""
    resolved = path.resolve()
    if not resolved.is_relative_to(WORK_ROOT.resolve()) or resolved == WORK_ROOT.resolve():
        raise ValueError(f'Refusing to remove a path outside this version build directory: {resolved}')
    if not path.exists():
        return

    def on_error(func, target, _exc):
        try:
            os.chmod(target, stat.S_IWRITE)
            func(target)
        except Exception:
            pass

    for attempt in range(5):
        try:
            shutil.rmtree(path, onexc=on_error)          # Python 3.12+
        except TypeError:
            shutil.rmtree(path, onerror=on_error)        # older Pythons
        if not path.exists():
            return
        time.sleep(1 + attempt)
    sys.exit(
        f"Could not delete {path}. "
        "Close any running 'RCA Wizard Suite.exe' / 'RCA Wizard Tools.exe' "
        "(and Explorer windows open in that folder), then run the build again."
    )


def suite_version() -> str:
    text = (WF / "core" / "embedding.py").read_text(encoding="utf-8")
    m = re.search(r'SUITE_VERSION\s*=\s*"([^"]+)"', text)
    return m.group(1) if m else "dev"


def ensure_icon() -> None:
    """Draw a simple icon so the exe and shortcut are recognisable."""
    if ICON.exists():
        return
    try:
        from PIL import Image, ImageDraw, ImageFont
    except Exception:
        return
    sizes = [256, 128, 64, 48, 32, 16]
    base = Image.new("RGBA", (256, 256), (0, 0, 0, 0))
    d = ImageDraw.Draw(base)
    d.rounded_rectangle((8, 8, 248, 248), radius=48, fill=(30, 58, 138, 255))
    d.rounded_rectangle((40, 56, 216, 200), radius=12, fill=(255, 255, 255, 255))
    for y in (88, 116, 144, 172):
        d.rectangle((60, y, 196, y + 10), fill=(203, 213, 225, 255))
    d.rectangle((60, 88, 130, 98), fill=(37, 99, 235, 255))
    try:
        font = ImageFont.truetype("arialbd.ttf", 54)
    except Exception:
        font = ImageFont.load_default()
    d.text((128, 226), "RCA", fill=(255, 255, 255, 255), font=font, anchor="mm")
    base.save(ICON, format="ICO", sizes=[(s, s) for s in sizes])


def check_environment() -> None:
    missing = []
    for mod in ("PyInstaller", "docx", "openpyxl", "pypdf", "PIL", "fitz", "lxml",
                "playwright", "win32com", "reportlab", "shapely", "numpy"):
        try:
            __import__(mod)
        except Exception:
            missing.append(mod)
    if missing:
        sys.exit(
            "This Python is missing packages needed for the build: "
            + ", ".join(missing)
            + "\nInstall them with:\n  "
            + f"{sys.executable} -m pip install pyinstaller pywin32 -r \"{WF / 'requirements.txt'}\""
        )


def run_pyinstaller(dist: Path, work: Path) -> Path:
    cmd = [sys.executable, "-m", "PyInstaller", "--noconfirm", "--clean",
           "--distpath", str(dist), "--workpath", str(work), str(SPEC)]
    print("\n> " + " ".join(f'"{c}"' if " " in c else c for c in cmd) + "\n")
    subprocess.run(cmd, check=True, cwd=str(SUITE_DIR))
    folder = dist / ('RCA_Wizard_Suite_' + suite_version())
    if not (folder / "RCA Wizard Suite.exe").exists():
        sys.exit("PyInstaller finished but 'RCA Wizard Suite.exe' was not produced.")
    return folder


def copy_tree(src: Path, dst: Path) -> None:
    def ignore(_dir, names):
        return [n for n in names if n in SKIP_DIR_NAMES or n.endswith((".pyc", ".tmp"))]
    shutil.copytree(src, dst, ignore=ignore, dirs_exist_ok=True)


def add_external_files(folder: Path, version: str) -> None:
    ext = folder / "Wizard_Files"
    ext.mkdir(exist_ok=True)
    for name in EXTERNAL_DIRS:
        if (WF / name).exists():
            copy_tree(WF / name, ext / name)
    for name in EXTERNAL_FILES:
        if (WF / name).exists():
            shutil.copy2(WF / name, ext / name)
    # Both launch methods are shipped in the portable ZIP, not just development folders.
    for name in ('Run Wizard.bat', 'Check Startup.bat'):
        shutil.copy2(SUITE_DIR / name, folder / name)
    # Suite-level docs travel with the package.
    for name in ("INSTALL_CITY_COMPUTERS.md", "TEMPLATE_UPDATE_GUIDE.md"):
        src = WF / "docs" / name
        if src.exists():
            shutil.copy2(src, folder / name)

    (folder / "README_FIRST.txt").write_text(
        f"City Planning RCA Wizard Suite {version} - packaged build\n"
        "=========================================================\n\n"
        "1. Keep this whole folder together (do not move the .exe files out of it).\n"
        "2. Double-click \"Run Wizard.bat\" to start.\n"
        "3. Optional: double-click \"Create Desktop Shortcut.bat\" for a Desktop icon.\n\n"
        "Nothing needs to be installed: no Python, no extra downloads, no admin rights.\n"
        "PDF packets use Microsoft Word/Excel on this computer (or LibreOffice if present).\n"
        "The HCAD downloader uses the Google Chrome already on this computer.\n\n"
        "Create Boundary Map is in Required setup in all four wizards. It uses the saved\n"
        "calculation sheet and public parcel/street data. Internet access is needed for\n"
        "new data; saved data can be reused. Review the preview, then choose Use This Map.\n\n"
        "NEW: In all four wizards' Required setup, use Import Application PDF / Fill\n"
        "Calculation Sheet. Review the application, select HCAD parcels and save a new\n"
        "workbook. See Wizard_Files\\docs\\APPLICATION_IMPORT.md for the workflow.\n\n"
        "Templates: Wizard_Files\\templates\\<wizard>\\ - or use Manage Templates on the\n"
        "home screen, which also runs the built-in check.\n\n"
        "Before generating an RCA, click Check PDF setup on the home screen to test\n"
        "desktop Word and Excel. Copy or save the report if either test fails.\n\n"
        "If Windows SmartScreen shows 'Windows protected your PC', click 'More info'\n"
        "then 'Run anyway' (the program is not code-signed). If the .exe is blocked\n"
        "outright by policy, ask IT to allow this folder or see INSTALL_CITY_COMPUTERS.md.\n",
        encoding="utf-8",
    )
    with (folder / 'README_FIRST.txt').open('a', encoding='utf-8') as readme:
        readme.write("\nSTARTUP HELP: Copy the entire extracted folder, including _internal and Wizard_Files.\n"
                     "Run Wizard.bat keeps startup errors visible. Check Startup.bat checks the actual home\n"
                     "screen and writes a report in your Windows temporary folder. Python is bundled.\n")
    (folder / "Create Desktop Shortcut.bat").write_text(
        "@echo off\r\n"
        "REM Creates a Desktop shortcut to RCA Wizard Suite.exe (no admin rights needed).\r\n"
        "set \"TARGET=%~dp0RCA Wizard Suite.exe\"\r\n"
        "powershell -NoProfile -ExecutionPolicy Bypass -Command "
        "\"$d=[Environment]::GetFolderPath('Desktop');"
        "$s=(New-Object -ComObject WScript.Shell).CreateShortcut($d+'\\RCA Wizard Suite.lnk');"
        "$s.TargetPath='%TARGET%';$s.WorkingDirectory='%~dp0';$s.IconLocation='%TARGET%';"
        "$s.Description='City Planning RCA Wizard Suite';$s.Save()\"\r\n"
        "echo Shortcut created on your Desktop.\r\n"
        "pause\r\n",
        encoding="utf-8",
    )


def main() -> int:
    parser = argparse.ArgumentParser(description="Build the RCA Wizard Suite package")
    parser.add_argument("--no-zip", action="store_true", help="skip the zip step")
    args = parser.parse_args()

    check_environment()
    ensure_icon()
    version = suite_version()
    tag = version.lstrip("v")
    t0 = time.time()

    # This candidate is separate from release/, which the planner launches.
    # Reuse one bounded work directory; successful builds remove it below.
    work = WORK_ROOT / 'work'
    remove_tree(work)
    dist = WORK_ROOT / "dist"
    remove_tree(dist)
    dist.mkdir(parents=True, exist_ok=True)

    folder = run_pyinstaller(dist, work)
    add_external_files(folder, version)

    final_name = f"RCA_Wizard_Suite_v{tag}"
    # PyInstaller creates the final versioned name directly. Renaming a newly
    # indexed executable directory can fail in a OneDrive-synced workspace.
    final = folder

    print(f"\nPackage folder: {final}")
    if not args.no_zip:
        OUT_DIR.mkdir(parents=True, exist_ok=True)
        zip_base = OUT_DIR / f"{final_name}_Windows"
        z = Path(str(zip_base) + ".zip")   # not with_suffix: the name contains "v1.11"
        print("Zipping (this takes a minute)...")
        # Finish and verify the replacement before touching a working ZIP.
        pending_base = OUT_DIR / (final_name + '_pending')
        pending = Path(shutil.make_archive(str(pending_base), "zip", root_dir=str(dist), base_dir=final_name))
        with zipfile.ZipFile(pending) as archive:
            if archive.testzip() is not None:
                raise RuntimeError('New package ZIP failed its integrity check.')
        if z.exists():
            rollback = MAINTENANCE_ROOT / 'Recovery' / SUITE_DIR.name / 'Previous_candidate'
            rollback.mkdir(parents=True, exist_ok=True)
            previous = rollback / 'Previous_Windows.zip'
            copy = rollback / 'Previous_Windows.pending'
            shutil.copy2(z, copy)
            with z.open('rb') as source, copy.open('rb') as target:
                if hashlib.file_digest(source, 'sha256').digest() != hashlib.file_digest(target, 'sha256').digest():
                    raise RuntimeError('Rollback ZIP verification failed; current ZIP preserved.')
            os.replace(copy, previous)
        os.replace(pending, z)
        print(f"Zip: {z}  ({z.stat().st_size / 1e6:.0f} MB)")
    remove_tree(work)
    print("Active program folder is unchanged. Test this candidate before deploying it.")
    print(f"Done in {time.time() - t0:.0f}s")
    return 0


if __name__ == "__main__":
    sys.exit(main())
