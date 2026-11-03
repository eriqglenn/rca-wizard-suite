"""Early startup reporting; imports only the standard library until logging is ready."""
from __future__ import annotations

import faulthandler
import os
import sys
import tempfile
import traceback
from datetime import datetime
from pathlib import Path


def log_file() -> Path:
    requested = os.environ.get('RCA_STARTUP_LOG')
    if requested:
        return Path(requested)
    base = Path(os.environ.get('LOCALAPPDATA') or tempfile.gettempdir())
    return base / 'RCAWizardSuite' / 'Logs' / 'startup.log'


def check_resources() -> None:
    """Report an incomplete copied folder before displaying a partial home screen."""
    if not getattr(sys, 'frozen', False):
        return
    folder = Path(sys.executable).resolve().parent
    missing = []
    for name in ('mls', 'mbl', 'mls_renewal', 'mbl_renewal'):
        for kind in ('templates', 'sample_input'):
            path = folder / 'Wizard_Files' / kind / name
            if not path.is_dir():
                missing.append(str(path.relative_to(folder)))
    if missing:
        raise RuntimeError('This wizard folder is incomplete. Copy or extract the complete Windows package. '
                           'Missing folders: ' + ', '.join(missing))


def show_failure(message: str) -> None:
    # This also works when tkinter itself failed to import. Batch launches keep
    # their error window open instead, and diagnostic checks must not hang.
    if os.environ.get('RCA_LAUNCHED_FROM_BATCH') == '1' or '--startup-check' in sys.argv:
        return
    if sys.platform == 'win32':
        try:
            import ctypes
            ctypes.windll.user32.MessageBoxW(None, message, 'Wizard could not start', 0x10)
        except Exception:
            pass


def launch() -> int:
    stream = None
    original_out, original_err = sys.stdout, sys.stderr
    report = log_file()
    try:
        try:
            report.parent.mkdir(parents=True, exist_ok=True)
            # A bounded log and one previous log, not an accumulating build log.
            if report.exists() and report.stat().st_size > 1024 * 1024:
                report.replace(report.with_suffix('.previous.log'))
            stream = report.open('a', encoding='utf-8', buffering=1)
            stream.write(f'\n[{datetime.now().isoformat(timespec="seconds")}] Starting wizard\n'
                         f'Executable: {sys.executable}\nPython: {sys.version}\n'
                         f'Frozen: {getattr(sys, "frozen", False)}\nWorking folder: {Path.cwd()}\n')
        except OSError:
            # A read-only log location must never stop a usable wizard.
            report = Path(tempfile.gettempdir()) / ('RCA_startup_' + str(os.getpid()) + '.log')
            stream = report.open('a', encoding='utf-8', buffering=1)
        if sys.stdout is None:
            sys.stdout = stream
        if sys.stderr is None:
            sys.stderr = stream
        try:
            faulthandler.enable(file=stream, all_threads=True)
        except (OSError, RuntimeError):
            pass
        check_resources()
        import app
        app.main()
        stream.write('Wizard exited normally.\n')
        return 0
    except SystemExit as exc:
        return int(exc.code or 0) if isinstance(exc.code, (int, type(None))) else 1
    except Exception:
        details = traceback.format_exc()
        if stream is not None:
            stream.write(details)
        if original_err is not None:
            original_err.write(details)
        show_failure('The wizard could not start.\n\n' + details.splitlines()[-1] +
                     '\n\nStartup report:\n' + str(report))
        return 1
    finally:
        try:
            faulthandler.disable()
        except Exception:
            pass
        sys.stdout, sys.stderr = original_out, original_err
        if stream is not None:
            stream.close()
