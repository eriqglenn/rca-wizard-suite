"""Centralized error logging for the RCA Wizard Suite.

When a generator fails, the dialog the planner sees only carries the message
text.  This module also appends the full traceback, timestamp, and context to
a plain-text log in the case output folder, so a reported problem can be
diagnosed from the log file instead of a screen share.

The logger must never raise: a logging failure must not mask or replace the
original error shown to the planner.
"""

from __future__ import annotations

import sys
import traceback

from .embedding import SUITE_VERSION
from datetime import datetime
from pathlib import Path

LOG_NAME = "wizard_error_log.txt"


def log_exception(wizard_label: str, output_folder: str | Path | None,
                  exc: BaseException | None = None,
                  context: str = "") -> Path | None:
    """Append the active exception's traceback to the output-folder log.

    Call from inside an ``except`` block.  Returns the log path on success or
    ``None`` if logging was not possible (the caller should still show the
    original error to the planner either way).
    """
    try:
        folder = Path(output_folder) if output_folder else (Path.home() / "Documents")
        folder.mkdir(parents=True, exist_ok=True)
        log_path = folder / LOG_NAME
        stamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        tb = traceback.format_exc()
        if tb.strip() == "NoneType: None" and exc is not None:
            tb = "".join(traceback.format_exception(type(exc), exc, exc.__traceback__))
        with open(log_path, "a", encoding="utf-8") as f:
            f.write(
                f"\n{'=' * 70}\n"
                f"[{stamp}] {wizard_label}"
                + (f" — {context}" if context else "")
                + "\n"
                f"{'-' * 70}\n"
                f"Suite version: {SUITE_VERSION}\n"
                f"Launch mode: {'packaged EXE' if getattr(sys, 'frozen', False) else 'Python source launcher'}\n"
                f"Runtime: {sys.executable}\n\n"
                f"{tb}\n"
            )
        return log_path
    except Exception:
        return None
