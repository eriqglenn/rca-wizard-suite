"""A separate temporary workspace for each RCA packet attempt."""

from contextlib import contextmanager
from pathlib import Path
import shutil
import tempfile


@contextmanager
def packet_workspace():
    """Never reuse or delete an earlier packet's folder in the case directory.

    Only the newly created child of the system temp folder is eligible for
    cleanup. An Office/file-system lock during cleanup must not hide a useful
    conversion error or turn a completed RCA into a reported failure.
    """
    temp_root = Path(tempfile.gettempdir()).resolve()
    work = Path(tempfile.mkdtemp(prefix="rca_packet_", dir=temp_root)).resolve()
    try:
        yield work
    finally:
        try:
            if (work.parent == temp_root and work.resolve() == work
                    and work.name.startswith("rca_packet_") and not work.is_symlink()):
                shutil.rmtree(work)
        except OSError:
            # It is safe to leave this one temporary folder for Windows cleanup.
            # No future packet attempt depends on removing it.
            pass
