"""Atomic file writing utility.

This module provides a context manager for writing files atomically, ensuring that
either the complete file is written or no file is created (preventing partial writes).
"""

import os
import shutil
import tempfile
from typing import TYPE_CHECKING, Any, TextIO

if TYPE_CHECKING:
    from types import TracebackType


class AtomicFileWriter:
    """Context manager for atomic file writing.

    Ensures that either the complete file is written or no file is created,
    preventing partial writes that could corrupt data.
    """

    def __init__(self, filename: str | os.PathLike[str], *args: Any, **kwargs: Any) -> None:
        """Initialize atomic file writer.

        :param filename: Target filename to write to
        :param args: Additional arguments passed to os.fdopen
        :param kwargs: Additional keyword arguments passed to os.fdopen
        """
        self.file: TextIO | None = None
        self.filename: str | os.PathLike[str] | None = None
        self.tmpfilename: str | None = None

        temp_fd, tmpfilename = tempfile.mkstemp()
        self.file = os.fdopen(temp_fd, *args, **kwargs)
        self.filename = filename

        self.tmpfilename = tmpfilename

    def __enter__(self) -> TextIO:
        """Enter context manager.

        :returns: File object for writing
        """
        return self.file

    def __exit__(
        self,
        exc_type: type[BaseException] | None,
        exc_val: BaseException | None,
        exc_tb: "TracebackType | None",
    ) -> None:
        """Exit context manager and atomically move file to final location.

        :param exc_type: Exception type if exception occurred
        :param exc_val: Exception value if exception occurred
        :param exc_tb: Exception traceback if exception occurred
        """
        # make sure that all data is on disk
        # see http://stackoverflow.com/questions/7433057/is-rename-without-fsync-safe
        self.file.flush()
        os.fsync(self.file.fileno())
        self.file.close()

        try:
            os.rename(self.tmpfilename, self.filename)
        except OSError:
            shutil.move(self.tmpfilename, self.filename)
