"""Atomic file writing utility.

This module provides a context manager for writing files atomically, ensuring that
either the complete file is written or no file is created (preventing partial writes).
"""

import os
import shutil
import tempfile


class AtomicFileWriter:
    def __init__(self, filename, *args, **kwargs):
        self.file = None
        self.filename = None
        self.tmpfilename = None

        temp_fd, tmpfilename = tempfile.mkstemp()
        self.file = os.fdopen(temp_fd, *args, **kwargs)
        self.filename = filename

        self.tmpfilename = tmpfilename

    def __enter__(self):
        return self.file

    def __exit__(self, type, value, traceback):
        # make sure that all data is on disk
        # see http://stackoverflow.com/questions/7433057/is-rename-without-fsync-safe
        self.file.flush()
        os.fsync(self.file.fileno())
        self.file.close()

        try:
            os.rename(self.tmpfilename, self.filename)
        except OSError:
            shutil.move(self.tmpfilename, self.filename)
