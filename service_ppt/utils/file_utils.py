"""File utility functions for service_ppt.

This module provides utility functions for file and directory operations.
"""

import errno
import os
from pathlib import Path
import re
import shutil


def rename_filename_to_zeropadded(dirname: str, num_digits: int) -> None:
    r"""Rename Slide(\d+).PNG to Slide%03d.png so that the length is same
    and they can be sorted properly.
    Powerpoint generates filename as 1-9 and 10, etc. that make sorting difficult.

    :param dirname: Directory containing files to rename
    :param num_digits: Number of digits to use for zero-padding
    """
    if num_digits <= 1:
        return

    def replace_format_3digits(m: re.Match[str]) -> str:
        if m.group(2).isdigit():
            num = int(m.group(2))
            fmt = r"%s%0" + str(num_digits) + r"d%s"

            ext = m.group(3)
            ext = ext.lower()
            if len(ext) == 5:
                if ext == ".jpeg":
                    ext = ".jpg"
                elif ext == ".tiff":
                    ext = ".tif"

            s = fmt % (m.group(1), num, ext)
        else:
            s = m.group(0)
        return s

    matching_fn_pattern = r"(.*?)(\d+)(\.(gif|jpg|jpeg|png|tif|tiff))"
    repl_func = replace_format_3digits

    fn_re = re.compile(matching_fn_pattern, re.IGNORECASE)

    files = os.listdir(dirname)
    for fn in files:
        new_fn = fn_re.sub(repl_func, fn)
        if new_fn == fn:
            continue

        old_fullname = os.path.join(dirname, fn)
        new_fullname = os.path.join(dirname, new_fn)
        os.rename(old_fullname, new_fullname)


def rmtree_except_self(dirname: str) -> None:
    """Remove all files and subdirectories in a directory, but not the directory itself.

    :param dirname: Directory to clean
    """
    try:
        files = os.listdir(dirname)
        for fn in files:
            fullname = os.path.join(dirname, fn)
            if os.path.isdir(fullname):
                shutil.rmtree(fullname)
            else:
                os.remove(fullname)
    except OSError as err:
        if err.errno == errno.ENOENT:
            return

        raise


def mkdir_if_not_exists(dirname: str) -> None:
    """Create a directory if it does not exist.

    :param dirname: Directory path to create
    """
    try:
        os.mkdir(dirname)
    except OSError as err:
        if err.errno == errno.EEXIST:
            return

        raise


def get_package_dir() -> Path:
    """Get the directory containing the service_ppt package.

    :returns: Path to the service_ppt package directory
    """
    # Go up two levels from utils/file_utils.py to get to service_ppt/
    return Path(__file__).parent.parent.absolute()


def get_image24_dir() -> str:
    """Get the image24 directory path.

    :returns: Path to the image24 directory as a string
    """
    return str(get_package_dir() / "image24")


def get_image32_dir() -> str:
    """Get the image32 directory path.

    :returns: Path to the image32 directory as a string
    """
    return str(get_package_dir() / "image32")


def get_locale_dir() -> Path:
    """Get the locale directory path.

    :returns: Path to the locale directory
    """
    return get_package_dir() / "locale"
