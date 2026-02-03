"""Service PowerPoint presentation generator.

A cross-platform application for generating PowerPoint slides for church services,
supporting Bible verses, lyrics, announcements, and various slide operations.
"""

from pathlib import Path


def get_package_dir():
    """Get the directory containing the service_ppt package."""
    return Path(__file__).parent.absolute()


def get_image24_dir():
    """Get the image24 directory path."""
    return get_package_dir() / "image24"


def get_image32_dir():
    """Get the image32 directory path."""
    return get_package_dir() / "image32"


def get_locale_dir():
    """Get the locale directory path."""
    return get_package_dir() / "locale"
