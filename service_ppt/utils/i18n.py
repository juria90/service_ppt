"""Internationalization support for service_ppt."""

import gettext
from pathlib import Path
from typing import TYPE_CHECKING

if TYPE_CHECKING:
    from gettext import NullTranslations


def _(s: str) -> str:
    """Default translation function (no-op)."""
    return s


def ngettext(s1: str, s2: str, n: int) -> str:
    """Default plural translation function (no-op)."""
    return s1 if n == 1 else s2


def pgettext(c: str, s: str) -> str:
    """Default context-aware translation function (no-op)."""
    return s


_translation: "NullTranslations | None" = None


def initialize_translation(locale_dir: str | Path, domain: str = "service_ppt") -> None:
    """Initialize translation for the application.

    Args:
        locale_dir: Directory containing translation files.
        domain: Translation domain name.
    """
    global _translation, _, ngettext, pgettext

    trans = gettext.translation(domain, localedir=str(locale_dir), fallback=True)
    _translation = trans
    _ = trans.gettext
    ngettext = trans.ngettext
    # Use pgettext if available, otherwise fall back to gettext
    pgettext = getattr(trans, "pgettext", lambda c, s: trans.gettext(s))


def get_translation() -> "NullTranslations | None":
    """Get the current translation object.

    Returns:
        The current translation object, or None if not initialized.
    """
    return _translation
