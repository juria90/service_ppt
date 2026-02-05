"""Internationalization support for service_ppt."""

import gettext
from pathlib import Path
from typing import TYPE_CHECKING

if TYPE_CHECKING:
    from gettext import NullTranslations


def _(s: str) -> str:
    """Return the string unchanged (default translation function).

    This is a no-op function used before translation is initialized.
    After initialization, this function is replaced with the actual translation function.

    :param s: String to translate
    :returns: The input string unchanged
    """
    return s


def ngettext(s1: str, s2: str, n: int) -> str:
    """Return singular or plural form based on count (default plural translation function).

    This is a no-op function used before translation is initialized.
    After initialization, this function is replaced with the actual translation function.

    :param s1: Singular form
    :param s2: Plural form
    :param n: Count to determine which form to use
    :returns: s1 if n == 1, otherwise s2
    """
    return s1 if n == 1 else s2


def pgettext(c: str, s: str) -> str:
    """Return the string unchanged (default context-aware translation function).

    This is a no-op function used before translation is initialized.
    After initialization, this function is replaced with the actual translation function.

    :param c: Translation context
    :param s: String to translate
    :returns: The input string unchanged
    """
    return s


_translation: "NullTranslations | None" = None


def initialize_translation(locale_dir: str | Path, domain: str = "service_ppt") -> None:
    """Initialize translation for the application.

    :param locale_dir: Directory containing translation files.
    :param domain: Translation domain name.
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

    :return: The current translation object, or None if not initialized.
    """
    return _translation
