"""Tests for i18n module.

This module contains unit tests for internationalization functions.
"""

import tempfile

from service_ppt.utils.i18n import (
    _,
    get_translation,
    initialize_translation,
    ngettext,
    pgettext,
)


class TestDefaultTranslationFunctions:
    """Test default translation functions before initialization."""

    def test_default_gettext_returns_unchanged(self):
        """Test that default _ function returns string unchanged."""
        assert _("Hello") == "Hello"
        assert _("Test string") == "Test string"

    def test_default_ngettext_returns_singular_or_plural(self):
        """Test that default ngettext returns correct form."""
        assert ngettext("item", "items", 1) == "item"
        assert ngettext("item", "items", 0) == "items"
        assert ngettext("item", "items", 2) == "items"

    def test_default_pgettext_returns_unchanged(self):
        """Test that default pgettext returns string unchanged."""
        assert pgettext("context", "Hello") == "Hello"


class TestInitializeTranslation:
    """Test initialize_translation function."""

    def test_initializes_with_fallback(self):
        """Test that initialize_translation works with fallback."""
        with tempfile.TemporaryDirectory() as tmpdir:
            # Should not raise even if locale directory doesn't exist
            initialize_translation(tmpdir, "test_domain")

            # Functions should still work
            assert _("test") == "test"

    def test_get_translation_returns_none_before_init(self):
        """Test that get_translation returns None before initialization."""
        # Reset translation state
        from service_ppt.utils import i18n
        i18n._translation = None

        result = get_translation()
        assert result is None

    def test_get_translation_returns_translation_after_init(self):
        """Test that get_translation returns translation after initialization."""
        with tempfile.TemporaryDirectory() as tmpdir:
            initialize_translation(tmpdir, "test_domain")
            result = get_translation()
            assert result is not None
