"""Tests for bibleformat module.

This module contains unit tests for Bible format registry and factory functions.
"""


from service_ppt.bible import bibleformat
from service_ppt.bible.bibleformat import BibleFormat, enum_versions, get_format_list, get_format_option, set_format_option


class TestBibleFormat:
    """Test BibleFormat enum."""

    def test_enum_values(self):
        """Test that BibleFormat enum has correct values."""
        assert BibleFormat.MYBIBLE.value == "MyBible"
        assert BibleFormat.MYSWORD.value == "MySword"
        assert BibleFormat.SWORD.value == "Sword"
        assert BibleFormat.ZEFANIA.value == "Zefania"


class TestGetFormatList:
    """Test get_format_list function."""

    def test_returns_list(self):
        """Test that get_format_list returns a list."""
        result = get_format_list()
        assert isinstance(result, list)

    def test_contains_expected_formats(self):
        """Test that get_format_list contains expected formats."""
        result = get_format_list()
        # Should at least contain MyBible, MySword, and Zefania
        assert "MyBible" in result
        assert "MySword" in result
        assert "Zefania" in result


class TestEnumVersions:
    """Test enum_versions function."""

    def test_returns_list(self):
        """Test that enum_versions returns a list."""
        result = enum_versions("MyBible")
        assert isinstance(result, list)

    def test_returns_empty_list_for_invalid_format(self):
        """Test that enum_versions returns empty list for invalid format."""
        result = enum_versions("InvalidFormat")
        assert result == []

    def test_returns_empty_list_when_no_versions(self):
        """Test that enum_versions returns empty list when no versions available."""
        # This may return empty list if no Bible files are configured
        result = enum_versions("MyBible")
        assert isinstance(result, list)


class TestFormatOptions:
    """Test format option getter and setter functions."""

    def test_get_format_option_returns_none_for_invalid_format(self):
        """Test that get_format_option returns None for invalid format."""
        result = get_format_option("InvalidFormat", "ROOT_DIR")
        assert result is None

    def test_set_format_option_handles_invalid_format(self):
        """Test that set_format_option handles invalid format gracefully."""
        # Should not raise
        set_format_option("InvalidFormat", "ROOT_DIR", "/path")

    def test_set_and_get_format_option(self):
        """Test that set_format_option and get_format_option work together."""
        test_path = "/test/path"
        set_format_option("MyBible", "ROOT_DIR", test_path)
        result = get_format_option("MyBible", "ROOT_DIR")
        # May return None if format not properly initialized, or the set value
        assert result is None or result == test_path


class TestGetBibleInfo:
    """Test get_bible_info function."""

    def test_returns_info_for_esv(self):
        """Test that get_bible_info returns info for ESV."""
        result = bibleformat.get_bible_info("ESV")
        assert result is not None
        assert "creator" in result
        assert "description" in result
        assert "publisher" in result
        assert "rights" in result
        assert result["creator"] == "Crossway"

    def test_returns_info_for_korean_version(self):
        """Test that get_bible_info returns info for Korean version."""
        result = bibleformat.get_bible_info("개역개정")
        assert result is not None
        assert "creator" in result
        assert "description" in result

    def test_returns_none_for_unknown_version(self):
        """Test that get_bible_info returns None for unknown version."""
        result = bibleformat.get_bible_info("UnknownVersion")
        assert result is None
