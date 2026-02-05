"""Tests for html_bible module.

This module contains unit tests for HTML Bible format writer.
"""

import os
import tempfile

from service_ppt.bible.bibcore import Bible, Book, Chapter, Verse
from service_ppt.bible.biblang import LANG_EN
from service_ppt.bible.html_bible import HTMLWriter


def create_test_bible() -> Bible:
    """Create a test Bible with sample data.

    :returns: A Bible object with Genesis chapter 1
    """
    book = Book()
    book.name = "Genesis"
    book.short_name = "Gen"

    chapter = Chapter()
    chapter.no = 1
    book.chapters.append(chapter)

    verse1 = Verse()
    verse1.set_no(1)
    verse1.text = "In the beginning, God created the heavens and the earth."
    chapter.verses.append(verse1)

    bible = Bible()
    bible.lang = LANG_EN
    bible.name = "Test Bible"
    bible.books.append(book)

    return bible


class TestHTMLWriter:
    """Test HTMLWriter class."""

    def test_get_extension_returns_html(self):
        """Test that _get_extension returns .html."""
        writer = HTMLWriter(False, False)
        assert writer._get_extension() == ".html"

    def test_write_bible_creates_files(self):
        """Test that write_bible creates HTML files."""
        bible = create_test_bible()
        writer = HTMLWriter(False, False)

        with tempfile.TemporaryDirectory() as tmpdir:
            writer.write_bible(tmpdir, bible)

            # Should create book files and index
            book_file = os.path.join(tmpdir, "book1.html")
            index_file = os.path.join(tmpdir, "index.html")

            assert os.path.exists(book_file)
            assert os.path.exists(index_file)

    def test_write_bible_with_dynamic_page(self):
        """Test that write_bible works with dynamic_page option."""
        bible = create_test_bible()
        writer = HTMLWriter(True, False)

        with tempfile.TemporaryDirectory() as tmpdir:
            writer.write_bible(tmpdir, bible)

            index_file = os.path.join(tmpdir, "index.html")
            assert os.path.exists(index_file)

            # Check that JavaScript is included
            with open(index_file, encoding="utf-8") as f:
                content = f.read()
                assert "changeBible" in content or "load_book" in content

    def test_write_bible_with_encoding(self):
        """Test that write_bible respects encoding parameter."""
        bible = create_test_bible()
        writer = HTMLWriter(False, False)

        with tempfile.TemporaryDirectory() as tmpdir:
            writer.write_bible(tmpdir, bible, encoding="utf-8")

            book_file = os.path.join(tmpdir, "book1.html")
            assert os.path.exists(book_file)
