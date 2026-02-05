"""Tests for csv_bible module.

This module contains unit tests for CSV Bible format reader and writer.
"""

import csv
import os
import tempfile

from service_ppt.bible.bibcore import Bible, Book, Chapter, Verse
from service_ppt.bible.biblang import LANG_EN
from service_ppt.bible.csv_bible import CSVWriter


def create_test_bible() -> Bible:
    """Create a test Bible with sample data.

    :returns: A Bible object with Genesis chapter 1
    """
    book = Book()
    book.name = "Genesis"
    book.short_name = "Gen"
    book.new_testament = False

    chapter = Chapter()
    chapter.no = 1
    book.chapters.append(chapter)

    verse1 = Verse()
    verse1.set_no(1)
    verse1.text = "In the beginning, God created the heavens and the earth."
    chapter.verses.append(verse1)

    verse2 = Verse()
    verse2.set_no(2)
    verse2.text = "The earth was without form and void."
    chapter.verses.append(verse2)

    bible = Bible()
    bible.lang = LANG_EN
    bible.name = "Test Bible"
    bible.books.append(book)

    return bible


class TestCSVWriter:
    """Test CSVWriter class."""

    def test_get_extension_returns_csv(self):
        """Test that _get_extension returns .csv."""
        writer = CSVWriter()
        assert writer._get_extension() == ".csv"

    def test_write_bible_creates_files(self):
        """Test that write_bible creates books.csv and verses.csv files."""
        bible = create_test_bible()
        writer = CSVWriter()

        with tempfile.TemporaryDirectory() as tmpdir:
            writer.write_bible(tmpdir, bible)

            books_file = os.path.join(tmpdir, "books.csv")
            verses_file = os.path.join(tmpdir, "verses.csv")

            assert os.path.exists(books_file)
            assert os.path.exists(verses_file)

    def test_write_books_creates_correct_format(self):
        """Test that _write_books creates CSV with correct columns."""
        bible = create_test_bible()
        writer = CSVWriter()

        with tempfile.TemporaryDirectory() as tmpdir:
            books_file = os.path.join(tmpdir, "books.csv")
            writer._write_books(books_file, bible)

            with open(books_file, encoding="utf-8") as f:
                reader = csv.reader(f)
                rows = list(reader)

                assert len(rows) == 1
                assert rows[0][0] == "1"  # book number
                assert rows[0][1] == "1"  # old/new testament (1 = old)
                assert rows[0][2] == "Genesis"  # book name
                assert rows[0][3] == "Gen"  # short name

    def test_write_verses_creates_correct_format(self):
        """Test that _write_verses creates CSV with correct columns."""
        bible = create_test_bible()
        writer = CSVWriter()

        with tempfile.TemporaryDirectory() as tmpdir:
            verses_file = os.path.join(tmpdir, "verses.csv")
            writer._write_verses(verses_file, bible)

            with open(verses_file, encoding="utf-8") as f:
                reader = csv.reader(f)
                rows = list(reader)

                assert len(rows) == 2  # Two verses
                assert rows[0][0] == "1"  # book number
                assert rows[0][1] == "1"  # chapter number
                assert rows[0][2] == "1"  # verse number
                assert "beginning" in rows[0][3]  # verse text

    def test_write_bible_with_encoding(self):
        """Test that write_bible respects encoding parameter."""
        bible = create_test_bible()
        writer = CSVWriter()

        with tempfile.TemporaryDirectory() as tmpdir:
            writer.write_bible(tmpdir, bible, encoding="utf-8-sig")

            books_file = os.path.join(tmpdir, "books.csv")
            assert os.path.exists(books_file)
