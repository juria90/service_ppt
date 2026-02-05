"""Tests for opensong_bible module.

This module contains unit tests for OpenSong Bible format writer.
"""

import os
import tempfile
import xml.etree.ElementTree as ET

from service_ppt.bible.bibcore import Bible, Book, Chapter, Verse
from service_ppt.bible.biblang import LANG_EN
from service_ppt.bible.opensong_bible import OpenSongXMLWriter


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

    verse2 = Verse()
    verse2.set_no(2)
    verse2.text = "The earth was without form and void."
    chapter.verses.append(verse2)

    bible = Bible()
    bible.lang = LANG_EN
    bible.name = "Test Bible"
    bible.books.append(book)

    return bible


class TestOpenSongXMLWriter:
    """Test OpenSongXMLWriter class."""

    def test_get_extension_returns_xml(self):
        """Test that _get_extension returns .xml."""
        writer = OpenSongXMLWriter()
        assert writer._get_extension() == ".xml"

    def test_write_bible_creates_xml_file(self):
        """Test that write_bible creates bible.xml file."""
        bible = create_test_bible()
        writer = OpenSongXMLWriter()

        with tempfile.TemporaryDirectory() as tmpdir:
            writer.write_bible(tmpdir, bible)

            xml_file = os.path.join(tmpdir, "bible.xml")
            assert os.path.exists(xml_file)

    def test_write_bible_creates_valid_xml(self):
        """Test that write_bible creates valid XML structure."""
        bible = create_test_bible()
        writer = OpenSongXMLWriter()

        with tempfile.TemporaryDirectory() as tmpdir:
            writer.write_bible(tmpdir, bible)

            xml_file = os.path.join(tmpdir, "bible.xml")
            # Try to parse as XML
            tree = ET.parse(xml_file)
            root = tree.getroot()

            assert root.tag == "bible"
            assert len(root) > 0  # Should have at least one book

    def test_write_bible_includes_book_elements(self):
        """Test that write_bible includes book elements."""
        bible = create_test_bible()
        writer = OpenSongXMLWriter()

        with tempfile.TemporaryDirectory() as tmpdir:
            writer.write_bible(tmpdir, bible)

            xml_file = os.path.join(tmpdir, "bible.xml")
            tree = ET.parse(xml_file)
            root = tree.getroot()

            books = root.findall("b")
            assert len(books) == 1
            assert books[0].get("n") == "Genesis"

    def test_write_bible_includes_chapter_elements(self):
        """Test that write_bible includes chapter elements."""
        bible = create_test_bible()
        writer = OpenSongXMLWriter()

        with tempfile.TemporaryDirectory() as tmpdir:
            writer.write_bible(tmpdir, bible)

            xml_file = os.path.join(tmpdir, "bible.xml")
            tree = ET.parse(xml_file)
            root = tree.getroot()

            book = root.find("b")
            chapters = book.findall("c")
            assert len(chapters) == 1
            assert chapters[0].get("n") == "1"

    def test_write_bible_includes_verse_elements(self):
        """Test that write_bible includes verse elements."""
        bible = create_test_bible()
        writer = OpenSongXMLWriter()

        with tempfile.TemporaryDirectory() as tmpdir:
            writer.write_bible(tmpdir, bible)

            xml_file = os.path.join(tmpdir, "bible.xml")
            tree = ET.parse(xml_file)
            root = tree.getroot()

            book = root.find("b")
            chapter = book.find("c")
            verses = chapter.findall("v")
            assert len(verses) == 2
            assert verses[0].get("n") == "1"
            assert "beginning" in verses[0].text

    def test_write_bible_with_encoding(self):
        """Test that write_bible respects encoding parameter."""
        bible = create_test_bible()
        writer = OpenSongXMLWriter()

        with tempfile.TemporaryDirectory() as tmpdir:
            writer.write_bible(tmpdir, bible, encoding="utf-8")

            xml_file = os.path.join(tmpdir, "bible.xml")
            assert os.path.exists(xml_file)
