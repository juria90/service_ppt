"""Tests for openlyrics module.

This module contains unit tests for OpenLyrics format reader and writer.
"""

import os
import tempfile
import xml.etree.ElementTree as ET

import pytest

from service_ppt.hymn.hymncore import Book, Line, Song, Verse
from service_ppt.hymn.openlyrics import OpenLyricsReader, OpenLyricsWriter


class TestOpenLyricsReader:
    """Tests for OpenLyricsReader class."""

    def test_get_extension(self):
        """Test _get_extension method."""
        reader = OpenLyricsReader()
        assert reader._get_extension() == ".xml"

    def test_get_element_text_with_valid_element(self):
        """Test _get_element_text with valid element."""
        root = ET.Element("root")
        child = ET.SubElement(root, "child")
        child.text = "Test text"
        result = OpenLyricsReader._get_element_text(root, "child")
        assert result == "Test text"

    def test_get_element_text_with_none_parent(self):
        """Test _get_element_text with None parent."""
        result = OpenLyricsReader._get_element_text(None, "child")
        assert result is None

    def test_get_element_text_with_missing_element(self):
        """Test _get_element_text with missing element."""
        root = ET.Element("root")
        result = OpenLyricsReader._get_element_text(root, "missing")
        assert result is None

    def test_get_element_text_with_empty_text(self):
        """Test _get_element_text with element having empty text."""
        root = ET.Element("root")
        child = ET.SubElement(root, "child")
        child.text = ""
        result = OpenLyricsReader._get_element_text(root, "child")
        assert result == ""

    def test_read_song_minimal(self):
        """Test reading a minimal valid song."""
        with tempfile.NamedTemporaryFile(mode="w", suffix=".xml", delete=False) as f:
            f.write(
                """<?xml version="1.0" encoding="UTF-8"?>
<song xmlns="http://openlyrics.info/namespace/2009/song" version="0.8">
  <properties>
    <titles>
      <title>Test Song</title>
    </titles>
  </properties>
  <lyrics>
    <verse name="v1">
      <lines>Line 1</lines>
    </verse>
  </lyrics>
</song>"""
            )
            filename = f.name

        try:
            reader = OpenLyricsReader()
            song = reader.read_song(filename)
            assert song is not None
            assert song.title == "Test Song"
            assert len(song.verses) == 1
            assert song.verses[0].no == "v1"
            assert len(song.verses[0].lines) == 1
            assert song.verses[0].lines[0].text == "Line 1"
        finally:
            os.unlink(filename)

    def test_read_song_with_authors(self):
        """Test reading a song with authors."""
        with tempfile.NamedTemporaryFile(mode="w", suffix=".xml", delete=False) as f:
            f.write(
                """<?xml version="1.0" encoding="UTF-8"?>
<song xmlns="http://openlyrics.info/namespace/2009/song" version="0.8">
  <properties>
    <titles>
      <title>Test Song</title>
    </titles>
    <authors>
      <author type="words">John Doe</author>
      <author type="music">Jane Smith</author>
    </authors>
  </properties>
  <lyrics>
    <verse name="v1">
      <lines>Line 1</lines>
    </verse>
  </lyrics>
</song>"""
            )
            filename = f.name

        try:
            reader = OpenLyricsReader()
            song = reader.read_song(filename)
            assert song is not None
            assert song.authors == [("words", "John Doe"), ("music", "Jane Smith")]
        finally:
            os.unlink(filename)

    def test_read_song_with_translation_author(self):
        """Test reading a song with translation author."""
        with tempfile.NamedTemporaryFile(mode="w", suffix=".xml", delete=False) as f:
            f.write(
                """<?xml version="1.0" encoding="UTF-8"?>
<song xmlns="http://openlyrics.info/namespace/2009/song" version="0.8">
  <properties>
    <titles>
      <title>Test Song</title>
    </titles>
    <authors>
      <author type="translation" lang="ko">Korean Translator</author>
    </authors>
  </properties>
  <lyrics>
    <verse name="v1">
      <lines>Line 1</lines>
    </verse>
  </lyrics>
</song>"""
            )
            filename = f.name

        try:
            reader = OpenLyricsReader()
            song = reader.read_song(filename)
            assert song is not None
            assert song.authors == [("translation/ko", "Korean Translator")]
        finally:
            os.unlink(filename)

    def test_read_song_with_verse_order(self):
        """Test reading a song with verse order."""
        with tempfile.NamedTemporaryFile(mode="w", suffix=".xml", delete=False) as f:
            f.write(
                """<?xml version="1.0" encoding="UTF-8"?>
<song xmlns="http://openlyrics.info/namespace/2009/song" version="0.8">
  <properties>
    <titles>
      <title>Test Song</title>
    </titles>
    <verseOrder>v2 v1</verseOrder>
  </properties>
  <lyrics>
    <verse name="v1">
      <lines>Verse 1</lines>
    </verse>
    <verse name="v2">
      <lines>Verse 2</lines>
    </verse>
  </lyrics>
</song>"""
            )
            filename = f.name

        try:
            reader = OpenLyricsReader()
            song = reader.read_song(filename)
            assert song is not None
            assert song.verse_order == "v2 v1"
        finally:
            os.unlink(filename)

    def test_read_song_with_optional_break(self):
        """Test reading a song with optional break."""
        with tempfile.NamedTemporaryFile(mode="w", suffix=".xml", delete=False) as f:
            f.write(
                """<?xml version="1.0" encoding="UTF-8"?>
<song xmlns="http://openlyrics.info/namespace/2009/song" version="0.8">
  <properties>
    <titles>
      <title>Test Song</title>
    </titles>
  </properties>
  <lyrics>
    <verse name="v1">
      <lines break="optional">Line 1</lines>
    </verse>
  </lyrics>
</song>"""
            )
            filename = f.name

        try:
            reader = OpenLyricsReader()
            song = reader.read_song(filename)
            assert song is not None
            assert song.verses[0].lines[0].optional_break is True
        finally:
            os.unlink(filename)

    def test_read_song_invalid_root(self):
        """Test reading a file with invalid root element."""
        with tempfile.NamedTemporaryFile(mode="w", suffix=".xml", delete=False) as f:
            f.write(
                """<?xml version="1.0" encoding="UTF-8"?>
<notsong xmlns="http://openlyrics.info/namespace/2009/song" version="0.8">
</notsong>"""
            )
            filename = f.name

        try:
            reader = OpenLyricsReader()
            song = reader.read_song(filename)
            assert song is None
        finally:
            os.unlink(filename)

    def test_read_song_no_verses(self):
        """Test reading a song with no verses."""
        with tempfile.NamedTemporaryFile(mode="w", suffix=".xml", delete=False) as f:
            f.write(
                """<?xml version="1.0" encoding="UTF-8"?>
<song xmlns="http://openlyrics.info/namespace/2009/song" version="0.8">
  <properties>
    <titles>
      <title>Test Song</title>
    </titles>
  </properties>
  <lyrics>
  </lyrics>
</song>"""
            )
            filename = f.name

        try:
            reader = OpenLyricsReader()
            song = reader.read_song(filename)
            assert song is None
        finally:
            os.unlink(filename)


class TestOpenLyricsWriter:
    """Tests for OpenLyricsWriter class."""

    def test_get_extension(self):
        """Test _get_extension method."""
        writer = OpenLyricsWriter()
        assert writer._get_extension() == ".xml"

    def test_write_song_basic(self):
        """Test writing a basic song."""
        song = Song()
        song.title = "Test Song"
        v1 = Verse()
        v1.no = "v1"
        v1.lines.append(Line("Line 1"))
        song.verses.append(v1)

        with tempfile.NamedTemporaryFile(mode="w", suffix=".xml", delete=False) as f:
            filename = f.name

        try:
            writer = OpenLyricsWriter()
            writer.write_song(filename, song)
            assert os.path.exists(filename)
            with open(filename, encoding="utf-8") as f:
                content = f.read()
                assert "Test Song" in content
                assert "v1" in content
                assert "Line 1" in content
        finally:
            os.unlink(filename)

    def test_write_song_with_authors(self):
        """Test writing a song with authors."""
        song = Song()
        song.title = "Test Song"
        song.authors = [("words", "John Doe"), ("music", "Jane Smith")]
        v1 = Verse()
        v1.no = "v1"
        v1.lines.append(Line("Line 1"))
        song.verses.append(v1)

        with tempfile.NamedTemporaryFile(mode="w", suffix=".xml", delete=False) as f:
            filename = f.name

        try:
            writer = OpenLyricsWriter()
            writer.write_song(filename, song)
            with open(filename, encoding="utf-8") as f:
                content = f.read()
                assert 'type="words"' in content
                assert "John Doe" in content
                assert 'type="music"' in content
                assert "Jane Smith" in content
        finally:
            os.unlink(filename)

    def test_write_song_with_translation_author(self):
        """Test writing a song with translation author."""
        song = Song()
        song.title = "Test Song"
        song.authors = [("translation/ko", "Korean Translator")]
        v1 = Verse()
        v1.no = "v1"
        v1.lines.append(Line("Line 1"))
        song.verses.append(v1)

        with tempfile.NamedTemporaryFile(mode="w", suffix=".xml", delete=False) as f:
            filename = f.name

        try:
            writer = OpenLyricsWriter()
            writer.write_song(filename, song)
            with open(filename, encoding="utf-8") as f:
                content = f.read()
                assert 'type="translation"' in content
                assert 'lang="ko"' in content
                assert "Korean Translator" in content
        finally:
            os.unlink(filename)

    def test_write_song_with_verse_order(self):
        """Test writing a song with verse order."""
        song = Song()
        song.title = "Test Song"
        song.verse_order = "v2 v1"
        v1 = Verse()
        v1.no = "v1"
        v1.lines.append(Line("Verse 1"))
        v2 = Verse()
        v2.no = "v2"
        v2.lines.append(Line("Verse 2"))
        song.verses = [v1, v2]

        with tempfile.NamedTemporaryFile(mode="w", suffix=".xml", delete=False) as f:
            filename = f.name

        try:
            writer = OpenLyricsWriter()
            writer.write_song(filename, song)
            with open(filename, encoding="utf-8") as f:
                content = f.read()
                assert "<verseOrder>v2 v1</verseOrder>" in content
        finally:
            os.unlink(filename)

    def test_write_song_with_optional_break(self):
        """Test writing a song with optional break."""
        song = Song()
        song.title = "Test Song"
        v1 = Verse()
        v1.no = "v1"
        v1.lines.append(Line("Line 1", True))
        song.verses.append(v1)

        with tempfile.NamedTemporaryFile(mode="w", suffix=".xml", delete=False) as f:
            filename = f.name

        try:
            writer = OpenLyricsWriter()
            writer.write_song(filename, song)
            with open(filename, encoding="utf-8") as f:
                content = f.read()
                assert 'break="optional"' in content
        finally:
            os.unlink(filename)

    def test_write_song_with_songbook(self):
        """Test writing a song with songbook."""
        song = Song()
        song.title = "Test Song"
        song.songbook = {"name": "Hymnal", "entry": "123"}
        v1 = Verse()
        v1.no = "v1"
        v1.lines.append(Line("Line 1"))
        song.verses.append(v1)

        with tempfile.NamedTemporaryFile(mode="w", suffix=".xml", delete=False) as f:
            filename = f.name

        try:
            writer = OpenLyricsWriter()
            writer.write_song(filename, song)
            with open(filename, encoding="utf-8") as f:
                content = f.read()
                assert 'name="Hymnal"' in content
                assert 'entry="123"' in content
        finally:
            os.unlink(filename)

    def test_write_hymn(self):
        """Test writing a hymn book."""
        book = Book()
        book.name = "Test Hymnal"
        song1 = Song()
        song1.title = "Song 1"
        v1 = Verse()
        v1.no = "v1"
        v1.lines.append(Line("Line 1"))
        song1.verses.append(v1)
        book.songs.append(song1)

        with tempfile.TemporaryDirectory() as tmpdir:
            writer = OpenLyricsWriter()
            writer.write_hymn(tmpdir, book)
            files = os.listdir(tmpdir)
            assert len(files) == 1
            assert files[0] == "Song 1.xml"

    def test_write_read_roundtrip(self):
        """Test writing and reading a song maintains data."""
        original_song = Song()
        original_song.title = "Test Song"
        original_song.authors = [("words", "John Doe")]
        original_song.verse_order = "v1"
        v1 = Verse()
        v1.no = "v1"
        v1.lines.append(Line("Line 1"))
        v1.lines.append(Line("Line 2", True))
        original_song.verses.append(v1)

        with tempfile.NamedTemporaryFile(mode="w", suffix=".xml", delete=False) as f:
            filename = f.name

        try:
            writer = OpenLyricsWriter()
            writer.write_song(filename, original_song)

            reader = OpenLyricsReader()
            read_song = reader.read_song(filename)

            assert read_song is not None
            assert read_song.title == original_song.title
            assert read_song.authors == original_song.authors
            assert read_song.verse_order == original_song.verse_order
            assert len(read_song.verses) == 1
            assert read_song.verses[0].no == "v1"
            assert len(read_song.verses[0].lines) == 2
            assert read_song.verses[0].lines[0].text == "Line 1"
            assert read_song.verses[0].lines[0].optional_break is False
            assert read_song.verses[0].lines[1].text == "Line 2"
            assert read_song.verses[0].lines[1].optional_break is True
        finally:
            os.unlink(filename)
