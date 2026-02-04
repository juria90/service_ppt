"""Tests for openlpservice module.

This module contains unit tests for OpenLP Service format writer.
"""

import json
import os
import tempfile
import zipfile

from service_ppt.hymn.hymncore import Book, Line, Song, Verse
from service_ppt.hymn.openlpservice import OpenLPServiceWriter


class TestOpenLPServiceWriter:
    """Tests for OpenLPServiceWriter class."""

    def test_get_extension(self):
        """Test _get_extension method."""
        writer = OpenLPServiceWriter()
        assert writer._get_extension() == ".osz"

    def test_write_single_song(self):
        """Test writing a single song to OpenLP format."""
        song = Song()
        song.title = "Test Song"
        v1 = Verse()
        v1.no = "v1"
        v1.lines.append(Line("Line 1"))
        v1.lines.append(Line("Line 2"))
        song.verses.append(v1)

        with tempfile.NamedTemporaryFile(suffix=".osz", delete=False) as f:
            filename = f.name

        try:
            writer = OpenLPServiceWriter()
            xml_content = "<song>Test XML</song>"
            writer.write(filename, [song], [xml_content])

            assert os.path.exists(filename)
            with zipfile.ZipFile(filename, "r") as zipf:
                files = zipf.namelist()
                assert len(files) == 1
                assert files[0].endswith(".osj")

                content = zipf.read(files[0]).decode("utf-8")
                data = json.loads(content)
                assert len(data) == 2  # Header + song
                assert "openlp_core" in data[0]
                assert "serviceitem" in data[1]
        finally:
            os.unlink(filename)

    def test_write_multiple_songs(self):
        """Test writing multiple songs to OpenLP format."""
        song1 = Song()
        song1.title = "Song 1"
        v1 = Verse()
        v1.no = "v1"
        v1.lines.append(Line("Line 1"))
        song1.verses.append(v1)

        song2 = Song()
        song2.title = "Song 2"
        v2 = Verse()
        v2.no = "v1"
        v2.lines.append(Line("Line 2"))
        song2.verses.append(v2)

        with tempfile.NamedTemporaryFile(suffix=".osz", delete=False) as f:
            filename = f.name

        try:
            writer = OpenLPServiceWriter()
            xml_contents = ["<song>Song 1 XML</song>", "<song>Song 2 XML</song>"]
            writer.write(filename, [song1, song2], xml_contents)

            with zipfile.ZipFile(filename, "r") as zipf:
                content = zipf.read(zipf.namelist()[0]).decode("utf-8")
                data = json.loads(content)
                assert len(data) == 3  # Header + 2 songs
                assert data[1]["serviceitem"]["header"]["title"] == "Song 1"
                assert data[2]["serviceitem"]["header"]["title"] == "Song 2"
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

        with tempfile.NamedTemporaryFile(suffix=".osz", delete=False) as f:
            filename = f.name

        try:
            writer = OpenLPServiceWriter()
            writer.write(filename, [song], ["<song>XML</song>"])

            with zipfile.ZipFile(filename, "r") as zipf:
                content = zipf.read(zipf.namelist()[0]).decode("utf-8")
                data = json.loads(content)
                serviceitem = data[1]["serviceitem"]
                data_list = serviceitem["data"]
                assert len(data_list) == 2
                assert data_list[0]["verseTag"] == "v2"
                assert data_list[1]["verseTag"] == "v1"
        finally:
            os.unlink(filename)

    def test_write_song_with_optional_break(self):
        """Test writing a song with optional break markers."""
        song = Song()
        song.title = "Test Song"
        v1 = Verse()
        v1.no = "v1"
        v1.lines.append(Line("Line 1", False))
        v1.lines.append(Line("Line 2", True))
        v1.lines.append(Line("Line 3", False))
        song.verses.append(v1)

        with tempfile.NamedTemporaryFile(suffix=".osz", delete=False) as f:
            filename = f.name

        try:
            writer = OpenLPServiceWriter()
            writer.write(filename, [song], ["<song>XML</song>"])

            with zipfile.ZipFile(filename, "r") as zipf:
                content = zipf.read(zipf.namelist()[0]).decode("utf-8")
                data = json.loads(content)
                serviceitem = data[1]["serviceitem"]
                raw_slide = serviceitem["data"][0]["raw_slide"]
                assert "[---]" in raw_slide
                assert "Line 2" in raw_slide
        finally:
            os.unlink(filename)

    def test_write_song_empty_verse(self):
        """Test writing a song with empty verse."""
        song = Song()
        song.title = "Test Song"
        v1 = Verse()
        v1.no = "v1"
        song.verses.append(v1)

        with tempfile.NamedTemporaryFile(suffix=".osz", delete=False) as f:
            filename = f.name

        try:
            writer = OpenLPServiceWriter()
            writer.write(filename, [song], ["<song>XML</song>"])

            with zipfile.ZipFile(filename, "r") as zipf:
                content = zipf.read(zipf.namelist()[0]).decode("utf-8")
                data = json.loads(content)
                serviceitem = data[1]["serviceitem"]
                data_list = serviceitem["data"]
                assert len(data_list) == 1
                assert data_list[0]["title"] == ""
        finally:
            os.unlink(filename)
