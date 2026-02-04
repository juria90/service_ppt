"""Tests for hymncore module.

This module contains unit tests for the core hymn data structures:
Line, Verse, Song, and Book.
"""

from service_ppt.hymn.hymncore import Book, Line, Song, Verse


class TestLine:
    """Tests for Line class."""

    def test_line_init(self):
        """Test Line initialization."""
        line = Line("Test line", False)
        assert line.text == "Test line"
        assert line.optional_break is False

    def test_line_with_optional_break(self):
        """Test Line with optional break."""
        line = Line("Test line", True)
        assert line.text == "Test line"
        assert line.optional_break is True

    def test_line_default_optional_break(self):
        """Test Line with default optional_break."""
        line = Line("Test line")
        assert line.text == "Test line"
        assert line.optional_break is False


class TestVerse:
    """Tests for Verse class."""

    def test_verse_init(self):
        """Test Verse initialization."""
        verse = Verse()
        assert verse.no is None
        assert verse.lines == []

    def test_verse_with_number(self):
        """Test Verse with number."""
        verse = Verse()
        verse.no = "v1"
        assert verse.no == "v1"

    def test_verse_with_lines(self):
        """Test Verse with lines."""
        verse = Verse()
        line1 = Line("Line 1")
        line2 = Line("Line 2", True)
        verse.lines.append(line1)
        verse.lines.append(line2)
        assert len(verse.lines) == 2
        assert verse.lines[0].text == "Line 1"
        assert verse.lines[1].optional_break is True


class TestSong:
    """Tests for Song class."""

    def test_song_init(self):
        """Test Song initialization."""
        song = Song()
        assert song.title is None
        assert song.authors is None
        assert song.verse_order is None
        assert song.songbook is None
        assert song.released is None
        assert song.keywords is None
        assert song.verses == []
        assert song.ordered_verses is None

    def test_song_with_metadata(self):
        """Test Song with metadata."""
        song = Song()
        song.title = "Amazing Grace"
        song.authors = [("words", "John Newton")]
        song.verse_order = "v1 v2"
        assert song.title == "Amazing Grace"
        assert song.authors == [("words", "John Newton")]
        assert song.verse_order == "v1 v2"

    def test_get_verses_by_order_no_order(self):
        """Test get_verses_by_order with no verse_order."""
        song = Song()
        v1 = Verse()
        v1.no = "v1"
        v2 = Verse()
        v2.no = "v2"
        song.verses = [v1, v2]
        result = song.get_verses_by_order()
        assert result == [v1, v2]

    def test_get_verses_by_order_with_order(self):
        """Test get_verses_by_order with verse_order."""
        song = Song()
        v1 = Verse()
        v1.no = "v1"
        v2 = Verse()
        v2.no = "v2"
        v3 = Verse()
        v3.no = "v3"
        song.verses = [v1, v2, v3]
        song.verse_order = "v3 v1"
        result = song.get_verses_by_order()
        assert result == [v3, v1]
        assert len(result) == 2

    def test_get_verses_by_order_caching(self):
        """Test get_verses_by_order caching."""
        song = Song()
        v1 = Verse()
        v1.no = "v1"
        song.verses = [v1]
        song.verse_order = "v1"
        result1 = song.get_verses_by_order()
        result2 = song.get_verses_by_order()
        assert result1 is result2  # Should return same cached list

    def test_get_verses_by_order_empty_order(self):
        """Test get_verses_by_order with empty verse_order."""
        song = Song()
        v1 = Verse()
        v1.no = "v1"
        song.verses = [v1]
        song.verse_order = ""
        result = song.get_verses_by_order()
        assert result == [v1]

    def test_get_lines_by_order(self):
        """Test get_lines_by_order."""
        song = Song()
        v1 = Verse()
        v1.no = "v1"
        v1.lines.append(Line("Line 1"))
        v1.lines.append(Line("Line 2"))
        v2 = Verse()
        v2.no = "v2"
        v2.lines.append(Line("Line 3"))
        song.verses = [v1, v2]
        lines = list(song.get_lines_by_order())
        assert len(lines) == 3
        assert lines[0].text == "Line 1"
        assert lines[1].text == "Line 2"
        assert lines[2].text == "Line 3"

    def test_get_lines_by_order_with_verse_order(self):
        """Test get_lines_by_order with verse_order."""
        song = Song()
        v1 = Verse()
        v1.no = "v1"
        v1.lines.append(Line("Line 1"))
        v2 = Verse()
        v2.no = "v2"
        v2.lines.append(Line("Line 2"))
        song.verses = [v1, v2]
        song.verse_order = "v2 v1"
        lines = list(song.get_lines_by_order())
        assert len(lines) == 2
        assert lines[0].text == "Line 2"
        assert lines[1].text == "Line 1"


class TestBook:
    """Tests for Book class."""

    def test_book_init(self):
        """Test Book initialization."""
        book = Book()
        assert book.name is None
        assert book.songs == []

    def test_book_with_songs(self):
        """Test Book with songs."""
        book = Book()
        book.name = "Hymnal"
        song1 = Song()
        song1.title = "Song 1"
        song2 = Song()
        song2.title = "Song 2"
        book.songs = [song1, song2]
        assert book.name == "Hymnal"
        assert len(book.songs) == 2
        assert book.songs[0].title == "Song 1"
        assert book.songs[1].title == "Song 2"
