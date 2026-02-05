"""MySword Bible format reader.

This module provides support for reading Bible text from MySword format files,
which use SQLite3 database format. MySword modules can be downloaded from
various sources or created using tools like bgmysword.

Format specification: https://www.mysword.info/modules-format
bgmysword tool: https://github.com/GreenRaccoon23/bgmysword
"""

import os
import re
import sqlite3

from service_ppt.bible import biblang
from service_ppt.bible.bibcore import Bible, BibleInfo, Book, Chapter, FileFormat, Verse


class MySwordReader:
    """Reader for MySword Bible format files.

    This class provides functionality to read Bible text from MySword format,
    which uses SQLite3 database format.
    """

    def __init__(self) -> None:
        """Initialize MySword reader."""
        self.conn: sqlite3.Connection | None = None
        self.cursor: sqlite3.Cursor | None = None
        self.remove_tags: bool = False

    def read_bible(self, conn: sqlite3.Connection, version: str, remove_tags: bool, load_all: bool = False) -> Bible:
        """Read Bible data from MySword SQLite database.

        :param conn: SQLite database connection
        :param version: Bible version name
        :param remove_tags: Whether to remove MySword formatting tags
        :param load_all: Whether to load all books immediately (defaults to lazy loading)
        :returns: Bible object
        """
        self.remove_tags = remove_tags

        bible = Bible()
        bible.name = version

        cursor = conn.cursor()
        if not load_all:
            self.conn = conn
            self.cursor = cursor
            bible.reader = self

        for book_no in range(self._get_book_count(cursor)):
            book = Book()
            book.new_testament = BibleInfo.is_new_testament(book_no)
            names = biblang.L18N.get_book_names(book_no)
            book.name = names[0]
            book.short_name = names[1]
            bible.books.append(book)

            if load_all:
                self._parse_chapters(book, cursor, book_no)

        bible.ensure_loaded(bible.books[0])
        bible.lang = biblang.detect_language(bible.books[0].chapters[0].verses[0].text)

        return bible

    def read_book(self, book: Book, book_no: int) -> None:
        """Read a single book from MySword database.

        :param book: Book object to populate with data
        :param book_no: Zero-based book number
        """
        self._parse_chapters(book, self.cursor, book_no)

    def _parse_chapters(self, book: Book, cursor: sqlite3.Cursor, book_no: int) -> None:
        """Parse chapters and verses from MySword database for a book.

        :param book: Book object to populate with chapters and verses
        :param cursor: SQLite cursor for database queries
        :param book_no: Zero-based book number
        """
        for chapter_no in range(self._get_chapter_count(cursor, book_no)):
            chapter = Chapter()
            chapter.no = chapter_no + 1
            book.chapters.append(chapter)

            cursor.execute(f"SELECT Verse, Scripture FROM Bible WHERE Book='{book_no + 1}' and Chapter='{chapter.no}';")
            verses = cursor.fetchall()
            for v in verses:
                v_text = v[1]
                # Treat 'Title Start' at the beginning as another verse.
                if v_text.startswith("<TS>"):
                    index = v_text.index("<Ts>")
                    if index != -1:
                        verse = Verse()
                        verse.text = v_text[4:index]
                        chapter.verses.append(verse)

                        v_text = v_text[index + 4 :]

                verse = Verse()
                verse.set_no(v[0])
                verse.text = self._cleanup_text(v_text)
                chapter.verses.append(verse)

    def _cleanup_text(self, text: str) -> str:
        """Remove or convert MySword formatting tags from text.

        Processes Bible-specific tags like CI, CM, FI/Fi, FO/Fo, FR/Fr, FU/Fu,
        PF#, PI#, RF/Rf, TS/Ts and either removes them or converts to HTML.

        :param text: Text containing MySword formatting tags
        :returns: Text with tags processed or removed
        """
        if self.remove_tags:
            text = re.sub("<CI>", " ", text)
            text = re.sub("<(CM|FI|Fi|FO|Fo|FR|Fr|FU|Fu|PF[0-7]|PI[0-7])>", "", text)

            # match non-greedy: https://docs.python.org/3/howto/regex.html#greedy-versus-non-greedy
            text = re.sub("<RF>.*?<Rf>", "", text)
            # Ignore 'Title Start' in the middle.
            text = re.sub("<TS>.*?<Ts>", "", text)

        return text.strip()

    def _has_book(self, cursor: sqlite3.Cursor, book: int) -> bool:
        """Check if a book exists in the database.

        :param cursor: SQLite cursor for database queries
        :param book: One-based book number
        :returns: True if book exists, False otherwise
        """
        cursor.execute(f"SELECT Book FROM Bible WHERE Book='{book}' and Chapter='1' and Verse='1';")
        rows = cursor.fetchall()
        return len(rows) > 0

    def _get_book_count(self, cursor: sqlite3.Cursor) -> int:
        """Get the actual number of books in the database.

        Searches for the highest book number that exists in the database,
        starting from the standard Bible book count.

        :param cursor: SQLite cursor for database queries
        :returns: Actual number of books in the database
        """
        book_count = BibleInfo.get_book_count()
        if self._has_book(cursor, book_count):
            while True:
                if not self._has_book(cursor, book_count + 1):
                    return book_count
                book_count += 1
        else:
            while True:
                if self._has_book(cursor, book_count - 1):
                    return book_count - 1
                book_count -= 1

    def _has_chapter(self, cursor: sqlite3.Cursor, book: int, chapter: int) -> bool:
        """Check if a chapter exists in the database for a book.

        :param cursor: SQLite cursor for database queries
        :param book: Zero-based book number
        :param chapter: Chapter number
        :returns: True if chapter exists, False otherwise
        """
        cursor.execute(f"SELECT Chapter FROM Bible WHERE Book='{book + 1}' and Chapter='{chapter}' and Verse='1';")
        rows = cursor.fetchall()
        return len(rows) > 0

    def _get_chapter_count(self, cursor: sqlite3.Cursor, book: int) -> int:
        """Get the actual number of chapters for a book in the database.

        Searches for the highest chapter number that exists in the database,
        starting from the standard chapter count for the book.

        :param cursor: SQLite cursor for database queries
        :param book: Zero-based book number
        :returns: Actual number of chapters for the book
        """
        chapter_count = BibleInfo.get_chapter_count(book)
        if self._has_chapter(cursor, book, chapter_count):
            while True:
                if not self._has_chapter(cursor, book, chapter_count + 1):
                    return chapter_count
                chapter_count += 1
        else:
            while True:
                if self._has_chapter(cursor, book, chapter_count - 1):
                    return chapter_count - 1
                chapter_count -= 1


class MySwordFormat(FileFormat):
    """File format handler for MySword format files.

    This class implements the FileFormat interface for reading MySword format
    Bible files from SQLite3 databases.
    """

    def __init__(self) -> None:
        """Initialize MySword format handler."""
        super().__init__()

        self.versions: dict[str, str] | None = None
        self.options: dict[str, str | None] = {"ROOT_DIR": "", "remove_bible_tags": None}

    def _get_root_dir(self) -> str:
        """Get the root directory for MySword format files.

        :returns: Path to the root directory containing MySword format files
        """
        dirname = self.get_option("ROOT_DIR")
        if not dirname:
            dirname = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "..", "bgmysword")

        return os.path.normpath(dirname)

    @staticmethod
    def _get_bible_name(filename: str) -> str | None:
        """Check if file is a valid MySword SQLite3 Bible database.

        Validates that the file contains both "Details" and "Bible" tables
        and extracts the Bible abbreviation from the Details table.

        :param filename: Path to SQLite3 file to check
        :returns: Bible name/abbreviation if valid, None otherwise
        """
        with sqlite3.connect(filename) as conn:
            cursor = conn.cursor()
            cursor.execute("SELECT name FROM sqlite_master WHERE type='table';")
            tables = cursor.fetchall()
            count = 0
            for t in tables:
                if t[0] in ("Details", "Bible"):
                    count = count + 1

            if count != 2:
                return None

            cursor.execute("select Abbreviation from Details;")
            rows = cursor.fetchall()

            if len(rows) == 1 and len(rows[0]):
                return rows[0][0]

        return None

    @staticmethod
    def _get_column_names(cursor: sqlite3.Cursor, tablename: str) -> list[str]:
        """Get list of column names from a SQLite table.

        Uses PRAGMA table_info to retrieve column information.

        :param cursor: SQLite cursor for database queries
        :param tablename: Name of the table to query
        :returns: List of column names
        """
        cursor.execute(f"pragma table_info({tablename});")
        columns = cursor.fetchall()
        return [c[1] for c in columns]

    def enum_versions(self) -> list[str]:
        """Enumerate available Bible versions in MySword format.

        :returns: List of available version names
        """
        dirname = self._get_root_dir()

        versions = {}
        for f in os.listdir(dirname):
            if os.path.isfile(os.path.join(dirname, f)):
                pathname = os.path.join(dirname, f)
                name = self._get_bible_name(pathname)
                if name:
                    versions[name] = pathname

        self.versions = versions

        return list(self.versions.keys())

    def read_version(self, version: str) -> Bible | None:
        """Read a specific Bible version from MySword format.

        :param version: Version name to read
        :returns: Bible object or None if version not found
        """
        if self.versions is None:
            self.enum_versions()

        if version not in self.versions:
            return None

        filename = self.versions[version]
        conn = sqlite3.connect(filename)

        reader = MySwordReader()
        remove_tags = False
        if self.options["remove_bible_tags"]:
            remove_tags = True
        return reader.read_bible(conn, version, remove_tags)
