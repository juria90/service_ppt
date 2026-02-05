"""MyBible format reader.

This module provides support for reading Bible text from MyBible format files,
which consist of an index.txt file and multiple book<DD>.txt files. The format
uses UTF-8 or UTF-16 encoding and stores Bible text in a structured text format.

MyBibleReader is a text file bible reader for MyBible written by James Lee.

The file is saved as utf-8 or utf-16 to save the space and the structure is like below.
It is composed of multiple files and the index.txt contains TOC information.
[index.txt]
<BOM>
INDEX FILE
NAME=<Bible Version Name>
BOOKCOUNT=<Number of Books>
[BOOK=<Book Name>
CHAPTERS=[<OFFSET OF FILE FOR EACH CAPTER START>] *
] *

<book1.txt to bookn.txt>
<BOM>
[
<verse number><TAB><verse text>
] *
"""

import os
import re
from typing import TextIO

from service_ppt.bible import biblang
from service_ppt.bible.bibcore import Bible, BibleInfo, Book, Chapter, FileFormat, Verse


def expect_string(buf: str, expect: str) -> str | None:
    """Extract value from a string that starts with expected prefix.

    :param buf: String buffer to parse
    :param expect: Expected prefix string (e.g., "NAME=")
    :returns: Remaining string after prefix (stripped) or None if prefix not found
    """
    if buf.startswith(expect):
        return buf[len(expect) :].strip()

    return None


class MyBibleReader:
    """Reader for MyBible format files.

    This class provides functionality to read Bible text from MyBible format,
    which consists of an index.txt file and multiple book<DD>.txt files.
    """

    def __init__(self) -> None:
        """Initialize MyBible reader."""
        self.dirname: str | None = None
        self.remove_chars: str | None = None

    def _get_extension(self) -> str:
        """Get the file extension for MyBible format.

        :returns: File extension string (".txt")
        """
        return ".txt"

    def read_bible(self, dirname: str, load_all: bool = False, remove_chars: str | None = None) -> Bible | None:
        """Read Bible data from MyBible format directory.

        :param dirname: Directory containing MyBible format files
        :param load_all: Whether to load all books immediately (defaults to lazy loading)
        :param remove_chars: Characters to remove from verse text
        :returns: Bible object or None if format is invalid
        """
        bible = None

        self.dirname = dirname
        self.remove_chars = remove_chars

        index_name = os.path.join(dirname, "index.txt")
        encoding = None
        try:
            encoding = biblang.detect_encoding(index_name)
        except FileNotFoundError:
            return None

        with open(index_name, encoding=encoding) as file:
            line = file.readline()
            if line[0] == biblang.UNICODE_BOM:
                line = line[1:]

            line = line.strip()
            if line != "INDEX FILE":
                return None

            bible = Bible()

            line = file.readline()
            bible.name = expect_string(line, "NAME=")
            if not bible.name:
                return None

            line = file.readline()
            book_count = expect_string(line, "BOOKCOUNT=")
            if not book_count:
                return None
            book_count = int(book_count)

            use_previous_line = False
            line = file.readline()
            lang = expect_string(line, "LANGUAGE=")
            if not lang:
                lang = biblang.LANG_EN
                use_previous_line = True

            if not load_all:
                bible.reader = self

            for b in range(book_count):
                if use_previous_line:
                    use_previous_line = False
                else:
                    line = file.readline()

                book_name = expect_string(line, "BOOK=")
                if not book_name:
                    return None
                book_names = book_name.split(",")

                line = file.readline()
                chapter_indices = expect_string(line, "CHAPTERS=")
                if not chapter_indices:
                    return None
                chapter_indices = chapter_indices.split(" ")

                book = Book()
                book.new_testament = BibleInfo.is_new_testament(b)
                book.name = book_names[0]
                if len(book_names) >= 2:
                    book.short_name = book_names[1]
                else:
                    book.short_name = biblang.L18N.get_short_book_name(b, lang=lang)

                bible.books.append(book)

                if load_all:
                    self.read_book(book, b)

        bible.ensure_loaded(bible.books[0])
        bible.lang = biblang.detect_language(bible.books[0].chapters[0].verses[0].text)

        return bible

    def read_book(self, book: Book, book_no: int) -> None:
        """Read a single book from MyBible format.

        :param book: Book object to populate with data
        :param book_no: Zero-based book number
        """
        extension = self._get_extension()
        bookname = f"book{book_no + 1}{extension}"
        book_filename = os.path.join(self.dirname, bookname)
        encoding = biblang.detect_encoding(book_filename)
        with open(book_filename, encoding=encoding) as bf:
            self._read_book_file(bf, book)

    def _read_book_file(self, file: TextIO, book: Book) -> None:
        """Parse book file and populate book with chapters and verses.

        :param file: Text file object containing book data
        :param book: Book object to populate
        """
        re_numbers = re.compile(r"^(\d+)(\-\d+)?(.*)")

        chapter = None
        c = 1
        for i, line in enumerate(file):
            if i == 0 and line[0] == biblang.UNICODE_BOM:
                line = line[1:]

            line = line.strip()
            if not line:
                chapter = None
                continue

            v1 = None
            no = None
            text = None
            m = re_numbers.match(line)
            if m:
                v1 = m.group(1)
                v2 = None
                no = v1
                if m.group(2):
                    v2 = m.group(2)[1:]
                    no = v1 + "-" + v2
                text = m.group(3).strip()
            else:
                text = line

            if chapter is None:
                chapter = Chapter()
                chapter.no = c
                book.chapters.append(chapter)

                c = c + 1

            if self.remove_chars:
                text = text.replace(self.remove_chars, "")

            verse = Verse()
            verse.set_no(no)
            verse.text = text
            chapter.verses.append(verse)


class MyBibleWriter:
    """Writer for MyBible format files.

    This class provides functionality to write Bible data to MyBible format,
    creating index.txt and book<DD>.txt files.
    """

    def __init__(self) -> None:
        """Initialize MyBible writer."""
        pass

    def _get_extension(self) -> str:
        """Get the file extension for MyBible format.

        :returns: File extension string (".txt")
        """
        return ".txt"

    def write_bible(self, dirname: str, bible: Bible, encoding: str = "utf-8") -> None:
        """Write Bible data to MyBible format files.

        :param dirname: Directory where MyBible files will be written
        :param bible: Bible object to write
        :param encoding: Character encoding (defaults to UTF-8)
        """
        extension = self._get_extension()

        indices = []
        for i, b in enumerate(bible.books):
            bible.ensure_loaded(b)

            filename = f"book{i + 1}{extension}"
            with open(os.path.join(dirname, filename), "wt", encoding=encoding) as file:
                chapter_indices = self._write_book(file, b)
                indices.append(chapter_indices)

        filename = f"index{extension}"
        with open(os.path.join(dirname, filename), "wt", encoding=encoding) as file:
            self._write_index(file, bible, indices)

    def _write_index(self, file: object, bible: Bible, indices: list[list[int]]) -> None:
        """Write index.txt file with Bible metadata and chapter offsets.

        :param file: File object to write to
        :param bible: Bible object containing metadata
        :param indices: List of chapter offset indices for each book
        """
        print(biblang.UNICODE_BOM, file=file, end="")
        print("INDEX FILE", file=file)

        print(f"NAME={bible.name}", file=file)

        book_count = len(bible.books)
        print(f"BOOKCOUNT={book_count}", file=file)

        print(f"LANGUAGE={bible.lang}", file=file)

        for i, book in enumerate(bible.books):
            if book.short_name:
                print(f"BOOK={book.name},{book.short_name}", file=file)
            else:
                print(f"BOOK={book.name}", file=file)

            chapter_indices = indices[i]
            string = ""
            for ind in chapter_indices:
                if string:
                    string = string + " "
                string = string + format(ind, "X")

            print(f"CHAPTERS={string}", file=file)

    def _write_book(self, file: object, book: Book) -> list[int]:
        """Write book file and return chapter offset indices.

        :param file: File object to write to
        :param book: Book object to write
        :returns: List of file positions for each chapter start
        """
        chapter_indices = []

        # BOM
        print(biblang.UNICODE_BOM, file=file, end="")

        for chapter in book.chapters:
            chapter_indices.append(file.tell())
            for verse in chapter.verses:
                verse_no = verse.no
                if not verse_no:
                    verse_no = ""
                else:
                    verse_no = str(verse_no) + "\t"
                print(f"{verse_no}{verse.text}", file=file)

            # blank line between each chapter
            print("\n", file=file)

        return chapter_indices


class MyBibleFormat(FileFormat):
    """File format handler for MyBible format.

    This class implements the FileFormat interface for reading and writing
    MyBible format Bible files.
    """

    def __init__(self) -> None:
        """Initialize MyBible format handler."""
        super().__init__()

        self.options: dict[str, str | bool] = {"ROOT_DIR": "", "remove_special_chars": True}

    def _get_root_dir(self) -> str:
        """Get the root directory for MyBible format files.

        :returns: Path to the root directory containing MyBible format files
        """
        dirname = self.get_option("ROOT_DIR")
        if not dirname:
            dirname = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "..", "Bible.text")

        return os.path.normpath(dirname)

    def enum_versions(self) -> list[str]:
        """Enumerate available Bible versions in MyBible format.

        :returns: List of available version names
        """
        dirname = self._get_root_dir()

        versions = []
        try:
            for d in os.listdir(dirname):
                if os.path.isdir(os.path.join(dirname, d)):
                    if os.path.exists(os.path.join(dirname, d, "index.txt")):
                        versions.append(d)
        except FileNotFoundError:
            # Directory doesn't exist, skip it
            pass

        return versions

    def read_version(self, version: str) -> Bible | None:
        """Read a specific Bible version from MyBible format.

        :param version: Version name to read
        :returns: Bible object or None if version not found
        """
        dirname = os.path.join(self._get_root_dir(), version)

        reader = MyBibleReader()
        remove_chars = None
        if "remove_special_chars" in self.options:
            remove_chars = "○"
        return reader.read_bible(dirname, remove_chars=remove_chars)
