"""Sword Bible format reader.

This module provides support for reading Bible text from Sword format modules
using the pysword Python library. Sword is a cross-platform Bible study
software framework.

Library: https://pypi.org/project/pysword/
"""

from typing import TYPE_CHECKING, Any

from service_ppt.bible import biblang
from service_ppt.bible.bibcore import Bible, Book, Chapter, FileFormat, Verse

if TYPE_CHECKING:
    pass


class SwordReader:
    def __init__(self) -> None:
        self.sw_bible: "Any | None" = None

    def read_bible(self, sw_bible: "Any", version: str, load_all: bool = False) -> Bible:
        bible = Bible()
        bible.name = version
        book_no = 0

        if not load_all:
            self.sw_bible = sw_bible
            bible.reader = self

        for testament, books in sw_bible.get_structure().get_books().items():
            for sw_book in books:
                book = Book()
                book.new_testament = testament == "nt"
                book.name = sw_book.name
                book.short_name = sw_book.preferred_abbreviation
                bible.books.append(book)
                book_no = book_no + 1

                if load_all:
                    self._parse_chapters(book, sw_bible, sw_book)

        bible.ensure_loaded(bible.books[0])
        bible.lang = biblang.detect_language(bible.books[0].chapters[0].verses[0].text)

        return bible

    def read_book(self, book: Book, book_no: int) -> None:
        ot_nt = "nt" if book.new_testament else "ot"
        sw_books = self.sw_bible.get_structure().get_books()[ot_nt]
        if book.new_testament:
            ot_len = len(self.sw_bible.get_structure().get_books()["ot"])
            book_no = book_no - ot_len
        self._parse_chapters(book, self.sw_bible, sw_books[book_no])

    def _parse_chapters(self, book: Book, sw_bible: "Any", sw_book: "Any") -> None:
        for chapter_no in range(sw_book.num_chapters):
            chapter = Chapter()
            chapter.no = chapter_no + 1
            book.chapters.append(chapter)

            verses = sw_bible.get_iter(books=[book.name.lower()], chapters=chapter_no + 1)
            for i, v in enumerate(verses):
                verse = Verse()
                verse.set_no(i + 1)
                verse.text = v
                chapter.verses.append(verse)


class SwordFormat(FileFormat):
    def __init__(self, modules: "Any", found_modules: "dict[str, Any]") -> None:
        super().__init__()

        self.modules: "Any" = modules
        self.found_modules: "dict[str, Any]" = found_modules
        self.versions: list[str] | None = None

    def enum_versions(self) -> list[str]:
        if self.versions is None:
            versions = []
            for m in self.found_modules:
                if "blocktype" in self.found_modules[m]:
                    if self.found_modules[m]["blocktype"] == "BOOK":
                        versions.append(m)

            self.versions = versions

        return self.versions

    def read_version(self, version: str) -> Bible | None:
        if self.versions is None:
            self.enum_versions()

        if version not in self.versions:
            return None

        sw_bible = self.modules.get_bible_from_module(version)
        reader = SwordReader()
        return reader.read_bible(sw_bible, version)
