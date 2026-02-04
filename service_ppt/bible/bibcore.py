#!/usr/bin/env python
"""Core Bible data structures and classes.

This module defines the core data structures (Verse, Chapter, Book, Bible) and
the FileFormat interface for reading Bible text from various formats.
"""

from typing import Any, ClassVar

from service_ppt.bible.biblang import L18N


class BibleInfo:
    """BibleInfo contains the King James Version's bible information
    regarding number of books, chapter.
    """

    BOOK_COUNT: ClassVar[int] = 66
    OLD_TESTAMENT_COUNT: ClassVar[int] = 39
    # fmt: off
    CHAPTER_COUNT: ClassVar[list[int]] = [
        # old testament
        50, 40, 27, 36, 34, 24, 21, 4, 31, 24,
        22, 25, 29, 36, 10, 13, 10, 41, 150, 31,
        12, 8, 66, 52, 5, 48, 12, 14, 3, 9,
        1, 4, 7, 3, 3, 3, 2, 14, 4,
        # new testament
        28, 16, 24, 21, 28, 16, 16, 13, 6, 6,
        4, 4, 5, 3, 6, 4, 3, 1, 13, 5,
        5, 3, 5, 1, 1, 1, 22,
    ]
    # fmt: on

    @staticmethod
    def get_book_count() -> int:
        return BibleInfo.BOOK_COUNT

    @staticmethod
    def get_chapter_count(book_no: int) -> int:
        return BibleInfo.CHAPTER_COUNT[book_no]

    @staticmethod
    def is_old_testament(book_no: int) -> bool:
        return book_no < BibleInfo.OLD_TESTAMENT_COUNT

    @staticmethod
    def is_new_testament(book_no: int) -> bool:
        return book_no >= BibleInfo.OLD_TESTAMENT_COUNT


class Verse:
    """Verse class contains verse number and text."""

    def __init__(self) -> None:
        self.no: str | int | None = None  # Can be a string based a single number('1') or two (i.e. '2-3')
        # with 1 based index
        # Or None for description or comments.
        self.text: str | None = None

        self._number1: int = 0
        self._number2: int | None = None

        self.chapter: Chapter | None = None  # link to parent chapter for extract_texts()
        self.book: Book | None = None  # link to parent book for extract_texts()

    def set_no(self, no: str | int | None) -> None:
        self.no = no
        if self.no is not None:
            if isinstance(self.no, int):
                self._number1 = self.no
            elif isinstance(self.no, str):
                index = self.no.find("-")
                if index == -1:
                    index = self.no.find(":")

                if index != -1:
                    self._number1 = int(self.no[:index])
                    self._number2 = int(self.no[index + 1 :])

                    # normalize concatenated multiline verse.
                    if self.no[index] == ":":
                        self.no = self.no[:index] + "-" + self.no[index + 1 :]
                else:
                    self._number1 = int(self.no)

    def in_range(self, v1: int, v2: int | None) -> bool:
        if self.no is None:
            return False

        if self._number2 is not None:
            if v2 is not None:
                return Verse._intersect_two(v1, v2, self._number1, self._number2)
            return Verse._intersect(v1, self._number1, self._number2)
        if v2 is not None:
            return Verse._intersect(self._number1, v1, v2)
        return v1 == self._number1

    def get_max_no(self) -> int | None:
        if self.no is None:
            return None

        if self._number2 is not None:
            return self._number2
        return self._number1

    @staticmethod
    def _intersect(v: int, v1: int, v2: int) -> bool:
        return v1 <= v and v <= v2

    @staticmethod
    def _intersect_two(u1: int, u2: int, v1: int, v2: int) -> bool:
        if v1 < u1:
            return v2 >= u1
        return v1 <= u2


class Chapter:
    """Chapter class contains list of verses."""

    def __init__(self) -> None:
        self.no: int | None = None  # int start from 1
        self.verses: list[Verse] = []


class Book:
    """Book class contains list of chapters."""

    def __init__(self) -> None:
        self.new_testament: bool = False  # either old or new testament
        self.name: str | None = None
        self.short_name: str | None = None
        self.chapters: list[Chapter] = []

    def is_loaded(self) -> bool:
        return len(self.chapters) > 0


class Bible:
    """Bible class contains list of books."""

    def __init__(self) -> None:
        self.lang: str | None = None  # ISO 639-1 codes
        self.name: str | None = None
        self.books: list[Book] = []

        self.reader: Any | None = None
        self.book_to_index_map: dict[str, int] | None = None

    def ensure_loaded(self, book: Book) -> None:
        """The ensure_loaded() loads book text, if it is not loaded yet.
        By providing this, each book can be delay loaded.
        """

        if not book.is_loaded():
            book_no = self.books.index(book)
            if book_no != -1:
                self.reader.read_book(book, book_no)

    def get_book_to_index_map(self) -> dict[str, int]:
        if self.book_to_index_map is None:
            self.book_to_index_map = {}
            for i, b in enumerate(self.books):
                self.book_to_index_map[b.name] = i
                self.book_to_index_map[b.short_name] = i

        return self.book_to_index_map

    def translate_to_bible_index(self, text_range: str) -> tuple[int, int, int, int | None, int | None] | None:
        """translate_to_bible_index() returns Book index, chapter, verse1/verse2 tuple.

        The text_range should be formatted as <Book> <Chapter>:<Verse1>[-<Verse2>],
        where Book can be long or short name, Chapter and Verse1/Verse2 are valid numbers.
        """
        bt, ct1, vs1, ct2, vs2 = L18N.parse_verse_range(self.lang, text_range)

        b2i_map = self.get_book_to_index_map()
        if bt in b2i_map:
            bi = b2i_map[bt]

            return bi, ct1, vs1, ct2, vs2

        return None

    def extract_texts_from_bible_index(self, bi: int, ct1: int, vs1: int, ct2: int | None, vs2: int | None) -> list[Verse] | None:
        """extract_texts() returns list of Verse within given bible index."""
        book = self.books[bi]
        verses = []

        if not book.is_loaded():
            self.reader.read_book(book, bi)

        for c in book.chapters:
            if c.no < ct1:
                continue

            # multi chapter
            if ct2 is not None:
                if c.no > ct2:
                    break
                if c.no == ct2:
                    for v in c.verses:
                        if v.in_range(1, vs2):
                            v.chapter = c
                            v.book = book
                            verses.append(v)
                elif c.no == ct1:
                    vmax = max([no for v in c.verses if (no := v.get_max_no())])
                    for v in c.verses:
                        if v.in_range(vs1, vmax):
                            v.chapter = c
                            v.book = book
                            verses.append(v)
                else:  # c.no < ct2
                    vmax = max([no for v in c.verses if (no := v.get_max_no())])
                    for v in c.verses:
                        if v.in_range(1, vmax):
                            v.chapter = c
                            v.book = book
                            verses.append(v)
            else:  # if ct2 is None
                for v in c.verses:
                    if v.in_range(vs1, vs2):
                        v.chapter = c
                        v.book = book
                        verses.append(v)

                if ct2 is None:
                    break

        if len(verses) == 0:
            return None

        return verses

    def extract_texts(self, text_range: str) -> list[Verse] | None:
        """extract_texts() returns list of Verse within given text_range.

        The text_range should be formatted as <Book> <Chapter>:<Verse1>[-<Verse2>],
        where Book can be long or short name, Chapter and Verse1/Verse2 are valid numbers.
        """

        result = self.translate_to_bible_index(text_range)
        if result is None:
            return None

        bi, ct1, vs1, ct2, vs2 = result
        return self.extract_texts_from_bible_index(bi, ct1, vs1, ct2, vs2)


class FileFormat:
    """FileFormat class provides handling different file format for the Bible text."""

    def __init__(self) -> None:
        self.options: dict[str, Any] = {}

    def get_option(self, key: str) -> Any:
        return self.options[key]

    def set_option(self, key: str, value: Any) -> None:
        if key in self.options:
            self.options[key] = value

    def get_options(self) -> dict[str, Any]:
        return self.options

    def set_options(self, options: dict[str, Any]) -> None:
        self.options = options

    def enum_versions(self) -> list[str]:
        return []

    def read_version(self, _version: str) -> "Bible | None":
        return None
