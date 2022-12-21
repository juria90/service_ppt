#!/usr/bin/env python
"""
"""
from wx.core import NO

from .biblang import L18N


class BibleInfo:
    """BibleInfo contains the King James Version's bible information
    regarding number of books, chapter.
    """

    BOOK_COUNT = 66
    OLD_TESTAMENT_COUNT = 39
    CHAPTER_COUNT = [
        # old testament
        50,
        40,
        27,
        36,
        34,
        24,
        21,
        4,
        31,
        24,
        22,
        25,
        29,
        36,
        10,
        13,
        10,
        41,
        150,
        31,
        12,
        8,
        66,
        52,
        5,
        48,
        12,
        14,
        3,
        9,
        1,
        4,
        7,
        3,
        3,
        3,
        2,
        14,
        4,
        # new testament
        28,
        16,
        24,
        21,
        28,
        16,
        16,
        13,
        6,
        6,
        4,
        4,
        5,
        3,
        6,
        4,
        3,
        1,
        13,
        5,
        5,
        3,
        5,
        1,
        1,
        1,
        22,
    ]

    @staticmethod
    def get_book_count():
        return BibleInfo.BOOK_COUNT

    @staticmethod
    def get_chapter_count(book_no):
        return BibleInfo.CHAPTER_COUNT[book_no]

    @staticmethod
    def is_old_testament(book_no):
        return book_no < BibleInfo.OLD_TESTAMENT_COUNT

    @staticmethod
    def is_new_testament(book_no):
        return book_no >= BibleInfo.OLD_TESTAMENT_COUNT


class Verse:
    """Verse class contains verse number and text."""

    def __init__(self):
        self.no = None  # Can be a string based a single number('1') or two (i.e. '2-3')
        # with 1 based index
        # Or None for description or comments.
        self.text = None

        self._number1 = None
        self._number2 = None

        self.chapter = None  # link to parent chapter for extract_texts()
        self.book = None  # link to parent book for extract_texts()

    def set_no(self, no):
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

    def in_range(self, v1, v2):
        if self.no is None:
            return False

        if self._number2 is not None:
            if v2 is not None:
                return Verse._intersect_two(v1, v2, self._number1, self._number2)
            else:
                return Verse._intersect(v1, self._number1, self._number2)
        else:
            if v2 is not None:
                return Verse._intersect(self._number1, v1, v2)
            else:
                return v1 == self._number1

    def get_max_no(self):
        if self.no is None:
            return None

        if self._number2 is not None:
            return self._number2
        else:
            return self._number1

    @staticmethod
    def _intersect(v, v1, v2):
        return v1 <= v and v <= v2

    @staticmethod
    def _intersect_two(u1, u2, v1, v2):
        if v1 < u1:
            return v2 >= u1
        elif v1 <= u2:
            return True
        else:
            return False


class Chapter:
    """Chapter class contains list of verses."""

    def __init__(self):
        self.no = None  # int start from 1
        self.verses = []


class Book:
    """Book class contains list of chapters."""

    def __init__(self):
        self.new_testament = False  # either old or new testament
        self.name = None
        self.short_name = None
        self.chapters = []

    def is_loaded(self):
        return len(self.chapters) > 0


class Bible:
    """Bible class contains list of books."""

    def __init__(self):
        self.lang = None  # ISO 639-1 codes
        self.name = None
        self.books = []

        self.reader = None
        self.book_to_index_map = None

    def ensure_loaded(self, book):
        """The ensure_loaded() loads book text, if it is not loaded yet.
        By providing this, each book can be delay loaded.
        """

        if not book.is_loaded():
            book_no = self.books.index(book)
            if book_no != -1:
                self.reader.read_book(book, book_no)

    def get_book_to_index_map(self):
        if self.book_to_index_map is None:
            self.book_to_index_map = {}
            for i, b in enumerate(self.books):
                self.book_to_index_map[b.name] = i
                self.book_to_index_map[b.short_name] = i

        return self.book_to_index_map

    def translate_to_bible_index(self, text_range):
        """translate_to_bible_index() returns Book index, chapter, verse1/verse2 tuple.

        The text_range should be formatted as <Book> <Chapter>:<Verse1>[-<Verse2>],
        where Book can be long or short name, Chapter and Verse1/Verse2 are valid numbers.
        """
        bt, ct1, v1t, ct2, v2t = L18N.parse_verse_range(self.lang, text_range)

        b2i_map = self.get_book_to_index_map()
        if bt in b2i_map:
            bi = b2i_map[bt]

            return bi, ct1, v1t, ct2, v2t

        return None

    def extract_texts_from_bible_index(self, bi, ct1, v1t, ct2, v2t):
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
                elif c.no == ct2:
                    for v in c.verses:
                        if v.in_range(1, v2t):
                            v.chapter = c
                            v.book = book
                            verses.append(v)
                elif c.no == ct1:
                    vmax = max([v.get_max_no() for v in c.verses if v.get_max_no()])
                    for v in c.verses:
                        if v.in_range(v1t, vmax):
                            v.chapter = c
                            v.book = book
                            verses.append(v)
                else:  # c.no < ct2
                    vmax = max([v.get_max_no() for v in c.verses if v.get_max_no()])
                    for v in c.verses:
                        if v.in_range(1, vmax):
                            v.chapter = c
                            v.book = book
                            verses.append(v)
            else:  # if ct2 is None
                for v in c.verses:
                    if v.in_range(v1t, v2t):
                        v.chapter = c
                        v.book = book
                        verses.append(v)

                if ct2 is None:
                    break

        if len(verses) == 0:
            return None

        return verses

    def extract_texts(self, text_range):
        """extract_texts() returns list of Verse within given text_range.

        The text_range should be formatted as <Book> <Chapter>:<Verse1>[-<Verse2>],
        where Book can be long or short name, Chapter and Verse1/Verse2 are valid numbers.
        """

        result = self.translate_to_bible_index(text_range)
        if result is None:
            return None

        bi, ct1, v1t, ct2, v2t = result
        verses = self.extract_texts_from_bible_index(bi, ct1, v1t, ct2, v2t)
        return verses


class FileFormat:
    """FileFormat class provides handling different file format for the Bible text."""

    def __init__(self):
        self.options = {}

    def get_option(self, key):
        return self.options[key]

    def set_option(self, key, value):
        if key in self.options:
            self.options[key] = value

    def get_options(self):
        return self.options

    def set_options(self, options):
        self.options = options

    def enum_versions(self):
        return []

    def read_version(self, _version):
        return None
