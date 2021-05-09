"""This file supports reading Sword Bible module by using Python sword_bible module.
https://pypi.org/project/pysword/
"""

from .bibcore import BibleInfo, Verse, Chapter, Book, Bible, FileFormat
from . import biblang


class SwordReader:
    def __init__(self):
        self.sw_bible = None

    def read_bible(self, sw_bible, version, load_all=False) -> Bible:
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

    def read_book(self, book, book_no):
        ot_nt = "nt" if book.new_testament else "ot"
        sw_books = self.sw_bible.get_structure().get_books()[ot_nt]
        if book.new_testament:
            ot_len = len(self.sw_bible.get_structure().get_books()["ot"])
            book_no = book_no - ot_len
        self._parse_chapters(book, self.sw_bible, sw_books[book_no])

    def _parse_chapters(self, book, sw_bible, sw_book):
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
    def __init__(self, modules, found_modules):
        super().__init__()

        self.modules = modules
        self.found_modules = found_modules
        self.versions = None

    def enum_versions(self):
        if self.versions is None:
            versions = []
            for m in self.found_modules:
                if "blocktype" in self.found_modules[m]:
                    if self.found_modules[m]["blocktype"] == "BOOK":
                        versions.append(m)

            self.versions = versions

        return self.versions

    def read_version(self, version):
        if self.versions is None:
            self.enum_versions()

        if version not in self.versions:
            return None

        sw_bible = self.modules.get_bible_from_module(version)
        reader = SwordReader()
        return reader.read_bible(sw_bible, version)
