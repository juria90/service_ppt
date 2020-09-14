'''This file supports reading "MySword" Bible modules, which is a sqlite3 file format.
https://www.mysword.info/modules-format

The module also can be downloaded by using "bgmysword" program that reads text from http://biblegateway.com.
https://github.com/GreenRaccoon23/bgmysword
'''

import os
import re
import sqlite3

from bibcore import BibleInfo, Verse, Chapter, Book, Bible, FileFormat
import biblang


class MySwordReader:

    def __init__(self):
        self.conn = None
        self.cursor = None
        self.remove_tags = False

    def read_bible(self, conn, version, remove_tags, load_all=False) -> Bible:
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

    def read_book(self, book, book_no):
        self._parse_chapters(book, self.cursor, book_no)

    def _parse_chapters(self, book, cursor, book_no):
        for chapter_no in range(self._get_chapter_count(cursor, book_no)):
            chapter = Chapter()
            chapter.no = chapter_no + 1
            book.chapters.append(chapter)

            cursor.execute(f"SELECT Verse, Scripture FROM Bible WHERE Book='{book_no+1}' and Chapter='{chapter.no}';")
            verses = cursor.fetchall()
            for v in verses:
                v_text = v[1]
                # Treat 'Title Start' at the beginning as another verse.
                if v_text.startswith('<TS>'):
                    index = v_text.index('<Ts>')
                    if index != -1:
                        verse = Verse()
                        verse.text = v_text[4:index]
                        chapter.verses.append(verse)

                        v_text = v_text[index+4:]

                verse = Verse()
                verse.set_no(v[0])
                verse.text = self._cleanup_text(v_text)
                chapter.verses.append(verse)

    def _cleanup_text(self, text):
        if self.remove_tags:
            text = re.sub('<CI>', ' ', text)
            text = re.sub('<(CM|FI|Fi|FO|Fo|FR|Fr|FU|Fu|PF[0-7]|PI[0-7])>', '', text)
            text = re.sub('<RF>.*<Rf>', '', text)
            # Ignore 'Title Start' in the middle.
            text = re.sub('<TS>.*<Ts>', '', text)

        text = text.strip()

        return text

    def _has_book(self, cursor, book):
        cursor.execute(f"SELECT Book FROM Bible WHERE Book='{book}' and Chapter='1' and Verse='1';")
        rows = cursor.fetchall()
        return len(rows) > 0

    def _get_book_count(self, cursor):
        book_count = BibleInfo.get_book_count()
        if self._has_book(cursor, book_count):
            while True:
                if not self._has_book(cursor, book_count + 1):
                    return book_count
                book_count = book_count + 1
        else:
            while True:
                if self._has_book(cursor, book_count - 1):
                    return book_count - 1
                book_count = book_count - 1

        return 0

    def _has_chapter(self, cursor, book, chapter):
        cursor.execute(f"SELECT Chapter FROM Bible WHERE Book='{book+1}' and Chapter='{chapter}' and Verse='1';")
        rows = cursor.fetchall()
        return len(rows) > 0

    def _get_chapter_count(self, cursor, book):
        chapter_count = BibleInfo.get_chapter_count(book)
        if self._has_chapter(cursor, book, chapter_count):
            while True:
                if not self._has_chapter(cursor, book, chapter_count + 1):
                    return chapter_count
                chapter_count = chapter_count + 1
        else:
            while True:
                if self._has_chapter(cursor, book, chapter_count - 1):
                    return chapter_count - 1
                chapter_count = chapter_count - 1

        return 0

class MySwordFormat(FileFormat):
    def __init__(self):
        super().__init__()

        self.versions = None
        self.options = {'ROOT_DIR': '',
                        'remove_bible_tags': None}

    def _get_root_dir(self):
        dirname = self.get_option('ROOT_DIR')
        if not dirname:
            dirname = os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', '..', 'bgmysword')

        dirname = os.path.normpath(dirname)

        return dirname

    @staticmethod
    def _get_bible_name(filename):
        '''Check whether it is a sqlite3 file with "Details" and "Bible" tables.

        return biblename if it is a valid bible.
        '''
        desc = None
        with sqlite3.connect(filename) as conn:
            cursor = conn.cursor()
            cursor.execute("SELECT name FROM sqlite_master WHERE type='table';")
            tables = cursor.fetchall()
            count = 0
            for t in tables:
                if t[0] in ('Details', 'Bible'):
                    count = count + 1

            if count != 2:
                return None

            cursor.execute("select Abbreviation from Details;")
            rows = cursor.fetchall()

            if len(rows) == 1 and len(rows[0]):
                desc = rows[0][0]

        return desc

        @staticmethod
        def _get_column_names(cursor, tablename):
            '''Return list of columns from "tablename".
            0|column_name|varchar|0||0
            '''
            cursor.execute(f"pragma table_info({tablename});")
            columns = cursor.fetchall()
            return [c[1] for c in columns]

    def enum_versions(self):
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

    def read_version(self, version):
        if self.versions is None:
            self.enum_versions()

        if version not in self.versions:
            return None

        filename = self.versions[version]
        conn = sqlite3.connect(filename)

        reader = MySwordReader()
        remove_tags = False
        if self.options['remove_bible_tags']:
            remove_tags = True
        return reader.read_bible(conn, version, remove_tags)
