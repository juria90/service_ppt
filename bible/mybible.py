'''This file supports reading "MyBible" Bible module that has index.txt and book<DD>.txt.

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
'''

import os
import re

from bibcore import Verse, Chapter, Book, Bible, BibleInfo, FileFormat
import biblang


def expect_string(buf, expect):
    if buf.startswith(expect):
        return buf[len(expect):].strip()

    return None


class MyBibleReader:
    def __init__(self):
        self.dirname = None
        self.remove_chars = None

    def _get_extension(self):
        return '.txt'

    def read_bible(self, dirname, load_all=False, remove_chars=None) -> Bible:
        bible = None

        self.dirname = dirname
        self.remove_chars = remove_chars

        index_name = os.path.join(dirname, 'index.txt')
        encoding = None
        try:
            encoding = biblang.detect_encoding(index_name)
        except FileNotFoundError:
            return None

        with open(index_name, encoding=encoding) as file:
            line = file.readline()
            if biblang.UNICODE_BOM == line[0]:
                line = line[1:]

            line = line.strip()
            if 'INDEX FILE' != line:
                return None

            bible = Bible()

            line = file.readline()
            bible.name = expect_string(line, 'NAME=')
            if not bible.name:
                return None

            line = file.readline()
            book_count = expect_string(line, 'BOOKCOUNT=')
            if not book_count:
                return None
            book_count = int(book_count)

            use_previous_line = False
            line = file.readline()
            lang = expect_string(line, 'LANGUAGE=')
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

                book_name = expect_string(line, 'BOOK=')
                if not book_name:
                    return None
                book_names = book_name.split(',')

                line = file.readline()
                chapter_indices = expect_string(line, 'CHAPTERS=')
                if not chapter_indices:
                    return None
                chapter_indices = chapter_indices.split(' ')

                book = Book()
                book.new_testament = BibleInfo.is_new_testament(b)
                book.name = book_names[0]
                if len(book_names) >= 2:
                    book.short_name = book_names[1]
                else:
                    book.short_name = biblang.L18N.get_short_book_name(
                        b, lang=lang)

                bible.books.append(book)

                if load_all:
                    self.read_book(book, b)

        bible.ensure_loaded(bible.books[0])
        bible.lang = biblang.detect_language(
            bible.books[0].chapters[0].verses[0].text)

        return bible

    def read_book(self, book, book_no):
        extension = self._get_extension()
        bookname = f'book{book_no+1}{extension}'
        book_filename = os.path.join(self.dirname, bookname)
        encoding = biblang.detect_encoding(book_filename)
        with open(book_filename, encoding=encoding) as bf:
            self._read_book_file(bf, book)

    def _read_book_file(self, file, book):
        chapter = None

        c = 1
        for i, line in enumerate(file):
            if i == 0 and biblang.UNICODE_BOM == line[0]:
                line = line[1:]

            line = line.strip()
            if not line:
                chapter = None
                continue

            v1 = None
            no = None
            text = None
            m = re.match(r'^(\d+)(\-\d+)?(.*)', line)
            if m:
                v1 = m.group(1)
                v2 = None
                no = v1
                if m.group(2):
                    v2 = m.group(2)[1:]
                    no = v1 + '-' + v2
                text = m.group(3).strip()
            else:
                text = line

            if chapter is None:
                chapter = Chapter()
                chapter.no = c
                book.chapters.append(chapter)

                c = c + 1

            if self.remove_chars:
                text = text.replace(self.remove_chars, '')

            verse = Verse()
            verse.set_no(no)
            verse.text = text
            chapter.verses.append(verse)


class MyBibleWriter:
    def __init__(self):
        pass

    def _get_extension(self):
        return '.txt'

    def write_bible(self, dirname, bible, encoding='utf-8'):
        extension = self._get_extension()

        indices = []
        for i, b in enumerate(bible.books):
            bible.ensure_loaded(b)

            filename = f'book{i+1}{extension}'
            with open(os.path.join(dirname, filename), 'wt', encoding=encoding) as file:
                chapter_indices = self._write_book(file, b)
                indices.append(chapter_indices)

        filename = f'index{extension}'
        with open(os.path.join(dirname, filename), 'wt', encoding=encoding) as file:
            self._write_index(file, bible, indices)

    def _write_index(self, file, bible, indices):

        print(biblang.UNICODE_BOM, file=file, end='')
        print('INDEX FILE', file=file)

        print(f'NAME={bible.name}', file=file)

        book_count = len(bible.books)
        print(f'BOOKCOUNT={book_count}', file=file)

        print(f'LANGUAGE={bible.lang}', file=file)

        for i, book in enumerate(bible.books):
            if book.short_name:
                print(f'BOOK={book.name},{book.short_name}', file=file)
            else:
                print(f'BOOK={book.name}', file=file)

            chapter_indices = indices[i]
            string = ''
            for ind in chapter_indices:
                if string:
                    string = string + ' '
                string = string + format(ind, 'X')

            print(f'CHAPTERS={string}', file=file)

    def _write_book(self, file, book):
        chapter_indices = []

        # BOM
        print(biblang.UNICODE_BOM, file=file, end='')

        for chapter in book.chapters:
            chapter_indices.append(file.tell())
            for verse in chapter.verses:
                verse_no = verse.no
                if not verse_no:
                    verse_no = ''
                else:
                    verse_no = str(verse_no) + '\t'
                print(f'{verse_no}{verse.text}', file=file)

            # blank line between each chapter
            print('\n', file=file)

        return chapter_indices


class MyBibleFormat(FileFormat):
    def __init__(self):
        super().__init__()

        self.options = {'ROOT_DIR': '',
                        'remove_special_chars': True}

    def _get_root_dir(self):
        dirname = self.get_option('ROOT_DIR')
        if not dirname:
            dirname = os.path.join(os.path.dirname(
                os.path.abspath(__file__)), '..', '..', 'Bible.text')

        dirname = os.path.normpath(dirname)

        return dirname

    def enum_versions(self):
        dirname = self._get_root_dir()

        versions = []
        for d in os.listdir(dirname):
            if os.path.isdir(os.path.join(dirname, d)):
                if os.path.exists(os.path.join(dirname, d, 'index.txt')):
                    versions.append(d)

        return versions

    def read_version(self, version):
        dirname = os.path.join(self._get_root_dir(), version)

        reader = MyBibleReader()
        remove_chars = None
        if 'remove_special_chars' in self.options:
            remove_chars = '○'
        bible = reader.read_bible(dirname, remove_chars=remove_chars)

        return bible
