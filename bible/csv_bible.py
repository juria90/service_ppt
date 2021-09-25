"""
"""

import csv
import os


class CSVWriter:
    def _get_extension(self):
        return ".csv"

    def write_bible(self, dirname, bible, encoding=None):
        if encoding and encoding.lower() == "utf-8":
            encoding = "utf-8-sig"
        booksname = os.path.join(dirname, "books.csv")
        self._write_books(booksname, bible, encoding)

        versesname = os.path.join(dirname, "verses.csv")
        self._write_verses(versesname, bible, encoding)

    def _write_books(self, booksname, bible, encoding=None):
        with open(booksname, "wt", newline="", encoding=encoding) as csvfile:
            f = csv.writer(csvfile)
            for i, b in enumerate(bible.books):
                book_no = i + 1
                old_new = "2" if b.new_testament else "1"
                line = [str(book_no), old_new, b.name, b.short_name]
                f.writerow(line)

    def _write_verses(self, versesname, bible, encoding=None):
        with open(versesname, "wt", newline="", encoding=encoding) as csvfile:
            f = csv.writer(csvfile)
            for b, book in enumerate(bible.books):
                bible.ensure_loaded(book)

                for c, chapter in enumerate(book.chapters):
                    for _v, verse in enumerate(chapter.verses):
                        line = [str(b + 1), str(c + 1), verse.no, verse.text]
                        f.writerow(line)
