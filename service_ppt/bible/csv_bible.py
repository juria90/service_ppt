"""CSV Bible format reader.

This module provides support for reading Bible text from CSV (Comma-Separated Values)
format files.
"""

import csv
import os

from service_ppt.bible.bibcore import Bible


class CSVWriter:
    """Writer for CSV Bible format.

    This class provides functionality to write Bible data to CSV format files,
    creating separate files for books and verses.
    """

    def _get_extension(self) -> str:
        """Get the file extension for CSV format.

        :returns: File extension string (".csv")
        """
        return ".csv"

    def write_bible(self, dirname: str, bible: Bible, encoding: str | None = None) -> None:
        """Write Bible data to CSV format files.

        :param dirname: Directory where CSV files will be written
        :param bible: Bible object to write
        :param encoding: Character encoding (defaults to UTF-8 with BOM if UTF-8)
        """
        if encoding and encoding.lower() == "utf-8":
            encoding = "utf-8-sig"
        booksname = os.path.join(dirname, "books.csv")
        self._write_books(booksname, bible, encoding)

        versesname = os.path.join(dirname, "verses.csv")
        self._write_verses(versesname, bible, encoding)

    def _write_books(self, booksname: str, bible: Bible, encoding: str | None = None) -> None:
        """Write book information to CSV file.

        :param booksname: Path to the books CSV file
        :param bible: Bible object containing book data
        :param encoding: Character encoding for the file
        """
        with open(booksname, "w", newline="", encoding=encoding) as csvfile:
            f = csv.writer(csvfile)
            for i, b in enumerate(bible.books):
                book_no = i + 1
                old_new = "2" if b.new_testament else "1"
                line = [str(book_no), old_new, b.name, b.short_name]
                f.writerow(line)

    def _write_verses(self, versesname: str, bible: Bible, encoding: str | None = None) -> None:
        """Write verse data to CSV file.

        :param versesname: Path to the verses CSV file
        :param bible: Bible object containing verse data
        :param encoding: Character encoding for the file
        """
        with open(versesname, "w", newline="", encoding=encoding) as csvfile:
            f = csv.writer(csvfile)
            for b, book in enumerate(bible.books):
                bible.ensure_loaded(book)

                for c, chapter in enumerate(book.chapters):
                    for _v, verse in enumerate(chapter.verses):
                        line = [str(b + 1), str(c + 1), verse.no, verse.text]
                        f.writerow(line)
