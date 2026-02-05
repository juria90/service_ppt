"""OpenSong Bible format reader.

This module provides support for reading Bible text from OpenSong format files.
OpenSong is a free software application for managing chords and lyrics.

OpenSong project: http://www.opensong.org/
"""

import os
import xml.sax.saxutils

from service_ppt.bible.bibcore import Bible


class OpenSongXMLWriter:
    """Writer for OpenSong XML Bible format.

    This class provides functionality to write Bible data to OpenSong XML format,
    which is used by the OpenSong application for managing Bible references.

    The format uses a simple XML structure with <bible>, <b>, <c>, and <v> tags
    for books, chapters, and verses respectively.
    """

    def _get_extension(self) -> str:
        """Get the file extension for XML format.

        :returns: File extension string (".xml")
        """
        return ".xml"

    def write_bible(self, dirname: str, bible: Bible, encoding: str = "utf-8") -> None:
        """Write Bible data to OpenSong XML format file.

        :param dirname: Directory where XML file will be written
        :param bible: Bible object to write
        :param encoding: Character encoding (defaults to UTF-8)
        """
        if encoding is None:
            encoding = "utf-8"

        filename = os.path.join(dirname, "bible.xml")
        with open(os.path.join(dirname, filename), "w", encoding=encoding) as file:
            self._write_xml_header(file, encoding)

            for book in bible.books:
                bible.ensure_loaded(book)

                print(f' <b n="{book.name}">', file=file)

                for chapter in book.chapters:
                    print(f'  <c n="{chapter.no}">', file=file)

                    for verse in chapter.verses:
                        verse_no = verse.no
                        if verse_no is None:
                            verse_no = ""
                        text = xml.sax.saxutils.escape(verse.text)
                        print(f'   <v n="{verse_no}">{text}</v>', file=file)

                    print("  </c>", file=file)

                print(" </b>", file=file)

            self._write_xml_footer(file)

    def _write_xml_header(self, file: object, encoding: str) -> None:
        """Write XML header to file.

        :param file: File object to write to
        :param encoding: Character encoding for the XML declaration
        """
        print(
            f"""<?xml version="1.0" encoding="{encoding}"?>
<bible>
""",
            file=file,
            end="",
        )

    def _write_xml_footer(self, file: object) -> None:
        """Write XML footer to file.

        :param file: File object to write to
        """
        print(
            """</bible>
""",
            file=file,
            end="",
        )
