"""OpenSong Bible format reader.

This module provides support for reading Bible text from OpenSong format files.
OpenSong is a free software application for managing chords and lyrics.
"""

import os
import xml.sax.saxutils

from service_ppt.bible.bibcore import Bible


class OpenSongXMLWriter:
    def _get_extension(self) -> str:
        return ".xml"

    def write_bible(self, dirname: str, bible: Bible, encoding: str = "utf-8") -> None:
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
        print(
            f"""<?xml version="1.0" encoding="{encoding}"?>
<bible>
""",
            file=file,
            end="",
        )

    def _write_xml_footer(self, file: object) -> None:
        print(
            """</bible>
""",
            file=file,
            end="",
        )
