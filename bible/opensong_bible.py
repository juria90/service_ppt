"""
"""

import os
import xml.sax.saxutils


class OpenSongXMLWriter:
    def _get_extension(self):
        return ".xml"

    def write_bible(self, dirname, bible, encoding="utf-8"):
        if encoding == None:
            encoding = "utf-8"

        filename = os.path.join(dirname, "bible.xml")
        with open(os.path.join(dirname, filename), "wt", encoding=encoding) as file:
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

                    print(f"  </c>", file=file)

                print(f" </b>", file=file)

            self._write_xml_footer(file)

    def _write_xml_header(self, file, encoding):
        print(
            f"""<?xml version="1.0" encoding="{encoding}"?>
<bible>
""",
            file=file,
            end="",
        )

    def _write_xml_footer(self, file):
        print(
            f"""</bible>
""",
            file=file,
            end="",
        )
