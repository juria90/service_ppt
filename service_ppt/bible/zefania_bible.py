"""Zefania Bible format reader.

This module provides support for reading Bible text from Zefania format files,
which store Bible content as XML. Zefania Bible modules can be downloaded
from the Zefania project.

Download location: https://sourceforge.net/projects/zefania-sharp/files/Bibles/
"""

from datetime import datetime
import os
import xml.etree.ElementTree as ET
import xml.sax.saxutils

from iso639 import Lang

from service_ppt.bible import biblang
from service_ppt.bible.bibcore import Bible, BibleInfo, Book, Chapter, FileFormat, Verse


class ZefaniaReader:
    """Reader for Zefania Bible format files.

    This class provides functionality to read Bible text from Zefania format,
    which stores Bible content as XML.
    """

    @staticmethod
    def _get_bible_name(filename: str, hint_line: int = 10) -> str | None:
        """Check whether file is a valid Zefania XML Bible file.

        :param filename: Path to XML file to check
        :param hint_line: Number of lines to read for checking (defaults to 10)
        :returns: Bible name if valid, None otherwise
        """
        lines = ""
        with open(filename, encoding="utf-8") as f:
            for _ in range(hint_line):
                lines += f.readline()

        parser = ET.XMLPullParser(["start", "end"])
        parser.feed(lines)
        for event, elem in parser.read_events():
            if event == "start" and elem.tag == "XMLBIBLE" and "biblename" in elem.attrib:
                return elem.attrib["biblename"]

        return None

    def read_bible(self, filename: str) -> Bible:
        """Read Bible data from Zefania XML file.

        :param filename: Path to Zefania XML file
        :returns: Bible object
        """
        tree = ET.parse(filename)
        root = tree.getroot()

        bible = Bible()
        bible.name = root.attrib["biblename"]
        for node in root:
            if node.tag == "INFORMATION":
                lang_node = node.find("language")
                if lang_node is not None:
                    lang_part2 = lang_node.text.strip().lower()
                    try:
                        lang = Lang(part2b=lang_part2)
                        bible.lang = lang.part1
                    except KeyError:
                        bible.lang = biblang.LANG_EN

            elif node.tag == "BIBLEBOOK":
                book = Book()
                book_no = int(node.attrib["bnumber"])
                book.new_testament = BibleInfo.is_new_testament(book_no - 1)
                if "bname" in node.attrib:
                    book.name = node.attrib["bname"]
                if "bsname" in node.attrib:
                    book.short_name = node.attrib["bsname"]
                if not book.name or not book.short_name:
                    names = biblang.L18N.get_book_names(book_no - 1, bible.lang)
                    if not book.name:
                        book.name = names[0]
                    if not book.short_name:
                        book.short_name = names[1]
                bible.books.append(book)

                self._parse_chapters(book, node)

        return bible

    def _parse_chapters(self, book: Book, book_node: ET.Element) -> None:
        """Parse chapters from book XML node.

        :param book: Book object to populate
        :param book_node: XML element containing book data
        """
        for node in book_node:
            if node.tag == "CHAPTER":
                chapter = Chapter()
                chapter.no = int(node.attrib["cnumber"])
                book.chapters.append(chapter)

                self._parse_verses(chapter, node)

    def _parse_verses(self, chapter: Chapter, chapter_node: ET.Element) -> None:
        """Parse verses from chapter XML node.

        :param chapter: Chapter object to populate
        :param chapter_node: XML element containing chapter data
        """
        for node in chapter_node:
            if node.tag == "VERS":
                verse = Verse()
                verse.set_no(node.attrib["vnumber"])
                verse.text = self._concat_children_text(node)
                chapter.verses.append(verse)

    def _concat_children_text(self, node: ET.Element) -> str:
        """Concatenate text from all child elements of an XML node.

        :param node: XML element to extract text from
        :returns: Concatenated text string
        """
        text = ""
        for it in node.itertext():
            t = it.strip()
            if t:
                if text:
                    text = text + " " + t
                else:
                    text = t

        return text


class ZefaniaWriter:
    """Writer for Zefania Bible format files.

    This class provides functionality to write Bible data to Zefania XML format.
    """

    def __init__(self) -> None:
        """Initialize Zefania writer."""
        pass

    def _get_extension(self) -> str:
        """Get the file extension for Zefania format.

        :returns: File extension string (".xml")
        """
        return ".xml"

    def write_bible(self, dirname: str, bible: Bible, encoding: str = "utf-8") -> None:
        """Write Bible data to Zefania XML format file.

        :param dirname: Directory where XML file will be written
        :param bible: Bible object to write
        :param encoding: Character encoding (defaults to UTF-8)
        """
        extension = self._get_extension()

        if encoding is None:
            encoding = "utf-8"

        filename = f"bible{extension}"
        with open(os.path.join(dirname, filename), "w", encoding=encoding) as file:
            self._write_header(file, bible)

            self._write_info(file, bible)

            for i, b in enumerate(bible.books):
                bible.ensure_loaded(b)

                self._write_book(file, i, b)

            self._write_footer(file)

    def _write_header(self, file: object, bible: Bible) -> None:
        """Write XML header with Bible metadata.

        :param file: File object to write to
        :param bible: Bible object containing metadata
        """
        print(
            f"""<?xml version="1.0" encoding="utf-8"?>
<!--Visit the online documentation for Zefania XML Markup-->
<!--http://bgfdb.de/zefaniaxml/bml/-->
<!--Download another Zefania XML files from-->
<!--http://sourceforge.net/projects/zefania-sharp-->
<XMLBIBLE xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xsi:noNamespaceSchemaLocation="zef2005.xsd" version="2.0.1.18" revision="1" status="v" biblename="{bible.name}" type="x-bible">
""",
            file=file,
            end="",
        )

    def _write_info(self, file: object, bible: Bible) -> None:
        """Write INFORMATION section with Bible details.

        :param file: File object to write to
        :param bible: Bible object containing information
        """
        from service_ppt.bible.bibleformat import get_bible_info

        dt_str = datetime.today().strftime("%Y-%m-%d")
        bible_name = xml.sax.saxutils.escape(bible.name)
        bible_lang = xml.sax.saxutils.escape(bible.lang)

        creator = ""
        description = ""
        publisher = ""
        rights = ""
        info_dict = get_bible_info(bible.name)
        if info_dict is not None:
            creator = info_dict.get("creator", "")
            description = info_dict.get("description", "")
            publisher = info_dict.get("publisher", "")
            rights = info_dict.get("rights", "")

        print(
            f""" <INFORMATION>
  <format>Zefania XML Bible Markup Language</format>
  <date>{dt_str}</date>
  <title>{bible_name}</title>
  <creator>{creator}</creator>
  <subject>Holy Bible</subject>
  <description>{description}</description>
  <publisher>{publisher}</publisher>
  <contributors></contributors>
  <type>bible</type>
  <identifier>{bible_name}</identifier>
  <source></source>
  <language>{bible_lang}</language>
  <coverage>Provide the bible to the world</coverage>
  <rights>{rights}</rights>
 </INFORMATION>
""",
            file=file,
            end="",
        )

    def _write_footer(self, file: object) -> None:
        """Write XML footer closing tag.

        :param file: File object to write to
        """
        print("</XMLBIBLE>", file=file)

    def _write_book(self, file: object, nth: int, book: Book) -> None:
        """Write book XML element with chapters and verses.

        :param file: File object to write to
        :param nth: Zero-based book number
        :param book: Book object to write
        """
        long_name = xml.sax.saxutils.escape(book.name)
        short_name = xml.sax.saxutils.escape(book.short_name)
        print(f""" <BIBLEBOOK bnumber="{nth + 1}" bname="{long_name}" bsname="{short_name}">""", file=file)

        for c, chapter in enumerate(book.chapters):
            print(f"""  <CHAPTER cnumber="{c + 1}">""", file=file)

            for verse in chapter.verses:
                verse_no = verse.no

                # Skip title verse
                if verse_no is None:
                    continue

                # Use the first verse no for range verse.
                if isinstance(verse_no, str) and "-" in verse_no:
                    verse_no = verse_no.split("-", maxsplit=1)[0]
                verse_no = str(verse_no)
                verse_text = xml.sax.saxutils.escape(verse.text)
                print(f"""   <VERS vnumber="{verse_no}">{verse_text}</VERS>""", file=file)

            print("  </CHAPTER>", file=file)

        print(" </BIBLEBOOK>", file=file)


class ZefaniaFormat(FileFormat):
    """File format handler for Zefania format files.

    This class implements the FileFormat interface for reading Zefania format
    Bible files from XML.
    """

    def __init__(self) -> None:
        """Initialize Zefania format handler."""
        super().__init__()

        self.versions: dict[str, str] | None = None
        self.options: dict[str, str] = {"ROOT_DIR": ""}

    def _get_root_dir(self) -> str:
        """Get the root directory for Zefania format files.

        :returns: Path to the root directory containing Zefania format files
        """
        dirname = self.get_option("ROOT_DIR")
        if not dirname:
            dirname = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "..", "Zefania-Bible-xml")

        dirname = os.path.normpath(dirname)

        return dirname

    def enum_versions(self) -> list[str]:
        """Enumerate available Bible versions in Zefania format.

        :returns: List of available version names
        """
        dirname = self._get_root_dir()

        versions = {}
        for f in os.listdir(dirname):
            if os.path.isfile(os.path.join(dirname, f)):
                pathname = os.path.join(dirname, f)
                name = ZefaniaReader._get_bible_name(pathname)
                if name:
                    versions[name] = pathname

        self.versions = versions

        return list(self.versions.keys())

    def read_version(self, version: str) -> Bible | None:
        """Read a specific Bible version from Zefania format.

        :param version: Version name to read
        :returns: Bible object or None if version not found
        """
        if self.versions is None:
            self.enum_versions()

        if version not in self.versions:
            return None

        filename = self.versions[version]
        reader = ZefaniaReader()
        bible = reader.read_bible(filename)

        return bible
