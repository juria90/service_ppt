"""This file supports reading Zefania Bible module that is stored as a XML file.
The Bible texts can be downloaded frmo https://sourceforge.net/projects/zefania-sharp/files/Bibles/.
"""

from datetime import datetime
import os
import xml.etree.ElementTree as ET
import xml.sax.saxutils

# pip install iso639-lang
from iso639 import Lang

from .bibcore import BibleInfo, Verse, Chapter, Book, Bible, FileFormat
from . import biblang


class ZefaniaReader:
    @staticmethod
    def _get_bible_name(filename, hint_line=10):
        """Check whether it is a XML file with following xml tag and attributes.
        <XMLBIBLE biblename="King James 2000" type="x-bible">

        return biblename if it is a valid bible.
        """
        lines = ""
        with open(filename, "r", encoding="utf-8") as f:
            for i in range(hint_line):
                lines += f.readline()

        parser = ET.XMLPullParser(["start", "end"])
        parser.feed(lines)
        for event, elem in parser.read_events():
            if event == "start" and elem.tag == "XMLBIBLE" and "biblename" in elem.attrib:
                return elem.attrib["biblename"]

            return None

    def read_bible(self, filename):
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

    def _parse_chapters(self, book, book_node):
        for node in book_node:
            if node.tag == "CHAPTER":
                chapter = Chapter()
                chapter.no = int(node.attrib["cnumber"])
                book.chapters.append(chapter)

                self._parse_verses(chapter, node)

    def _parse_verses(self, chapter, chapter_node):
        for node in chapter_node:
            if node.tag == "VERS":
                verse = Verse()
                verse.set_no(node.attrib["vnumber"])
                verse.text = self._concat_children_text(node)
                chapter.verses.append(verse)

    def _concat_children_text(self, node):
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
    def __init__(self):
        pass

    def _get_extension(self):
        return ".xml"

    def write_bible(self, dirname, bible, encoding="utf-8"):
        extension = self._get_extension()

        if encoding is None:
            encoding = "utf-8"

        filename = f"bible{extension}"
        with open(os.path.join(dirname, filename), "wt", encoding=encoding) as file:
            self._write_header(file, bible)

            self._write_info(file, bible)

            for i, b in enumerate(bible.books):
                bible.ensure_loaded(b)

                self._write_book(file, i, b)

            self._write_footer(file)

    def _write_header(self, file, bible):

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

    def _write_info(self, file, bible):
        from .fileformat import get_bible_info

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

    def _write_footer(self, file):
        print(f"</XMLBIBLE>", file=file)

    def _write_book(self, file, nth, book):
        long_name = xml.sax.saxutils.escape(book.name)
        short_name = xml.sax.saxutils.escape(book.short_name)
        print(f""" <BIBLEBOOK bnumber="{nth+1}" bname="{long_name}" bsname="{short_name}">""", file=file)

        for c, chapter in enumerate(book.chapters):
            print(f"""  <CHAPTER cnumber="{c+1}">""", file=file)

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

            print(f"  </CHAPTER>", file=file)

        print(f" </BIBLEBOOK>", file=file)


class ZefaniaFormat(FileFormat):
    def __init__(self):
        super().__init__()

        self.versions = None
        self.options = {"ROOT_DIR": ""}

    def _get_root_dir(self):
        dirname = self.get_option("ROOT_DIR")
        if not dirname:
            dirname = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "..", "Zefania-Bible-xml")

        dirname = os.path.normpath(dirname)

        return dirname

    def enum_versions(self):
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

    def read_version(self, version):
        if self.versions is None:
            self.enum_versions()

        if version not in self.versions:
            return None

        filename = self.versions[version]
        reader = ZefaniaReader()
        bible = reader.read_bible(filename)

        return bible
