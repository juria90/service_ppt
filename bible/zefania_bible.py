'''This file supports reading Zefania Bible module that is stored as a XML file.
The Bible texts can be downloaded frmo https://sourceforge.net/projects/zefania-sharp/files/Bibles/.
'''

import os
import xml.etree.ElementTree as ET

# pip install iso-639
from iso639 import languages

from bibcore import BibleInfo, Verse, Chapter, Book, Bible, FileFormat
import biblang


class ZefaniaFormat(FileFormat):
    def __init__(self):
        super().__init__()

        self.versions = None
        self.options = {'ROOT_DIR': ''}

    def _get_root_dir(self):
        dirname = self.get_option('ROOT_DIR')
        if not dirname:
            dirname = os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', '..', 'Zefania-Bible-xml')

        dirname = os.path.normpath(dirname)

        return dirname

    @staticmethod
    def _get_bible_name(filename, hint_line=10):
        '''Check whether it is a XML file with following xml tag and attributes.
        <XMLBIBLE biblename="King James 2000" type="x-bible">

        return biblename if it is a valid bible.
        '''
        lines = ''
        with open(filename, 'r', encoding='utf-8') as f:
            for i in range(hint_line):
                lines += f.readline()

        parser = ET.XMLPullParser(['start', 'end'])
        parser.feed(lines)
        for event, elem in parser.read_events():
            if event == 'start' and elem.tag == 'XMLBIBLE' and 'biblename' in elem.attrib:
                return elem.attrib['biblename']

            return None

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
        tree = ET.parse(filename)
        root = tree.getroot()

        bible = Bible()
        bible.name = root.attrib['biblename']
        for node in root:
            if node.tag == 'INFORMATION':
                lang_node = node.find('language')
                if lang_node is not None:
                    lang_part2 = lang_node.text.strip().lower()
                    try:
                        lang = languages.get(part2b=lang_part2)
                        bible.lang = lang.part1
                    except KeyError:
                        bible.lang = biblang.LANG_EN

            elif node.tag == 'BIBLEBOOK':
                book = Book()
                book_no = int(node.attrib['bnumber'])
                book.new_testament = BibleInfo.is_new_testament(book_no-1)
                if 'bname' in node.attrib:
                    book.name = node.attrib['bname']
                if 'bsname' in node.attrib:
                    book.short_name = node.attrib['bsname']
                if not book.name or not book.short_name:
                    names = biblang.L18N.get_book_names(book_no-1, bible.lang)
                    if not book.name:
                        book.name = names[0]
                    if not book.short_name:
                        book.short_name = names[1]
                bible.books.append(book)

                self._parse_chapters(book, node)

        return bible

    def _parse_chapters(self, book, book_node):
        for node in book_node:
            if node.tag == 'CHAPTER':
                chapter = Chapter()
                chapter.no = int(node.attrib['cnumber'])
                book.chapters.append(chapter)

                self._parse_verses(chapter, node)

    def _parse_verses(self, chapter, chapter_node):
        for node in chapter_node:
            if node.tag == 'VERS':
                verse = Verse()
                verse.set_no(node.attrib['vnumber'])
                verse.text = self._concat_children_text(node)
                chapter.verses.append(verse)

    def _concat_children_text(self, node):
        text = ''
        for it in node.itertext():
            t = it.strip()
            if t:
                if text:
                    text = text + ' ' + t
                else:
                    text = t

        return text
