#!/usr/bin/env python
"""OpenLyrics format reader and writer.

This module provides support for reading and writing hymn lyrics in the OpenLyrics
format (https://docs.openlyrics.org/), an open standard for storing song lyrics.
"""

from datetime import datetime
import os
from typing import TYPE_CHECKING, TextIO
import xml.etree.ElementTree as ET
import xml.sax.saxutils

if TYPE_CHECKING:
    from service_ppt.hymn.hymncore import Book

from service_ppt.hymn.hymncore import Line, Song, Verse


class OpenLyricsReader:
    """Reader for OpenLyrics format XML files."""

    def _get_extension(self) -> str:
        """Get the file extension for OpenLyrics format files.

        :returns: The file extension string (".xml").
        """
        return ".xml"

    @staticmethod
    def _get_element_text(parent: ET.Element | None, tag: str, namespaces: dict[str, str] | None = None) -> str | None:
        """Find and extract text from a child XML element if it exists and has text.

        :param parent: Parent XML element to search in, or None
        :param tag: Tag name of the child element to find
        :param namespaces: Optional namespace dictionary for the find operation
        :returns: The element's text if the element exists and has text, None otherwise
        """
        if parent is None:
            return None
        elem = parent.find(tag, namespaces) if namespaces else parent.find(tag)
        if elem is not None and elem.text is not None:
            return elem.text
        return None

    def read_song(self, filename: str) -> Song | None:
        """Read a song from an OpenLyrics format XML file.

        :param filename: Path to the OpenLyrics XML file to read
        :returns: A Song object if the file contains valid song data, None otherwise
        """
        ns = {"ns": "http://openlyrics.info/namespace/2009/song"}

        tree = ET.parse(filename)
        root = tree.getroot()
        if not root.tag.endswith("song"):
            return None

        song = Song()
        props = root.find("ns:properties", ns)
        if props is not None:
            titles = props.find("ns:titles", ns)
            if (title_text := self._get_element_text(titles, "ns:title", ns)) is not None:
                song.title = title_text

            authors = props.find("ns:authors", ns)
            if authors is not None:
                author_list: list[tuple[str, str]] = []
                for author_elem in authors.findall("ns:author", ns):
                    author_name = author_elem.text if author_elem.text else ""
                    author_type = author_elem.get("type", "words")
                    if author_type == "translation":
                        lang = author_elem.get("lang", "")
                        if lang:
                            author_type = f"translation/{lang}"
                    author_list.append((author_type, author_name))
                if author_list:
                    song.authors = author_list

            if (verse_order_text := self._get_element_text(props, "ns:verseOrder", ns)) is not None:
                song.verse_order = verse_order_text

        lyrics = root.find("ns:lyrics", ns)
        v: Verse | None = None
        if lyrics is not None:
            for elem in lyrics.iter():
                if elem.tag.endswith("verse"):
                    name = elem.get("name")
                    v = Verse()
                    v.no = name
                    song.verses.append(v)
                elif elem.tag.endswith("lines") and v is not None:
                    text = "\n".join([line for line in elem.itertext() if line])
                    optional_break = elem.get("break") == "optional"
                    line = Line(text, optional_break)

                    v.lines.append(line)

        if len(song.verses) != 0:
            return song
        return None


class OpenLyricsWriter:
    """Writer for OpenLyrics format XML files."""

    def _get_extension(self) -> str:
        """Get the file extension for OpenLyrics format files.

        :returns: The file extension string (".xml").
        """
        return ".xml"

    def write_hymn(self, dirname: str, book: "Book", encoding: str = "utf-8") -> None:
        """Write all songs from a book to OpenLyrics format XML files.

        :param dirname: Directory path where XML files will be written
        :param book: Book object containing songs to write
        :param encoding: Character encoding for the XML files (default: "utf-8")
        """
        if encoding is None:
            encoding = "utf-8"

        for song in book.songs:
            filename = song.title + ".xml"
            filename = os.path.join(dirname, filename)
            self.write_song(filename, song, encoding)

    def write_song(self, filename: str, song: Song, encoding: str = "utf-8") -> None:
        """Write a single song to an OpenLyrics format XML file.

        :param filename: Path where the XML file will be written
        :param song: Song object to write
        :param encoding: Character encoding for the XML file (default: "utf-8")
        """
        with open(filename, "w", encoding=encoding) as file:
            self._write_xml_header(file, encoding)

            self._write_properties(file, song)

            print(" <lyrics>\n", file=file, end="")

            for v, verse in enumerate(song.verses):
                verse_name = None
                if verse.no is None:
                    verse_name = f"v{v + 1}"
                elif isinstance(verse.no, int):
                    verse_name = f"v{verse.no}"
                else:
                    verse_name = verse.no

                print(f'  <verse name="{verse_name}">\n', file=file, end="")

                for line in verse.lines:
                    optional_break = ' break="optional"' if line.optional_break else ""
                    line_text = xml.sax.saxutils.escape(line.text)
                    print(f"   <lines{optional_break}>{line_text}</lines>\n", file=file, end="")

                print("  </verse>\n", file=file, end="")

            print(" </lyrics>\n", file=file, end="")

            self._write_xml_footer(file)

    def _write_properties(self, file: TextIO, song: Song) -> None:
        """Write song properties section to XML file.

        :param file: Text file object to write to
        :param song: Song object containing properties to write
        """
        print(" <properties>", file=file)

        if song.title:
            print("  <titles>", file=file)
            title = xml.sax.saxutils.escape(song.title)
            print(f"<title>{title}</title>", file=file)
            print("  </titles>", file=file)

        if isinstance(song.authors, list) and len(song.authors) > 0:
            print("  <authors>", file=file)
            for author_type, value in song.authors:
                if not author_type.startswith("translation/"):
                    author_name = xml.sax.saxutils.escape(value)
                    print(f'   <author type="{author_type}">{author_name}</author>', file=file)
                else:
                    lang = author_type[len("translation/") :]
                    author_name = xml.sax.saxutils.escape(value)
                    print(f'   <author type="translation" lang="{lang}">{author_name}</author>', file=file)
            print("  </authors>", file=file)

        if song.verse_order:
            print(f"  <verseOrder>{song.verse_order}</verseOrder>", file=file)

        if isinstance(song.songbook, dict) and "name" in song.songbook:
            print("  <songbooks>", file=file)
            name = song.songbook["name"]
            if "entry" in song.songbook:
                entry = song.songbook["entry"]
                print(f'   <songbook name="{name}" entry="{entry}"/>', file=file)
            else:
                print(f'   <songbook name="{name}"/>', file=file)
            print("  </songbooks>", file=file)

        if song.released:
            print(f"  <released>{song.released}</released>", file=file)

        if song.keywords:
            print(f"  <keywords>{song.keywords}</keywords>", file=file)

        print(" </properties>", file=file)

    def _write_xml_header(self, file: TextIO, encoding: str) -> None:
        """Write XML header and root element to file.

        :param file: Text file object to write to
        :param encoding: Character encoding used in the XML file
        """
        dt_now = datetime.now()
        dt_now_str = dt_now.isoformat(timespec="seconds")
        print(
            f"""<?xml version="1.0" encoding="{encoding}"?>
<song xmlns="http://openlyrics.info/namespace/2009/song" version="0.8" createdIn="hymnconv 1.0" modifiedIn="hymnconv 1.0" modifiedDate="{dt_now_str}">
""",
            file=file,
            end="",
        )

    def _write_xml_footer(self, file: TextIO) -> None:
        """Write XML closing tag to file.

        :param file: Text file object to write to
        """
        print(
            """</song>
""",
            file=file,
            end="",
        )
