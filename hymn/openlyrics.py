#!/usr/bin/env python
"""
"""
from datetime import datetime
import os
import xml.sax.saxutils
import xml.etree.ElementTree as ET

from hymncore import Verse, Song


class OpenLyricsReader:
    def _get_extension(self):
        return ".xml"

    def read_song(self, filename):
        ns = {"ns": "http://openlyrics.info/namespace/2009/song"}

        tree = ET.parse(filename)
        root = tree.getroot()
        if not root.tag.endswith("song"):
            return

        song = Song()
        props = root.find("ns:properties", ns)
        if props:
            titles = props.find("ns:titles", ns)
            if titles:
                song.title = titles.find("ns:title", ns).text

            authors = props.find("ns:authors", ns)
            if authors:
                song.authors = authors.find("ns:author", ns).text

            verse_order = props.find("ns:verseOrder", ns)
            if verse_order is not None:
                song.verse_order = verse_order.text

        lyrics = root.find("ns:lyrics", ns)
        v = None
        for elem in lyrics.iter():
            if elem.tag.endswith("verse"):
                name = elem.get("name", ns)
                v = Verse()
                v.no = name
                song.verses.append(v)

            elif elem.tag.endswith("lines"):
                line = "\n".join([l for l in elem.itertext() if l])

                v.lines.append(line)

        if len(song.verses) != 0:
            return song
        else:
            return None


class OpenLyricsWriter:
    def _get_extension(self):
        return ".xml"

    def write_hymn(self, dirname, book, encoding="utf-8"):
        if encoding == None:
            encoding = "utf-8"

        for song in book.songs:
            filename = song.title + ".xml"
            filename = os.path.join(dirname, filename)
            self.write_song(filename, song, encoding)

    def write_song(self, filename, song, encoding="utf-8"):
        with open(filename, "wt", encoding=encoding) as file:
            self._write_xml_header(file, encoding)

            self._write_properties(file, song)

            print(f" <lyrics>\n", file=file, end="")

            for v, verse in enumerate(song.verses):
                verse_name = None
                if verse.no is None:
                    verse_name = f"v{v+1}"
                elif isinstance(verse.no, int):
                    verse_name = f"v{verse.no}"
                else:
                    verse_name = verse.no

                print(f'  <verse name="{verse_name}">\n', file=file, end="")

                whole_line = ""
                for line in verse.lines:
                    if whole_line:
                        whole_line = whole_line + "<br/>"
                    line = xml.sax.saxutils.escape(line)
                    whole_line = whole_line + line
                print(f"   <lines>{whole_line}</lines>\n", file=file, end="")

                print(f"  </verse>\n", file=file, end="")

            print(f" </lyrics>\n", file=file, end="")

            self._write_xml_footer(file)

    def _write_properties(self, file, song):
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

    def _write_xml_header(self, file, encoding):
        dt_now = datetime.now()
        dt_now_str = dt_now.isoformat(timespec="seconds")
        print(
            f"""<?xml version="1.0" encoding="{encoding}"?>
<song xmlns="http://openlyrics.info/namespace/2009/song" version="0.8" createdIn="hymnconv 1.0" modifiedIn="hymnconv 1.0" modifiedDate="{dt_now_str}">
""",
            file=file,
            end="",
        )

    def _write_xml_footer(self, file):
        print(
            f"""</song>
""",
            file=file,
            end="",
        )
