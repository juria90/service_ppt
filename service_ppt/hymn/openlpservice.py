#!/usr/bin/env python
"""OpenLP Service format writer.

This module provides support for writing hymn lyrics in OpenLP Service format,
which is used by the OpenLP presentation software for church services.
"""

import json
import os
from typing import Any
import zipfile

from service_ppt.hymn.hymncore import Line, Song


class OpenLPServiceWriter:
    """Writer for OpenLP Service format files.

    This class writes songs in the OpenLP Service format (.osz), which
    is a ZIP file containing a JSON service file used by OpenLP.
    """

    def _get_extension(self) -> str:
        """Get the file extension for OpenLP Service format files.

        :returns: The file extension string (".osz").
        """
        return ".osz"

    def write(self, zipfilename: str, song_list: list[Song], xml_list: list[str]) -> None:
        """Write songs to an OpenLP Service format ZIP file.

        :param zipfilename: Path where the .osz ZIP file will be written
        :param song_list: List of Song objects to include in the service
        :param xml_list: List of XML content strings corresponding to each song
        """
        _osj_list = self._osj_from_files(song_list, xml_list)
        data = json.dumps(_osj_list)
        with zipfile.ZipFile(zipfilename, "w", zipfile.ZIP_DEFLATED) as zipf:
            basename = os.path.basename(zipfilename)
            osj_filename, _ = os.path.splitext(basename)
            osj_filename += ".osj"
            zipf.writestr(osj_filename, data)

    def _osj_from_files(self, song_list: list[Song], xml_list: list[str]) -> list[dict[str, Any]]:
        """Convert songs and XML content to OpenLP Service JSON format.

        :param song_list: List of Song objects to convert
        :param xml_list: List of XML content strings corresponding to each song
        :returns: List of dictionaries representing the OpenLP Service JSON structure
        """
        osj_list = []

        # first header
        osj_list.append({"openlp_core": {"service-theme": "Transparent", "lite-service": False}})

        def merge_lines(lines: list[Line]) -> str:
            text = ""
            for line in lines:
                if text:
                    text = text + "\n"
                text = text + line.text
                if line.optional_break:
                    text = text + "\n[---]"

            return text

        def serviceitem_from_file(song: Song, xml_content: str) -> dict[str, Any]:
            header_dict = {
                "start_time": 0,
                "search": "",
                "icon": ":/plugins/plugin_songs.png",
                "will_auto_start": False,
                "footer": [f"{song.title}"],
                "auto_play_slides_loop": False,
                "title": f"{song.title}",
                "xml_version": xml_content,
                "theme": None,
                "from_plugin": False,
                "data": {"title": f"{song.title} @"},
                "media_length": 0,
                "capabilities": [2, 1, 5, 8, 9, 13],
                "processor": None,
                "auto_play_slides_once": False,
                "end_time": 0,
                "audit": [f"{song.title}", [], "", ""],
                "name": "songs",
                "theme_overwritten": False,
                "type": 1,
                "background_audio": [],
                "plugin": "songs",
                "notes": "",
                "timed_slide_interval": 0,
            }

            data_list: list[dict[str, Any]] = []
            for v in song.get_verses_by_order():
                line0: Line | None = None
                if len(v.lines) > 0:
                    line0 = v.lines[0]
                title_text = line0.text if line0 else ""
                data_list.append({"verseTag": v.no, "title": title_text, "raw_slide": merge_lines(v.lines)})

            return {"serviceitem": {"header": header_dict, "data": data_list}}

        for i, song in enumerate(song_list):
            xml_content = xml_list[i]
            osj = serviceitem_from_file(song, xml_content)
            osj_list.append(osj)

        return osj_list
