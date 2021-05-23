#!/usr/bin/env python
"""
"""
import json
import os
import zipfile


class OpenLPServiceWriter:
    def _get_extension(self):
        return ".osz"

    def write(self, zipfilename, song_list, xml_list):
        _osj_list = self._osj_from_files(song_list, xml_list)
        data = json.dumps(_osj_list)
        with zipfile.ZipFile(zipfilename, "w", zipfile.ZIP_DEFLATED) as zipf:
            basename = os.path.basename(zipfilename)
            osj_filename, _ = os.path.splitext(basename)
            osj_filename += ".osj"
            zipf.writestr(osj_filename, data)

    def _osj_from_files(self, song_list, xml_list):
        osj_list = []

        # first header
        osj_list.append({"openlp_core": {"service-theme": "Transparent", "lite-service": False}})

        def merge_lines(lines):
            text = ""
            for l in lines:
                if text:
                    text = text + "\n"
                text = text + l.text
                if l.optional_break:
                    text = text + "\n[---]"

            return text

        def serviceitem_from_file(song, xml_content):
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

            data_list = []
            for v in song.get_verses_by_order():
                line0 = ""
                if len(v.lines) > 0:
                    line0 = v.lines[0]
                data_list.append({"verseTag": v.no, "title": line0.text, "raw_slide": merge_lines(v.lines)})

            return {"serviceitem": {"header": header_dict, "data": data_list}}

        for i, song in enumerate(song_list):
            xml_content = xml_list[i]
            osj = serviceitem_from_file(song, xml_content)
            osj_list.append(osj)

        return osj_list
