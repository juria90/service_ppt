"""LyricManager class for service_ppt.

This module contains the LyricManager class that manages lyric file operations,
including reading songs, searching for lyric files, and managing lyric file collections.
"""

import errno
import os
from typing import Any

from service_ppt.hymn.openlyrics import OpenLyricsReader
from service_ppt.utils.i18n import _


class LyricManager:
    """Manages lyric file operations for service_ppt.

    LyricManager handles reading lyric files, searching for lyric files,
    and maintaining collections of lyric files for archiving.
    """

    def __init__(self, cm: Any) -> None:
        """Initialize LyricManager.

        :param cm: CommandManager instance for error reporting
        """
        self.cm: Any = cm
        self.reader: OpenLyricsReader = OpenLyricsReader()
        self.lyric_file_map: dict[str, Any] = {}
        self.all_lyric_files: list[dict[str, Any] | str] = []
        self.lyric_search_path: str | None = None

    def reset_exec_vars(self) -> None:
        """Reset execution variables to initial state."""
        self.lyric_file_map = {}
        self.all_lyric_files = []

    def read_song(self, filename: str) -> Any:
        """Read a song from a lyric file.

        :param filename: Path to the lyric file
        :returns: Song object
        :raises FileNotFoundError: If the file doesn't exist
        """
        if filename in self.lyric_file_map:
            return self.lyric_file_map[filename]

        if not os.path.exists(filename):
            self.cm.error_message(_("Cannot open a lyric file '{filename}'.").format(filename=filename))
            raise FileNotFoundError(errno.ENOENT, os.strerror(errno.ENOENT), filename)

        try:
            song = self.reader.read_song(filename)
        except Exception:
            self.cm.error_message(_("Cannot open a lyric file '{filename}'.").format(filename=filename))
            raise

        self.lyric_file_map[filename] = song
        return song

    def read_songs(self, filelist: list[str]) -> list[Any]:
        """Read multiple songs from a list of files.

        :param filelist: List of file paths to read
        :returns: List of song objects
        """
        songs: list[Any] = []
        for filename in filelist:
            song = self.read_song(filename)
            songs.append(song)

        return songs

    def search_lyric_file(self, filename: str) -> str | None:
        """Search for a lyric file corresponding to a given filename.

        :param filename: Base filename to search for
        :returns: Path to lyric file if found, or None
        """
        _dir, fn = os.path.split(filename)
        xml_pathname = os.path.splitext(filename)[0] + ".xml"
        file_exist = os.path.exists(xml_pathname)
        if file_exist:
            return xml_pathname

        # search xml file from search path
        if self.lyric_search_path:
            xml_filename = os.path.splitext(fn)[0] + ".xml"
            searched_xml_pathname = os.path.join(self.lyric_search_path, xml_filename)
            file_exist = os.path.exists(searched_xml_pathname)
            if file_exist:
                return searched_xml_pathname

        return xml_pathname

    def add_lyric_file(self, filelist: list[str] | dict[str, Any] | str) -> None:
        """Add lyric file(s) to the collection.

        :param filelist: Lyric file path(s) or dictionary to add
        """
        if isinstance(filelist, list):
            _songs = self.read_songs(filelist)
            self.all_lyric_files.extend(filelist)
        elif isinstance(filelist, dict):
            self.all_lyric_files.append(filelist)
        else:
            _song = self.read_song(filelist)
            self.all_lyric_files.append(filelist)
