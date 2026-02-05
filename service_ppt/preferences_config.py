"""Application preferences configuration.

This module provides the PreferencesConfig class for managing application
settings such as window position, Bible format preferences, and directory paths.
"""

import json

import wx

# Window
SW_RESTORED = 0  # Window is in normal, restore status
SW_ICONIZED = 1  # Window is in Iconized status
SW_MAXIMIZED = 2  # Window is in Maximized status


class PreferencesConfig:
    """PreferencesConfig class contains all configuration used in PreferencesDialog.
    It has methods to read and write config.
    """

    def __init__(self) -> None:
        self.current_bible_format: str = ""
        self.bible_rootdir: str = ""
        self.current_bible_version: str = ""

        self.lyric_open_textfile: bool = False
        self.lyric_search_path: str = ""
        self.lyric_copy_from_template: bool = False
        self.lyric_application_pathname: str = ""
        self.lyric_template_filename: str = ""

        self.dir_dict: dict[str, str] = {}

    def _read_one_bool(self, config: wx.ConfigBase, label: str, default_value: bool) -> bool:
        value = default_value

        try:
            value = config.ReadBool(label, default_value)
        except ValueError:
            # Invalid boolean value in config, use default
            pass

        return value

    def _read_one_string(self, config: wx.ConfigBase, label: str, default_value: str) -> str:
        value = default_value

        try:
            value = config.Read(label, default_value)
        except ValueError:
            # Invalid string value in config, use default
            pass

        return value

    def read_config(self, config: wx.ConfigBase) -> None:
        """read_config reads all configuration from config class.

        :param config: wx configuration object to read from
        """
        self.current_bible_format = self._read_one_string(config, "current_bible_format", "")
        self.bible_rootdir = self._read_one_string(config, "bible_rootdir", "")
        self.current_bible_version = self._read_one_string(config, "current_bible_version", "")

        self.lyric_open_textfile = self._read_one_bool(config, "lyric_open_textfile", False)
        self.lyric_search_path = self._read_one_string(config, "lyric_search_path", "")
        self.lyric_copy_from_template = self._read_one_bool(config, "lyric_copy_from_template", False)
        self.lyric_application_pathname = self._read_one_string(config, "lyric_application_pathname", "")
        self.lyric_template_filename = self._read_one_string(config, "lyric_template_filename", "")

        dir_dict = self._read_one_string(config, "dir_dict", "{}")
        dir_dict = json.loads(dir_dict)
        if isinstance(dir_dict, dict):
            self.dir_dict = dir_dict
        else:
            self.dir_dict = {}

    def write_config(self, config: wx.ConfigBase) -> None:
        """write_config writes all configuration to config class.

        :param config: wx configuration object to write to
        """
        config.Write("current_bible_format", self.current_bible_format)
        config.Write("bible_rootdir", self.bible_rootdir)
        config.Write("current_bible_version", self.current_bible_version)

        config.WriteBool("lyric_open_textfile", self.lyric_open_textfile)
        config.Write("lyric_search_path", self.lyric_search_path)
        config.WriteBool("lyric_copy_from_template", self.lyric_copy_from_template)
        config.Write("lyric_application_pathname", self.lyric_application_pathname)
        config.Write("lyric_template_filename", self.lyric_template_filename)

        dir_symbols = json.dumps(self.dir_dict)
        config.Write("dir_dict", dir_symbols)

    def read_window_rect(self, config: wx.ConfigBase) -> tuple[int, tuple[int, int, int, int]] | None:
        s = self._read_one_string(config, "window_rect", "")
        numbers = [int(n) for n in s.split()]
        if len(numbers) != 5:
            return None

        return numbers[0], tuple(numbers[1:])

    def write_window_rect(self, config: wx.ConfigBase, sw: int, rc: tuple[int, int, int, int]) -> None:
        """Write window rectangle configuration.

        :param config: wx configuration object to write to
        :param sw: Window state (SW_RESTORED, SW_ICONIZED, or SW_MAXIMIZED)
        :param rc: Window rectangle as (x, y, width, height)
        """
        s = "%d %d %d %d %d" % (sw, rc[0], rc[1], rc[2], rc[3])
        config.Write("window_rect", s)
