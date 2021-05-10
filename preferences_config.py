"""This file contains PreferencesConfig class.
"""
import typing
import wx


# Window
SW_RESTORED = 0  # Window is in normal, restore status
SW_ICONIZED = 1  # Window is in Iconized status
SW_MAXIMIZED = 2  # Window is in Maximized status


class PreferencesConfig:
    """PreferencesConfig class contains all configuration used in PreferencesDialog.
    It has methods to read and write config.
    """

    def __init__(self):
        self.current_bible_format = ""
        self.bible_rootdir = ""
        self.current_bible_version = ""

    def _read_one_string(self, config: wx.ConfigBase, label: str, default_value: str) -> str:
        value = default_value

        try:
            value = config.Read(label, default_value)
        except ValueError:
            pass

        return value

    def _read_one_integer(self, config: wx.ConfigBase, label: str, default_value: int) -> int:
        value = default_value

        try:
            value = config.ReadInt(label, default_value)
        except ValueError:
            pass

        return value

    def read_config(self, config: wx.ConfigBase):
        """read_config reads all configuration from config class."""
        self.current_bible_format = self._read_one_string(config, "current_bible_format", "")
        self.bible_rootdir = self._read_one_string(config, "bible_rootdir", "")
        self.current_bible_version = self._read_one_string(config, "current_bible_version", "")

    def write_config(self, config):
        """write_config writes all configuration to config class."""
        config.Write("current_bible_format", self.current_bible_format)
        config.Write("bible_rootdir", self.bible_rootdir)
        config.Write("current_bible_version", self.current_bible_version)

    def read_window_rect(self, config: wx.ConfigBase) -> typing.Optional[typing.Tuple[int, typing.List[int]]]:
        s = self._read_one_string(config, "window_rect", "")
        numbers = [int(n) for n in s.split()]
        if len(numbers) != 5:
            return None

        return numbers[0], numbers[1:]

    def write_window_rect(self, config: wx.ConfigBase, sw: int, rc: typing.List[int]):
        s = "%d %d %d %d %d" % (sw, rc[0], rc[1], rc[2], rc[3])
        config.Write("window_rect", s)
