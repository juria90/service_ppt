"""Bible file format registry and factory.

This module provides a registry of available Bible file formats and a factory
function to create appropriate format readers based on the format type.
"""

from enum import Enum
from typing import TYPE_CHECKING, Any

from service_ppt.bible.bibcore import FileFormat
from service_ppt.bible.mybible import MyBibleFormat
from service_ppt.bible.mysword_bible import MySwordFormat
from service_ppt.bible.sword_bible import SwordFormat
from service_ppt.bible.zefania_bible import ZefaniaFormat

if TYPE_CHECKING:
    from service_ppt.bible.bibcore import Bible


class BibleFormat(str, Enum):
    """Enumeration of supported Bible file formats."""

    MYBIBLE = "MyBible"
    MYSWORD = "MySword"
    SWORD = "Sword"
    ZEFANIA = "Zefania"


def _import_bible_format() -> dict[str, FileFormat]:
    prog_dict: dict[str, FileFormat] = {}

    prog_dict[BibleFormat.MYBIBLE.value] = MyBibleFormat()

    prog_dict[BibleFormat.MYSWORD.value] = MySwordFormat()

    try:
        from pysword.modules import SwordModules

        modules = SwordModules()
        found_modules = modules.parse_modules()
        prog_dict[BibleFormat.SWORD.value] = SwordFormat(modules, found_modules)
    except ImportError:
        # Optional dependency not available, skip
        pass
    except FileNotFoundError:
        # File not found, skip
        pass

    prog_dict[BibleFormat.ZEFANIA.value] = ZefaniaFormat()

    return prog_dict


FORMAT_LIST = _import_bible_format()


def get_format_list() -> list[str]:
    return list(FORMAT_LIST)


def get_format_option(fileformat: str, key: str) -> "Any | None":
    if fileformat in FORMAT_LIST:
        format_obj = FORMAT_LIST[fileformat]
        return format_obj.get_option(key)

    return None


def set_format_option(fileformat: str, key: str, value: "Any") -> None:
    if fileformat in FORMAT_LIST:
        format_obj = FORMAT_LIST[fileformat]
        format_obj.set_option(key, value)


def enum_versions(fileformat: str) -> list[str] | None:
    if fileformat in FORMAT_LIST:
        format_obj = FORMAT_LIST[fileformat]
        return format_obj.enum_versions()

    return None


def read_version(fileformat: str, version: str) -> "Bible | None":
    bible = None
    if fileformat in FORMAT_LIST:
        format_obj = FORMAT_LIST[fileformat]
        bible = format_obj.read_version(version)

    return bible


def get_bible_info(version: str) -> dict[str, str] | None:
    info_dict = {
        "ESV": {
            "creator": "Crossway",
            "description": "The English Standard Version (ESV) is an essentially literal translation of the Bible in contemporary English.",
            "publisher": "Crossway",
            "rights": "Crossway",
        },
        "개역개정": {
            "creator": "재단법인 대한성서공회",
            "description": "대한성서공회가 발표한 성경으로 개역한글판을 1998년에 개정한 한국어 성경",
            "publisher": "재단법인 대한성서공회",
            "rights": "재단법인 대한성서공회",
        },
    }

    return info_dict.get(version)
