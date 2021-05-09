"""
"""
import os

from mybible import MyBibleFormat
from mysword_bible import MySwordFormat
from sword_bible import SwordFormat
from zefania_bible import ZefaniaFormat

# Currently supported Bible Program.
FORMAT_MYBIBLE = "MyBible"
FORMAT_MYSWORD = "MySword"
FORMAT_SWORD = "Sword"
FORMAT_ZEFANIA = "Zefania"
FORMAT_LIST = []


def _import_bible_format():
    prog_dict = {}

    prog_dict[FORMAT_MYBIBLE] = MyBibleFormat()

    prog_dict[FORMAT_MYSWORD] = MySwordFormat()

    try:
        from pysword.modules import SwordModules

        modules = SwordModules()
        found_modules = modules.parse_modules()
        prog_dict[FORMAT_SWORD] = SwordFormat(modules, found_modules)
    except ImportError:
        pass
    except FileNotFoundError:
        pass

    prog_dict[FORMAT_ZEFANIA] = ZefaniaFormat()

    return prog_dict


FORMAT_LIST = _import_bible_format()


def get_format_list():
    return [p for p in FORMAT_LIST]


def get_format_option(fileformat, key):
    if fileformat in FORMAT_LIST:
        format_obj = FORMAT_LIST[fileformat]
        return format_obj.get_option(key)

    return None


def set_format_option(fileformat, key, value):
    if fileformat in FORMAT_LIST:
        format_obj = FORMAT_LIST[fileformat]
        format_obj.set_option(key, value)


def enum_versions(fileformat):
    if fileformat in FORMAT_LIST:
        format_obj = FORMAT_LIST[fileformat]
        return format_obj.enum_versions()

    return None


def read_version(fileformat, version):
    bible = None
    if fileformat in FORMAT_LIST:
        format_obj = FORMAT_LIST[fileformat]
        bible = format_obj.read_version(version)

    return bible
