"""
"""
from .mybible import MyBibleFormat
from .mysword_bible import MySwordFormat
from .sword_bible import SwordFormat
from .zefania_bible import ZefaniaFormat

# Currently supported Bible Program.
FORMAT_MYBIBLE = "MyBible"
FORMAT_MYSWORD = "MySword"
FORMAT_SWORD = "Sword"
FORMAT_ZEFANIA = "Zefania"


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


def get_bible_info(version):
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

    return info_dict.get(version, None)
