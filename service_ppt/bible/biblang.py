#!/usr/bin/env python
"""Bible language and internationalization support.

This module provides language detection, translation support, and book name
mapping for different Bible translations and languages.
"""

import gettext
import re

from langdetect import detect

# iso 639-1
LANG_EN = "en"

UNICODE_BOM = "\ufeff"

_ = lambda s: s
pgettext = lambda c, s: s


def check_bom(data: bytes) -> str:
    UTF_16_BE_BOM = b"\xfe\xff"
    UTF_16_LE_BOM = b"\xff\xfe"
    UTF_8_BOM = b"\xef\xbb\xbf"
    UTF_32_BE_BOM = b"\x00\x00\xfe\xff"
    UTF_32_LE_BOM = b"\xff\xfe\x00\x00"

    size = len(data)
    if size >= 3:
        if data[:3] == UTF_8_BOM:
            return "utf-8"

    if size >= 4:
        if data[:4] == UTF_32_LE_BOM:
            return "utf-32-le"
        if data[:4] == UTF_32_BE_BOM:
            return "utf-32-be"

    if size >= 2:
        if data[:2] == UTF_16_LE_BOM:
            return "utf-16-le"
        if data[:2] == UTF_16_BE_BOM:
            return "utf-16-be"

    return None


def detect_encoding(file):
    bom = None
    with open(file, "rb") as f:
        b = f.read(4)
        bom = check_bom(b)

    return bom


def detect_language(text):
    """https://pypi.org/project/langdetect/"""

    # returns ISO 639-1 codes
    return detect(text)


class L18N:
    OLD_NEW_TESTAMENT = (_("Old Testament"), _("New Testament"))
    VERSE_PARSING_PATTERN = _(r"(.*?) ?(\d+):(\d+)(-((\d+):)?(\d+))?")
    BOOK_CHAPTER_FORMAT = _("{book} {chapter}")

    BOOK_NAME = [
        # old testament
        (_("Genesis"), _("Gen")),
        (_("Exodus"), _("Ex")),
        (_("Leviticus"), _("Lev")),
        (_("Numbers"), _("Num")),
        (_("Deuteronomy"), _("Deut")),
        (_("Joshua"), _("Josh")),
        (_("Judges"), _("Judg")),
        (_("Ruth"), pgettext("short_name", "Ruth")),
        (_("1 Samuel"), _("1 Sam")),
        (_("2 Samuel"), _("2 Sam")),
        (_("1 Kings"), pgettext("short_name", "1 Kings")),
        (_("2 Kings"), pgettext("short_name", "2 Kings")),
        (_("1 Chronicles"), _("1 Chron")),
        (_("2 Chronicles"), _("2 Chron")),
        (_("Ezra"), pgettext("short_name", "Ezra")),
        (_("Nehemiah"), _("Neh")),
        (_("Esther"), _("Est")),
        (_("Job"), pgettext("short_name", "Job")),
        (_("Psalms"), _("Ps")),
        (_("Proverbs"), _("Prov")),
        (_("Ecclesiastes"), _("Eccles")),
        (_("Song of Solomon"), _("Song")),
        (_("Isaiah"), _("Isa")),
        (_("Jeremiah"), _("Jer")),
        (_("Lamentations"), _("Lam")),
        (_("Ezekiel"), _("Ezek")),
        (_("Daniel"), _("Dan")),
        (_("Hosea"), _("Hos")),
        (_("Joel"), pgettext("short_name", "Joel")),
        (_("Amos"), pgettext("short_name", "Amos")),
        (_("Obadiah"), _("Obad")),
        (_("Jonah"), pgettext("short_name", "Jonah")),
        (_("Micah"), _("Mic")),
        (_("Nahum"), _("Nah")),
        (_("Habakkuk"), _("Hab")),
        (_("Zephaniah"), _("Zeph")),
        (_("Haggai"), _("Hag")),
        (_("Zechariah"), _("Zech")),
        (_("Malachi"), _("Mal")),
        # new testament
        (_("Matthew"), _("Matt")),
        (_("Mark"), pgettext("short_name", "Mark")),
        (_("Luke"), pgettext("short_name", "Luke")),
        (_("John"), pgettext("short_name", "John")),
        (_("Acts"), pgettext("short_name", "Acts")),
        (_("Romans"), _("Rom")),
        (_("1 Corinthians"), _("1 Cor")),
        (_("2 Corinthians"), _("2 Cor")),
        (_("Galatians"), _("Gal")),
        (_("Ephesians"), _("Eph")),
        (_("Philippians"), _("Phil")),
        (_("Colossians"), _("Col")),
        (_("1 Thessalonians"), _("1 Thess")),
        (_("2 Thessalonians"), _("2 Thess")),
        (_("1 Timothy"), _("1 Tim")),
        (_("2 Timothy"), _("2 Tim")),
        (_("Titus"), pgettext("short_name", "Titus")),
        (_("Philemon"), _("Philem")),
        (_("Hebrews"), _("Heb")),
        (_("James"), pgettext("short_name", "James")),
        (_("1 Peter"), _("1 Pet")),
        (_("2 Peter"), _("2 Pet")),
        (_("1 John"), pgettext("short_name", "1 John")),
        (_("2 John"), pgettext("short_name", "2 John")),
        (_("3 John"), pgettext("short_name", "3 John")),
        (_("Jude"), pgettext("short_name", "Jude")),
        (_("Revelation of John"), _("Rev")),
    ]

    LOCALE_DIR = None
    TRANSLATIONS = {}

    @staticmethod
    def get_translation(lang):
        if lang in L18N.TRANSLATIONS:
            return L18N.TRANSLATIONS[lang]

        dirname = L18N.LOCALE_DIR
        if dirname is None:
            from service_ppt import get_locale_dir

            dirname = str(get_locale_dir())

        trans = None
        try:
            trans = gettext.translation("bible", localedir=dirname, languages=[lang], fallback=True)
        except OSError:
            pass

        L18N.TRANSLATIONS[lang] = trans
        return L18N.TRANSLATIONS[lang]

    @staticmethod
    def get_testament_name(new_testament, lang=None):
        name = L18N.OLD_NEW_TESTAMENT[1 if new_testament else 0]
        if lang is not None:
            trans = L18N.get_translation(lang)
            name = trans.gettext(name)

        return name

    @staticmethod
    def get_english_book_name(book_no):
        return L18N.BOOK_NAME[book_no]

    @staticmethod
    def get_short_english_book_name(book_no):
        return L18N.BOOK_NAME[book_no][1]

    @staticmethod
    def get_book_names(book_no, lang=None):
        """get_book_names() returns localized bible book's long and short names.

        @param: book_no is a zero based index.
        @param: lang iso 639-1 language.
        """
        name = L18N.BOOK_NAME[book_no][0]
        short_name = L18N.BOOK_NAME[book_no][1]

        if lang is not None:
            trans = L18N.get_translation(lang)
            name = trans.gettext(name)
            short_name = trans.pgettext("short_name", short_name)

        return (name, short_name)

    @staticmethod
    def get_short_book_name(book_no, lang=None):
        """get_short_book_name() returns localized bible book's short name.

        @param: book_no is a zero based index.
        @param: lang iso 639-1 language.
        """
        short_name = L18N.BOOK_NAME[book_no][1]

        if lang is not None:
            trans = L18N.get_translation(lang)
            short_name = trans.pgettext("short_name", short_name)

        return short_name

    @staticmethod
    def get_book_chapter_name(lang, book, chapter):
        trans = L18N.get_translation(lang)
        name = trans.gettext(L18N.BOOK_CHAPTER_FORMAT)
        name = name.format(book=book, chapter=chapter)
        return name

    @staticmethod
    def parse_verse_range(lang, text_range):
        """parse_verse_range() returns (book, chapter1, verse1, chapter2, verse2}) tuple based on text_range.

        The text_range should be formatted as <Book> <Chapter1>:<Verse1>[[<Chapter2>:]-<Verse2>],
        where Book can be long or short name, Chapter1/Chapter2 and Verse1/Verse2 are valid numbers.
        """

        bt = None
        ct1 = None
        vs1 = None
        ct2 = None
        vs2 = None
        m = None

        trans = L18N.get_translation(lang)

        # try localized reg pattern
        pattern = trans.gettext(L18N.VERSE_PARSING_PATTERN)
        m = re.match(pattern, text_range)
        if m is not None:
            bt = m.group(1)
            ct1 = m.group(2)
            vs1 = m.group(3)
            ct2 = m.group(6)
            vs2 = m.group(7)

        if pattern != L18N.VERSE_PARSING_PATTERN and m is None:
            m = re.match(L18N.VERSE_PARSING_PATTERN, text_range)
            if m is not None:
                bt = m.group(1)
                ct1 = m.group(2)
                vs1 = m.group(3)
                ct2 = m.group(6)
                vs2 = m.group(7)

        if m is None:
            raise ValueError("Invalid text range: %s")

        bt = bt.strip()
        ct1 = int(ct1)
        vs1 = int(vs1)
        if ct2 is not None:
            ct2 = int(ct2)
        if vs2 is not None:
            vs2 = int(vs2)

        if ct2 is not None:
            if ct1 > ct2 or (ct1 == ct2 and vs1 > vs2):
                raise ValueError("Invalid text range: %s")

        return bt, ct1, vs1, ct2, vs2
