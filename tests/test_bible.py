"""Tests for Bible module.

This module contains unit tests for the Bible reading functionality,
including tests for various Bible formats and data structures.
"""

from pytest import mark

from service_ppt.bible.bibcore import Bible, Book, Chapter, Verse
from service_ppt.bible.biblang import L18N, LANG_EN

test_data = [
    ("Genesis 1:1", "Genesis", 1, 1, None, None),
    ("Genesis 1:1-2", "Genesis", 1, 1, None, 2),
    ("Genesis 1:1-2:2", "Genesis", 1, 1, 2, 2),
]


@mark.parametrize("text_range,bt,ct1,vs1,ct2,vs2", test_data)
def test_parse_verse_range(text_range, bt, ct1, vs1, ct2, vs2):
    ac_bt, ac_ct1, ac_vs1, ac_ct2, ac_vs2 = L18N.parse_verse_range(LANG_EN, text_range)
    assert ac_bt == bt, f"{ac_bt}!={bt} doesn't match!"
    assert ac_ct1 == ct1, f"{ac_ct1}!={ct1} doesn't match!"
    assert ac_vs1 == vs1, f"{ac_vs1}!={vs1} doesn't match!"
    assert ac_ct2 == ct2, f"{ac_ct2}!={ct2} doesn't match!"
    assert ac_vs2 == vs2, f"{ac_vs2}!={vs2} doesn't match!"


b1 = {
    1: {
        1: "In the beginning, God created the heavens and the earth.",
        2: "The earth was without form and void, and darkness was over the face of the deep. And the Spirit of God was hovering over the face of the waters.",
        3: "And God said, “Let there be light,” and there was light.",
    },
    2: {
        1: "Thus the heavens and the earth were finished, and all the host of them.",
        2: "And on the seventh day God finished his work that he had done, and he rested on the seventh day from all his work that he had done.",
        3: "So God blessed the seventh day and made it holy, because on it God rested from all his work that he had done in creation.",
    },
}


def populate_bible():
    book = Book()
    book.name = "Genesis"
    book.short_name = "Gen"
    for c_no, verses in b1.items():
        c = Chapter()
        c.no = c_no
        book.chapters.append(c)
        for v_no, v_text in verses.items():
            v = Verse()
            v.set_no(v_no)
            v.text = v_text
            c.verses.append(v)

    bible = Bible()
    bible.lang = LANG_EN
    bible.name = "ESV"
    bible.books.append(book)

    return bible


test_data = [
    ("Genesis 1:1", 1),
    ("Genesis 1:1-2", 2),
    ("Genesis 1:1-2:2", 5),
]


@mark.parametrize("text_range,verse_count", test_data)
def test_extract_texts_from_bible_index(text_range, verse_count):
    bible = populate_bible()
    bible_index = bible.translate_to_bible_index(text_range)
    verses = bible.extract_texts_from_bible_index(*bible_index)
    assert verse_count == len(verses), f"{verse_count}!={len(verses)} doesn't match!"
