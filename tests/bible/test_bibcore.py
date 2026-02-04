"""Tests for bibcore module.

This module contains unit tests for core Bible data structures and classes,
including Bible, Book, Chapter, and Verse.
"""

from pytest import mark

from service_ppt.bible.bibcore import Bible, Book, Chapter, Verse
from service_ppt.bible.biblang import LANG_EN

b1 = {
    1: {
        1: "In the beginning, God created the heavens and the earth.",
        2: "The earth was without form and void, and darkness was over the face of the deep. And the Spirit of God was hovering over the face of the waters.",
        3: 'And God said, "Let there be light," and there was light.',
    },
    2: {
        1: "Thus the heavens and the earth were finished, and all the host of them.",
        2: "And on the seventh day God finished his work that he had done, and he rested on the seventh day from all his work that he had done.",
        3: "So God blessed the seventh day and made it holy, because on it God rested from all his work that he had done in creation.",
    },
}


def populate_bible():
    """Create a test Bible with sample data.

    :returns: A Bible object populated with Genesis chapters 1-2
    """
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
    """Test extracting verses from Bible using index.

    :param text_range: Verse range string to parse
    :param verse_count: Expected number of verses to extract
    """
    bible = populate_bible()
    bible_index = bible.translate_to_bible_index(text_range)
    verses = bible.extract_texts_from_bible_index(*bible_index)
    assert verse_count == len(verses), f"{verse_count}!={len(verses)} doesn't match!"
