"""Tests for biblang module.

This module contains unit tests for Bible language and internationalization
functionality, including verse range parsing.
"""

from pytest import mark

from service_ppt.bible.biblang import L18N, LANG_EN

test_data = [
    ("Genesis 1:1", "Genesis", 1, 1, None, None),
    ("Genesis 1:1-2", "Genesis", 1, 1, None, 2),
    ("Genesis 1:1-2:2", "Genesis", 1, 1, 2, 2),
]


@mark.parametrize("text_range,bt,ct1,vs1,ct2,vs2", test_data)
def test_parse_verse_range(text_range, bt, ct1, vs1, ct2, vs2):
    """Test parsing verse range strings.

    :param text_range: Input verse range string
    :param bt: Expected book title
    :param ct1: Expected first chapter number
    :param vs1: Expected first verse number
    :param ct2: Expected second chapter number (if range spans chapters)
    :param vs2: Expected second verse number (if range)
    """
    ac_bt, ac_ct1, ac_vs1, ac_ct2, ac_vs2 = L18N.parse_verse_range(LANG_EN, text_range)
    assert ac_bt == bt, f"{ac_bt}!={bt} doesn't match!"
    assert ac_ct1 == ct1, f"{ac_ct1}!={ct1} doesn't match!"
    assert ac_vs1 == vs1, f"{ac_vs1}!={vs1} doesn't match!"
    assert ac_ct2 == ct2, f"{ac_ct2}!={ct2} doesn't match!"
    assert ac_vs2 == vs2, f"{ac_vs2}!={vs2} doesn't match!"
