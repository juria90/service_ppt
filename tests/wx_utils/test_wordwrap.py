"""Tests for wordwrap module."""

from typing import TYPE_CHECKING

import pytest

if TYPE_CHECKING:
    import wx

from service_ppt.wx_utils.wordwrap import WordWrap

SAMPLE_TEXT_ENGLISH = """In the beginning, God created the heavens and the earth.
The earth was without form and void, and darkness was over the face of the deep. And the Spirit of God was hovering over the face of the waters.
And God said, “Let there be light,” and there was light.
And God saw that the light was good. And God separated the light from the darkness.
God called the light Day, and the darkness he called Night. And there was evening and there was morning, the first day."""

# Korean Pangram.
SAMPLE_TEXT_KOREAN = """키스의 고유 조건은 입술끼리 만나야하고, 특별한 기술은 필요치 않다.. 1234567890"""


@pytest.mark.parametrize(
    "sample_text,min_line_count",
    [(SAMPLE_TEXT_ENGLISH, 5), (SAMPLE_TEXT_KOREAN, 1)],
)
def test_wordwrap(wx_app: "wx.App", sample_text: str, min_line_count: int):
    """Test word wrapping functionality.

    :param wx_app: wx application instance (ensures wx is initialized)
    :param sample_text: Text sample to test word wrapping with
    :param min_line_count: Minimum number of lines in the wrapped text
    """
    import wx

    fi = wx.FontInfo(54).FaceName("나눔고딕 ExtraBold").Bold()

    page_width = 1200  # pixel

    ww = WordWrap(fi)
    wrapped_text = ww.wordwrap(sample_text, page_width)
    assert isinstance(wrapped_text, str)
    assert len(wrapped_text.splitlines()) > min_line_count
