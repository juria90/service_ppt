"""Text word wrapping and justification utilities.

This module provides functionality for wrapping text to fit within specified widths,
taking into account font metrics and margins. Used for formatting text in PowerPoint slides.
"""

from typing import TYPE_CHECKING

import wx

if TYPE_CHECKING:
    from types import TracebackType


class WordWrap:
    def __init__(self, fi: wx.FontInfo) -> None:
        """Initialize word wrap utility.

        :param fi: Font information for text rendering
        """
        self.dc: wx.MemoryDC | None = None
        self.fi = fi  # wx.FontInfo(54).FaceName("나눔고딕 ExtraBold").Bold()
        self.font: wx.Font | None = None

    def __del__(self) -> None:
        """Cleanup resources on deletion."""
        self.close()

    def __enter__(self) -> "WordWrap":
        """Enter context manager.

        :returns: Self
        """
        return self

    def __exit__(
        self,
        exc_type: type[BaseException] | None,
        exc_value: BaseException | None,
        exc_tb: "TracebackType | None",
    ) -> None:
        """Exit context manager.

        :param exc_type: Exception type if exception occurred
        :param exc_value: Exception value if exception occurred
        :param exc_tb: Exception traceback if exception occurred
        """
        self.close()

    def close(self) -> None:
        """Close and cleanup resources."""
        if self.font is not None:
            del self.font
        self.font = None

        if self.dc is not None:
            del self.dc
        self.dc = None

    def set_font_info(self, fi: wx.FontInfo) -> None:
        """Set font information for text rendering.

        :param fi: Font information
        """
        self.fi = fi
        if self.font is not None:
            del self.font
        self.font = None

    # https://github.com/wxWidgets/wxPython-Classic/blob/master/wx/lib/wordwrap.py
    def _wordwrap_dc(
        self,
        text: str,
        width: int,
        dc: wx.DC,
        break_long_words: bool = True,
        margin: int = 0,
        break_chars: str = " ",
    ) -> str:
        """Wrap text to fit within specified width.

        Returns a copy of text with newline characters inserted where long
        lines should be broken such that they will fit within the given
        width, with the given margin left and right, on the given `wx.DC`
        using its current font settings.  By default words that are wider
        than the margin-adjusted width will be broken at the nearest
        character boundary, but this can be disabled by passing ``False``
        for the ``break_long_words`` parameter.

        :param text: Text to wrap
        :param width: Maximum width in pixels
        :param dc: Device context for text measurement
        :param break_long_words: If True, break long words at character boundaries
        :param margin: Margin in characters
        :param break_chars: Characters that can be used for line breaks
        :returns: Wrapped text with newlines
        """

        wrapped_lines = []
        space_width = dc.GetTextExtent(" ")[0]
        text = text.split("\n")
        for line in text:
            pte = dc.GetPartialTextExtents(line)
            wid = width - (2 * margin + 1) * space_width - max([0] + [pte[i] - pte[i - 1] for i in range(1, len(pte))])
            idx = 0
            start = 0
            start_idx = 0
            spc_idx = -1
            while idx < len(pte):
                # remember the last seen space
                if line[idx] in break_chars:
                    spc_idx = idx

                # have we reached the max width?
                if pte[idx] - start > wid and (spc_idx != -1 or break_long_words):
                    if spc_idx != -1:
                        idx = min(spc_idx + 1, len(pte) - 1)
                    wrapped_lines.append(" " * margin + line[start_idx:idx] + " " * margin)
                    start = pte[idx]
                    start_idx = idx
                    spc_idx = -1

                idx += 1

            wrapped_lines.append(" " * margin + line[start_idx:idx] + " " * margin)

        return "\n".join(wrapped_lines)

    def wordwrap(self, text: str | list[str], page_width: int) -> str | list[str]:
        """Wrap text to fit within page width.

        :param text: Text string or list of text strings to wrap
        :param page_width: Page width in pixels
        :returns: Wrapped text (same type as input)
        """
        if self.dc is None:
            self.dc = wx.MemoryDC()
        dc = self.dc

        if self.font is None:
            self.font = wx.Font(self.fi)

        dc.SetFont(self.font)

        if isinstance(text, str):
            wrapped_text = self._wordwrap_dc(text, page_width, dc)
        elif isinstance(text, list):
            wrapped_text = []
            for t in text:
                new_text = self._wordwrap_dc(t, page_width, dc)
                wrapped_text.append(new_text)

        dc.SetFont(wx.NullFont)

        return wrapped_text
