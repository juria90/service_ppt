#
# text justification
#

import wx


class WordWrap:
    def __init__(self, fi):
        self.dc = None
        self.fi = fi  # wx.FontInfo(54).FaceName("나눔고딕 ExtraBold").Bold()
        self.font = None

    def __del__(self):
        self.close()

    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc_value, traceback):
        self.close()

    def close(self):
        if self.font is not None:
            del self.font
        self.font = None

        if self.dc is not None:
            del self.dc
        self.dc = None

    def set_font_info(self, fi: wx.FontInfo):
        self.fi = fi
        if self.font is not None:
            del self.font
        self.font = None

    # https://github.com/wxWidgets/wxPython-Classic/blob/master/wx/lib/wordwrap.py
    def _wordwrap_dc(self, text, width, dc, breakLongWords=True, margin=0, breakChars=" "):
        """
        Returns a copy of text with newline characters inserted where long
        lines should be broken such that they will fit within the given
        width, with the given margin left and right, on the given `wx.DC`
        using its current font settings.  By default words that are wider
        than the margin-adjusted width will be broken at the nearest
        character boundary, but this can be disabled by passing ``False``
        for the ``breakLongWords`` parameter.
        """

        wrapped_lines = []
        space_width = dc.GetTextExtent(" ")[0]
        text = text.split("\n")
        for line in text:
            pte = dc.GetPartialTextExtents(line)
            wid = width - (2 * margin + 1) * space_width - max([0] + [pte[i] - pte[i - 1] for i in range(1, len(pte))])
            idx = 0
            start = 0
            startIdx = 0
            spcIdx = -1
            while idx < len(pte):
                # remember the last seen space
                if line[idx] in breakChars:
                    spcIdx = idx

                # have we reached the max width?
                if pte[idx] - start > wid and (spcIdx != -1 or breakLongWords):
                    if spcIdx != -1:
                        idx = min(spcIdx + 1, len(pte) - 1)
                    wrapped_lines.append(" " * margin + line[startIdx:idx] + " " * margin)
                    start = pte[idx]
                    startIdx = idx
                    spcIdx = -1

                idx += 1

            wrapped_lines.append(" " * margin + line[startIdx:idx] + " " * margin)

        return "\n".join(wrapped_lines)

    def wordwrap(self, text: str, page_width: int):
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


def test_wordwrap():
    sample_text = """Dynamic programming is both a mathematical optimization method and a computer programming method. \
    The method was developed by Richard Bellman in the 1950s and has found applications in numerous fields, from aerospace engineering to economics.
    In both contexts it refers to simplifying a complicated problem by breaking it down into simpler sub-problems in a recursive manner. \
    While some decision problems cannot be taken apart this way, decisions that span several points in time do often break apart recursively. \
    Likewise, in computer science, if a problem can be solved optimally by breaking it into sub-problems and then recursively finding the optimal solutions to the sub-problems, \
    then it is said to have optimal substructure.
    If sub-problems can be nested recursively inside larger problems, so that dynamic programming methods are applicable, \
    then there is a relation between the value of the larger problem and the values of the sub-problems. \
    In the optimization literature this relationship is called the Bellman equation.
    """

    sample_text1 = """8월 5일(목)~8월 7일(토) 오후 4~7시까지 열리는 2021 VBS (여름성경학교)의 참가신청서가 본당 앞에 비치되어 있습니다.\n작성하셔서 교육부장(심언택 장로)님께 제출해 주시기 바랍니다."""

    _app = wx.App(False)
    fi = wx.FontInfo(54).FaceName("나눔고딕 ExtraBold").Bold()

    page_width = 1200  # pixel
    page_height = 300  # pixel

    ww = WordWrap()
    # ww.set_font_info(fi)
    wrapped_text = ww.wordwrap(sample_text1, page_width)
    print(wrapped_text)


if __name__ == "__main__":
    test_wordwrap()
