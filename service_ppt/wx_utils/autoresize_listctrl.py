"""Auto-resizing list control widget.

This module provides a wxPython ListCtrl widget that automatically resizes its
columns to fit the available width.
"""

import wx
import wx.lib.mixins.listctrl


class AutoresizeListCtrl(wx.ListCtrl, wx.lib.mixins.listctrl.ListCtrlAutoWidthMixin):
    def __init__(
        self,
        parent,
        ID=wx.ID_ANY,
        pos=wx.DefaultPosition,
        size=wx.DefaultSize,
        style=wx.LC_ICON,
        validator=wx.DefaultValidator,
        name=wx.ListCtrlNameStr,
    ):
        wx.ListCtrl.__init__(self, parent, ID, pos, size, style, validator, name)
        wx.lib.mixins.listctrl.ListCtrlAutoWidthMixin.__init__(self)
        self.setResizeColumn(0)
