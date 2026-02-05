"""Auto-resizing list control widget.

This module provides a wxPython ListCtrl widget that automatically resizes its
columns to fit the available width.
"""

from typing import TYPE_CHECKING

import wx
import wx.lib.mixins.listctrl

if TYPE_CHECKING:
    from wx import Validator


class AutoresizeListCtrl(wx.ListCtrl, wx.lib.mixins.listctrl.ListCtrlAutoWidthMixin):
    def __init__(
        self,
        parent: wx.Window,
        id: int = wx.ID_ANY,
        pos: wx.Point = wx.DefaultPosition,
        size: wx.Size = wx.DefaultSize,
        style: int = wx.LC_ICON,
        validator: "Validator" = wx.DefaultValidator,
        name: str = wx.ListCtrlNameStr,
    ) -> None:
        """Initialize auto-resizing list control.

        :param parent: Parent window
        :param id: Window ID (defaults to wx.ID_ANY)
        :param pos: Window position
        :param size: Window size
        :param style: List control style flags
        :param validator: Window validator
        :param name: Window name
        """
        wx.ListCtrl.__init__(self, parent, id, pos, size, style, validator, name)
        wx.lib.mixins.listctrl.ListCtrlAutoWidthMixin.__init__(self)
        self.setResizeColumn(0)
