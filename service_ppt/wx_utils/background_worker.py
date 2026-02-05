"""Background worker thread for asynchronous operations.

This module provides a worker thread implementation for executing commands
asynchronously in the background, allowing the GUI to remain responsive during
long-running operations like slide generation.
"""

from collections.abc import Callable
from threading import Thread
from typing import TYPE_CHECKING, Any

import wx

from service_ppt.wx_utils.autoresize_listctrl import AutoresizeListCtrl

if TYPE_CHECKING:
    pass

DEFAULT_BORDER = 5

# Define notification event for thread completion
EVT_RESULT_ID = wx.NewEventType()

# Create event binder for the custom event
EVT_RESULT = wx.PyEventBinder(EVT_RESULT_ID, 1)


class ResultEvent(wx.PyEvent):
    """Simple event to carry arbitrary result data."""

    def __init__(self, data: Any) -> None:
        """Init Result Event.

        :param data: Data to carry with the event
        """
        wx.PyEvent.__init__(self)
        self.SetEventType(EVT_RESULT_ID)
        self.data = data


# Thread class that executes the commands
class WorkerThread(Thread):
    """Worker Thread Class."""

    def __init__(self, notify_window: wx.Window, bkgnd_handler: Callable[[wx.Window], None]) -> None:
        """Init Worker Thread Class.

        :param notify_window: Window to notify when thread completes
        :param bkgnd_handler: Function to execute in background thread
        """
        Thread.__init__(self)
        self._notify_window = notify_window
        self._bkgnd_handler = bkgnd_handler
        # This starts the thread running on creation, but you could
        # also make the GUI thread responsible for calling this
        self.start()

    def run(self) -> None:
        """Run Worker Thread."""

        try:
            self._bkgnd_handler(self._notify_window)
        except Exception:
            # Exceptions are already handled and logged by the command handler
            # Don't print traceback here to avoid duplicate output
            pass

        # Notify to the main window.
        wx.PostEvent(self._notify_window, ResultEvent(0))


LABEL_MORE = "More"
LABEL_LESS = "Less"


class BkgndProgressDialog(wx.Dialog):
    """BkgndProgressDialog class displays the current progress status."""

    def __init__(self, parent: wx.Window | None, title: str, message: str, bkgnd_handler: Callable[[wx.Window], None]) -> None:
        """Initialize background progress dialog.

        :param parent: Parent window
        :param title: Dialog title
        :param message: Initial message to display
        :param bkgnd_handler: Function to execute in background thread
        """
        self.maximum = 100
        self.value = 0
        self.cancelled = False

        self.subrange_min = 0
        self.subrange_max = 100
        self.subrange_value = 0

        self.message = message

        self.message_ctrl: wx.StaticText | None = None
        self.guage_ctrl: wx.Gauge | None = None
        self.message_listctrl: AutoresizeListCtrl | None = None

        style = wx.CAPTION
        wx.Dialog.__init__(self, parent, title=title, style=style)
        self.Bind(wx.EVT_CLOSE, self.on_close)

        self.init_controls(message)

        # Set up event handler for any worker thread results
        self.Bind(EVT_RESULT, self.on_result)

        self.worker = WorkerThread(self, bkgnd_handler)

    def init_controls(self, message: str) -> None:
        """Initialize dialog controls.

        :param message: Initial message to display
        """
        sizer = wx.BoxSizer(wx.VERTICAL)

        text_for_size = (("W" * 40) + "\n") * 3
        self.message_ctrl = wx.StaticText(self, label=text_for_size, style=wx.ST_NO_AUTORESIZE)
        sizer.Add(self.message_ctrl, proportion=1, flag=wx.ALL | wx.FIXED_MINSIZE | wx.EXPAND, border=DEFAULT_BORDER)

        self.guage_ctrl = wx.Gauge(self, range=100, style=wx.GA_HORIZONTAL | wx.GA_PROGRESS)
        sizer.Add(self.guage_ctrl, proportion=0, flag=wx.LEFT | wx.BOTTOM | wx.RIGHT | wx.EXPAND, border=DEFAULT_BORDER)

        self.cp = cp = wx.CollapsiblePane(self, label=LABEL_MORE, style=wx.CP_DEFAULT_STYLE | wx.CP_NO_TLW_RESIZE)
        self.Bind(wx.EVT_COLLAPSIBLEPANE_CHANGED, self.on_pane_changed, cp)
        self.create_pane_content(cp.GetPane())
        sizer.Add(cp, 0, wx.ALL | wx.EXPAND, border=DEFAULT_BORDER)

        btnsizer = self.CreateSeparatedButtonSizer(wx.OK | wx.CANCEL)
        sizer.Add(btnsizer, 0, wx.ALL | wx.EXPAND, border=DEFAULT_BORDER)

        self.SetSizer(sizer)
        self.Fit()
        self.Center()

        self.message_ctrl.SetLabelText(message)

        close_btn = self.FindWindowById(wx.ID_OK)
        close_btn.Show(False)

    def create_pane_content(self, pane: wx.Window) -> None:
        """Create content for collapsible pane.

        :param pane: Pane window to add content to
        """
        style = wx.LC_REPORT | wx.LC_NO_HEADER | wx.LC_SINGLE_SEL
        self.message_listctrl = AutoresizeListCtrl(pane, style=style, name="messages")
        self.message_listctrl.InsertColumn(0, "Messages", width=wx.LIST_AUTOSIZE)
        border = wx.BoxSizer()
        border.Add(self.message_listctrl, 1, wx.EXPAND | wx.ALL, 5)
        pane.SetSizer(border)

    def on_pane_changed(self, _: wx.CommandEvent) -> None:
        """Handle collapsible pane state change.

        :param _: Event object (unused)
        """
        # redo the layout
        self.Layout()

        # and also change the labels
        if self.cp.IsExpanded():
            self.cp.SetLabel(LABEL_LESS)
            self.Fit()
        else:
            self.cp.SetLabel(LABEL_MORE)

    def on_close(self, _: wx.CloseEvent) -> None:
        """Event handler for Close.

        :param _: Event object (unused)
        """
        result = wx.ID_CANCEL if self.cancelled else wx.ID_OK
        self.EndModal(result)

    def on_btn_close(self, _: wx.CommandEvent) -> None:
        """Event handler for Close.

        :param _: Event object (unused)
        """
        self.EndModal(wx.ID_OK)

    def on_cancel(self, _: wx.CommandEvent) -> None:
        """Event handler for Cancel.

        :param _: Event object (unused)
        """
        self.cancelled = True

    # functions that will be called by worker thread.
    def on_result(self, _event: ResultEvent) -> None:
        """Worker thread is done and notified.

        :param _event: Result event from worker thread (unused)
        """

        # the worker is done: hide cancel
        # result = wx.ID_CANCEL if self.cancelled else wx.ID_OK
        # self.EndModal(result)
        close_btn = self.FindWindowById(wx.ID_OK)
        close_btn.Show(True)

        cancel_btn = self.FindWindowById(wx.ID_CANCEL)
        cancel_btn.Show(False)

        self.GetSizer().Layout()

    def progress_message(self, subrange_value: float, message: str) -> bool:
        """Update progress message and gauge.

        :param subrange_value: Progress value (0-100)
        :param message: Progress message to display
        :returns: True if not cancelled, False if cancelled
        """
        changed = False

        if subrange_value < 0:
            subrange_value = 0
        elif subrange_value > 100:
            subrange_value = 100

        if self.subrange_value != subrange_value:
            self.subrange_value = subrange_value
            changed = True

        value = subrange_value * (self.subrange_max - self.subrange_min) / self.maximum + self.subrange_min

        if self.value != value:
            self.value = value
            changed = True

        if not changed and message:
            changed = True

        if changed:
            self.message_ctrl.SetLabelText(message)
            self.guage_ctrl.SetValue(int(value))

        return not self.cancelled

    def error_message(self, message: str) -> None:
        """Add error message to the message list.

        :param message: Error message to display
        """
        count = self.message_listctrl.GetItemCount()
        self.message_listctrl.InsertItem(count, message)
        if not self.cp.IsExpanded():
            self.cp.Expand()
            evt = wx.CollapsiblePaneEvent(self, self.cp.GetId(), False)
            wx.PostEvent(self.cp, evt)

    def get_subrange(self) -> tuple[float, float]:
        """Get current subrange values.

        :returns: Tuple of (min, max) subrange values
        """
        return (self.subrange_min, self.subrange_max)

    def set_subrange(self, sub_min: float, sub_max: float) -> None:
        """Set subrange for progress calculation.

        Convert 0 to 100 to min & max.

        :param sub_min: Minimum subrange value
        :param sub_max: Maximum subrange value
        """
        self.subrange_min = sub_min
        self.subrange_max = sub_max

        # apply new subrange
        self.progress_message(0, "")
