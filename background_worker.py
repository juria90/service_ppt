'''
'''

import sys
import time
from threading import *
import traceback
import wx

from autoresize_listctrl import AutoresizeListCtrl


DEFAULT_BORDER = 5

# Define notification event for thread completion
EVT_RESULT_ID = wx.NewId()

def EVT_RESULT(win, func):
    """Define Result Event."""
    win.Connect(-1, -1, EVT_RESULT_ID, func)


class ResultEvent(wx.PyEvent):
    """Simple event to carry arbitrary result data."""
    def __init__(self, data):
        """Init Result Event."""
        wx.PyEvent.__init__(self)
        self.SetEventType(EVT_RESULT_ID)
        self.data = data


# Thread class that executes the commands
class WorkerThread(Thread):
    """Worker Thread Class."""
    def __init__(self, notify_window, bkgnd_handler):
        """Init Worker Thread Class."""
        Thread.__init__(self)
        self._notify_window = notify_window
        self._bkgnd_handler = bkgnd_handler
        # This starts the thread running on creation, but you could
        # also make the GUI thread responsible for calling this
        self.start()

    def run(self):
        """Run Worker Thread."""

        try:
            self._bkgnd_handler(self._notify_window)
        except:
            # e = sys.exc_info()[0]
            # print("Error: %s" % e)
            traceback.print_exc()

        # Notify to the main window.
        wx.PostEvent(self._notify_window, ResultEvent(0))


label1 = "More"
label2 = "Less"


class BkgndProgressDialog(wx.Dialog):
    '''BkgndProgressDialog class displays the current progress status.
    '''

    def __init__(self, parent, title, message, bkgnd_handler):
        self.maximum = 100
        self.value = 0
        self.cancelled = False

        self.subrange_min = 0
        self.subrange_max = 100
        self.subrange_value = 0

        self.message = message

        self.message_ctrl = None
        self.guage_ctrl = None
        self.message_listctrl = None

        style = wx.CAPTION
        wx.Dialog.__init__(self, parent, title=title, style=style)
        self.Bind(wx.EVT_CLOSE, self.on_close)

        self.init_controls(message)

        # Set up event handler for any worker thread results
        EVT_RESULT(self, self.on_result)

        self.worker = WorkerThread(self, bkgnd_handler)

    def init_controls(self, message):
        sizer = wx.BoxSizer(wx.VERTICAL)

        text_for_size = (('W'*40)+'\n') * 3
        self.message_ctrl = wx.StaticText(self, label=text_for_size, style=wx.ST_NO_AUTORESIZE)
        sizer.Add(self.message_ctrl, proportion=1, flag=wx.ALL|wx.FIXED_MINSIZE|wx.EXPAND, border=DEFAULT_BORDER)

        self.guage_ctrl = wx.Gauge(self, range=100, style=wx.GA_HORIZONTAL|wx.GA_PROGRESS)
        sizer.Add(self.guage_ctrl, proportion=0, flag=wx.LEFT|wx.BOTTOM|wx.RIGHT|wx.EXPAND, border=DEFAULT_BORDER)

        self.cp = cp = wx.CollapsiblePane(self, label=label1,
                                          style=wx.CP_DEFAULT_STYLE|wx.CP_NO_TLW_RESIZE)
        self.Bind(wx.EVT_COLLAPSIBLEPANE_CHANGED, self.on_pane_changed, cp)
        self.create_pane_content(cp.GetPane())
        sizer.Add(cp, 0, wx.ALL|wx.EXPAND, border=DEFAULT_BORDER)

        btnsizer = self.CreateSeparatedButtonSizer(wx.OK|wx.CANCEL)
        sizer.Add(btnsizer, 0, wx.ALL|wx.EXPAND, border=DEFAULT_BORDER)

        self.SetSizer(sizer)
        self.Fit()
        self.Center()

        self.message_ctrl.SetLabelText(message)

        close_btn = self.FindWindowById(wx.ID_OK)
        close_btn.Show(False)

    def create_pane_content(self, pane):
        style = wx.LC_REPORT|wx.LC_NO_HEADER|wx.LC_SINGLE_SEL
        self.message_listctrl = AutoresizeListCtrl(pane, style=style, name='messages')
        self.message_listctrl.InsertColumn(0, 'Messages', width=wx.LIST_AUTOSIZE)
        border = wx.BoxSizer()
        border.Add(self.message_listctrl, 1, wx.EXPAND|wx.ALL, 5)
        pane.SetSizer(border)

    def on_pane_changed(self, _):
        # redo the layout
        self.Layout()

        # and also change the labels
        if self.cp.IsExpanded():
            self.cp.SetLabel(label2)
            self.Fit()
        else:
            self.cp.SetLabel(label1)

    def on_close(self, _):
        '''Event handler for Close.
        '''
        result = wx.ID_CANCEL if self.cancelled else wx.ID_OK
        self.EndModal(result)

    def on_btn_close(self, _):
        '''Event handler for Close.
        '''
        self.EndModal(wx.ID_OK)

    def on_cancel(self, _):
        '''Event handler for Cancel.
        '''
        self.cancelled = True

    # functions that will be called by worker thread.
    def on_result(self, _event):
        """Worker thread is done and notified."""

        # the worker is done: hide cancel
        # result = wx.ID_CANCEL if self.cancelled else wx.ID_OK
        # self.EndModal(result)
        close_btn = self.FindWindowById(wx.ID_OK)
        close_btn.Show(True)

        cancel_btn = self.FindWindowById(wx.ID_CANCEL)
        cancel_btn.Show(False)

        self.GetSizer().Layout()

    def progress_message(self, subrange_value, message):
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
            self.guage_ctrl.SetValue(value)

        result = not self.cancelled
        return result

    def error_message(self, message):
        count = self.message_listctrl.GetItemCount()
        self.message_listctrl.InsertItem(count, message)
        if not self.cp.IsExpanded():
            self.cp.Expand()
            evt = wx.CollapsiblePaneEvent(self, self.cp.GetId(), False)
            wx.PostEvent(self.cp, evt)

    def get_subrange(self):
        return (self.subrange_min, self.subrange_max)

    def set_subrange(self, sub_min, sub_max):
        '''Convert 0 to 100 to min & max
        '''
        self.subrange_min = sub_min
        self.subrange_max = sub_max

        # apply new subrange
        self.progress_message(0, '')
