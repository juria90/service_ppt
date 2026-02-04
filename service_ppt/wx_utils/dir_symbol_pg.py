"""Directory symbol property grid widget.

This module provides a property grid widget for managing directory symbols,
which are used to define variable names that map to directory paths for use
in service definition files.
"""

import wx
import wx.propgrid as wxpg

from service_ppt.utils.i18n import _


class DirSymbolPG(wxpg.PropertyGrid):
    VARNAME_D = _("Symbol %d")
    VARVALUE_D = _("Directory %d")

    TIMER_ID = 100

    def __init__(self, *args, **kw):
        super().__init__(*args, **kw)
        self.dynamic_count = 0

        self.timer = wx.Timer(self, self.TIMER_ID)
        self.Bind(wx.EVT_TIMER, self.on_timer)

        self.Bind(wxpg.EVT_PG_CHANGED, self.on_property_changed, self)
        # self.ui.Unbind(wxpg.EVT_PG_CHANGED, handler=self.on_property_changed)

    def add_dir_symbol(self, name: str, dirname: str):
        values = [name, dirname]
        for i in range(2):
            prop = self.get_dynamic_property(self.dynamic_count + i)
            prop.SetValue(values[i])
            self.Append(prop)
        self.dynamic_count += 2

    def populate_dir_symbols(self, dir_symbols):
        for key, value in dir_symbols.items():
            self.add_dir_symbol(key, value)
            # prop_grid.add_dir_symbol("ONLINE_SERVICE_DIR", r"C:\Users\juria\Church\OnlineService")

        # Add blank one to add more
        self.add_dir_symbol("", "")

    def get_dir_symbols(self):
        dir_dict = {}
        i = 0
        while True:
            prop_name = self.get_dynamic_label(i)
            symbol = self.GetPropertyValue(prop_name)
            if not symbol:
                break
            i += 1

            prop_name = self.get_dynamic_label(i)
            directory = self.GetPropertyValue(prop_name)
            if not directory:
                break
            i += 1

            dir_dict[symbol] = directory

        return dir_dict

    def get_dynamic_property(self, index):
        prop_name = self.get_dynamic_label(index)
        if (index % 2) == 0:
            return wxpg.StringProperty(prop_name)
        return wxpg.DirProperty(prop_name)

    def get_dynamic_label(self, index):
        if (index % 2) == 0:
            return DirSymbolPG.VARNAME_D % (index / 2 + 1)
        return DirSymbolPG.VARVALUE_D % (index / 2 + 1)

    def on_property_changed(self, event):
        name = event.GetPropertyName()
        value = event.GetPropertyValue()
        label = self.get_dynamic_label(self.dynamic_count - 1)
        if value and name == label:
            # if last file and non empty value entered,
            # create new entry.
            # But, creating new entry here, causes recursion somehow,
            # so, do it in timer callback.
            self.timer.StartOnce(100)

    def on_timer(self, _event):
        self.add_dir_symbol("", "")
