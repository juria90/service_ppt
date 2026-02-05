"""Preferences dialog window.

This module provides the PreferencesDialog class, a wxPython dialog for
displaying and editing application preferences and settings.
"""

from collections.abc import Callable
from typing import TYPE_CHECKING

import wx
import wx.adv
import wx.propgrid as wxpg

from service_ppt.bible import bibleformat
from service_ppt.utils.i18n import _
from service_ppt.wx_utils.dir_symbol_pg import DirSymbolPG

if TYPE_CHECKING:
    from service_ppt.preferences_config import PreferencesConfig

DEFAULT_SPAN = (1, 1)
BORDER_STYLE_EXCEPT_LEFT = wx.TOP | wx.RIGHT | wx.BOTTOM
BORDER_STYLE_EXCEPT_TOP = wx.LEFT | wx.RIGHT | wx.BOTTOM
BORDER_STYLE_RIGHT_BOTTOM = wx.LEFT | wx.RIGHT | wx.BOTTOM
DEFAULT_BORDER = 5


class PreferencesDialog(wx.adv.PropertySheetDialog):
    """PreferencesDialog class displays preferences settings, so that user can view and update the settings."""

    def __init__(self, parent: wx.Window, config: "PreferencesConfig", *args: object, **kwargs: object) -> None:
        self.is_modified: bool = False

        self.config: "PreferencesConfig" = config

        self.dir_dict: dict[str, str] = config.dir_dict

        self.current_bible_format: str = config.current_bible_format
        self.bible_rootdir: str = config.bible_rootdir
        self.current_bible_version: str = config.current_bible_version

        self.lyric_open_textfile: bool = config.lyric_open_textfile
        self.lyric_search_path: str = config.lyric_search_path
        self.lyric_copy_from_template: bool = config.lyric_copy_from_template
        self.lyric_application_pathname: str = config.lyric_application_pathname
        self.lyric_template_filename: str = config.lyric_template_filename

        self._dir_symbols_ctrl: DirSymbolPG | None = None

        self._bible_format_combo: wx.ComboBox | None = None
        self._bible_rootdir_stext: wx.StaticText | None = None
        self._bible_rootdir_ctrl: wx.DirPickerCtrl | None = None
        self._bible_version_combo: wx.ComboBox | None = None

        self._lyric_search_path_stext: wx.StaticText | None = None
        self._lyric_search_path_picker: wx.DirPickerCtrl | None = None
        self._lyric_open_check: wx.CheckBox | None = None
        self._lyric_copy_check: wx.CheckBox | None = None
        self._lyric_template_picker: wx.FilePickerCtrl | None = None

        resize_border = wx.RESIZE_BORDER
        wx.adv.PropertySheetDialog.__init__(
            self, parent, title=_("Preferences"), style=wx.DEFAULT_DIALOG_STYLE | resize_border, *args, **kwargs
        )

        self.initialize_controls()

        self.LayoutDialog()

        self.Centre()

    def initialize_controls(self) -> None:
        """initialize_controls creates all child controls and initialize positions."""
        self.CreateButtons(wx.OK | wx.CANCEL | wx.HELP)

        notebook = self.GetBookCtrl()

        directory_panel = self.create_directory_settings_page(notebook)
        tab_image1 = -1
        notebook.AddPage(directory_panel, _("Directory settings"), True, tab_image1)

        bible_panel = self.create_bible_settings_page(notebook)
        tab_image2 = -1
        notebook.AddPage(bible_panel, _("Bible settings"), True, tab_image2)

        lyric_panel = self.create_lyric_settings_page(notebook)
        tab_image3 = -1
        notebook.AddPage(lyric_panel, _("Lyric settings"), True, tab_image3)

        ok_button = self.FindWindow(wx.ID_OK)
        ok_button.Bind(wx.EVT_BUTTON, self.on_ok)

        cancel_button = self.FindWindow(wx.ID_CANCEL)
        cancel_button.Bind(wx.EVT_BUTTON, self.on_cancel)

    def on_ok(self, _: wx.CommandEvent) -> None:
        """Event handler for OK."""
        if self._dir_symbols_ctrl is None:
            return
        dir_dict = self._dir_symbols_ctrl.get_dir_symbols()
        if self.dir_dict != dir_dict:
            self.is_modified = True
            self.dir_dict = dir_dict

        self.Destroy()

    def on_cancel(self, _: wx.CommandEvent) -> None:
        """Event handler for Cancel."""
        self.is_modified = False
        self.Destroy()

    def create_bible_settings_page(self, parent: wx.Window) -> wx.Panel:
        """Create bible settings page."""
        panel = wx.Panel(parent)
        sizer = wx.GridBagSizer()
        row = 0

        # Bible Format
        _stext, self._bible_format_combo = self._create_combobox(panel, sizer, row, _("Bible &Format:"), self.on_bible_format_changed)
        row += 1

        # MyBible root directory
        label = _("Root &directory for MyBible, MySword, Zefania:")
        _stext, _ctrl = self._create_directory_picker(panel, sizer, row, label, self.bible_rootdir, self.on_bible_dir_changed)
        self._bible_rootdir_stext, self._bible_rootdir_ctrl = _stext, _ctrl
        row += 1

        # current bible version in combo
        label = _("Current Bible &Version:")
        _stext, self._bible_version_combo = self._create_combobox(panel, sizer, row, label, self.on_bible_version_changed)
        row += 1

        self._fill_bible_format()
        self._fill_bible_version()

        panel.SetSizerAndFit(sizer)

        return panel

    def _fill_bible_format(self) -> None:
        if self._bible_format_combo is None:
            return
        formats = bibleformat.get_format_list()
        self._bible_format_combo.Clear()
        self._bible_format_combo.AppendItems(formats)
        if self.current_bible_format is None and len(formats) > 0:
            self.current_bible_format = formats[0]
        self._bible_format_combo.SetStringSelection(self.current_bible_format)

    def on_bible_format_changed(self, _: wx.CommandEvent) -> None:
        if self._bible_format_combo is None or self._bible_rootdir_stext is None or self._bible_rootdir_ctrl is None:
            return
        self.current_bible_format = self._bible_format_combo.GetStringSelection()

        enable_dir = False
        if self.current_bible_format in [
            bibleformat.BibleFormat.MYBIBLE.value,
            bibleformat.BibleFormat.MYSWORD.value,
            bibleformat.BibleFormat.ZEFANIA.value,
        ]:
            enable_dir = True

        self._bible_rootdir_stext.Enable(enable_dir)
        self._bible_rootdir_ctrl.Enable(enable_dir)

        if enable_dir:
            self.on_bible_dir_changed(None)
        else:
            self._fill_bible_version()

        self.is_modified = True

    def on_bible_dir_changed(self, _: wx.FileDirPickerEvent | None) -> None:
        if self._bible_rootdir_ctrl is None:
            return
        self.bible_rootdir = self._bible_rootdir_ctrl.GetPath()
        bibleformat.set_format_option(self.current_bible_format, "ROOT_DIR", self.bible_rootdir)
        self._fill_bible_version()
        self.is_modified = True

    def _fill_bible_version(self) -> None:
        if self._bible_version_combo is None:
            return
        versions: list[str] = []
        if self.current_bible_format:
            versions = bibleformat.enum_versions(self.current_bible_format)
        self._bible_version_combo.Clear()
        self._bible_version_combo.AppendItems(versions)
        self._bible_version_combo.SetStringSelection(self.current_bible_version)

        # if nothing is selected, select the first one.
        if self._bible_version_combo.GetSelection() == wx.NOT_FOUND and self._bible_version_combo.GetCount():
            self._bible_version_combo.SetSelection(0)

    def on_bible_version_changed(self, _: wx.CommandEvent) -> None:
        if self._bible_version_combo is None:
            return
        self.current_bible_version = self._bible_version_combo.GetStringSelection()
        self.is_modified = True

    def create_directory_settings_page(self, parent: wx.Window) -> wx.Panel:
        """Create directory settings page."""
        panel = wx.Panel(parent)
        sizer = wx.BoxSizer(wx.VERTICAL)

        style = wxpg.PG_SPLITTER_AUTO_CENTER | wxpg.PG_TOOLBAR
        prop_grid = DirSymbolPG(panel, style=style)
        # Show help as tooltips
        prop_grid.SetExtraStyle(wxpg.PG_EX_HELP_AS_TOOLTIPS)
        prop_grid.populate_dir_symbols(self.config.dir_dict)
        self._dir_symbols_ctrl = prop_grid

        sizer.Add(prop_grid, proportion=100, flag=wx.EXPAND | wx.ALL, border=DEFAULT_BORDER)

        panel.SetSizerAndFit(sizer)

        return panel

    def create_lyric_settings_page(self, parent: wx.Window) -> wx.Panel:
        """Create lyric settings page."""
        panel = wx.Panel(parent)
        sizer = wx.GridBagSizer()
        row = 0

        # Lyric Search Path
        label = _("Lyric &search pathname:")
        _stext, _picker = self._create_directory_picker(panel, sizer, row, label, self.lyric_search_path, self.on_lyric_search_path_changed)
        self._lyric_search_path_stext, self._lyric_search_path_picker = _stext, _picker
        row += 1

        # Open Lyric Text file automatically?
        label = _("&Open lyric text file automatically.")
        self._lyric_open_check = self._create_checkbox(panel, sizer, row, label, self.lyric_open_textfile, self.on_lyric_open_check_changed)
        row += 1

        # Application path
        label = _("&Application pathname:")
        _stext, _picker = self._create_file_picker(
            panel, sizer, row, label, self.lyric_application_pathname, self.on_application_pathname_changed
        )
        self._lyric_application_picker: wx.FilePickerCtrl = _picker
        row += 1

        # Use template file if the lyric file doesn't exist yet.
        label = _("Copy from a &template file if the file not exists yet.")
        _check = self._create_checkbox(panel, sizer, row, label, self.lyric_copy_from_template, self.on_lyric_copy_check_changed)
        self._lyric_copy_check = _check
        row += 1

        # Template file path
        label = _("Template &filename:")
        _stext, _picker = self._create_file_picker(
            panel, sizer, row, label, self.lyric_template_filename, self.on_template_filename_changed
        )
        self._lyric_template_picker = _picker
        row += 1

        panel.SetSizerAndFit(sizer)

        return panel

    def _create_checkbox(
        self,
        panel: wx.Panel,
        sizer: wx.GridBagSizer,
        row: int,
        label: str,
        initial_value: bool,
        handler: "Callable[[wx.CommandEvent], None]",
    ) -> wx.CheckBox:
        checkbox = wx.CheckBox(panel, label=label)
        checkbox.SetValue(initial_value)
        checkbox.Bind(wx.EVT_CHECKBOX, handler, checkbox)
        sizer.Add(checkbox, pos=(row, 0), span=DEFAULT_SPAN, flag=wx.ALIGN_LEFT | wx.ALIGN_CENTER_VERTICAL | wx.ALL, border=DEFAULT_BORDER)

        return checkbox

    def _create_combobox(
        self, panel: wx.Panel, sizer: wx.GridBagSizer, row: int, label: str, handler: "Callable[[wx.CommandEvent], None]"
    ) -> tuple[wx.StaticText, wx.ComboBox]:
        stext = wx.StaticText(panel, label=label)
        sizer.Add(stext, pos=(row, 0), span=DEFAULT_SPAN, flag=wx.ALIGN_LEFT | wx.ALIGN_CENTER_VERTICAL | wx.ALL, border=DEFAULT_BORDER)

        cbox = wx.ComboBox(panel, style=wx.CB_READONLY)
        cbox.Bind(wx.EVT_COMBOBOX, handler, cbox)
        flag = wx.ALIGN_RIGHT | wx.ALIGN_CENTER_VERTICAL | wx.EXPAND | BORDER_STYLE_EXCEPT_LEFT
        sizer.Add(cbox, pos=(row, 1), span=DEFAULT_SPAN, flag=flag, border=DEFAULT_BORDER)

        return stext, cbox

    def _create_directory_picker(
        self,
        panel: wx.Panel,
        sizer: wx.GridBagSizer,
        row: int,
        label: str,
        initial_value: str,
        handler: "Callable[[wx.FileDirPickerEvent], None]",
    ) -> tuple[wx.StaticText, wx.DirPickerCtrl]:
        stext = wx.StaticText(panel, label=label)
        flag = wx.ALIGN_LEFT | wx.ALIGN_CENTER_VERTICAL | BORDER_STYLE_EXCEPT_TOP
        sizer.Add(stext, pos=(row, 0), span=DEFAULT_SPAN, flag=flag, border=DEFAULT_BORDER)

        picker = wx.DirPickerCtrl(panel, path=initial_value)
        picker.Bind(wx.EVT_DIRPICKER_CHANGED, handler, picker)
        flag = wx.ALIGN_RIGHT | wx.ALIGN_CENTER_VERTICAL | wx.EXPAND | BORDER_STYLE_RIGHT_BOTTOM
        sizer.Add(picker, pos=(row, 1), span=DEFAULT_SPAN, flag=flag, border=DEFAULT_BORDER)
        pickctrl = picker.GetPickerCtrl()
        if pickctrl:
            pickctrl.SetLabel(_("&Browse..."))

        return stext, picker

    def _create_file_picker(
        self,
        panel: wx.Panel,
        sizer: wx.GridBagSizer,
        row: int,
        label: str,
        initial_value: str,
        handler: "Callable[[wx.FileDirPickerEvent], None]",
    ) -> tuple[wx.StaticText, wx.FilePickerCtrl]:
        stext = wx.StaticText(panel, label=label)
        flag = wx.ALIGN_LEFT | wx.ALIGN_CENTER_VERTICAL | BORDER_STYLE_EXCEPT_TOP
        sizer.Add(stext, pos=(row, 0), span=DEFAULT_SPAN, flag=flag, border=DEFAULT_BORDER)

        picker = wx.FilePickerCtrl(panel, path=initial_value)
        picker.Bind(wx.EVT_FILEPICKER_CHANGED, handler, picker)
        flag = wx.ALIGN_RIGHT | wx.ALIGN_CENTER_VERTICAL | wx.EXPAND | BORDER_STYLE_RIGHT_BOTTOM
        sizer.Add(picker, pos=(row, 1), span=DEFAULT_SPAN, flag=flag, border=DEFAULT_BORDER)
        pickctrl = picker.GetPickerCtrl()
        if pickctrl:
            pickctrl.SetLabel(_("&Browse..."))

        return stext, picker

    def on_lyric_open_check_changed(self, _: wx.CommandEvent) -> None:
        if self._lyric_open_check is None:
            return
        self.lyric_open_textfile = self._lyric_open_check.GetValue()
        self.is_modified = True

    def on_lyric_search_path_changed(self, _: wx.FileDirPickerEvent) -> None:
        if self._lyric_search_path_picker is None:
            return
        self.lyric_search_path = self._lyric_search_path_picker.GetPath()
        self.is_modified = True

    def on_lyric_copy_check_changed(self, _: wx.CommandEvent) -> None:
        if self._lyric_copy_check is None:
            return
        self.lyric_copy_from_template = self._lyric_copy_check.GetValue()
        self.is_modified = True

    def on_application_pathname_changed(self, _: wx.FileDirPickerEvent) -> None:
        if not hasattr(self, "_lyric_application_picker") or self._lyric_application_picker is None:
            return
        self.lyric_application_pathname = self._lyric_application_picker.GetPath()
        self.is_modified = True

    def on_template_filename_changed(self, _: wx.FileDirPickerEvent) -> None:
        if self._lyric_template_picker is None:
            return
        self.lyric_template_filename = self._lyric_template_picker.GetPath()
        self.is_modified = True
