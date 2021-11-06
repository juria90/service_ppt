"""This file contains PreferencesDialog class.
"""

import wx
import wx.adv

from bible import fileformat as bibfileformat


_ = lambda s: s

DEFAULT_SPAN = (1, 1)
BORDER_STYLE_EXCEPT_LEFT = wx.TOP | wx.RIGHT | wx.BOTTOM
BORDER_STYLE_EXCEPT_TOP = wx.LEFT | wx.RIGHT | wx.BOTTOM
BORDER_STYLE_RIGHT_BOTTOM = wx.LEFT | wx.RIGHT | wx.BOTTOM
DEFAULT_BORDER = 5


def set_translation(trans):
    global _
    _ = trans.gettext


class PreferencesDialog(wx.adv.PropertySheetDialog):
    """PreferencesDialog class displays preferences settings
    so that user can view and update the settings.
    """

    def __init__(self, parent, config, *args, **kwargs):
        self.is_modified = False

        self.config = config
        self.current_bible_format = config.current_bible_format
        self.bible_rootdir = config.bible_rootdir
        self.current_bible_version = config.current_bible_version

        self.lyric_open_textfile = config.lyric_open_textfile
        self.lyric_copy_from_template = config.lyric_copy_from_template
        self.lyric_application_pathname = config.lyric_application_pathname
        self.lyric_template_filename = config.lyric_template_filename

        self._bible_format_combo = None
        self._bible_rootdir_stext = None
        self._bible_rootdir_ctrl = None
        self._bible_version_combo = None

        self._lyric_open_check = None
        self._lyric_copy_check = None
        self._lyric_template_picker = None

        resize_border = wx.RESIZE_BORDER
        wx.adv.PropertySheetDialog.__init__(
            self, parent, title=_("Preferences"), style=wx.DEFAULT_DIALOG_STYLE | resize_border, *args, **kwargs
        )

        self.initialize_controls()

        self.LayoutDialog()

        self.Centre()

    def initialize_controls(self):
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

    def on_ok(self, _):
        """Event handler for OK."""
        self.Destroy()

    def on_cancel(self, _):
        """Event handler for Cancel."""
        self.is_modified = False
        self.Destroy()

    def create_bible_settings_page(self, parent):
        """Create bible settings page."""
        panel = wx.Panel(parent)

        sizer = wx.GridBagSizer()
        row = 0

        # Bible Format
        stext = wx.StaticText(panel, label=_("Bible &Format:"))
        sizer.Add(
            stext,
            pos=(row, 0),
            span=DEFAULT_SPAN,
            flag=wx.ALIGN_LEFT | wx.ALIGN_CENTER_VERTICAL | wx.ALL,
            border=DEFAULT_BORDER,
        )

        cbox = wx.ComboBox(panel, style=wx.CB_READONLY)
        cbox.Bind(wx.EVT_COMBOBOX, self.on_bible_format_changed, cbox)
        sizer.Add(
            cbox,
            pos=(row, 1),
            span=DEFAULT_SPAN,
            flag=wx.ALIGN_RIGHT | wx.ALIGN_CENTER_VERTICAL | wx.EXPAND | BORDER_STYLE_EXCEPT_LEFT,
            border=DEFAULT_BORDER,
        )
        self._bible_format_combo = cbox

        # MyBible root directory
        row = row + 1
        self._bible_rootdir_stext = wx.StaticText(panel, label=_("Root &directory for MyBible, MySword, Zefania:"))
        sizer.Add(
            self._bible_rootdir_stext,
            pos=(row, 0),
            span=DEFAULT_SPAN,
            flag=wx.ALIGN_LEFT | wx.ALIGN_CENTER_VERTICAL | BORDER_STYLE_EXCEPT_TOP,
            border=DEFAULT_BORDER,
        )

        self._bible_rootdir_ctrl = wx.DirPickerCtrl(panel, path=self.bible_rootdir)
        self._bible_rootdir_ctrl.Bind(
            wx.EVT_DIRPICKER_CHANGED,
            self.on_bible_dir_changed,
            self._bible_rootdir_ctrl,
        )
        sizer.Add(
            self._bible_rootdir_ctrl,
            pos=(row, 1),
            span=DEFAULT_SPAN,
            flag=wx.ALIGN_RIGHT | wx.ALIGN_CENTER_VERTICAL | wx.EXPAND | BORDER_STYLE_RIGHT_BOTTOM,
            border=DEFAULT_BORDER,
        )
        pickctrl = self._bible_rootdir_ctrl.GetPickerCtrl()
        if pickctrl:
            pickctrl.SetLabel(_("&Browse..."))

        # current bible version in combo
        row = row + 1
        stext = wx.StaticText(panel, label=_("Current Bible &Version:"))
        sizer.Add(
            stext,
            pos=(row, 0),
            span=DEFAULT_SPAN,
            flag=wx.ALIGN_LEFT | wx.ALIGN_CENTER_VERTICAL | BORDER_STYLE_EXCEPT_TOP,
            border=DEFAULT_BORDER,
        )

        cbox = wx.ComboBox(panel, style=wx.CB_READONLY)
        cbox.Bind(wx.EVT_COMBOBOX, self.on_bible_version_changed, cbox)
        sizer.Add(
            cbox,
            pos=(row, 1),
            span=DEFAULT_SPAN,
            flag=wx.ALIGN_RIGHT | wx.ALIGN_CENTER_VERTICAL | wx.EXPAND | BORDER_STYLE_RIGHT_BOTTOM,
            border=DEFAULT_BORDER,
        )
        self._bible_version_combo = cbox

        self._fill_bible_format()
        self._fill_bible_version()

        panel.SetSizerAndFit(sizer)

        return panel

    def _fill_bible_format(self):
        formats = bibfileformat.get_format_list()
        self._bible_format_combo.Clear()
        self._bible_format_combo.AppendItems(formats)
        if self.current_bible_format == None and len(formats) > 0:
            self.current_bible_format = formats[0]
        self._bible_format_combo.SetStringSelection(self.current_bible_format)

    def on_bible_format_changed(self, _):
        self.current_bible_format = self._bible_format_combo.GetStringSelection()

        enable_dir = False
        if self.current_bible_format in [
            bibfileformat.FORMAT_MYBIBLE,
            bibfileformat.FORMAT_MYSWORD,
            bibfileformat.FORMAT_ZEFANIA,
        ]:
            enable_dir = True

        self._bible_rootdir_stext.Enable(enable_dir)
        self._bible_rootdir_ctrl.Enable(enable_dir)

        if enable_dir:
            self.on_bible_dir_changed(None)
        else:
            self._fill_bible_version()

        self.is_modified = True

    def on_bible_dir_changed(self, _):
        self.bible_rootdir = self._bible_rootdir_ctrl.GetPath()
        bibfileformat.set_format_option(self.current_bible_format, "ROOT_DIR", self.bible_rootdir)
        self._fill_bible_version()
        self.is_modified = True

    def _fill_bible_version(self):
        versions = []
        if self.current_bible_format:
            versions = bibfileformat.enum_versions(self.current_bible_format)
        self._bible_version_combo.Clear()
        self._bible_version_combo.AppendItems(versions)
        self._bible_version_combo.SetStringSelection(self.current_bible_version)

        # if nothing is selected, select the first one.
        if self._bible_version_combo.GetSelection() == wx.NOT_FOUND and self._bible_version_combo.GetCount():
            self._bible_version_combo.SetSelection(0)

    def on_bible_version_changed(self, _):
        self.current_bible_version = self._bible_version_combo.GetStringSelection()
        self.is_modified = True

    def create_directory_settings_page(self, parent):
        """Create directory settings page."""
        panel = wx.Panel(parent)

        return panel

    def create_lyric_settings_page(self, parent):
        """Create lyric settings page."""
        panel = wx.Panel(parent)

        sizer = wx.GridBagSizer()
        row = 0

        # Open Lyric Text file automatically?
        self._lyric_open_check = wx.CheckBox(panel, label=_("&Open lyric text file automatically."))
        self._lyric_open_check.SetValue(self.lyric_open_textfile)
        self._lyric_open_check.Bind(
            wx.EVT_CHECKBOX,
            self.on_lyric_open_check_changed,
            self._lyric_open_check,
        )
        sizer.Add(
            self._lyric_open_check,
            pos=(row, 0),
            span=DEFAULT_SPAN,
            flag=wx.ALIGN_LEFT | wx.ALIGN_CENTER_VERTICAL | wx.ALL,
            border=DEFAULT_BORDER,
        )

        # Application path
        row = row + 1
        stext = wx.StaticText(panel, label=_("&Application pathname:"))
        sizer.Add(
            stext,
            pos=(row, 0),
            span=DEFAULT_SPAN,
            flag=wx.ALIGN_LEFT | wx.ALIGN_CENTER_VERTICAL | BORDER_STYLE_EXCEPT_TOP,
            border=DEFAULT_BORDER,
        )

        self._lyric_application_picker = wx.FilePickerCtrl(panel, path=self.lyric_application_pathname)
        self._lyric_application_picker.Bind(
            wx.EVT_FILEPICKER_CHANGED,
            self.on_application_pathname_changed,
            self._lyric_application_picker,
        )
        sizer.Add(
            self._lyric_application_picker,
            pos=(row, 1),
            span=DEFAULT_SPAN,
            flag=wx.ALIGN_RIGHT | wx.ALIGN_CENTER_VERTICAL | wx.EXPAND | BORDER_STYLE_RIGHT_BOTTOM,
            border=DEFAULT_BORDER,
        )
        pickctrl = self._lyric_application_picker.GetPickerCtrl()
        if pickctrl:
            pickctrl.SetLabel(_("&Browse..."))

        # Use template file if the lyric file doesn't exist yet.
        row = row + 1
        self._lyric_copy_check = wx.CheckBox(panel, label=_("Copy from a &template file if the file not exists yet."))
        self._lyric_copy_check.SetValue(self.lyric_copy_from_template)
        self._lyric_copy_check.Bind(
            wx.EVT_CHECKBOX,
            self.on_lyric_copy_check_changed,
            self._lyric_copy_check,
        )
        sizer.Add(
            self._lyric_copy_check,
            pos=(row, 0),
            span=DEFAULT_SPAN,
            flag=wx.ALIGN_LEFT | wx.ALIGN_CENTER_VERTICAL | wx.ALL,
            border=DEFAULT_BORDER,
        )

        # Template file path
        row = row + 1
        stext = wx.StaticText(panel, label=_("Template &filename:"))
        sizer.Add(
            stext,
            pos=(row, 0),
            span=DEFAULT_SPAN,
            flag=wx.ALIGN_LEFT | wx.ALIGN_CENTER_VERTICAL | BORDER_STYLE_EXCEPT_TOP,
            border=DEFAULT_BORDER,
        )

        self._lyric_template_picker = wx.FilePickerCtrl(panel, path=self.lyric_template_filename)
        self._lyric_template_picker.Bind(
            wx.EVT_FILEPICKER_CHANGED,
            self.on_template_filename_changed,
            self._lyric_template_picker,
        )
        sizer.Add(
            self._lyric_template_picker,
            pos=(row, 1),
            span=DEFAULT_SPAN,
            flag=wx.ALIGN_RIGHT | wx.ALIGN_CENTER_VERTICAL | wx.EXPAND | BORDER_STYLE_RIGHT_BOTTOM,
            border=DEFAULT_BORDER,
        )
        pickctrl = self._lyric_template_picker.GetPickerCtrl()
        if pickctrl:
            pickctrl.SetLabel(_("&Browse..."))

        panel.SetSizerAndFit(sizer)

        return panel

    def on_lyric_open_check_changed(self, _):
        self.lyric_open_textfile = self._lyric_open_check.GetValue()
        self.is_modified = True

    def on_lyric_copy_check_changed(self, _):
        self.lyric_copy_from_template = self._lyric_copy_check.GetValue()
        self.is_modified = True

    def on_application_pathname_changed(self, _):
        self.lyric_application_pathname = self._lyric_application_picker.GetPath()
        self.is_modified = True

    def on_template_filename_changed(self, _):
        self.lyric_template_filename = self._lyric_template_picker.GetPath()
        self.is_modified = True
