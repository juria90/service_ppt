"""This file contains main window derived from Frame class.
"""

from enum import IntEnum
import gettext
import json
import os
import shutil
import typing

import wx

from autoresize_listctrl import AutoresizeListCtrl
from background_worker import BkgndProgressDialog
from bible import fileformat as bibfileformat
import command_ui as cmdui
import preferences_config as pc
import preferences_dialog as pd


_ = lambda s: s

DEFAULT_SPAN = (1, 1)
DEFAULT_BORDER = 5


class CommandEnum(IntEnum):
    ADD = 100
    DELETE = 101
    UP = 102
    DOWN = 103


class ILIDEnum(IntEnum):
    DUPLICATE_SLIDES = 0
    EXPORT_SLIDE_IMAGES = 1
    EXPORT_SHAPES_IMAGES = 2
    FIND_REPLACE_TEXTS = 3
    GENERATE_BIBLE_VERSES = 4
    INSERT_LYRICS = 5
    INSERT_SLIDES = 6
    OPEN_FILE = 7
    POPUP_MESSAGE = 8
    SAVE_FILE = 9


COMMAND_INFO = [
    # ILID,                 Image File,          UI String,               UI Class
    (ILIDEnum.DUPLICATE_SLIDES, "slide_duplicate.png", _("Duplicate slides"), cmdui.DuplicateWithTextUI),
    (ILIDEnum.EXPORT_SLIDE_IMAGES, "Save picture.png", _("Export slides as images"), cmdui.ExportSlidesUI),
    (ILIDEnum.EXPORT_SHAPES_IMAGES, "save shape.png", _("Export shapes in slide as images"), cmdui.ExportShapesUI),
    (ILIDEnum.FIND_REPLACE_TEXTS, "slide_text.png", _("Find and replace texts"), cmdui.SetVariablesUI),
    (ILIDEnum.GENERATE_BIBLE_VERSES, "slide_bible.png", _("Generate Bible verses slides"), cmdui.GenerateBibleVerseUI),
    (ILIDEnum.INSERT_LYRICS, "slide_note.png", _("Insert lyrics from files"), cmdui.InsertLyricsUI),
    (ILIDEnum.INSERT_SLIDES, "slide_insert.png", _("Insert slides from files"), cmdui.InsertSlidesUI),
    (ILIDEnum.OPEN_FILE, "Open.png", _("Open a presentation file"), cmdui.OpenFileUI),
    (ILIDEnum.POPUP_MESSAGE, "message.png", _("Pop up a message"), cmdui.PopupMessageUI),
    (ILIDEnum.SAVE_FILE, "Save.png", _("Save the presentation and other files"), cmdui.SaveFilesUI),
]

UICLS_TO_ILID_MAP = {pi[3]: pi[0] for pi in COMMAND_INFO}
ILID_TO_UICLS_MAP = {pi[0]: pi[3] for pi in COMMAND_INFO}


class Frame(wx.Frame):
    """Frame is the main wx.Frame class for service_ppt."""

    def __init__(self):
        # os.environ['LANGUAGE'] = 'ko'
        trans = gettext.translation("service_ppt", "locale", fallback=True)
        trans.install()
        global _
        _ = trans.gettext

        cmdui.set_translation(trans)
        pd.set_translation(trans)

        self.image_path24 = os.path.join(os.path.dirname(os.path.abspath(__file__)), "image24")
        self.image_path32 = os.path.join(os.path.dirname(os.path.abspath(__file__)), "image32")

        self.command_toolbar = None
        self.command_ctrl = None
        self.command_imglist = None
        self.settings_panel = None
        self.uimgr = cmdui.UIManager()

        self._filename = ""
        self.m_save = None

        self.config = wx.FileConfig("ServicePPT", vendorName="EMC")
        self.pconfig = pc.PreferencesConfig()
        self.pconfig.read_config(self.config)
        cmdui.GenerateBibleVerseUI.current_bible_format = self.pconfig.current_bible_format
        bibfileformat.set_format_option(self.pconfig.current_bible_format, "ROOT_DIR", self.pconfig.bible_rootdir)

        self.filehistory = wx.FileHistory(8)
        self.filehistory.Load(self.config)
        self.m_recent = None

        self.window_rect = None
        sw = pc.SW_RESTORED
        window_rect = None
        wp = self.pconfig.read_window_rect(self.config)
        if wp:
            sw = wp[0]
            window_rect = wp[1]
        else:
            sw = pc.SW_RESTORED
            disp = wx.Display(0)
            rc = disp.GetClientArea()
            tl = rc.GetTopLeft()
            wh = rc.GetSize()
            window_rect = (tl[0] + wh[0] // 4, tl[1] + wh[1] // 4, wh[0] // 2, wh[1] // 2)

        app_display_name = _("Service Presentation Generator")
        style = wx.DEFAULT_DIALOG_STYLE | wx.SYSTEM_MENU | wx.CLOSE_BOX | wx.RESIZE_BORDER | wx.MAXIMIZE_BOX | wx.MINIMIZE_BOX
        wx.Frame.__init__(self, None, title=app_display_name, style=style)
        wx.App.Get().SetAppDisplayName(app_display_name)
        self.Bind(wx.EVT_MOVE, self.on_move)
        self.Bind(wx.EVT_SIZE, self.on_move)

        self.Bind(wx.EVT_CLOSE, self.on_close)

        self.SetIcon(wx.Icon(os.path.join(self.image_path32, "bible.png")))

        self.create_menubar()
        self.create_toolbar()
        self.create_statusbar()
        self.init_controls()
        self._set_title()

        if self.window_rect:
            self.SetSize(*window_rect)
        if sw == pc.SW_ICONIZED:
            self.Iconize()
        elif sw == pc.SW_MAXIMIZED:
            self.Maximize()

    def create_menubar(self):
        """Creates default Menubar."""
        menubar = wx.MenuBar()
        self.SetMenuBar(menubar)

        m_file = wx.Menu()
        menubar.Append(m_file, _("&File"))

        m_open = m_file.Append(wx.ID_OPEN, _("&Open...\tCtrl+O"), _("Open an existing file."))
        self.Bind(wx.EVT_MENU, self.on_open, m_open)

        m_recent = wx.Menu()
        m_file.Append(wx.ID_ANY, _("Open &recent"), m_recent, _("Open a recently used file."))
        self.filehistory.UseMenu(m_recent)
        self.filehistory.AddFilesToMenu()
        self.Bind(wx.EVT_MENU_RANGE, self.on_file_history, id=wx.ID_FILE1, id2=wx.ID_FILE9)
        self.m_recent = m_recent

        self.m_save = m_file.Append(wx.ID_SAVE, _("&Save\tCtrl+S"), _("Save to a file."))
        self.Bind(wx.EVT_MENU, self.on_save, self.m_save)

        m_saveas = m_file.Append(wx.ID_SAVEAS, _("Save &as...\tCtrl+Shift+S"), _("Save as a file."))
        self.Bind(wx.EVT_MENU, self.on_saveas, m_saveas)

        m_file.AppendSeparator()

        m_run = m_file.Append(wx.ID_EXECUTE, _("&Execute...\tCtrl+E"), _("Executes the commands."))
        self.Bind(wx.EVT_MENU, self.on_execute, m_run)

        m_file.AppendSeparator()

        m_pref = m_file.Append(wx.ID_PREFERENCES, _("&Preferences...\tCtrl+,"), _("Show and edit preferences."))
        self.Bind(wx.EVT_MENU, self.on_preferences, m_pref)

        m_file.AppendSeparator()

        m_exit = m_file.Append(wx.ID_EXIT, _("E&xit\tCtrl+Q"), _("Close window and exit the program."))
        self.Bind(wx.EVT_MENU, self.on_close, m_exit)

        self.update_menu()

    def update_menu(self):
        enable = False
        if self._filename:
            enable = True
        if self.m_save:
            self.m_save.Enable(enable)

    def create_toolbar(self):
        """Creates default Toolbar."""
        toolbar = self.CreateToolBar()
        _t = toolbar.AddTool(wx.ID_OPEN, _("Open"), wx.Bitmap(os.path.join(self.image_path32, "Open.png")))
        _t = toolbar.AddTool(wx.ID_SAVE, _("Save"), wx.Bitmap(os.path.join(self.image_path32, "Save.png")))
        _t = toolbar.AddTool(wx.ID_EXECUTE, _("Execute"), wx.Bitmap(os.path.join(self.image_path32, "Play.png")))
        _t = toolbar.AddSeparator()
        _t = toolbar.AddTool(wx.ID_EXIT, _("Exit"), wx.Bitmap(os.path.join(self.image_path32, "Exit.png")))
        toolbar.Realize()

        self.toolbar = toolbar

    def create_statusbar(self):
        """Creates default Statusbar."""
        self.statusbar = super(Frame, self).CreateStatusBar()

    def create_command_toolbar(self, parent: wx.Window):
        toolbar = wx.ToolBar(parent, size=(200, -1), style=wx.TB_HORIZONTAL | wx.TB_BOTTOM)

        bitmap = wx.Bitmap(os.path.join(self.image_path24, "Add.png"))
        toolbar.AddTool(CommandEnum.ADD, _("Add Command"), bitmap)
        self.Bind(wx.EVT_TOOL, self.on_add_command, id=CommandEnum.ADD)

        bitmap = wx.Bitmap(os.path.join(self.image_path24, "Delete.png"))
        toolbar.AddTool(CommandEnum.DELETE, _("Delete Command"), bitmap)
        self.Bind(wx.EVT_TOOL, self.on_delete_command, id=CommandEnum.DELETE)

        bitmap = wx.Bitmap(os.path.join(self.image_path24, "Down.png"))
        toolbar.AddTool(CommandEnum.DOWN, _("Move Down"), bitmap)
        self.Bind(wx.EVT_TOOL, self.on_move_down_command, id=CommandEnum.DOWN)

        bitmap = wx.Bitmap(os.path.join(self.image_path24, "Up.png"))
        toolbar.AddTool(CommandEnum.UP, _("Move Up"), bitmap)
        self.Bind(wx.EVT_TOOL, self.on_move_up_command, id=CommandEnum.UP)

        toolbar.Realize()

        return toolbar

    def create_command_type_imagelist(self):
        il = None
        for pi in COMMAND_INFO:
            fn = pi[1]
            bitmap = wx.Bitmap(os.path.join(self.image_path32, fn))
            if il is None:
                il = wx.ImageList(bitmap.Width, bitmap.Height, mask=True)
            il.Add(bitmap)

        return il

    def init_controls(self):
        panel = wx.Panel(self)
        parent = panel
        sizer = wx.GridBagSizer(DEFAULT_BORDER, DEFAULT_BORDER)
        sizer.SetFlexibleDirection(wx.BOTH)
        sizer.SetNonFlexibleGrowMode(wx.FLEX_GROWMODE_SPECIFIED)

        # command panel
        self.command_toolbar = self.create_command_toolbar(parent)
        sizer.Add(self.command_toolbar, pos=(0, 0), span=DEFAULT_SPAN, flag=wx.ALL, border=DEFAULT_BORDER)

        command_style = wx.LC_REPORT | wx.LC_EDIT_LABELS | wx.LC_NO_HEADER | wx.LC_SINGLE_SEL
        self.command_ctrl = AutoresizeListCtrl(parent, size=wx.Size(200, -1), style=command_style, name="CommandList")
        self.command_ctrl.InsertColumn(0, "Command", width=wx.LIST_AUTOSIZE)
        # self.command_ctrl.EnableAlternateRowColours()
        self.command_imglist = self.create_command_type_imagelist()
        self.command_ctrl.SetImageList(self.command_imglist, wx.IMAGE_LIST_SMALL)
        self.Bind(wx.EVT_LIST_ITEM_FOCUSED, self.on_command_focused, self.command_ctrl)
        self.Bind(wx.EVT_LIST_END_LABEL_EDIT, self.on_command_end_labeledit, self.command_ctrl)
        sizer.Add(self.command_ctrl, pos=(1, 0), span=DEFAULT_SPAN, flag=wx.LEFT | wx.RIGHT | wx.BOTTOM | wx.EXPAND, border=DEFAULT_BORDER)

        # command settings
        self.settings_panel = wx.Panel(panel)
        ui_sizer = wx.BoxSizer(wx.VERTICAL)
        self.settings_panel.SetSizer(ui_sizer)

        sizer.Add(self.settings_panel, pos=(1, 1), span=DEFAULT_SPAN, flag=wx.RIGHT | wx.BOTTOM | wx.EXPAND, border=DEFAULT_BORDER)

        sizer.AddGrowableCol(1)
        sizer.AddGrowableRow(1)

        panel.SetSizerAndFit(sizer)

    def on_move(self, event: wx.Event):
        rc = self.GetRect()
        # print("window rect = (%d, %d, %d, %d)" % (rc.GetX(), rc.GetY(), rc.GetWidth(), rc.GetHeight()))
        if not self.IsIconized() and not self.IsMaximized():
            self.window_rect = (rc.GetX(), rc.GetY(), rc.GetWidth(), rc.GetHeight())

        event.Skip()

    def on_close(self, _event: wx.Event):
        """Event handler for EVT_CLOSE."""
        can_close = True
        can_save = False
        if self.uimgr.check_modified():
            dlg = wx.MessageDialog(
                self,
                _("The document is modified.\nDo you want to save it?"),
                wx.AppConsole.GetInstance().GetAppDisplayName(),
                wx.YES_NO | wx.CANCEL | wx.ICON_QUESTION,
            )
            result = dlg.ShowModal()
            dlg.Destroy()
            can_close = result != wx.ID_CANCEL
            can_save = result == wx.ID_YES

        if can_save:
            self.on_save(wx.Event())

        if can_close:
            sw = pc.SW_RESTORED
            if self.IsIconized():
                sw = pc.SW_ICONIZED
            elif self.IsMaximized():
                sw = pc.SW_MAXIMIZED

            if self.window_rect:
                self.pconfig.write_window_rect(self.config, sw, self.window_rect)

            self.Destroy()

    def on_open(self, _event: wx.Event):
        """Event handler for the File Open command."""
        self.open_file("", True)

    def on_file_history(self, event: wx.Event):
        fileNum = event.GetId() - wx.ID_FILE1
        path = self.filehistory.GetHistoryFile(fileNum)

        self.open_file(path, True)

    def open_file(self, filename: str, prompt: bool):
        if self.uimgr.check_modified() and prompt:
            result = wx.MessageBox(
                _("The document is modified.\nDo you want to continue?"),
                wx.App.Get().GetAppDisplayName(),
                style=wx.OK | wx.CANCEL | wx.ICON_QUESTION,
            )
            if result != wx.ID_OK:
                return

        if not filename:
            defDir, defFile = "", ""
            dlg = wx.FileDialog(
                self, _("Open File"), defDir, defFile, _("Service definition files (*.sdf)|*.sdf"), wx.FD_OPEN | wx.FD_FILE_MUST_EXIST
            )
            result = dlg.ShowModal()
            filename = dlg.GetPath()
            dlg.Destroy()
            if result == wx.ID_CANCEL:
                return
        try:
            self.uimgr.open(filename, self.pconfig.dir_dict)
        except (json.JSONDecodeError, Exception) as e:
            title = wx.App.Get().GetAppDisplayName()
            wx.MessageBox(_("Failed to open file '{filename}'.".format(filename=filename)), caption=title, style=wx.OK | wx.ICON_STOP)
            return

        self.populate_ui_items()

        self.update_file_history(filename)

        self.uimgr.set_modified(False)
        self._filename = filename
        self._set_title()
        self.update_menu()

    def populate_ui_items(self):
        self.command_ctrl.DeleteAllItems()

        for index, ui in enumerate(self.uimgr.command_ui_list):
            command_type = UICLS_TO_ILID_MAP[ui.__class__]
            self.command_ctrl.InsertItem(index, ui.name, command_type)

        if len(self.uimgr.command_ui_list) > 0:
            item = 0
            self.command_ctrl.SetItemState(
                item, wx.LIST_STATE_FOCUSED | wx.LIST_STATE_SELECTED, wx.LIST_STATE_FOCUSED | wx.LIST_STATE_SELECTED
            )

    def on_save(self, _event: wx.Event):
        self.save_file(False, _("Save to a file"))

    def on_saveas(self, _event: wx.Event):
        self.save_file(True, _("Save as a file"))

    def save_file(self, prompt: bool, title: str):
        filename = self._filename
        if prompt or not self._filename:
            defDir, defFile = "", ""
            if not self._filename:
                defDir, defFile = os.path.split(self._filename)

            dlg = wx.FileDialog(
                self, title, defDir, defFile, _("Service definition files (*.sdf)|*.sdf"), wx.FD_SAVE | wx.FD_OVERWRITE_PROMPT
            )
            result = dlg.ShowModal()
            filename = dlg.GetPath()
            dlg.Destroy()
            if result == wx.ID_CANCEL:
                return

        self.uimgr.save(filename, self.pconfig.dir_dict)

        self.update_file_history(filename)
        self.filehistory.Save(self.config)

        self.uimgr.set_modified(False)
        self._filename = filename
        self._set_title()
        self.update_menu()

    def update_file_history(self, filename: str):
        self.filehistory.AddFileToHistory(filename)
        self.filehistory.RemoveMenu(self.m_recent)
        self.filehistory.AddFilesToMenu()

    def on_execute(self, _event: wx.Event):
        def bkgnd_handler(window):
            self.uimgr.execute_commands(window, self.pconfig)

        dialog = BkgndProgressDialog(self, _("Running commands."), _("Intializing commands."), bkgnd_handler)
        dialog.ShowModal()
        dialog.Destroy()

    def _set_title(self):
        filename_title = _("(Untitled)")
        if self._filename:
            __, filename_title = os.path.split(self._filename)

        self.SetTitle(filename_title + " - " + wx.App.Get().GetAppDisplayName())

    def on_preferences(self, _event: wx.Event):
        """Event handler for the Preferences command."""
        dlg = pd.PreferencesDialog(self, self.pconfig)
        result = dlg.ShowModal()
        dlg.Destroy()
        if result != wx.ID_OK:
            pass

        if dlg.is_modified:
            self.pconfig.current_bible_format = dlg.current_bible_format
            self.pconfig.bible_rootdir = dlg.bible_rootdir
            self.pconfig.current_bible_version = dlg.current_bible_version

            self.pconfig.lyric_open_textfile = dlg.lyric_open_textfile
            self.pconfig.lyric_search_path = dlg.lyric_search_path
            self.pconfig.lyric_copy_from_template = dlg.lyric_copy_from_template
            self.pconfig.lyric_application_pathname = dlg.lyric_application_pathname
            self.pconfig.lyric_template_filename = dlg.lyric_template_filename

            self.pconfig.dir_dict = dlg.dir_dict

            self.pconfig.write_config(self.config)

    def get_lyric_file(self, filename: str) -> typing.Optional[str]:
        # empty basename when the filename is deleted.
        _dir, fn = os.path.split(filename)
        if not os.path.splitext(fn)[0]:
            return None

        xml_pathname = os.path.splitext(filename)[0] + ".xml"
        file_exist = os.path.exists(xml_pathname)
        if file_exist:
            return xml_pathname

        # seach xml file from search path
        if self.pconfig.lyric_search_path:
            xml_filename = os.path.splitext(fn)[0] + ".xml"
            searched_xml_pathname = os.path.join(self.pconfig.lyric_search_path, xml_filename)
            file_exist = os.path.exists(searched_xml_pathname)
            if file_exist:
                return searched_xml_pathname

        # copy template file and return it
        if self.pconfig.lyric_copy_from_template and self.pconfig.lyric_template_filename:
            shutil.copyfile(self.pconfig.lyric_template_filename, xml_pathname)
            file_exist = os.path.exists(xml_pathname)
            if file_exist:
                return xml_pathname

        return None

    def open_lyric_file(self, filename):
        xml_pathname = self.get_lyric_file(filename)
        if xml_pathname is None:
            return

        cmd = ""
        if self.pconfig.lyric_application_pathname:
            cmd = f'"{self.pconfig.lyric_application_pathname}" "{xml_pathname}"'
        else:
            ft = wx.TheMimeTypesManager.GetFileTypeFromExtension("txt")
            if ft:
                cmd = ft.GetOpenCommand(wx.FileType.MessageParameters(xml_pathname, ""))

        if cmd:
            wx.Execute(cmd, wx.EXEC_ASYNC)

    def on_add_command(self, _event: wx.Event):
        """Event handler for the CommandEnum.ADD command."""
        dlg = wx.SingleChoiceDialog(self, _("Please choose a slide command."), _("Add slide command"), [x[2] for x in COMMAND_INFO])
        result = dlg.ShowModal()
        command_index = dlg.GetSelection()
        command_name = COMMAND_INFO[command_index][2]
        dlg.Destroy()
        if result != wx.ID_OK:
            return

        index = self.command_ctrl.GetFirstSelected()
        if index < 0:
            index = self.command_ctrl.GetItemCount()
        self.insert_command(index, command_index, command_name)

    def on_delete_command(self, _event: wx.Event):
        """Event handler for the CommandEnum.DELETE command."""
        index = self.command_ctrl.GetFirstSelected()
        self.command_ctrl.DeleteItem(index)
        self.uimgr.delete_item(index)

    def on_move_down_command(self, _event: wx.Event):
        """Event handler for the CommandEnum.DOWN command."""
        index = self.command_ctrl.GetFirstSelected()
        if index + 1 < self.command_ctrl.GetItemCount():
            ui = self.uimgr.command_ui_list[index]
            self.uimgr.move_down_item(index)

            self.command_ctrl.DeleteItem(index)
            command_type = UICLS_TO_ILID_MAP[ui.__class__]
            self.command_ctrl.InsertItem(index + 1, ui.name, command_type)
            self.command_ctrl.Select(index + 1)

    def on_move_up_command(self, _event: wx.Event):
        """Event handler for the CommandEnum.UP command."""
        index = self.command_ctrl.GetFirstSelected()
        if index > 0:
            ui = self.uimgr.command_ui_list[index]
            self.uimgr.move_up_item(index)

            self.command_ctrl.DeleteItem(index)
            command_type = UICLS_TO_ILID_MAP[ui.__class__]
            self.command_ctrl.InsertItem(index - 1, ui.name, command_type)
            self.command_ctrl.Select(index - 1)

    def on_command_focused(self, event: wx.Event):
        """Event handler for the command_ui_list's EVT_LIST_ITEM_FOCUSED."""
        ui = self.uimgr.command_ui_list[event.GetIndex()]
        self.uimgr.activate(self.settings_panel, ui)

    def on_command_end_labeledit(self, event: wx.Event):
        index = event.GetIndex()
        name = event.GetText()
        self.uimgr.set_item_name(index, name)

    def insert_command(self, index: int, command_type: int, command_name: str):
        self.command_ctrl.InsertItem(index, command_name, command_type)

        uicls = ILID_TO_UICLS_MAP[command_type]
        ui = uicls(self.uimgr, command_name)
        self.uimgr.insert_item(index, ui)
