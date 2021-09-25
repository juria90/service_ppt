"""This file contains classes for Command UI that can edit the command settings.
"""

from io import StringIO
import json
import pickle
import re

import wx
import wx.propgrid as wxpg

from atomicfile import AtomicFileWriter
from bible import fileformat as bibfileformat
import command as cmd


def _(s):
    return s


DEFAULT_SPAN = (1, 1)
BORDER_STYLE_EXCEPT_TOP = wx.LEFT | wx.RIGHT | wx.BOTTOM
DEFAULT_BORDER = 5

POWERPOINT_FILES_WILDCARD = _("Powerpoint files (*.ppt;*.pptx)|*.ppt;*.pptx")


def set_translation(trans):
    global _
    _ = trans.gettext

    cmd.set_translation(trans)


class StringBuilder:
    _file_str = None

    def __init__(self):
        self._file_str = StringIO()

    def __str__(self):
        return self._file_str.getvalue()

    def append(self, str):
        self._file_str.write(str)


def unescape_backslash(s):
    """unescape_backslash() unescapes \ x into x for wxpg.LongStringProperty."""
    sb = StringBuilder()
    escape = False
    for i in range(len(s)):
        ch = s[i]
        if escape:
            if ch == "n":
                ch = "\n"
            elif ch == "r":
                ch = "\r"
            escape = False
            sb.append(ch)
        elif ch == "\\":
            escape = True
        else:
            escape = False
            sb.append(ch)

    result = str(sb)
    return result


def escape_backslash(s):
    """escape_backslash() escapes x into \ x for wxpg.LongStringProperty."""
    sb = StringBuilder()
    escape = False
    for i in range(len(s)):
        ch = s[i]
        if ch == "\\":
            sb.append("\\\\")
        else:
            sb.append(ch)

    result = str(sb)
    return result


class CommandUI(wx.EvtHandler):
    def __init__(self, uimgr, name, proc=None):
        super().__init__()

        self.uimgr = uimgr
        self.name = name
        self.command = proc
        self.ui = None

    def get_flattened_dict(self):
        return {
            "type": self.__class__.proc_class.__name__,
            "name": self.name,
            "data": self.command,
        }

    def activate(self, parent):
        if self.ui is None:
            self.construct_ui(parent)

        # append or use existing spacer at the end if needed.
        sizer = parent.GetSizer()
        count = sizer.GetItemCount()
        if self.need_stretch_spacer():
            si = sizer.GetItem(count - 1)
            if not si.IsSpacer():
                sizer.AddStretchSpacer()
        else:
            si = sizer.GetItem(count - 1)
            if si.IsSpacer():
                sizer.Remove(count - 1)

        self.TransferToWindow()
        self.ui.Show()
        return True

    def deactivate(self):
        if not self.TransferFromWindow():
            return False

        self.ui.Hide()
        return True

    def construct_ui(self, parent):
        self.ui = self.uimgr.get_ui_mapping(self.__class__.__name__)
        if self.ui is None:
            self.ui = self.construct_toplevel_ui(parent)
            sizer = parent.GetSizer()
            # insert ui at the beginning
            sizer.Prepend(self.ui, 1, wx.ALIGN_TOP | wx.EXPAND)
            self.uimgr.set_ui_mapping(self.__class__.__name__, self.ui)

    def construct_toplevel_ui(self, _parent):
        """Override this function in derived class to provide the own UI."""
        result = None
        return result

    def need_stretch_spacer(self):
        return False

    def TransferFromWindow(self):
        return True

    def TransferToWindow(self):
        return True

    def set_modified(self, a, b):
        """set_modified() calls set_modified if different and return the latter."""
        if a != b:
            self.uimgr.set_modified()
        return b

    @staticmethod
    def create_openfile_property(prop_name, wildcard):
        file_prop = wxpg.FileProperty(prop_name)
        # PG_DIALOG_TITLE
        file_prop.SetAttribute(wxpg.PG_FILE_DIALOG_STYLE, wx.FD_OPEN | wx.FD_FILE_MUST_EXIST)
        file_prop.SetAttribute(wxpg.PG_FILE_WILDCARD, wildcard)

        return file_prop

    @staticmethod
    def create_savefile_property(prop_name, wildcard):
        file_prop = wxpg.FileProperty(prop_name)
        # PG_DIALOG_TITLE
        file_prop.SetAttribute(wxpg.PG_FILE_DIALOG_STYLE, wx.FD_SAVE)
        file_prop.SetAttribute(wxpg.PG_FILE_WILDCARD, wildcard)

        return file_prop

    @staticmethod
    def create_datetime_property(prop_name):
        date_prop = wxpg.DateProperty(prop_name)
        # date_prop.SetAttribute(wxpg.PG_DATE_PICKER_STYLE, wx.DP_DEFAULT | DP_SHOWCENTURY)
        # date_prop.SetAttribute(wxpg.PG_DATE_FORMAT, wildcard)

        return date_prop


class PopupMessage(cmd.Command):
    def __init__(self, message):
        super().__init__()

        self.message = message

    def execute(self, cm, prs):
        cm.progress_message(0, _("Displaying PopupMessage."))

        title = wx.App.Get().GetAppDisplayName()
        wx.MessageBox(
            self.message,
            caption=title,
            style=wx.OK | wx.ICON_INFORMATION | wx.STAY_ON_TOP,
        )

        prs.check_modified()


class PopupMessageUI(CommandUI):
    proc_class = PopupMessage

    def __init__(self, uimgr, name, proc=None):
        super().__init__(uimgr, name, proc=proc)

        if self.command is None:
            self.command = PopupMessage("")

    def construct_toplevel_ui(self, parent):
        return wx.TextCtrl(parent)

    def TransferFromWindow(self):
        self.command.message = self.set_modified(self.command.message, self.ui.GetValue())
        return True

    def TransferToWindow(self):
        self.ui.SetValue(self.command.message)
        return True


class PropertyGridUI(CommandUI):
    TIMER_ID = 100

    def __init__(self, uimgr, name, proc=None):
        super().__init__(uimgr, name, proc=proc)

        self.dynamic_count = 0

        self.timer = wx.Timer(self, self.TIMER_ID)
        self.Bind(wx.EVT_TIMER, self.on_timer)

    def construct_toplevel_ui(self, parent):
        # Difference between using PropertyGridManager vs PropertyGrid is that
        # the manager supports multiple pages and a description box.
        style = wxpg.PG_SPLITTER_AUTO_CENTER | wxpg.PG_TOOLBAR
        prop_grid = wxpg.PropertyGrid(parent, style=style)

        # Show help as tooltips
        prop_grid.SetExtraStyle(wxpg.PG_EX_HELP_AS_TOOLTIPS)

        prop_grid.Bind(wxpg.EVT_PG_CHANGED, self.on_property_changed, prop_grid)

        self.initialize_fixed_properties(prop_grid)

        self.initialize_dynamic_properties(prop_grid)

        return prop_grid

    def initialize_fixed_properties(self, pg):
        pass

    def initialize_dynamic_properties(self, pg):
        pass

    def get_dynamic_label(self, _index):
        return ""

    def get_dynamic_property(self, _index):
        return None

    def get_dynamic_properties_from_window(self):
        plist = []
        i = 0
        while True:
            name = self.get_dynamic_label(i)
            if self.ui.GetProperty(name) is None:
                break
            plist.append(self.ui.GetPropertyValueAsString(name))
            i += 1

        while len(plist) and not plist[-1]:
            del plist[-1]

        return plist

    def set_dynamic_properties_to_window(self, plist):
        for i in range(len(plist) + 1):
            name = self.get_dynamic_label(i)
            if self.ui.GetProperty(name) is None:
                self.ui.Append(self.get_dynamic_property(i))

            self.ui.SetPropertyValueString(name, plist[i] if i < len(plist) else "")

        i = len(plist) + 1
        while True:
            name = self.get_dynamic_label(i)
            if self.ui.GetProperty(name) is None:
                break

            self.ui.RemoveProperty(name)
            i += 1

        self.dynamic_count = len(plist) + 1

    def SetPropertyValueString(self, name, value):
        if value is None:
            value = ""
        self.ui.SetPropertyValueString(name, value)

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
        self.ui.Append(self.get_dynamic_property(self.dynamic_count))
        self.dynamic_count = self.dynamic_count + 1


class OpenFileUI(PropertyGridUI):
    proc_class = cmd.OpenFile

    PRESENTATION_FILE = _("Presentation file")
    NOTES_FILE = _("Notes file")

    def __init__(self, uimgr, name, proc=None):
        super().__init__(uimgr, name, proc=proc)

        if self.command is None:
            self.command = cmd.OpenFile("", "")
        self.wildcard = POWERPOINT_FILES_WILDCARD

    def initialize_fixed_properties(self, pg):
        pg.Append(wxpg.PropertyCategory(_("File paths")))

        file_prop = self.create_openfile_property(self.PRESENTATION_FILE, self.wildcard)
        pg.Append(file_prop)

        file_prop = self.create_openfile_property(self.NOTES_FILE, _("Text files (*.txt)|*.txt"))
        pg.Append(file_prop)

    def TransferFromWindow(self):
        self.command.filename = self.set_modified(
            self.command.filename,
            self.ui.GetPropertyValueAsString(self.PRESENTATION_FILE),
        )
        self.command.notes_filename = self.set_modified(
            self.command.notes_filename,
            self.ui.GetPropertyValueAsString(self.NOTES_FILE),
        )
        return True

    def TransferToWindow(self):
        self.ui.SetPropertyValueString(self.PRESENTATION_FILE, self.command.filename)
        self.ui.SetPropertyValueString(self.NOTES_FILE, self.command.notes_filename)
        return True


class SaveFilesUI(PropertyGridUI):
    proc_class = cmd.SaveFiles

    PRESENTATION_FILE = _("Presentation file")
    LYRICS_ARCHIVE_FILE = _("Lyrics archive file")
    NOTES_FILE = _("Notes file")
    VERSES_FILE = _("Verses file")

    def __init__(self, uimgr, name, proc=None):
        super().__init__(uimgr, name, proc=proc)

        if self.command is None:
            self.command = cmd.SaveFiles("", "", "", "")
        self.wildcard = POWERPOINT_FILES_WILDCARD

    def initialize_fixed_properties(self, pg):
        pg.Append(wxpg.PropertyCategory(_("File paths")))

        file_prop = self.create_savefile_property(self.PRESENTATION_FILE, self.wildcard)
        pg.Append(file_prop)

        file_prop = self.create_savefile_property(self.LYRICS_ARCHIVE_FILE, _("OpenLP service files (*.osz)|*.osz|Zip files (*.zip)|*.zip"))
        pg.Append(file_prop)

        file_prop = self.create_savefile_property(self.NOTES_FILE, _("Text files (*.txt)|*.txt"))
        pg.Append(file_prop)

        file_prop = self.create_savefile_property(self.VERSES_FILE, _("Text files (*.txt)|*.txt"))
        pg.Append(file_prop)

    def TransferFromWindow(self):
        self.command.filename = self.set_modified(
            self.command.filename,
            self.ui.GetPropertyValueAsString(self.PRESENTATION_FILE),
        )
        self.command.lyrics_archive_filename = self.set_modified(
            self.command.lyrics_archive_filename,
            self.ui.GetPropertyValueAsString(self.LYRICS_ARCHIVE_FILE),
        )
        self.command.notes_filename = self.set_modified(
            self.command.notes_filename,
            self.ui.GetPropertyValueAsString(self.NOTES_FILE),
        )
        self.command.verses_filename = self.set_modified(
            self.command.verses_filename,
            self.ui.GetPropertyValueAsString(self.VERSES_FILE),
        )
        return True

    def TransferToWindow(self):
        self.ui.SetPropertyValueString(self.PRESENTATION_FILE, self.command.filename)
        self.ui.SetPropertyValueString(self.LYRICS_ARCHIVE_FILE, self.command.lyrics_archive_filename)
        self.ui.SetPropertyValueString(self.NOTES_FILE, self.command.notes_filename)
        self.ui.SetPropertyValueString(self.VERSES_FILE, self.command.verses_filename)
        return True


class SetVariablesUI(PropertyGridUI):
    proc_class = cmd.SetVariables

    DATETIME_NAME = _("Datetime variable name")
    DATETIME_VALUE = _("Datetime value")
    VARNAME_D = _("Variable name %d")
    VARVALUE_D = _("Variable value %d")

    def __init__(self, uimgr, name, proc=None):
        super().__init__(uimgr, name, proc=proc)

        if self.command is None:
            self.command = cmd.SetVariables()

    def get_dynamic_label(self, index):
        if (index % 2) == 0:
            return self.VARNAME_D % (index / 2 + 1)
        else:
            return self.VARVALUE_D % (index / 2 + 1)

    def get_dynamic_property(self, index):
        return wxpg.StringProperty(self.get_dynamic_label(index))

    def initialize_fixed_properties(self, pg):
        pg.Append(wxpg.PropertyCategory(_("Datetime")))

        pg.Append(wxpg.StringProperty(self.DATETIME_NAME))

        date_prop = self.create_datetime_property(self.DATETIME_VALUE)
        pg.Append(date_prop)

        pg.Append(wxpg.PropertyCategory(_("Variable names and values to be used for find and replace")))

    def initialize_dynamic_properties(self, pg):
        pg.Append(self.get_dynamic_property(0))
        pg.Append(self.get_dynamic_property(1))
        self.dynamic_count = 2

    def get_dynamic_properties_from_window(self):
        plist = []
        i = 0
        while True:
            name = self.get_dynamic_label(i)
            if self.ui.GetProperty(name) is None:
                break
            plist.append(self.ui.GetPropertyValueAsString(name))
            i += 1

        if len(plist) > 2 and not plist[-2] and not plist[-1]:
            del plist[-2]
            del plist[-1]
        return plist

    def set_dynamic_properties_to_window(self, plist):
        for i in range(len(plist) + 2):
            name = self.get_dynamic_label(i)
            if self.ui.GetProperty(name) is None:
                self.ui.Append(self.get_dynamic_property(i))

            self.ui.SetPropertyValueString(name, plist[i] if i < len(plist) else "")

        i = len(plist) + 2
        while True:
            name = self.get_dynamic_label(i)
            if self.ui.GetProperty(name) is None:
                break

            self.ui.RemoveProperty(name)
            i += 1

        self.dynamic_count = len(plist) + 2

    def TransferFromWindow(self):
        format_dict = {}
        name = self.ui.GetPropertyValueAsString(self.DATETIME_NAME)
        if name:
            dt_str = ""
            try:
                dt_value = self.ui.GetPropertyValueAsDateTime(self.DATETIME_VALUE)
                dt_str = dt_value.FormatISODate()
            except:
                pass

            # format comes from find_string
            format_dict[name] = cmd.DateTimeFormat(value=dt_str)

        self.command.format_dict = self.set_modified(self.command.format_dict, format_dict)

        plist = self.get_dynamic_properties_from_window()
        texts = {}
        for i in range(0, len(plist), 2):
            f = plist[i]
            r = plist[i + 1]
            if f:
                texts[f] = r

        self.command.str_dict = self.set_modified(self.command.str_dict, texts)

        return True

    def TransferToWindow(self):
        if len(self.command.format_dict) == 1:
            dt_name = next(iter(self.command.format_dict))
            self.SetPropertyValueString(self.DATETIME_NAME, dt_name)

            fobj = self.command.format_dict[dt_name]
            dt_value = cmd.DateTimeFormat.datetime_from_c_locale(fobj.value)
            self.ui.SetPropertyValue(self.DATETIME_VALUE, dt_value)

        plist = []
        for k, v in self.command.str_dict.items():
            plist.append(k)
            plist.append(v)

        self.set_dynamic_properties_to_window(plist)

        return True

    def on_property_changed(self, event):
        name = event.GetPropertyName()
        value = event.GetPropertyValue()
        label = self.get_dynamic_label(self.dynamic_count - 2)
        if value and name == label:
            # if last file and non empty value entered,
            # create new entry.
            self.ui.Append(self.get_dynamic_property(self.dynamic_count))
            self.ui.Append(self.get_dynamic_property(self.dynamic_count + 1))
            self.dynamic_count = self.dynamic_count + 2


class InsertSlidesUI(PropertyGridUI):
    proc_class = cmd.InsertSlides

    INSERT_LOCATION = _("Insert location")
    SEPARATOR_SLIDES = _("Separator slides")
    FILE_D = _("File %d")

    def __init__(self, uimgr, name, proc=None):
        super().__init__(uimgr, name, proc=proc)

        if self.command is None:
            self.command = cmd.InsertSlides(None, None, [])

    def get_dynamic_label(self, index):
        return self.FILE_D % (index + 1)

    def get_dynamic_property(self, index):
        return self.create_openfile_property(self.get_dynamic_label(index), POWERPOINT_FILES_WILDCARD)

    def initialize_fixed_properties(self, pg):
        pg.Append(wxpg.PropertyCategory(_("1 - Basic Settings")))
        pg.Append(wxpg.StringProperty(self.INSERT_LOCATION))
        pg.Append(wxpg.StringProperty(self.SEPARATOR_SLIDES))

        pg.Append(wxpg.PropertyCategory(_("2 - List of files to insert")))

    def initialize_dynamic_properties(self, pg):
        pg.Append(self.get_dynamic_property(0))
        self.dynamic_count = 1

    def TransferFromWindow(self):
        self.command.insert_location = self.set_modified(
            self.command.insert_location,
            self.ui.GetPropertyValueAsString(self.INSERT_LOCATION),
        )
        self.command.separator_slides = self.set_modified(
            self.command.separator_slides,
            self.ui.GetPropertyValueAsString(self.SEPARATOR_SLIDES),
        )

        self.command.filelist = self.set_modified(self.command.filelist, self.get_dynamic_properties_from_window())

        return True

    def TransferToWindow(self):
        self.SetPropertyValueString(self.INSERT_LOCATION, self.command.insert_location)
        self.SetPropertyValueString(self.SEPARATOR_SLIDES, self.command.separator_slides)

        self.set_dynamic_properties_to_window(self.command.filelist)

        return True


class InsertLyricsUI(PropertyGridUI):
    proc_class = cmd.InsertLyrics

    FILE_TYPE = _("Lyric file types")
    FILE_TYPE_LIST = [
        _("Lyric slide files"),
        _("Lyric text files"),
        _("Lyric slide and text files"),
    ]

    SLIDE_REPEAT_RANGE = _("Score repeat range")
    SLIDE_SEPARATOR_SLIDES = _("Score separator slides")

    LYRIC_REPEAT_RANGE = _("Lyric repeat range")
    LYRIC_SEPARATOR_SLIDES = _("Lyric separator slides")
    LYRIC_PATTERN = _("Lyric pattern")
    ARCHIVE_LYRIC_FILES = _("Archive lyric files")

    FILE_D = _("File %d")
    LYRIC_FILES_WILDCARD = _(
        "Powerpoint/Lyric files (*.ppt;*.pptx;*.xml)|*.ppt;*.pptx;*.xml|Powerpoint files (*.ppt;*.pptx)|*.ppt;*.pptx|Lyric xml files (*.xml)|*.xml"
    )

    def __init__(self, uimgr, name, proc=None):
        super().__init__(uimgr, name, proc=proc)

        if self.command is None:
            self.command = cmd.InsertLyrics("", None, "", None, "", False, [], 2)

    def get_dynamic_label(self, index):
        return self.FILE_D % (index + 1)

    def get_dynamic_property(self, index):
        return self.create_openfile_property(self.get_dynamic_label(index), self.LYRIC_FILES_WILDCARD)

    def initialize_fixed_properties(self, pg):
        pg.Append(wxpg.EnumProperty(self.FILE_TYPE, labels=self.FILE_TYPE_LIST, value=2))

        pg.Append(wxpg.PropertyCategory(_("1 - Slide settings")))
        pg.Append(wxpg.StringProperty(self.SLIDE_REPEAT_RANGE))
        pg.Append(wxpg.StringProperty(self.SLIDE_SEPARATOR_SLIDES))

        pg.Append(wxpg.PropertyCategory(_("2 - Lyric settings")))
        pg.Append(wxpg.StringProperty(self.LYRIC_REPEAT_RANGE))
        pg.Append(wxpg.StringProperty(self.LYRIC_SEPARATOR_SLIDES))
        pg.Append(wxpg.StringProperty(self.LYRIC_PATTERN))

        pg.Append(wxpg.BoolProperty(self.ARCHIVE_LYRIC_FILES))

        pg.Append(wxpg.PropertyCategory(_("3 - List of lyric files to insert")))

    def initialize_dynamic_properties(self, pg):
        pg.Append(self.get_dynamic_property(0))
        self.dynamic_count = 1

    def TransferFromWindow(self):
        index = self.ui.GetPropertyValueAsInt(self.FILE_TYPE)
        self.command.flags = self.set_modified(self.command.flags, index + 1)

        self.command.slide_insert_location = self.set_modified(
            self.command.slide_insert_location,
            self.ui.GetPropertyValueAsString(self.SLIDE_REPEAT_RANGE),
        )
        self.command.slide_separator_slides = self.set_modified(
            self.command.slide_separator_slides,
            self.ui.GetPropertyValueAsString(self.SLIDE_SEPARATOR_SLIDES),
        )

        self.command.lyric_insert_location = self.set_modified(
            self.command.lyric_insert_location,
            self.ui.GetPropertyValueAsString(self.LYRIC_REPEAT_RANGE),
        )
        self.command.lyric_separator_slides = self.set_modified(
            self.command.lyric_separator_slides,
            self.ui.GetPropertyValueAsString(self.LYRIC_SEPARATOR_SLIDES),
        )
        self.command.lyric_pattern = self.set_modified(
            self.command.lyric_pattern,
            self.ui.GetPropertyValueAsString(self.LYRIC_PATTERN),
        )
        self.command.archive_lyric_file = self.set_modified(
            self.command.archive_lyric_file,
            self.ui.GetPropertyValueAsBool(self.ARCHIVE_LYRIC_FILES),
        )

        self.command.filelist = self.set_modified(self.command.filelist, self.get_dynamic_properties_from_window())

        return True

    def TransferToWindow(self):
        index = 0
        if self.command.flags > 0:
            index = self.command.flags - 1
        self.SetPropertyValueString(self.FILE_TYPE, self.FILE_TYPE_LIST[index])

        self.SetPropertyValueString(self.SLIDE_REPEAT_RANGE, self.command.slide_insert_location)
        self.SetPropertyValueString(self.SLIDE_SEPARATOR_SLIDES, self.command.slide_separator_slides)

        self.SetPropertyValueString(self.LYRIC_REPEAT_RANGE, self.command.lyric_insert_location)
        self.SetPropertyValueString(self.LYRIC_SEPARATOR_SLIDES, self.command.lyric_separator_slides)
        self.SetPropertyValueString(self.LYRIC_PATTERN, self.command.lyric_pattern)

        self.ui.SetPropertyValue(self.ARCHIVE_LYRIC_FILES, self.command.archive_lyric_file)

        self.set_dynamic_properties_to_window(self.command.filelist)

        return True


class DuplicateWithTextUI(PropertyGridUI):
    proc_class = cmd.DuplicateWithText

    SLIDE_RANGE = _("Slide range")
    REPEAT_RANGE = _("Repeat range")
    FIND_TEXT = _("Text to find")
    REPLACE_TEXT = _("Texts to replace (a blank line for a separation)")
    PREPROCESS_SCRIPT = _("Preprocessing script before replace (Use input and output for input/output variables)")
    ARCHIVE_LYRIC_FILES = _("Archive as lyric files")
    OPTION_SPLIT_LINE_AT_EVERY = _("Optional split lines at every nth line")
    ENABLE_WORDWRAP = _("Enable word wrap")
    WORDWRAP_FONT = _("Font")
    PAGE_WITDH = _("Page width")

    def __init__(self, uimgr, name, proc=None):
        super().__init__(uimgr, name, proc=proc)

        if self.command is None:
            self.command = cmd.DuplicateWithText("", "", "", [], "", False, 0)

    def initialize_fixed_properties(self, pg):
        pg.Append(wxpg.PropertyCategory(_("1 - Range specification")))
        pg.Append(wxpg.StringProperty(self.SLIDE_RANGE))
        pg.Append(wxpg.StringProperty(self.REPEAT_RANGE))

        pg.Append(wxpg.PropertyCategory(_("2 - Text to find and replace")))
        pg.Append(wxpg.StringProperty(self.FIND_TEXT))
        pg.Append(wxpg.LongStringProperty(self.REPLACE_TEXT))
        pg.Append(wxpg.LongStringProperty(self.PREPROCESS_SCRIPT))

        pg.Append(wxpg.PropertyCategory(_("3 - Archive processing")))
        pg.Append(wxpg.BoolProperty(self.ARCHIVE_LYRIC_FILES))
        pg.Append(wxpg.StringProperty(self.OPTION_SPLIT_LINE_AT_EVERY))

        pg.Append(wxpg.PropertyCategory(_("4 - Word wrap archive text")))
        pg.Append(wxpg.BoolProperty(self.ENABLE_WORDWRAP))
        pg.Append(wxpg.FontProperty(self.WORDWRAP_FONT))
        pg.Append(wxpg.IntProperty(self.PAGE_WITDH))

    def TransferFromWindow(self):
        self.command.slide_range = self.set_modified(self.command.slide_range, self.ui.GetPropertyValueAsString(self.SLIDE_RANGE))
        self.command.repeat_range = self.set_modified(
            self.command.repeat_range,
            self.ui.GetPropertyValueAsString(self.REPEAT_RANGE),
        )

        self.command.find_text = self.ui.GetPropertyValueAsString(self.FIND_TEXT)
        text = self.ui.GetPropertyValueAsString(self.REPLACE_TEXT)
        text = unescape_backslash(text)
        lines = re.split(r"(\n){2}", text)
        lines = [l.strip() for l in lines if l.strip()]
        self.command.replace_texts = self.set_modified(self.command.replace_texts, lines)

        text = self.ui.GetPropertyValueAsString(self.PREPROCESS_SCRIPT)
        text = unescape_backslash(text)
        self.command.preprocessing_script = self.set_modified(self.command.preprocessing_script, text)

        self.command.archive_lyric_file = self.set_modified(
            self.command.archive_lyric_file,
            self.ui.GetPropertyValueAsBool(self.ARCHIVE_LYRIC_FILES),
        )
        optional_line_break = self.ui.GetPropertyValueAsString(self.OPTION_SPLIT_LINE_AT_EVERY)
        optional_line_break = int(optional_line_break) if optional_line_break.isdigit() else 0
        self.command.optional_line_break = self.set_modified(self.command.optional_line_break, optional_line_break)

        self.command.enable_wordwrap = self.set_modified(
            self.command.enable_wordwrap,
            self.ui.GetPropertyValueAsBool(self.ENABLE_WORDWRAP),
        )

        font_property = self.ui.GetPropertyByLabel(self.WORDWRAP_FONT)
        font_obj = font_property.GetValue()
        wordwrap_font = font_obj.GetNativeFontInfoDesc()
        self.command.wordwrap_font = self.set_modified(
            self.command.wordwrap_font,
            wordwrap_font,
        )
        wordwrap_pagewidth = int(self.ui.GetPropertyValueAsLongLong(self.PAGE_WITDH))
        self.command.wordwrap_pagewidth = self.set_modified(
            self.command.wordwrap_pagewidth,
            wordwrap_pagewidth,
        )

        return True

    def TransferToWindow(self):
        self.SetPropertyValueString(self.SLIDE_RANGE, self.command.slide_range)
        self.SetPropertyValueString(self.REPEAT_RANGE, self.command.repeat_range)

        self.ui.SetPropertyValueString(self.FIND_TEXT, self.command.find_text)
        replace_texts = "\n\n".join(self.command.replace_texts)
        replace_texts = escape_backslash(replace_texts)
        self.ui.SetPropertyValueString(self.REPLACE_TEXT, replace_texts)

        preprocessing_script = self.command.preprocessing_script
        preprocessing_script = escape_backslash(preprocessing_script)
        self.ui.SetPropertyValueString(self.PREPROCESS_SCRIPT, preprocessing_script)

        self.ui.SetPropertyValue(self.ARCHIVE_LYRIC_FILES, self.command.archive_lyric_file)
        self.ui.SetPropertyValueString(self.OPTION_SPLIT_LINE_AT_EVERY, str(self.command.optional_line_break))

        self.ui.SetPropertyValue(self.ENABLE_WORDWRAP, self.command.enable_wordwrap)

        font_property = self.ui.GetPropertyByLabel(self.WORDWRAP_FONT)
        font_obj = wx.Font(str(self.command.wordwrap_font))
        font_property.SetValue(font_obj)

        self.ui.SetPropertyValue(self.PAGE_WITDH, self.command.wordwrap_pagewidth)

        return True


class GenerateBibleVerseUI(PropertyGridUI):
    proc_class = cmd.GenerateBibleVerse
    current_bible_format = bibfileformat.FORMAT_MYBIBLE

    BIBLE_VERSION = _("Bible Version")
    MAIN_VERSE_NAME = _("Main Bible verse name")
    EACH_VERSE_NAME = _("Each Bible verse name")
    MAIN_VERSES = _("Main Bible verses")
    ADDITONAL_VERSES = _("Additional Bible verses")
    REPEAT_RANGE = _("Repeating slides range")

    def __init__(self, uimgr, name, proc=None):
        super().__init__(uimgr, name, proc=proc)

        if self.command is None:
            self.command = cmd.GenerateBibleVerse(GenerateBibleVerseUI.current_bible_format, "", "", "", "", "", "")

    def initialize_fixed_properties(self, pg):
        pg.Append(wxpg.PropertyCategory(_("1 - Bible specification")))
        versions = bibfileformat.enum_versions(GenerateBibleVerseUI.current_bible_format)
        pg.Append(wxpg.EnumProperty(self.BIBLE_VERSION, labels=versions, value=0))
        pg.Append(wxpg.StringProperty(self.MAIN_VERSE_NAME))
        pg.Append(wxpg.StringProperty(self.EACH_VERSE_NAME))
        pg.Append(wxpg.StringProperty(self.MAIN_VERSES))
        pg.Append(wxpg.StringProperty(self.ADDITONAL_VERSES))
        pg.Append(wxpg.StringProperty(self.REPEAT_RANGE))

    def TransferFromWindow(self):
        self.command.bible_version = self.set_modified(
            self.command.bible_version,
            self.ui.GetPropertyValueAsString(self.BIBLE_VERSION),
        )
        self.command.main_verse_name = self.set_modified(
            self.command.main_verse_name,
            self.ui.GetPropertyValueAsString(self.MAIN_VERSE_NAME),
        )
        self.command.each_verse_name = self.set_modified(
            self.command.each_verse_name,
            self.ui.GetPropertyValueAsString(self.EACH_VERSE_NAME),
        )
        self.command.main_verses = self.set_modified(self.command.main_verses, self.ui.GetPropertyValueAsString(self.MAIN_VERSES))
        self.command.additional_verses = self.set_modified(
            self.command.additional_verses,
            self.ui.GetPropertyValueAsString(self.ADDITONAL_VERSES),
        )
        self.command.repeat_range = self.set_modified(
            self.command.repeat_range,
            self.ui.GetPropertyValueAsString(self.REPEAT_RANGE),
        )

        return True

    def TransferToWindow(self):
        self.SetPropertyValueString(self.BIBLE_VERSION, self.command.bible_version)
        self.SetPropertyValueString(self.MAIN_VERSE_NAME, self.command.main_verse_name)
        self.SetPropertyValueString(self.EACH_VERSE_NAME, self.command.each_verse_name)
        self.SetPropertyValueString(self.MAIN_VERSES, self.command.main_verses)
        self.SetPropertyValueString(self.ADDITONAL_VERSES, self.command.additional_verses)
        self.SetPropertyValueString(self.REPEAT_RANGE, self.command.repeat_range)

        return True


class ExportSlidesUI(PropertyGridUI):
    proc_class = cmd.ExportSlides

    SLIDE_RANGE = _("Slides to export")
    IMAGE_TYPE = _("Image type")
    OUTPUT_DIR = _("Output directory")
    CLEANUP_OUTPUT_DIR = _("Clean up output directory")
    TRANSPARENT_IMAGE = _("Generate transparent image")
    TRANSPARENT_COLOR = _("Color to make transparent")

    def __init__(self, uimgr, name, proc=None):
        super().__init__(uimgr, name, proc=proc)

        if self.command is None:
            self.command = cmd.ExportSlides("", "", "PNG", 0, "#FFFFFF")

    def initialize_fixed_properties(self, pg):
        pg.Append(wxpg.StringProperty(self.SLIDE_RANGE))
        pg.Append(wxpg.EnumProperty(self.IMAGE_TYPE, labels=["GIF", "JPEG", "PNG", "TIF"], value=1))
        pg.Append(wxpg.DirProperty(self.OUTPUT_DIR))

        pg.Append(wxpg.BoolProperty(self.CLEANUP_OUTPUT_DIR))
        pg.Append(wxpg.BoolProperty(self.TRANSPARENT_IMAGE))
        pg.Append(wxpg.ColourProperty(self.TRANSPARENT_COLOR))

    def TransferFromWindow(self):
        self.command.slide_range = self.set_modified(self.command.slide_range, self.ui.GetPropertyValueAsString(self.SLIDE_RANGE))
        self.command.image_type = self.set_modified(self.command.image_type, self.ui.GetPropertyValueAsString(self.IMAGE_TYPE))
        self.command.out_dirname = self.set_modified(self.command.out_dirname, self.ui.GetPropertyValueAsString(self.OUTPUT_DIR))

        cleanup_output_dir = self.ui.GetPropertyValueAsBool(self.CLEANUP_OUTPUT_DIR)
        transparent_image = self.ui.GetPropertyValueAsBool(self.TRANSPARENT_IMAGE)
        flags = 0
        if cleanup_output_dir:
            flags = flags | cmd.Export_CleanupFiles
        if transparent_image:
            flags = flags | cmd.Export_Transparent
        self.command.flags = self.set_modified(self.command.flags, flags)
        color = self.ui.GetPropertyValue(self.TRANSPARENT_COLOR)
        str_color = color.GetAsString(wx.C2S_HTML_SYNTAX)
        self.command.color = self.set_modified(self.command.color, str_color)

        return True

    def TransferToWindow(self):
        self.SetPropertyValueString(self.SLIDE_RANGE, self.command.slide_range)
        self.SetPropertyValueString(self.IMAGE_TYPE, self.command.image_type)
        self.SetPropertyValueString(self.OUTPUT_DIR, self.command.out_dirname)

        value = (self.command.flags | cmd.Export_CleanupFiles) != 0
        self.ui.SetPropertyValue(self.CLEANUP_OUTPUT_DIR, value)
        value = (self.command.flags | cmd.Export_Transparent) != 0
        self.ui.SetPropertyValue(self.TRANSPARENT_IMAGE, value)
        self.SetPropertyValueString(self.TRANSPARENT_COLOR, self.command.color)

        return True


class ExportShapesUI(PropertyGridUI):
    proc_class = cmd.ExportShapes

    SLIDE_RANGE = _("Slides to export")
    IMAGE_TYPE = _("Image type")
    OUTPUT_DIR = _("Output directory")
    CLEANUP_OUTPUT_DIR = _("Clean up output directory")

    def __init__(self, uimgr, name, proc=None):
        super().__init__(uimgr, name, proc=proc)

        if self.command is None:
            self.command = cmd.ExportShapes("", "", "PNG", 0)

    def initialize_fixed_properties(self, pg):
        pg.Append(wxpg.StringProperty(self.SLIDE_RANGE))
        pg.Append(wxpg.EnumProperty(self.IMAGE_TYPE, labels=["GIF", "PNG"], value=1))
        pg.Append(wxpg.DirProperty(self.OUTPUT_DIR))

        pg.Append(wxpg.BoolProperty(self.CLEANUP_OUTPUT_DIR))

    def TransferFromWindow(self):
        self.command.slide_range = self.set_modified(self.command.slide_range, self.ui.GetPropertyValueAsString(self.SLIDE_RANGE))
        self.command.image_type = self.set_modified(self.command.image_type, self.ui.GetPropertyValueAsString(self.IMAGE_TYPE))
        self.command.out_dirname = self.set_modified(self.command.out_dirname, self.ui.GetPropertyValueAsString(self.OUTPUT_DIR))

        cleanup_output_dir = self.ui.GetPropertyValueAsBool(self.CLEANUP_OUTPUT_DIR)
        flags = 0
        if cleanup_output_dir:
            flags = flags | cmd.Export_CleanupFiles
        self.command.flags = self.set_modified(self.command.flags, flags)

        return True

    def TransferToWindow(self):
        self.SetPropertyValueString(self.SLIDE_RANGE, self.command.slide_range)
        self.SetPropertyValueString(self.IMAGE_TYPE, self.command.image_type)
        self.SetPropertyValueString(self.OUTPUT_DIR, self.command.out_dirname)

        value = (self.command.flags | cmd.Export_CleanupFiles) != 0
        self.ui.SetPropertyValue(self.CLEANUP_OUTPUT_DIR, value)

        return True


class CommandEncoder(json.JSONEncoder):

    proc_ui_list = [
        DuplicateWithTextUI,
        ExportSlidesUI,
        ExportShapesUI,
        SetVariablesUI,
        GenerateBibleVerseUI,
        InsertLyricsUI,
        InsertSlidesUI,
        OpenFileUI,
        PopupMessageUI,
        SaveFilesUI,
    ]
    proc_map = {ui.proc_class.__name__: ui for ui in proc_ui_list}

    format_list = [cmd.BibleVerseFormat, cmd.DateTimeFormat]

    format_map = {fo.__name__: fo for fo in format_list}

    def default(self, o):
        if isinstance(o, CommandUI):
            return o.get_flattened_dict()
        elif isinstance(o, cmd.Command) or isinstance(o, cmd.FormatObj):
            func = getattr(o, "get_flattened_dict", None)
            if callable(func):
                return o.get_flattened_dict()
            else:
                return o.__dict__
        else:
            super().default(o)

        return o

    @staticmethod
    def decoder(o):
        if "type" in o:
            proc_type = o["type"]
            ui_name = o["name"]
            enabled = True
            if "enabled" in o:
                enabled = o["enabled"]
            data = o["data"]
            if proc_type in CommandEncoder.proc_map:
                uicls = CommandEncoder.proc_map[proc_type]
                ui = uicls(None, ui_name)

                ui.command.__dict__.update(data)
                ui.command.enabled = enabled

                return ui
            else:
                # unsupported type
                return None
        elif "format_type" in o:
            format_type = o["format_type"]
            del o["format_type"]
            if format_type in CommandEncoder.format_map:
                fobj_cls = CommandEncoder.format_map[format_type]
                fobj = fobj_cls(None)

                fobj.__dict__.update(o)

                return fobj

        return o


class UIManager:
    def __init__(self):
        self.ui_map = {}
        self.active_ui = None
        self.command_ui_list = []
        self.modified = False

    def get_ui_mapping(self, name):
        if name in self.ui_map:
            return self.ui_map[name]

        return None

    def set_ui_mapping(self, name, ui):
        self.ui_map[name] = ui

    def activate(self, parent, ui):
        if self.active_ui == ui:
            return

        if self.active_ui is not None:
            if not self.active_ui.deactivate():
                return

        if ui is not None:
            self.active_ui = ui
            self.active_ui.activate(parent)

        sizer = parent.GetSizer()
        sizer.Layout()

    def deactivate(self):
        if self.active_ui is not None:
            if not self.active_ui.deactivate():
                return

        self.active_ui = None

    def get_modified(self):
        return self.modified

    def check_modified(self):
        if self.active_ui is not None:
            self.active_ui.TransferFromWindow()

        return self.modified

    def set_modified(self, modified=True):
        self.modified = modified

    def set_item_name(self, index, name):
        self.command_ui_list[index].name = name
        self.set_modified()

    def insert_item(self, index, ui):
        self.command_ui_list.insert(index, ui)
        self.set_modified()

    def delete_item(self, index):
        del self.command_ui_list[index]
        self.set_modified()

    def move_down_item(self, index):
        ui = self.command_ui_list.pop(index)
        self.command_ui_list.insert(index + 1, ui)
        self.set_modified()

    def move_up_item(self, index):
        ui = self.command_ui_list.pop(index)
        self.command_ui_list.insert(index - 1, ui)
        self.set_modified()

    def open(self, filename):
        command_ui_list = []
        with open(filename, "r", encoding="utf-8") as f:
            command_ui_list = json.load(f, object_hook=CommandEncoder.decoder)

        for ui in command_ui_list:
            ui.uimgr = self

        self.command_ui_list = command_ui_list

    def open_pickle(self, filename):
        with open(filename, "rb") as f:
            proc_list = pickle.load(f)

        command_ui_list = []
        for pair in proc_list:
            name, proc = pair
            for uicls in CommandEncoder.proc_ui_list:
                if isinstance(proc, uicls.proc_class):
                    ui = uicls(self, name, proc)
                    command_ui_list.append(ui)

        self.command_ui_list = command_ui_list

    def save(self, filename):
        self.check_modified()

        with AtomicFileWriter(filename, "w", encoding="utf-8") as f:
            json.dump(
                self.command_ui_list,
                f,
                indent=2,
                cls=CommandEncoder,
                ensure_ascii=False,
            )

    def save_pickle(self, filename):
        proc_list = [(x.name, x.command) for x in self.command_ui_list]
        with open(filename, "wb") as f:
            pickle.dump(proc_list, f, protocol=pickle.HIGHEST_PROTOCOL)

    def execute_commands(self, monitor):
        self.check_modified()

        proc_list = [x.command for x in self.command_ui_list]
        cm = cmd.CommandManager()
        cm.execute_commands(proc_list, monitor=monitor)
