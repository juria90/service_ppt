"""Command UI classes for editing command settings.

This module provides user interface classes for each command type, allowing
users to configure and edit command parameters through dialog boxes and
property grids.
"""

from __future__ import annotations

import copy
from io import StringIO
import json
import re
from typing import TYPE_CHECKING, Any, ClassVar

import wx
import wx.propgrid as wxpg

from service_ppt.bible import bibleformat
from service_ppt.cmd.cmd import (
    BibleVerseFormat,
    Command,
    DateTimeFormat,
    DirSymbols,
    DuplicateWithText,
    ExportFlag,
    ExportShapes,
    ExportSlides,
    GenerateBibleVerse,
    InsertLyrics,
    InsertSlides,
    OpenFile,
    SaveFiles,
    SetVariables,
)
from service_ppt.cmd.cmdmgr import CommandManager
from service_ppt.utils.atomicfile import AtomicFileWriter
from service_ppt.utils.i18n import _

if TYPE_CHECKING:
    from service_ppt.cmd.cmd import Command

DEFAULT_SPAN: tuple[int, int] = (1, 1)
BORDER_STYLE_EXCEPT_TOP: int = wx.LEFT | wx.RIGHT | wx.BOTTOM
DEFAULT_BORDER: int = 5

POWERPOINT_FILES_WILDCARD: str = _("Powerpoint files (*.ppt;*.pptx)|*.ppt;*.pptx")


class StringBuilder:
    """String builder utility for constructing strings incrementally.

    This class provides a simple interface for building strings by appending
    content, similar to Java's StringBuilder class.
    """

    _file_str: StringIO | None = None

    def __init__(self) -> None:
        """Initialize a new StringBuilder instance."""
        self._file_str = StringIO()

    def __str__(self) -> str:
        """Return the accumulated string content.

        :returns: The complete string that has been built
        """
        return self._file_str.getvalue()

    def append(self, text: str) -> None:
        """Append text to the string builder.

        :param text: The text to append
        """
        self._file_str.write(text)


def unescape_backslash(s: str) -> str:
    r"""Unescape backslash sequences in a string.

    Converts escape sequences like \n and \r to their actual characters.
    Used for processing text from wxpg.LongStringProperty.

    :param s: The string to unescape
    :returns: The unescaped string
    """
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

    return str(sb)


def escape_backslash(s: str) -> str:
    r"""Escape backslashes in a string.

    Escapes backslash characters by doubling them. Used for preparing
    text for wxpg.LongStringProperty.

    :param s: The string to escape
    :returns: The escaped string
    """
    sb = StringBuilder()
    for i in range(len(s)):
        ch = s[i]
        if ch == "\\":
            sb.append("\\\\")
        else:
            sb.append(ch)

    return str(sb)


class MyFileProperty(wxpg.FileProperty):
    """File property that handles filename changes by opening lyric files.

    This class extends wxpg.FileProperty to automatically open lyric files
    when the filename is changed by the user, unless explicitly disabled.
    """

    ignore_open_file: bool = False

    def OnSetValue(self) -> None:
        """Handle value change event.

        Opens the lyric file in the main frame when a new filename is set,
        unless ignore_open_file is True.
        """
        if self.ignore_open_file:
            return

        tlws = wx.GetTopLevelWindows()
        if len(tlws) > 0 and isinstance(tlws[0], wx.Frame):
            # call mainframe.Frame.open_lyric_file()
            tlws[0].open_lyric_file(self.m_value)


class CommandUI(wx.EvtHandler):
    """Base class for command user interface components.

    This class provides the foundation for all command UI classes, handling
    activation, deactivation, and data transfer between UI and command objects.
    """

    proc_class: type[Command]  # type: ignore[misc]

    def __init__(self, uimgr: UIManager | None, name: str, proc: Command | None = None) -> None:
        """Initialize a CommandUI instance.

        :param uimgr: The UI manager instance, or None
        :param name: The name of this command UI
        :param proc: The command object to associate with this UI, or None
        """
        super().__init__()

        self.uimgr: UIManager | None = uimgr
        self.name: str = name
        self.command: Command | None = proc
        self.ui: wx.Window | None = None

    def get_flattened_dict(self) -> dict[str, Any]:
        """Get a dictionary representation of this command UI.

        :returns: Dictionary containing type, name, and data
        """
        return {
            "type": self.__class__.proc_class.__name__,
            "name": self.name,
            "data": self.command,
        }

    def activate(self, parent: wx.Window) -> bool:
        """Activate and show the UI for this command.

        :param parent: The parent window to attach the UI to
        :returns: True if activation was successful
        """
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
        self.post_activate_ui()
        return True

    def post_activate_ui(self) -> None:
        """Called after the UI is activated.

        Override in derived classes to perform post-activation setup.
        """
        pass

    def deactivate(self) -> bool:
        """Deactivate and hide the UI for this command.

        :returns: True if deactivation was successful
        """
        if not self.TransferFromWindow():
            return False

        self.pre_deactivate_ui()
        self.ui.Hide()
        return True

    def pre_deactivate_ui(self) -> None:
        """Called before the UI is deactivated.

        Override in derived classes to perform pre-deactivation cleanup.
        """
        pass

    def construct_ui(self, parent: wx.Window) -> None:
        """Construct or retrieve the UI component for this command.

        :param parent: The parent window to attach the UI to
        """
        if self.uimgr is not None:
            self.ui = self.uimgr.get_ui_mapping(self.__class__.__name__)
        if self.ui is None:
            self.ui = self.construct_toplevel_ui(parent)
            sizer = parent.GetSizer()
            # insert ui at the beginning
            sizer.Prepend(self.ui, 1, wx.ALIGN_TOP | wx.EXPAND)
            if self.uimgr is not None:
                self.uimgr.set_ui_mapping(self.__class__.__name__, self.ui)

    def construct_toplevel_ui(self, _parent: wx.Window) -> wx.Window:
        """Construct the top-level UI component.

        Override this function in derived class to provide the own UI.

        :param _parent: The parent window
        :returns: The constructed UI window
        """
        result: wx.Window | None = None
        return result  # type: ignore[return-value]

    def need_stretch_spacer(self) -> bool:
        """Check if a stretch spacer is needed in the layout.

        :returns: True if a stretch spacer should be added
        """
        return False

    def TransferFromWindow(self) -> bool:
        """Transfer data from the UI to the command object.

        :returns: True if transfer was successful
        """
        return True

    def TransferToWindow(self) -> bool:
        """Transfer data from the command object to the UI.

        :returns: True if transfer was successful
        """
        return True

    def set_modified(self, a: Any, b: Any) -> Any:
        """Set modified flag if values differ and return the new value.

        :param a: The old value
        :param b: The new value
        :returns: The new value (b)
        """
        if a != b and self.uimgr is not None:
            self.uimgr.set_modified()
        return b

    @staticmethod
    def create_openfile_property(prop_name: str, wildcard: str) -> wxpg.FileProperty:
        """Create a file property for opening files.

        :param prop_name: The property name
        :param wildcard: The file wildcard pattern
        :returns: A configured FileProperty instance
        """
        file_prop = wxpg.FileProperty(prop_name)
        # PG_DIALOG_TITLE
        file_prop.SetAttribute(wxpg.PG_FILE_DIALOG_STYLE, wx.FD_OPEN | wx.FD_FILE_MUST_EXIST)
        file_prop.SetAttribute(wxpg.PG_FILE_WILDCARD, wildcard)

        return file_prop

    @staticmethod
    def create_savefile_property(prop_name: str, wildcard: str) -> wxpg.FileProperty:
        """Create a file property for saving files.

        :param prop_name: The property name
        :param wildcard: The file wildcard pattern
        :returns: A configured FileProperty instance
        """
        file_prop = wxpg.FileProperty(prop_name)
        # PG_DIALOG_TITLE
        file_prop.SetAttribute(wxpg.PG_FILE_DIALOG_STYLE, wx.FD_SAVE)
        file_prop.SetAttribute(wxpg.PG_FILE_WILDCARD, wildcard)

        return file_prop

    @staticmethod
    def create_datetime_property(prop_name: str) -> wxpg.DateProperty:
        """Create a date/time property.

        :param prop_name: The property name
        :returns: A configured DateProperty instance
        """
        date_prop = wxpg.DateProperty(prop_name)
        # date_prop.SetAttribute(wxpg.PG_DATE_PICKER_STYLE, wx.DP_DEFAULT | DP_SHOWCENTURY)
        # date_prop.SetAttribute(wxpg.PG_DATE_FORMAT, wildcard)

        return date_prop


class PopupMessage(Command):
    """Command to display a popup message dialog to the user."""

    def __init__(self, message: str) -> None:
        """Initialize a PopupMessage command.

        :param message: The message text to display
        """
        super().__init__()

        self.message: str = message

    def execute(self, cm: CommandManager, prs: Any) -> None:
        """Execute the popup message command.

        :param cm: The command manager instance
        :param prs: The presentation object
        """
        cm.progress_message(0, _("Displaying PopupMessage."))

        app = wx.App.Get()
        title = app.GetAppDisplayName()
        dlg = wx.RichMessageDialog(app.GetTopWindow(), self.message, caption=title, style=wx.OK | wx.ICON_INFORMATION | wx.STAY_ON_TOP)
        dlg.ShowCheckBox("Refresh page content.", checked=True)
        dlg.ShowModal()

        cm.progress_message(0, _("Refreshing page content."))
        prs.refresh_page_cache(dlg.IsCheckBoxChecked())


class PopupMessageUI(CommandUI):
    """UI for editing popup message commands."""

    proc_class = PopupMessage

    def __init__(self, uimgr: UIManager | None, name: str, proc: PopupMessage | None = None) -> None:
        """Initialize a PopupMessageUI instance.

        :param uimgr: The UI manager instance, or None
        :param name: The name of this command UI
        :param proc: The PopupMessage command, or None
        """
        super().__init__(uimgr, name, proc=proc)

        if self.command is None:
            self.command = PopupMessage("")

    def construct_toplevel_ui(self, parent: wx.Window) -> wx.TextCtrl:
        """Construct a text control for editing the message.

        :param parent: The parent window
        :returns: A TextCtrl instance
        """
        return wx.TextCtrl(parent)

    def TransferFromWindow(self) -> bool:
        """Transfer the message text from the UI to the command.

        :returns: True if transfer was successful
        """
        self.command.message = self.set_modified(self.command.message, self.ui.GetValue())
        return True

    def TransferToWindow(self) -> bool:
        """Transfer the message text from the command to the UI.

        :returns: True if transfer was successful
        """
        self.ui.SetValue(self.command.message)
        return True


class PropertyGridUI(CommandUI):
    """Base class for command UIs using wxPropertyGrid.

    This class provides functionality for managing property grids with both
    fixed and dynamic properties that can be added or removed at runtime.
    """

    TIMER_ID: int = 100

    def __init__(self, uimgr: UIManager | None, name: str, proc: Command | None = None) -> None:
        """Initialize a PropertyGridUI instance.

        :param uimgr: The UI manager instance, or None
        :param name: The name of this command UI
        :param proc: The command object, or None
        """
        super().__init__(uimgr, name, proc=proc)

        self.dynamic_count: int = 0

        self.timer = wx.Timer(self, self.TIMER_ID)
        self.Bind(wx.EVT_TIMER, self.on_timer)

    def construct_toplevel_ui(self, parent: wx.Window) -> wxpg.PropertyGrid:
        """Construct a PropertyGrid for this UI.

        Difference between using PropertyGridManager vs PropertyGrid is that
        the manager supports multiple pages and a description box.

        :param parent: The parent window
        :returns: A configured PropertyGrid instance
        """
        style = wxpg.PG_SPLITTER_AUTO_CENTER | wxpg.PG_TOOLBAR
        prop_grid = wxpg.PropertyGrid(parent, style=style)

        # Show help as tooltips
        prop_grid.SetExtraStyle(wxpg.PG_EX_HELP_AS_TOOLTIPS)

        self.initialize_fixed_properties(prop_grid)

        self.initialize_dynamic_properties(prop_grid)

        return prop_grid

    def initialize_fixed_properties(self, pg: wxpg.PropertyGrid) -> None:
        """Initialize fixed properties in the property grid.

        Override in derived classes to add fixed properties.

        :param pg: The property grid to add properties to
        """
        pass

    def initialize_dynamic_properties(self, pg: wxpg.PropertyGrid) -> None:
        """Initialize dynamic properties in the property grid.

        Override in derived classes to add initial dynamic properties.

        :param pg: The property grid to add properties to
        """
        pass

    def post_activate_ui(self) -> None:
        """Bind property change events after UI activation."""
        self.ui.Bind(wxpg.EVT_PG_CHANGED, self.on_property_changed, self.ui)

    def pre_deactivate_ui(self) -> None:
        """Unbind property change events before UI deactivation."""
        self.ui.Unbind(wxpg.EVT_PG_CHANGED, handler=self.on_property_changed)

    def get_dynamic_label(self, _index: int) -> str:
        """Get the label for a dynamic property at the given index.

        Override in derived classes to provide custom labels.

        :param _index: The index of the dynamic property
        :returns: The property label
        """
        return ""

    def get_dynamic_property(self, _index: int) -> wxpg.PGProperty | None:
        """Get a dynamic property object for the given index.

        Override in derived classes to provide custom properties.

        :param _index: The index of the dynamic property
        :returns: A property object, or None
        """
        return None

    def get_dynamic_properties_from_window(self) -> list[str]:
        """Get all dynamic property values from the property grid.

        :returns: List of property values, with empty trailing values removed
        """
        plist: list[str] = []
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

    def set_dynamic_properties_to_window(self, plist: list[str]) -> None:
        """Set dynamic property values in the property grid.

        :param plist: List of property values to set
        """
        for i in range(len(plist) + 1):
            name = self.get_dynamic_label(i)
            if self.ui.GetProperty(name) is None:
                prop = self.get_dynamic_property(i)
                if prop is not None:
                    self.ui.Append(prop)

            self.ui.SetPropertyValueString(name, plist[i] if i < len(plist) else "")

        i = len(plist) + 1
        while True:
            name = self.get_dynamic_label(i)
            if self.ui.GetProperty(name) is None:
                break

            self.ui.RemoveProperty(name)
            i += 1

        self.dynamic_count = len(plist) + 1

    def SetPropertyValueString(self, name: str, value: str | None) -> None:
        """Set a property value as a string, handling None values.

        :param name: The property name
        :param value: The value to set, or None to set empty string
        """
        if value is None:
            value = ""
        self.ui.SetPropertyValueString(name, value)

    def on_property_changed(self, event: wxpg.PropertyGridEvent) -> None:
        """Handle property change events.

        When the last dynamic property is filled, schedules creation of a new
        property via timer to avoid recursion issues.

        :param event: The property change event
        """
        name = event.GetPropertyName()
        value = event.GetPropertyValue()
        label = self.get_dynamic_label(self.dynamic_count - 1)
        if value and name == label:
            # if last file and non empty value entered,
            # create new entry.
            # But, creating new entry here, causes recursion somehow,
            # so, do it in timer callback.
            self.timer.StartOnce(100)

    def on_timer(self, _event: wx.TimerEvent) -> None:
        """Handle timer events to create new dynamic properties.

        :param _event: The timer event (unused)
        """
        prop = self.get_dynamic_property(self.dynamic_count)
        if prop is not None:
            self.ui.Append(prop)
        self.dynamic_count = self.dynamic_count + 1


class OpenFileUI(PropertyGridUI):
    """UI for editing open file commands."""

    proc_class = OpenFile

    PRESENTATION_FILE: str = _("Presentation file")
    NOTES_FILE: str = _("Notes file")

    def __init__(self, uimgr: UIManager | None, name: str, proc: OpenFile | None = None) -> None:
        """Initialize an OpenFileUI instance.

        :param uimgr: The UI manager instance, or None
        :param name: The name of this command UI
        :param proc: The OpenFile command, or None
        """
        super().__init__(uimgr, name, proc=proc)

        if self.command is None:
            self.command = OpenFile("", "")
        self.wildcard: str = POWERPOINT_FILES_WILDCARD

    def initialize_fixed_properties(self, pg: wxpg.PropertyGrid) -> None:
        """Initialize fixed properties for file paths.

        :param pg: The property grid to add properties to
        """
        pg.Append(wxpg.PropertyCategory(_("File paths")))

        file_prop = self.create_openfile_property(self.PRESENTATION_FILE, self.wildcard)
        pg.Append(file_prop)

        file_prop = self.create_openfile_property(self.NOTES_FILE, _("Text files (*.txt)|*.txt"))
        pg.Append(file_prop)

    def TransferFromWindow(self) -> bool:
        """Transfer file paths from the UI to the command.

        :returns: True if transfer was successful
        """
        self.command.filename = self.set_modified(
            self.command.filename,
            self.ui.GetPropertyValueAsString(self.PRESENTATION_FILE),
        )
        self.command.notes_filename = self.set_modified(
            self.command.notes_filename,
            self.ui.GetPropertyValueAsString(self.NOTES_FILE),
        )
        return True

    def TransferToWindow(self) -> bool:
        """Transfer file paths from the command to the UI.

        :returns: True if transfer was successful
        """
        self.ui.SetPropertyValueString(self.PRESENTATION_FILE, self.command.filename)
        self.ui.SetPropertyValueString(self.NOTES_FILE, self.command.notes_filename)
        return True


class SaveFilesUI(PropertyGridUI):
    """UI for editing save files commands."""

    proc_class = SaveFiles

    PRESENTATION_FILE: str = _("Presentation file")
    LYRICS_ARCHIVE_FILE: str = _("Lyrics archive file")
    NOTES_FILE: str = _("Notes file")
    VERSES_FILE: str = _("Verses file")

    def __init__(self, uimgr: UIManager | None, name: str, proc: SaveFiles | None = None) -> None:
        """Initialize a SaveFilesUI instance.

        :param uimgr: The UI manager instance, or None
        :param name: The name of this command UI
        :param proc: The SaveFiles command, or None
        """
        super().__init__(uimgr, name, proc=proc)

        if self.command is None:
            self.command = SaveFiles("", "", "", "")
        self.wildcard: str = POWERPOINT_FILES_WILDCARD

    def initialize_fixed_properties(self, pg: wxpg.PropertyGrid) -> None:
        """Initialize fixed properties for file paths.

        :param pg: The property grid to add properties to
        """
        pg.Append(wxpg.PropertyCategory(_("File paths")))

        file_prop = self.create_savefile_property(self.PRESENTATION_FILE, self.wildcard)
        pg.Append(file_prop)

        file_prop = self.create_savefile_property(self.LYRICS_ARCHIVE_FILE, _("OpenLP service files (*.osz)|*.osz|Zip files (*.zip)|*.zip"))
        pg.Append(file_prop)

        file_prop = self.create_savefile_property(self.NOTES_FILE, _("Text files (*.txt)|*.txt"))
        pg.Append(file_prop)

        file_prop = self.create_savefile_property(self.VERSES_FILE, _("Text files (*.txt)|*.txt"))
        pg.Append(file_prop)

    def TransferFromWindow(self) -> bool:
        """Transfer file paths from the UI to the command.

        :returns: True if transfer was successful
        """
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

    def TransferToWindow(self) -> bool:
        """Transfer file paths from the command to the UI.

        :returns: True if transfer was successful
        """
        self.ui.SetPropertyValueString(self.PRESENTATION_FILE, self.command.filename)
        self.ui.SetPropertyValueString(self.LYRICS_ARCHIVE_FILE, self.command.lyrics_archive_filename)
        self.ui.SetPropertyValueString(self.NOTES_FILE, self.command.notes_filename)
        self.ui.SetPropertyValueString(self.VERSES_FILE, self.command.verses_filename)
        return True


class SetVariablesUI(PropertyGridUI):
    """UI for editing set variables commands."""

    proc_class = SetVariables

    DATETIME_NAME: str = _("Datetime variable name")
    DATETIME_VALUE: str = _("Datetime value")
    VARNAME_D: str = _("Variable name %d")
    VARVALUE_D: str = _("Variable value %d")

    def __init__(self, uimgr: UIManager | None, name: str, proc: SetVariables | None = None) -> None:
        """Initialize a SetVariablesUI instance.

        :param uimgr: The UI manager instance, or None
        :param name: The name of this command UI
        :param proc: The SetVariables command, or None
        """
        super().__init__(uimgr, name, proc=proc)

        if self.command is None:
            self.command = SetVariables()

    def get_dynamic_label(self, index: int) -> str:
        """Get the label for a dynamic property at the given index.

        :param index: The index of the dynamic property
        :returns: The property label
        """
        if (index % 2) == 0:
            return self.VARNAME_D % (index / 2 + 1)
        return self.VARVALUE_D % (index / 2 + 1)

    def get_dynamic_property(self, index: int) -> wxpg.StringProperty:
        """Get a dynamic property object for the given index.

        :param index: The index of the dynamic property
        :returns: A StringProperty instance
        """
        return wxpg.StringProperty(self.get_dynamic_label(index))

    def initialize_fixed_properties(self, pg: wxpg.PropertyGrid) -> None:
        """Initialize fixed properties for datetime and variables.

        :param pg: The property grid to add properties to
        """
        pg.Append(wxpg.PropertyCategory(_("Datetime")))

        pg.Append(wxpg.StringProperty(self.DATETIME_NAME))

        date_prop = self.create_datetime_property(self.DATETIME_VALUE)
        pg.Append(date_prop)

        pg.Append(wxpg.PropertyCategory(_("Variable names and values to be used for find and replace")))

    def initialize_dynamic_properties(self, pg: wxpg.PropertyGrid) -> None:
        """Initialize initial dynamic properties for variable pairs.

        :param pg: The property grid to add properties to
        """
        pg.Append(self.get_dynamic_property(0))
        pg.Append(self.get_dynamic_property(1))
        self.dynamic_count = 2

    def get_dynamic_properties_from_window(self) -> list[str]:
        """Get all dynamic property values from the property grid.

        :returns: List of property values, with empty trailing pairs removed
        """
        plist: list[str] = []
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

    def set_dynamic_properties_to_window(self, plist: list[str]) -> None:
        """Set dynamic property values in the property grid.

        :param plist: List of property values to set (pairs of name/value)
        """
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

    def TransferFromWindow(self) -> bool:
        """Transfer variable values from the UI to the command.

        :returns: True if transfer was successful
        """
        format_dict: dict[str, DateTimeFormat] = {}
        name = self.ui.GetPropertyValueAsString(self.DATETIME_NAME)
        if name:
            dt_str = ""
            try:
                dt_value = self.ui.GetPropertyValueAsDateTime(self.DATETIME_VALUE)
                dt_str = dt_value.FormatISODate()
            except (AttributeError, ValueError, TypeError):
                # Property may not be a datetime type, use default empty string
                pass

            # format comes from find_string
            format_dict[name] = DateTimeFormat(value=dt_str)

        self.command.format_dict = self.set_modified(self.command.format_dict, format_dict)

        plist = self.get_dynamic_properties_from_window()
        texts: dict[str, str] = {}
        for i in range(0, len(plist), 2):
            f = plist[i]
            r = plist[i + 1]
            if f:
                texts[f] = r

        self.command.str_dict = self.set_modified(self.command.str_dict, texts)

        return True

    def TransferToWindow(self) -> bool:
        """Transfer variable values from the command to the UI.

        :returns: True if transfer was successful
        """
        if len(self.command.format_dict) == 1:
            dt_name = next(iter(self.command.format_dict))
            self.SetPropertyValueString(self.DATETIME_NAME, dt_name)

            fobj = self.command.format_dict[dt_name]
            dt_value = DateTimeFormat.datetime_from_c_locale(fobj.value)
            self.ui.SetPropertyValue(self.DATETIME_VALUE, dt_value)

        plist: list[str] = []
        for k, v in self.command.str_dict.items():
            plist.append(k)
            plist.append(v)

        self.set_dynamic_properties_to_window(plist)

        return True

    def on_property_changed(self, event: wxpg.PropertyGridEvent) -> None:
        """Handle property change events to add new variable pairs.

        :param event: The property change event
        """
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
    """UI for editing insert slides commands."""

    proc_class = InsertSlides

    INSERT_LOCATION: str = _("Insert location")
    SEPARATOR_SLIDES: str = _("Separator slides")
    FILE_D: str = _("File %d")

    def __init__(self, uimgr: UIManager | None, name: str, proc: InsertSlides | None = None) -> None:
        """Initialize an InsertSlidesUI instance.

        :param uimgr: The UI manager instance, or None
        :param name: The name of this command UI
        :param proc: The InsertSlides command, or None
        """
        super().__init__(uimgr, name, proc=proc)

        if self.command is None:
            self.command = InsertSlides(None, None, [])

    def get_dynamic_label(self, index: int) -> str:
        """Get the label for a dynamic property at the given index.

        :param index: The index of the dynamic property
        :returns: The property label
        """
        return self.FILE_D % (index + 1)

    def get_dynamic_property(self, index: int) -> wxpg.FileProperty:
        """Get a dynamic property object for the given index.

        :param index: The index of the dynamic property
        :returns: A FileProperty instance
        """
        return self.create_openfile_property(self.get_dynamic_label(index), POWERPOINT_FILES_WILDCARD)

    def initialize_fixed_properties(self, pg: wxpg.PropertyGrid) -> None:
        """Initialize fixed properties for insert location and separator.

        :param pg: The property grid to add properties to
        """
        pg.Append(wxpg.PropertyCategory(_("1 - Basic Settings")))
        pg.Append(wxpg.StringProperty(self.INSERT_LOCATION))
        pg.Append(wxpg.StringProperty(self.SEPARATOR_SLIDES))

        pg.Append(wxpg.PropertyCategory(_("2 - List of files to insert")))

    def initialize_dynamic_properties(self, pg: wxpg.PropertyGrid) -> None:
        """Initialize initial dynamic properties for file list.

        :param pg: The property grid to add properties to
        """
        pg.Append(self.get_dynamic_property(0))
        self.dynamic_count = 1

    def TransferFromWindow(self) -> bool:
        """Transfer settings from the UI to the command.

        :returns: True if transfer was successful
        """
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

    def TransferToWindow(self) -> bool:
        """Transfer settings from the command to the UI.

        :returns: True if transfer was successful
        """
        self.SetPropertyValueString(self.INSERT_LOCATION, self.command.insert_location)
        self.SetPropertyValueString(self.SEPARATOR_SLIDES, self.command.separator_slides)

        self.set_dynamic_properties_to_window(self.command.filelist)

        return True


class InsertLyricsUI(PropertyGridUI):
    """UI for editing insert lyrics commands."""

    proc_class = InsertLyrics

    FILE_TYPE: str = _("Lyric file types")
    FILE_TYPE_LIST: ClassVar[list[str]] = [
        _("Lyric slide files"),
        _("Lyric text files"),
        _("Lyric slide and text files"),
    ]

    SLIDE_REPEAT_RANGE: str = _("Score repeat range")
    SLIDE_SEPARATOR_SLIDES: str = _("Score separator slides")

    LYRIC_REPEAT_RANGE: str = _("Lyric repeat range")
    LYRIC_SEPARATOR_SLIDES: str = _("Lyric separator slides")
    LYRIC_PATTERN: str = _("Lyric pattern")
    ARCHIVE_LYRIC_FILES: str = _("Archive lyric files")

    FILE_D: str = _("File %d")
    LYRIC_FILES_WILDCARD: str = _(
        "Powerpoint/Lyric files (*.ppt;*.pptx;*.xml)|*.ppt;*.pptx;*.xml|Powerpoint files (*.ppt;*.pptx)|*.ppt;*.pptx|Lyric xml files (*.xml)|*.xml"
    )

    def __init__(self, uimgr: UIManager | None, name: str, proc: InsertLyrics | None = None) -> None:
        """Initialize an InsertLyricsUI instance.

        :param uimgr: The UI manager instance, or None
        :param name: The name of this command UI
        :param proc: The InsertLyrics command, or None
        """
        super().__init__(uimgr, name, proc=proc)

        if self.command is None:
            self.command = InsertLyrics("", None, "", None, "", False, [], 2)

    @staticmethod
    def create_openfile_property(prop_name: str, wildcard: str) -> MyFileProperty:
        """Create a file property for opening lyric files.

        :param prop_name: The property name
        :param wildcard: The file wildcard pattern
        :returns: A configured MyFileProperty instance
        """
        file_prop = MyFileProperty(prop_name)
        # PG_DIALOG_TITLE
        file_prop.SetAttribute(wxpg.PG_FILE_DIALOG_STYLE, wx.FD_OPEN | wx.FD_FILE_MUST_EXIST)
        file_prop.SetAttribute(wxpg.PG_FILE_WILDCARD, wildcard)

        return file_prop

    def get_dynamic_label(self, index: int) -> str:
        """Get the label for a dynamic property at the given index.

        :param index: The index of the dynamic property
        :returns: The property label
        """
        return self.FILE_D % (index + 1)

    def get_dynamic_property(self, index: int) -> MyFileProperty:
        """Get a dynamic property object for the given index.

        :param index: The index of the dynamic property
        :returns: A MyFileProperty instance
        """
        return self.create_openfile_property(self.get_dynamic_label(index), self.LYRIC_FILES_WILDCARD)

    def initialize_fixed_properties(self, pg: wxpg.PropertyGrid) -> None:
        """Initialize fixed properties for lyric file settings.

        :param pg: The property grid to add properties to
        """
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

    def initialize_dynamic_properties(self, pg: wxpg.PropertyGrid) -> None:
        """Initialize initial dynamic properties for file list.

        :param pg: The property grid to add properties to
        """
        pg.Append(self.get_dynamic_property(0))
        self.dynamic_count = 1

    def TransferFromWindow(self) -> bool:
        """Transfer settings from the UI to the command.

        :returns: True if transfer was successful
        """
        MyFileProperty.ignore_open_file = True

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

        MyFileProperty.ignore_open_file = False

        return True

    def TransferToWindow(self) -> bool:
        """Transfer settings from the command to the UI.

        :returns: True if transfer was successful
        """
        MyFileProperty.ignore_open_file = True

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

        MyFileProperty.ignore_open_file = False

        return True


class DuplicateWithTextUI(PropertyGridUI):
    proc_class = DuplicateWithText

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
            self.command = DuplicateWithText("", "", "", [], "", False, 0)

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
        lines = [line.strip() for line in lines if line.strip()]
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
    proc_class = GenerateBibleVerse
    current_bible_format = bibleformat.BibleFormat.MYBIBLE.value

    BIBLE_VERSION1 = _("Bible Version 1")
    MAIN_VERSE_NAME1 = _("Main Bible verse name 1")
    EACH_VERSE_NAME1 = _("Each Bible verse name 1")
    BIBLE_VERSION2 = _("Bible Version 2")
    MAIN_VERSE_NAME2 = _("Main Bible verse name 2")
    EACH_VERSE_NAME2 = _("Each Bible verse name 2")
    MAIN_VERSES = _("Main Bible verses")
    ADDITONAL_VERSES = _("Additional Bible verses")
    REPEAT_RANGE = _("Repeating slides range")

    def __init__(self, uimgr, name, proc=None):
        super().__init__(uimgr, name, proc=proc)

        if self.command is None:
            self.command = GenerateBibleVerse(GenerateBibleVerseUI.current_bible_format)

    def initialize_fixed_properties(self, pg):
        pg.Append(wxpg.PropertyCategory(_("1 - Bible #1")))
        versions = bibleformat.enum_versions(GenerateBibleVerseUI.current_bible_format)
        pg.Append(wxpg.EnumProperty(self.BIBLE_VERSION1, labels=versions, value=0))
        pg.Append(wxpg.StringProperty(self.MAIN_VERSE_NAME1))
        pg.Append(wxpg.StringProperty(self.EACH_VERSE_NAME1))
        pg.Append(wxpg.PropertyCategory(_("2 - Bible #2")))
        pg.Append(wxpg.EnumProperty(self.BIBLE_VERSION2, labels=versions, value=-1))
        pg.Append(wxpg.StringProperty(self.MAIN_VERSE_NAME2))
        pg.Append(wxpg.StringProperty(self.EACH_VERSE_NAME2))
        pg.Append(wxpg.PropertyCategory(_("3 - Other Bible specification")))
        pg.Append(wxpg.StringProperty(self.MAIN_VERSES))
        pg.Append(wxpg.StringProperty(self.ADDITONAL_VERSES))
        pg.Append(wxpg.StringProperty(self.REPEAT_RANGE))

    def TransferFromWindow(self):
        self.command.bible_version1 = self.set_modified(
            self.command.bible_version1,
            self.ui.GetPropertyValueAsString(self.BIBLE_VERSION1),
        )
        self.command.main_verse_name1 = self.set_modified(
            self.command.main_verse_name1,
            self.ui.GetPropertyValueAsString(self.MAIN_VERSE_NAME1),
        )
        self.command.each_verse_name1 = self.set_modified(
            self.command.each_verse_name1,
            self.ui.GetPropertyValueAsString(self.EACH_VERSE_NAME1),
        )

        self.command.bible_version2 = self.set_modified(
            self.command.bible_version2,
            self.ui.GetPropertyValueAsString(self.BIBLE_VERSION2),
        )
        self.command.main_verse_name2 = self.set_modified(
            self.command.main_verse_name2,
            self.ui.GetPropertyValueAsString(self.MAIN_VERSE_NAME2),
        )
        self.command.each_verse_name2 = self.set_modified(
            self.command.each_verse_name2,
            self.ui.GetPropertyValueAsString(self.EACH_VERSE_NAME2),
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
        self.SetPropertyValueString(self.BIBLE_VERSION1, self.command.bible_version1)
        self.SetPropertyValueString(self.MAIN_VERSE_NAME1, self.command.main_verse_name1)
        self.SetPropertyValueString(self.EACH_VERSE_NAME1, self.command.each_verse_name1)

        self.SetPropertyValueString(self.BIBLE_VERSION2, self.command.bible_version2)
        self.SetPropertyValueString(self.MAIN_VERSE_NAME2, self.command.main_verse_name2)
        self.SetPropertyValueString(self.EACH_VERSE_NAME2, self.command.each_verse_name2)

        self.SetPropertyValueString(self.MAIN_VERSES, self.command.main_verses)
        self.SetPropertyValueString(self.ADDITONAL_VERSES, self.command.additional_verses)
        self.SetPropertyValueString(self.REPEAT_RANGE, self.command.repeat_range)

        return True


class ExportSlidesUI(PropertyGridUI):
    proc_class = ExportSlides

    SLIDE_RANGE = _("Slides to export")
    IMAGE_TYPE = _("Image type")
    OUTPUT_DIR = _("Output directory")
    CLEANUP_OUTPUT_DIR = _("Clean up output directory")
    TRANSPARENT_IMAGE = _("Generate transparent image")
    TRANSPARENT_COLOR = _("Color to make transparent")

    def __init__(self, uimgr, name, proc=None):
        super().__init__(uimgr, name, proc=proc)

        if self.command is None:
            self.command = ExportSlides("", "", "PNG", 0, "#FFFFFF")

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
            flags = flags | ExportFlag.CLEANUP_FILES
        if transparent_image:
            flags = flags | ExportFlag.TRANSPARENT
        self.command.flags = self.set_modified(self.command.flags, flags)
        color = self.ui.GetPropertyValue(self.TRANSPARENT_COLOR)
        str_color = color.GetAsString(wx.C2S_HTML_SYNTAX)
        self.command.color = self.set_modified(self.command.color, str_color)

        return True

    def TransferToWindow(self):
        self.SetPropertyValueString(self.SLIDE_RANGE, self.command.slide_range)
        self.SetPropertyValueString(self.IMAGE_TYPE, self.command.image_type)
        self.SetPropertyValueString(self.OUTPUT_DIR, self.command.out_dirname)

        value = (self.command.flags | ExportFlag.CLEANUP_FILES) != 0
        self.ui.SetPropertyValue(self.CLEANUP_OUTPUT_DIR, value)
        value = (self.command.flags | ExportFlag.TRANSPARENT) != 0
        self.ui.SetPropertyValue(self.TRANSPARENT_IMAGE, value)
        self.SetPropertyValueString(self.TRANSPARENT_COLOR, self.command.color)

        return True


class ExportShapesUI(PropertyGridUI):
    proc_class = ExportShapes

    SLIDE_RANGE = _("Slides to export")
    IMAGE_TYPE = _("Image type")
    OUTPUT_DIR = _("Output directory")
    CLEANUP_OUTPUT_DIR = _("Clean up output directory")

    def __init__(self, uimgr, name, proc=None):
        super().__init__(uimgr, name, proc=proc)

        if self.command is None:
            self.command = ExportShapes("", "", "PNG", 0)

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
            flags = flags | ExportFlag.CLEANUP_FILES
        self.command.flags = self.set_modified(self.command.flags, flags)

        return True

    def TransferToWindow(self):
        self.SetPropertyValueString(self.SLIDE_RANGE, self.command.slide_range)
        self.SetPropertyValueString(self.IMAGE_TYPE, self.command.image_type)
        self.SetPropertyValueString(self.OUTPUT_DIR, self.command.out_dirname)

        value = (self.command.flags | ExportFlag.CLEANUP_FILES) != 0
        self.ui.SetPropertyValue(self.CLEANUP_OUTPUT_DIR, value)

        return True


class SymboledDirectory(Command):
    """Command for managing directory symbols."""

    def __init__(self, dir_dict: dict[str, str]) -> None:
        """Initialize a SymboledDirectory command.

        :param dir_dict: Dictionary mapping symbol names to directory paths
        """
        super().__init__()

        self.dir_dict: dict[str, str] = dir_dict

    def get_flattened_dict(self) -> dict[str, str]:
        """Get a flattened dictionary representation.

        :returns: The directory dictionary
        """
        return self.dir_dict

    def load_directory_symbols(self) -> dict[str, str]:
        """Load directory symbols from the command's attributes.

        :returns: Dictionary of directory symbols excluding internal attributes
        """
        d = copy.deepcopy(self.__dict__)

        keys_to_remove = {"dir_dict", "enabled"}
        for k in keys_to_remove:
            del d[k]

        return d  # type: ignore[return-value]

    def execute(self, cm: CommandManager, prs: Any) -> None:
        """Execute the command (no-op for directory symbols).

        :param cm: The command manager instance
        :param prs: The presentation object
        """
        pass


class SymboledDirectoryUI(CommandUI):
    """UI for editing directory symbol commands."""

    proc_class = SymboledDirectory

    def __init__(self, uimgr: UIManager | None, name: str, proc: SymboledDirectory | None = None) -> None:
        """Initialize a SymboledDirectoryUI instance.

        :param uimgr: The UI manager instance, or None
        :param name: The name of this command UI
        :param proc: The SymboledDirectory command, or None
        """
        super().__init__(uimgr, name, proc=proc)

        self.command = SymboledDirectory([])


class CommandEncoder(json.JSONEncoder):
    """JSON encoder for command UI objects."""

    proc_ui_list: ClassVar[list[type[CommandUI]]] = [
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
        SymboledDirectoryUI,
    ]
    proc_map: ClassVar[dict[str, type[CommandUI]]] = {ui.proc_class.__name__: ui for ui in proc_ui_list}

    format_list: ClassVar[list[type[Any]]] = [BibleVerseFormat, DateTimeFormat]

    format_map: ClassVar[dict[str, type[Any]]] = {fo.__name__: fo for fo in format_list}

    def default(self, o: Any) -> Any:
        """Encode an object to JSON.

        :param o: The object to encode
        :returns: The encoded object
        """
        if (func := getattr(o, "get_flattened_dict", None)) and callable(func):
            return o.get_flattened_dict()

        return o.__dict__

    @staticmethod
    def decoder(o: dict[str, Any]) -> Any:
        """Decode a JSON object to a command UI or format object.

        :param o: The dictionary to decode
        :returns: A CommandUI instance, format object, or the original object
        """
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
        if "format_type" in o:
            format_type = o["format_type"]
            del o["format_type"]
            if format_type in CommandEncoder.format_map:
                fobj_cls = CommandEncoder.format_map[format_type]
                fobj = fobj_cls()

                fobj.__dict__.update(o)

                return fobj

        return o


class UIManager:
    """Manager for command UI components.

    This class manages the lifecycle of command UI components, including
    activation, deactivation, loading, saving, and execution.
    """

    def __init__(self) -> None:
        """Initialize a UIManager instance."""
        self.ui_map: dict[str, wx.Window] = {}
        self.active_ui: CommandUI | None = None
        self.command_ui_list: list[CommandUI] = []
        self.modified: bool = False

    def get_ui_mapping(self, name: str) -> wx.Window | None:
        """Get a UI component by name.

        :param name: The name of the UI component
        :returns: The UI component, or None if not found
        """
        if name in self.ui_map:
            return self.ui_map[name]

        return None

    def set_ui_mapping(self, name: str, ui: wx.Window) -> None:
        """Set a UI component mapping.

        :param name: The name of the UI component
        :param ui: The UI component to map
        """
        self.ui_map[name] = ui

    def activate(self, parent: wx.Window, ui: CommandUI | None) -> None:
        """Activate a command UI component.

        :param parent: The parent window
        :param ui: The command UI to activate, or None
        """
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

    def deactivate(self) -> None:
        """Deactivate the currently active UI component."""
        if self.active_ui is not None:
            if not self.active_ui.deactivate():
                return

        self.active_ui = None

    def get_modified(self) -> bool:
        """Get the modified flag.

        :returns: True if the UI has been modified
        """
        return self.modified

    def check_modified(self) -> bool:
        """Check and update the modified flag.

        :returns: True if the UI has been modified
        """
        if self.active_ui is not None:
            self.active_ui.TransferFromWindow()

        return self.modified

    def set_modified(self, modified: bool = True) -> None:
        """Set the modified flag.

        :param modified: The new modified state
        """
        self.modified = modified

    def set_item_name(self, index: int, name: str) -> None:
        """Set the name of a command UI item.

        :param index: The index of the item
        :param name: The new name
        """
        self.command_ui_list[index].name = name
        self.set_modified()

    def insert_item(self, index: int, ui: CommandUI) -> None:
        """Insert a command UI item at the specified index.

        :param index: The index to insert at
        :param ui: The command UI to insert
        """
        self.command_ui_list.insert(index, ui)
        self.set_modified()

    def delete_item(self, index: int) -> None:
        """Delete a command UI item at the specified index.

        :param index: The index of the item to delete
        """
        del self.command_ui_list[index]
        self.set_modified()

    def move_down_item(self, index: int) -> None:
        """Move a command UI item down in the list.

        :param index: The index of the item to move
        """
        ui = self.command_ui_list.pop(index)
        self.command_ui_list.insert(index + 1, ui)
        self.set_modified()

    def move_up_item(self, index: int) -> None:
        """Move a command UI item up in the list.

        :param index: The index of the item to move
        """
        ui = self.command_ui_list.pop(index)
        self.command_ui_list.insert(index - 1, ui)
        self.set_modified()

    def open(self, filename: str, dir_dict: dict[str, str]) -> None:
        """Open and load command UI list from a file.

        :param filename: The file to load from
        :param dir_dict: Dictionary of directory symbols
        """
        command_ui_list: list[CommandUI] = []
        with open(filename, encoding="utf-8") as f:
            command_ui_list = json.load(f, object_hook=CommandEncoder.decoder)  # type: ignore[assignment]

        dir_symbol: SymboledDirectoryUI | None = None
        if len(command_ui_list) > 0 and isinstance(command_ui_list[0], SymboledDirectoryUI):
            dir_symbol = command_ui_list[0]
            command_ui_list = command_ui_list[1:]

            mycmd = dir_symbol.command
            loaded_dss = mycmd.load_directory_symbols()

            loaded_dss.update(dir_dict)
            dir_dict = loaded_dss

        dss = DirSymbols(dir_dict)
        for ui in command_ui_list:
            ui.uimgr = self

            if dir_symbol:
                ui.command.translate_dir_symbols(dss, to_symbol=False)

        self.command_ui_list = command_ui_list
        self.set_modified(False)

    def save(self, filename: str, dir_dict: dict[str, str]) -> None:
        """Save command UI list to a file.

        :param filename: The file to save to
        :param dir_dict: Dictionary of directory symbols
        """
        self.check_modified()

        dss = DirSymbols(dir_dict)
        dir_symbol: SymboledDirectoryUI | None = None
        modified = self.modified
        if len(dir_dict) > 0:
            dir_symbol = SymboledDirectoryUI(self, "directories")
            dir_symbol.command.dir_dict = dir_dict
            self.insert_item(0, dir_symbol)

            # translate absolute path to symbolized path.
            for ui in self.command_ui_list[1:]:
                mycmd = ui.command
                mycmd.translate_dir_symbols(dss, to_symbol=True)

        with AtomicFileWriter(filename, "w", encoding="utf-8") as f:
            json.dump(self.command_ui_list, f, indent=2, cls=CommandEncoder, ensure_ascii=False)

        if len(dir_dict) > 0:
            # revert back the paths.
            for ui in self.command_ui_list[1:]:
                mycmd = ui.command
                mycmd.translate_dir_symbols(dss, to_symbol=False)

            self.delete_item(0)
            self.set_modified(modified)

    def execute_commands(self, monitor: Any, pconfig: Any) -> None:
        """Execute all commands in the command UI list.

        :param monitor: Progress monitor object
        :param pconfig: Presentation configuration object
        """
        self.check_modified()

        proc_list = [x.command for x in self.command_ui_list]
        cm = CommandManager()
        cm.lyric_manager.lyric_search_path = pconfig.lyric_search_path
        cm.execute_commands(proc_list, monitor=monitor)
