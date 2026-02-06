"""CommandManager class for service_ppt.

This module contains the CommandManager class that orchestrates command execution
and manages state during PowerPoint slide generation operations.
"""

import re
import shutil
import sys
import tempfile
import traceback
from typing import TYPE_CHECKING, Any

from service_ppt.utils.file_utils import rename_filename_to_zeropadded
from service_ppt.utils.i18n import _

if TYPE_CHECKING:
    from service_ppt.cmd.cmd import Command, FormatObj

# Import at runtime to avoid circular import
from service_ppt.cmd.cmd import Command, FormatObj, replace_all_notes_text
from service_ppt.cmd.lyricmgr import LyricManager

# Use platform-specific implementations only when needed (Windows for COM interface)
if sys.platform.startswith("win32"):
    # On Windows, use win32 COM interface for full PowerPoint integration
    import service_ppt.ppt_slide.powerpoint_win32 as PowerPoint
else:
    # On macOS and Linux, use python-pptx (cross-platform, no PowerPoint required)
    try:
        import service_ppt.ppt_slide.powerpoint_pptx as PowerPoint
    except ImportError:
        # Fallback to OSX AppleScript implementation if python-pptx is not available
        import service_ppt.ppt_slide.powerpoint_osx as PowerPoint


class CommandManager:
    """Manages command execution and state for PowerPoint slide generation.

    CommandManager orchestrates the execution of Command objects, manages
    variable substitution, handles lyric files, and coordinates with the
    PowerPoint application.
    """

    def __init__(self) -> None:
        self.running_already: bool = False
        self.powerpoint: Any = None
        self.prs: Any = None

        self.string_variables: dict[str, str] = {}
        self.format_variables: dict[str, FormatObj] = {}
        self.var_dict: dict[str, str] = {}
        self.varname_re: re.Pattern[str] | None = None

        self.notes: str = ""

        self.monitor: Any = None
        self._continue: bool = True

        self.image_dir: dict[str, str] = {}
        self.num_digits: int = 0

        self.bible_verse: Any = None

        self.lyric_manager: LyricManager = LyricManager(self)

    def __del__(self) -> None:
        self.close()

    def close(self) -> None:
        """Close the presentation and PowerPoint application if needed."""
        if self.prs:
            self.prs.close()
            self.prs = None

        if not self.running_already and self.powerpoint:
            self.powerpoint.quit()

        self.powerpoint = None

        self.remove_cache()

    def reset_exec_vars(self) -> None:
        """Reset execution variables to initial state."""
        self.running_already = False
        self.powerpoint = None
        self.prs = None

        self.string_variables = {}
        self.format_variables = {}
        self.var_dict = {}
        self.varname_re = None

        self.notes = ""

        self.monitor = None
        self._continue = True

        self.lyric_manager.reset_exec_vars()

    def set_presentation(self, prs: Any) -> None:
        """Set the current presentation object.

        :param prs: Presentation object to use
        """
        self.prs = prs

    def add_global_variables(self, str_dict: dict[str, str], format_dict: dict[str, FormatObj] | None = None) -> None:
        """Add global variables for string and format substitution.

        :param str_dict: Dictionary of string variable names to values
        :param format_dict: Dictionary of format variable names to FormatObj instances
        """
        self.string_variables.update(str_dict)
        if format_dict:
            self.format_variables.update(format_dict)

        self.var_dict = {}
        for var in self.string_variables:
            new_var = var
            if not new_var:
                continue
            if new_var[0] != "{":
                new_var = "{" + new_var
            if new_var[-1] != "}":
                new_var = new_var + "}"

            self.var_dict[new_var] = self.string_variables[var]

        varname_pattern = ""
        self.varname_re = None
        if len(self.format_variables):
            varnames = self.format_variables.keys()
            varname_pattern = FormatObj.build_format_pattern(varnames)
            if varname_pattern:
                self.varname_re = re.compile(varname_pattern)

    def process_format_var(self, var: list[str] | str) -> None:
        """Process format variables in a string or list of strings.

        :param var: String or list of strings to process
        """
        if isinstance(var, list):
            for elem in var:
                self.process_format_var(elem)
        elif isinstance(var, str):
            if var in self.var_dict:
                return

            if self.varname_re:
                for m in self.varname_re.finditer(var):
                    varname = m.group(1)
                    if varname not in self.format_variables:
                        continue

                    format_obj = self.format_variables[varname]

                    format = m.group(2)
                    if len(format):
                        format = format[1:]  # remove FormatObj.VAR_FORMAT_SEP
                    key = "{" + varname + m.group(2) + "}"
                    value = format_obj.build_replace_string(format)
                    self.var_dict[key] = value

    def replace_format_vars(self, custom_str: str) -> str:
        """Replace format variables in a custom string.

        :param custom_str: String containing format variables
        :returns: String with variables replaced
        """
        if self.varname_re:
            self.process_format_var(custom_str)
        _, custom_str = replace_all_notes_text(custom_str, self.var_dict)
        return custom_str

    def process_variable_substitution(self) -> None:
        """Process variable substitution in all slides and notes."""
        if self.varname_re:
            self.process_format_var(self.prs.get_text_in_all_slides(False))
        self.prs.replace_all_slides_texts(self.var_dict)

        if self.notes:
            if self.varname_re:
                self.process_format_var(self.notes)
            _, self.notes = replace_all_notes_text(self.notes, self.var_dict)

    def get_notes(self) -> str:
        """Get the current notes text.

        :returns: Notes text string
        """
        return self.notes

    def set_notes(self, notes: str) -> None:
        """Set the notes text.

        :param notes: Notes text to set
        """
        self.notes = notes

    def set_bible_verse(self, bible_verse: Any) -> None:
        """Set the bible verse object.

        :param bible_verse: Bible verse object to set
        """
        self.bible_verse = bible_verse

    def read_songs(self, filelist: list[str]) -> list[Any]:
        """Read songs from a list of files.

        :param filelist: List of file paths to read
        :returns: List of song objects
        """
        return self.lyric_manager.read_songs(filelist)

    def search_lyric_files(self, filelist: list[str]) -> list[str | None]:
        """Search for lyric files corresponding to a list of filenames.

        :param filelist: List of filenames to search for
        :returns: List of lyric file paths (or None if not found)
        """
        return [self.lyric_manager.search_lyric_file(file) for file in filelist]

    def add_lyric_file(self, lyric_file: list[str] | dict[str, Any] | str) -> None:
        """Add lyric file(s) to the lyric manager.

        :param lyric_file: Lyric file path(s) or dictionary to add
        """
        self.lyric_manager.add_lyric_file(lyric_file)

    def progress_message(self, progress: int, message: str) -> None:
        """Send a progress message to the monitor.

        :param progress: Progress percentage (0-100)
        :param message: Progress message text
        """
        self._continue = self.monitor.progress_message(progress, message)

    def error_message(self, message: str) -> None:
        """Send an error message to the monitor.

        :param message: Error message text
        """
        self.monitor.error_message(message)

    # slide image file caching functions.
    def remove_cache(self) -> None:
        """Remove cached image files."""
        for dirname in self.image_dir.values():
            shutil.rmtree(dirname)
        self.image_dir = {}

    def normalize_image_type(self, image_type: str) -> str:
        """Normalize image type string to standard format.

        :param image_type: Image type string (e.g., "jpeg", "tiff")
        :returns: Normalized image type (e.g., "jpg", "tif")
        """
        image_type = image_type.lower()
        if image_type == "jpeg":
            image_type = "jpg"
        elif image_type == "tiff":
            image_type = "tif"
        return image_type

    def generate_image_files(self, image_type: str, color: str | None = None) -> str:
        """Generate image files from the presentation.

        :param image_type: Type of image to generate (e.g., "png", "jpg")
        :param color: Optional color for transparency conversion
        :returns: Directory path containing generated images
        """
        if self.prs:
            slide_count = self.prs.slide_count()
            self.num_digits = len(f"{slide_count + 1}")
        else:
            self.num_digits = 0

        image_type = self.normalize_image_type(image_type)

        if image_type not in self.image_dir:
            outdir = tempfile.mkdtemp(prefix="slides")

            self.prs.saveas_format(outdir, image_type)

            rename_filename_to_zeropadded(outdir, self.num_digits)

            self.image_dir[image_type] = outdir

        return self.image_dir[image_type]

    def get_filename_from_slideno(self, image_type: str, slideno: int) -> str:
        """Get filename for a slide number.

        :param image_type: Image type extension
        :param slideno: Slide number (0-based)
        :returns: Filename string
        """
        image_type = self.normalize_image_type(image_type)

        num = slideno + 1
        fmt = r"%0" + str(self.num_digits) + r"d"
        return "Slide" + (fmt % num) + "." + image_type

    def execute_commands(self, instructions: list[Command], monitor: Any) -> None:
        """Execute a list of commands.

        This is the main function calling each command's execute() method
        to achieve the goal specified in each command.

        :param instructions: List of Command objects to execute
        :param monitor: Monitor object for progress and error reporting
        """
        self.reset_exec_vars()

        self.monitor = monitor
        self._continue = True
        self.progress_message(0, _("Start processing commands."))

        self.running_already = PowerPoint.App.is_running()
        self.powerpoint = PowerPoint.App()

        self.prs = None
        count = len(instructions)
        for i, bi in enumerate(instructions):
            if not self._continue:
                break

            if not bi.enabled:
                continue

            monitor.set_subrange(i * 100 / count, (i + 1) * 100 / count)
            try:
                bi.execute(self, self.prs)
            except Exception as e:
                traceback.print_exc()
                self.error_message(str(e))

        self.progress_message(100, _("Cleaning up after running all the commands."))
        self.close()

        self.progress_message(100, _("Processing is done successfully."))
