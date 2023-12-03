"""
"""
import datetime
import errno
import locale
import math
import os
import re
import shutil
import sys
import tempfile
import traceback
import typing
import zipfile

from PIL import ImageColor

from bible import bibcore
from bible import biblang
from bible import fileformat as bibfileformat
from hymn import hymncore
from hymn.openlpservice import OpenLPServiceWriter
from hymn.openlyrics import OpenLyricsReader, OpenLyricsWriter
from wordwrap import WordWrap

if sys.platform.startswith("win32"):
    import powerpoint_win32 as PowerPoint
else:
    import powerpoint_osx as PowerPoint

from make_transparent import color_to_transparent


def _(s):
    return s


def ngettext(s1, s2, c):
    return s1 if c == 1 else s2


Export_None = 0
Export_CleanupFiles = 1
Export_Transparent = 2


def set_translation(trans):
    global _, ngettext
    _ = trans.gettext
    ngettext = trans.ngettext


def get_contiguous_range(r):
    if isinstance(r, int):
        return [r, r]

    if not isinstance(r, list) or len(r) == 0:
        return None

    start = None
    end = None
    for i, n in enumerate(r):
        if start is None:
            start = n
            end = n
        else:
            if start + i != n:
                break
            end = n

    return [start, end]


def get_contiguous_ranges(r):
    if isinstance(r, int):
        yield [r, r]

    if not isinstance(r, list) or len(r) == 0:
        return None

    start = None
    end = None
    for i, n in enumerate(r):
        if start is None:
            start = n
            end = n
        else:
            if start + i != n:
                yield [start, end]
                start = n

            end = n

    yield [start, end]


def replace_all_notes_text(notes, text_dict):
    count = 0
    for from_text, to_text in text_dict.items():
        if from_text in notes:
            notes = notes.replace(from_text, to_text)
            count = count + 1

    return count, notes


class EvalShape:
    """class EvalShape is to support slide matching logic that supports
    slide and note with contains_text method accepting string that returns True.
    """

    def __init__(self, prs, slide_index, note_shapes):
        self.prs = prs
        self.slide_index = slide_index
        self.note_shapes = note_shapes

    def contains_text(self, text, ignore_case=False, whole_words=False):
        return self.prs.find_text_in_slide(self.slide_index, self.note_shapes, text, ignore_case, whole_words)


def populate_slide_dict(prs, slide_index):
    """populate_slide_dict() construct dict that will be used in eval() function
    which matches to a slide.
    """

    sdict = {
        "slide": EvalShape(prs, slide_index, False),
        "note": EvalShape(prs, slide_index, True),
    }

    return sdict


def evaluate_to_single_slide(prs, expr):
    if not expr:
        return None

    # if the expr works with empty dict meaning it doesn't have dependency to slide dict,
    # return it.
    try:
        gdict = {}
        result = eval(expr, gdict, None)
        return result
    except NameError:
        # print("Error: %s" % e)
        pass

    # Use each slide's dict to evaluate the expr.
    for index in range(prs.slide_count()):
        gdict = populate_slide_dict(prs, index)

        try:
            eval_result = eval(expr, gdict, None)
            if eval_result:
                return index
        except NameError:
            # print("Error: %s" % e)
            pass


def evaluate_to_multiple_slide(prs, expr):
    if expr is None:
        return None

    # if the expr works with empty dict meaning it doesn't have dependency to slide dict,
    # return it.
    try:
        gdict = {}
        result = eval(expr, gdict, None)
        return result
    except NameError:
        # print("Error: %s" % e)
        pass

    # Use each slide's dict to evaluate the expr.
    result = []
    for index in range(prs.slide_count()):
        gdict = populate_slide_dict(prs, index)

        try:
            eval_result = eval(expr, gdict, None)
            if eval_result:
                result.append(index)
        except NameError:
            # print("Error: %s" % e)
            pass

    if len(result) == 0:
        return None

    return result


class DirSymbols(dict):
    def update_map(self):
        pass


class Command:
    def __init__(self):
        """Command base class."""
        self.enabled = True

    def get_enabled(self):
        return self.enabled

    def set_enabled(self, enabled):
        self.enabled = enabled

    @staticmethod
    def translate_one_path(value, dir_dict, to_symbol=True):
        new_value = None
        # convert absolute path to symbolized one.
        if to_symbol:
            for symbol, directory in dir_dict.items():
                if value.startswith(directory):
                    new_value = f"$({symbol})" + value[len(directory) :]
                    break
        # convert symbolized path to absolute one.
        else:
            for symbol, directory in dir_dict.items():
                symbol_str = f"$({symbol})"
                if value.startswith(symbol_str):
                    new_value = directory + value[len(symbol_str) :]
                    break

        return new_value

    def translate_dir_symbols(self, dir_dict, to_symbol=True):
        for attr in self.__dict__:
            if attr.endswith(("filename", "dirname")):
                value = getattr(self, attr)
                new_value = Command.translate_one_path(value, dir_dict, to_symbol)
                if new_value:
                    setattr(self, attr, new_value)
            elif attr.endswith("filelist"):
                filelist = getattr(self, attr)
                for i, file in enumerate(filelist):
                    new_file = Command.translate_one_path(file, dir_dict, to_symbol)
                    if new_file:
                        filelist[i] = new_file


class OpenFile(Command):
    def __init__(self, filename, notes_filename=None):
        """OpenFile opens the template ppt file to operate on."""
        super().__init__()

        self.filename = filename
        self.notes_filename = notes_filename

    def execute(self, cm, prs):
        if not self.filename:
            cm.progress_message(0, _("Creating a new presentation."))

            prs = cm.powerpoint.new_presentation()
        else:
            cm.progress_message(0, _("Opening a template presentation file '{filename}'.").format(filename=self.filename))

            if not os.path.exists(self.filename):
                cm.error_message(_("Cannot open a template presentation file '{filename}'.").format(filename=self.filename))
                raise FileNotFoundError(errno.ENOENT, os.strerror(errno.ENOENT), self.filename)

            prs = cm.powerpoint.open_presentation(self.filename)

        cm.set_presentation(prs)

        if self.notes_filename:
            cm.progress_message(90, _("Opening a template notes file '{filename}'.").format(filename=self.notes_filename))

            if not os.path.exists(self.notes_filename):
                cm.error_message(_("Cannot open a template notes file '{filename}'.").format(filename=self.notes_filename))
                raise FileNotFoundError(errno.ENOENT, os.strerror(errno.ENOENT), self.notes_filename)

            notes = ""
            with open(self.notes_filename, "rt", encoding="utf-8") as f:
                notes = f.read()

            cm.set_notes(notes)


class SaveFiles(Command):
    def __init__(self, filename, lyrics_archive_filename=None, notes_filename=None, verses_filename=None):
        """SaveFiles saves the current processed presentation to a given filename."""
        super().__init__()

        self.filename = filename
        self.lyrics_archive_filename = lyrics_archive_filename
        self.notes_filename = notes_filename
        self.verses_filename = verses_filename

    def execute(self, cm, prs):
        filename = cm.replace_format_vars(self.filename)
        cm.progress_message(0, _("Saving the presentation to the file '{filename}'.").format(filename=filename))
        prs.saveas(filename)

        all_lyric_files = cm.lyric_manager.all_lyric_files
        if self.lyrics_archive_filename and len(all_lyric_files):
            archive_filename = cm.replace_format_vars(self.lyrics_archive_filename)
            filename, ext = os.path.splitext(archive_filename)
            ext = ext.lower()
            cm.progress_message(80, _("Saving the lyrics to the archive file '{filename}'.").format(filename=archive_filename))
            if ext == ".zip":
                self.create_zip_lyric_files(archive_filename, all_lyric_files)
            elif ext == ".osz":
                self.create_osz_lyric_files(cm, archive_filename, all_lyric_files)
            else:
                cm.error_message(_("Unknown file extension for lyrics archive file '{filename}'.").format(filename=archive_filename))

        if self.notes_filename:
            notes_filename = cm.replace_format_vars(self.notes_filename)
            cm.progress_message(90, _("Saving the notes to the file '{filename}'.").format(filename=notes_filename))

            try:
                notes = cm.get_notes()
                with open(notes_filename, "wt", encoding="utf-8") as f:
                    f.write(notes)
            except FileNotFoundError:
                cm.error_message(_("Cannot save the notes file '{filename}'.").format(filename=notes_filename))

        if self.verses_filename:
            verses_filename = cm.replace_format_vars(self.verses_filename)
            cm.progress_message(95, _("Saving Bible verses to file '{filename}'.").format(filename=verses_filename))

            main_verses = cm.bible_verse.main_verses
            verses_text = cm.bible_verse._verses_text

            # Save bible verse to text file.
            try:
                with open(verses_filename, "wt", encoding="utf-8") as f:
                    f.write(biblang.UNICODE_BOM)
                    print(f"{main_verses}", file=f)
                    for d in verses_text:
                        v = d[cm.bible_verse.each_verse_name1]
                        print(f"{v.no} {v.text}", file=f)
            except FileNotFoundError:
                cm.error_message(_("Cannot save bible verses to the file '{filename}'.").format(filename=verses_filename))

    def create_zip_lyric_files(self, zipfilename, files):
        with zipfile.ZipFile(zipfilename, "w", zipfile.ZIP_DEFLATED) as zipf:
            for file in files:
                if isinstance(file, dict):
                    song = file["song"]
                    key = song.title
                    xml_content = file["xml_content"]
                    zipf.writestr(key, xml_content)
                else:
                    zipf.write(file, os.path.basename(file))

    def create_osz_lyric_files(self, cm, zipfilename, filelist):
        song_list = []
        xml_list = []
        for filename in filelist:
            if isinstance(filename, dict):
                song = filename["song"]
                xml_content = filename["xml_content"]
            else:
                song = cm.lyric_manager.read_song(filename)

                xml_content = ""
                with open(filename, "r", encoding="utf-8") as file:
                    xml_content = file.read().replace("\n", "")

            song_list.append(song)
            xml_list.append(xml_content)

        writer = OpenLPServiceWriter()
        writer.write(zipfilename, song_list, xml_list)


class InsertSlides(Command):
    def __init__(self, insert_location, separator_slides, filelist):
        """InsertSlides inserts all slides from each file in filelist after insert_location.
        insert_location is the expression that the inserted slide will located.
        separator_slides is separator slides between each inserted slides, and last one will not be inserted.
        """
        super().__init__()

        self.insert_location = insert_location
        self.separator_slides = separator_slides
        self.filelist = filelist

    def execute(self, cm, prs):
        file_count = len(self.filelist)
        insert_format = ngettext("Inserting slides from {file_count} file.", "Inserting slides from {file_count} files.", file_count)
        cm.progress_message(0, insert_format.format(file_count=file_count))

        insert_location = evaluate_to_single_slide(prs, self.insert_location)
        if insert_location is None:
            print("Insert location({self.insert_location}) is not found!")
            insert_location = prs.slide_count()

        separator_slides = None
        if self.separator_slides:
            separator_slides = evaluate_to_multiple_slide(prs, self.separator_slides)
            separator_slides = prs.slide_index_to_ID(separator_slides)

        for i, filename in enumerate(self.filelist):
            percent = 100 * i / file_count
            cm.progress_message(percent, _("Inserting slides from file '{filename}'.").format(filename=filename))

            added_count = prs.insert_file_slides(insert_location, filename)
            insert_location = insert_location + added_count

            # Add separator except last file.
            if separator_slides and i + 1 < len(self.filelist):
                separator_slides2 = prs.slide_ID_to_index(separator_slides)
                added_count = prs.duplicate_slides(separator_slides2, insert_location + 1)
                insert_location = insert_location + added_count


class InsertLyrics(Command):
    INSERT_LYRIC_SLIDE = 1
    INSERT_LYRIC_TEXT = 2
    INSERT_LYRIC_BOTH = 3

    def __init__(
        self,
        slide_insert_location,
        slide_separator_slides,
        lyric_insert_location,
        lyric_separator_slides,
        lyric_pattern,
        archive_lyric_file,
        filelist,
        flags=0,
    ):
        """InsertLyrics has two operations into one class.
        The reason for having two operations in one class is to manage one filelist that can handle both operations.

        It inserts score slides from each file in filelist after slide_insert_location.
        slide_insert_location is the expression that the inserted slide will located.
        slide_separator_slides is separator slides between each inserted slides, and last one will not be inserted.

        It inserts lyrics from each file in filelist after lyric_insert_location.
        lyric_insert_location is the expression to the location of existing slide that will be duplicated.
        lyric_separator_slides is separator slides between each inserted slides, and last one will not be inserted.
        archive_lyric_file is a bool flag to archive the xml files into .zip or .osz(OpenLP service file) specified in the SaveFiles command.

        filelist is the ppt filename and for lyric file, the extension .xml will be used to get the lyrics.

        flags is an option how to handle both score and lyric slides.
        """
        super().__init__()

        self.slide_insert_location = slide_insert_location
        self.slide_separator_slides = slide_separator_slides

        self.lyric_insert_location = lyric_insert_location
        self.lyric_separator_slides = lyric_separator_slides
        self.lyric_pattern = lyric_pattern
        self.archive_lyric_file = archive_lyric_file

        self.filelist = filelist
        self.flags = flags

    def get_filelist(self, cm, filetype):
        filelist = []
        if filetype == self.INSERT_LYRIC_SLIDE:
            for fn in self.filelist:
                base, ext = os.path.splitext(fn)
                ext = ext.lower()
                if ext == ".pptx" or ext == ".ppt":
                    pass
                else:
                    fn = base + ".pptx"
                    if os.path.exists(fn):
                        pass
                    else:
                        # do not check, so the execute() function can throw exception.
                        fn = base + ".ppt"
                filelist.append(fn)
        elif filetype == self.INSERT_LYRIC_TEXT:
            filelist = [cm.lyric_manager.search_lyric_file(fn) for fn in self.filelist]

        return filelist

    def execute(self, cm, prs):
        if self.flags & self.INSERT_LYRIC_SLIDE:
            lyric_filelist = self.get_filelist(cm, self.INSERT_LYRIC_SLIDE)
            slides = InsertSlides(self.slide_insert_location, self.slide_separator_slides, lyric_filelist)
            slides.execute(cm, prs)
            del slides

        if self.flags & self.INSERT_LYRIC_TEXT:
            lyric_filelist = self.get_filelist(cm, self.INSERT_LYRIC_TEXT)
            self.execute_lyric_files(cm, prs, lyric_filelist)

        if self.archive_lyric_file:
            lyric_filelist = self.get_filelist(cm, self.INSERT_LYRIC_TEXT)
            cm.add_lyric_file(lyric_filelist)

    def execute_lyric_files(self, cm, prs, filelist):
        file_count = len(filelist)
        insert_format = ngettext("Inserting lyrics from {file_count} file.", "Inserting lyrics from {file_count} files.", file_count)
        cm.progress_message(0, insert_format.format(file_count=file_count))

        lyric_insert_location = evaluate_to_single_slide(prs, self.lyric_insert_location)
        if lyric_insert_location is None:
            cm.progress_message(0, _("No repeatable slides are found. Aborting the command."))
            return

        lyric_separator_slides = None
        separator_slide_count = 0
        if self.lyric_separator_slides:
            lyric_separator_slides = evaluate_to_multiple_slide(prs, self.lyric_separator_slides)
            lyric_separator_slides = prs.slide_index_to_ID(lyric_separator_slides)
            if isinstance(lyric_separator_slides, list):
                separator_slide_count = len(lyric_separator_slides)

        songs = cm.read_songs(filelist)

        # Because the original slides are updated with song, duplicate slides first.
        self.duplicate_slides(prs, lyric_insert_location, lyric_separator_slides, separator_slide_count, songs)

        last_index = len(songs) - 1
        for i, song in enumerate(songs):
            percent = 100 * i / file_count
            cm.progress_message(percent, _("Inserting lyric from '{filename}'.").format(filename=filelist[i]))

            lines = list(song.get_lines_by_order())
            for j, l in enumerate(lines):
                text_dict = {self.lyric_pattern: l.text}
                prs.replace_one_slide_texts(lyric_insert_location + j, text_dict)

            added_count = len(lines)
            lyric_insert_location = lyric_insert_location + added_count

            # Skip separator
            if separator_slide_count != 0 and i < last_index:
                added_count = separator_slide_count
                lyric_insert_location = lyric_insert_location + added_count

    def duplicate_slides(self, prs, source_location, lyric_separator_slides, separator_slide_count, songs):
        """Duplicate slide based on count_of_lyric1, separators, count_of_lyric2, separators, ...
        songs is a two dimensional string array.
        """
        lyric_insert_location = source_location
        last_index = len(songs) - 1
        for i, song in enumerate(songs):
            lines = list(song.get_lines_by_order())
            if i == last_index:
                duplicate_count = len(lines) - 1
            else:
                duplicate_count = len(lines)

            added_count = prs.duplicate_slides(source_location, lyric_insert_location, duplicate_count)
            lyric_insert_location = lyric_insert_location + added_count

            # Add separator except th first lyric.
            if separator_slide_count != 0 and i < last_index:
                separator_slide_indices = prs.slide_ID_to_index(lyric_separator_slides)
                added_count = prs.duplicate_slides(separator_slide_indices, lyric_insert_location, separator_slide_count)
                lyric_insert_location = lyric_insert_location + added_count


class FormatObj(object):
    VAR_FORMAT_SEP = ":"  # Change it to semi-colon(';')

    def __init__(self, format="", value: str = ""):
        # self.fobj_type = self.__class__.__name__
        self.format = format  # class specific format string
        self.value = value  # class specific value object

    def get_flattened_dict(self):
        return {"format_type": self.__class__.__name__, "value": self.value}

    @staticmethod
    def build_format_pattern(varnames):
        varname_pattern = ""
        if isinstance(varnames, str) and str:
            varname_pattern = r"\{(" + re.escape(varnames) + r")(" + FormatObj.VAR_FORMAT_SEP + "[^\}]+)?" + r"\}"
        else:
            try:
                for var in varnames:
                    if varname_pattern:
                        varname_pattern = varname_pattern + "|"
                    varname_pattern = varname_pattern + re.escape(var)
            except TypeError as e:
                print(varnames + " is not iterable.")

            if varname_pattern:
                varname_pattern = r"\{(" + varname_pattern + r")(" + FormatObj.VAR_FORMAT_SEP + "[^\}]+)?" + r"\}"

        return varname_pattern

    def __eq__(self, other: object) -> bool:
        return self.format == other.format and self.value == other.value


class BibleVerseFormat(FormatObj):
    def __init__(self, format: str = "", value: str = ""):
        super().__init__(format, value)

        # self.format : '%B' : long book name, '%b': short book name, '%c': chapter.no, '%v': v.no, %t: verse text
        # self.value : bible_core.Verse object

    @staticmethod
    def build_fr_dict(slide_text: typing.Any, verses: typing.Dict[str, bibcore.Verse]):
        fr_dict = {}
        for each_verse_name, verse in verses.items():
            varname_pattern = FormatObj.build_format_pattern(each_verse_name)
            varname_re = re.compile(varname_pattern)

            def process_format_var(var: typing.Any):
                if isinstance(var, list):
                    for elem in var:
                        process_format_var(elem)
                elif isinstance(var, str):
                    for m in varname_re.finditer(var):
                        format = m.group(2)
                        if len(format):
                            format = format[1:]  # remove FormatObj.VAR_FORMAT_SEP
                        key = "{" + m.group(1) + m.group(2) + "}"
                        value = BibleVerseFormat.translate_verse(format, verse)
                        if key not in fr_dict:
                            fr_dict[key] = value

            process_format_var(slide_text)

        return fr_dict

    @staticmethod
    def translate_verse(format: str, verse: bibcore.Verse):
        s = format
        s = s.replace("%B", verse.book.name)
        s = s.replace("%b", verse.book.short_name)
        s = s.replace("%c", str(verse.chapter.no))
        s = s.replace("%v", str(verse.no))
        s = s.replace("%t", verse.text)
        return s


class DateTimeFormat(FormatObj):
    def __init__(self, format: str = "", value: str = ""):
        super().__init__(format, value)

        # self.format : https://docs.python.org/3/library/datetime.html#strftime-strptime-behavior

        # self.value : https://docs.python.org/3/library/datetime.html#datetime.datetime

    @staticmethod
    def datetime_from_c_locale(str_value: str, format: str = "%Y-%m-%d"):
        saved = None
        dt_value = None
        try:
            saved = locale.setlocale(locale.LC_ALL, "C")
            dt_value = datetime.datetime.strptime(str_value, format)
        except:
            if saved:
                locale.setlocale(locale.LC_ALL, saved)

        return dt_value

    def build_replace_string(self, format: str):
        # https://docs.python.org/3/library/datetime.html#strftime-and-strptime-format-codes

        # If the format contains # btw % and alphabet, it will remove leading zero.
        # Mimick the behavior mentioned in https://stackoverflow.com/questions/904928/python-strftime-date-without-leading-0
        dt_value = DateTimeFormat.datetime_from_c_locale(self.value)
        value = format
        value = value.replace("%Y", str(dt_value.year).zfill(4))
        value = value.replace("%y", str(dt_value.year % 100).zfill(2))
        value = value.replace("%m", str(dt_value.month).zfill(2))
        value = value.replace("%#m", str(dt_value.month))
        value = value.replace("%d", str(dt_value.day).zfill(2))
        value = value.replace("%#d", str(dt_value.day))
        value = value.replace("%W", dt_value.strftime("%W"))
        value = value.replace("%w", dt_value.strftime("%w"))
        value = value.replace("%U", dt_value.strftime("%U"))
        value = value.replace("%j", dt_value.strftime("%j"))
        value = value.replace("%B", dt_value.strftime("%B"))
        value = value.replace("%b", dt_value.strftime("%b"))
        value = value.replace("%A", dt_value.strftime("%A"))
        value = value.replace("%a", dt_value.strftime("%a"))
        return value


class SetVariables(Command):
    def __init__(self, format_dict: dict = {}, str_dict: dict = {}):
        """SetVariables finds all occurrence of find_text and replace it with replace_text,
        which are (k, v) pair in dictionary str_dict.
        """
        super().__init__()

        self.format_dict = format_dict
        self.str_dict = str_dict

    def execute(self, cm: typing.Any, prs: typing.Any):
        text_count = len(self.str_dict)
        replace_format = ngettext("Replacing {text_count} text.", "Replacing {text_count} texts.", text_count)
        cm.progress_message(0, replace_format.format(text_count=text_count))

        cm.add_global_variables(self.str_dict, self.format_dict)
        cm.process_variable_substitution()


class DuplicateWithText(Command):
    def __init__(
        self,
        slide_range,
        repeat_range,
        find_text,
        replace_texts,
        preprocessing_script,
        archive_lyric_file,
        optional_line_break,
    ):
        """DuplicateWithText finds all occurrence of find_text and replace it replace_texts,
        which are list of text.
        Because replace_texts are a list, slides in repeat_range will be duplicated
        to match with len(replace_texts) == len(slide_range) + repeated len(repeat_range).
        """
        super().__init__()

        self.slide_range = slide_range
        self.repeat_range = repeat_range
        self.find_text = find_text
        self.replace_texts = replace_texts
        self.preprocessing_script = preprocessing_script
        self.archive_lyric_file = archive_lyric_file
        self.optional_line_break = optional_line_break
        self.enable_wordwrap = False
        self.wordwrap_font = ""
        self.wordwrap_pagewidth = 0

    def execute(self, cm, prs):
        cm.progress_message(0, _("Duplicating slides and replacing texts."))

        slide_ranges = evaluate_to_multiple_slide(prs, self.slide_range)
        slide_ranges = reversed(list(get_contiguous_ranges(slide_ranges)))

        repeat_ranges = evaluate_to_multiple_slide(prs, self.repeat_range)
        repeat_ranges = reversed(list(get_contiguous_ranges(repeat_ranges)))

        while True:
            try:
                slide_range = next(slide_ranges)
            except StopIteration:
                break

            try:
                repeat_range = next(repeat_ranges)
            except StopIteration:
                break

            # incorrect ranges are specified.
            if not (slide_range[0] <= repeat_range[0] and repeat_range[1] <= slide_range[1]):
                return

            # The input template pptx contains slide_range[1] - slide_range[0] + 1 slides for the operation.
            # From that, repeat_range[1] - repeat_range[0] + 1 slides are repeatable and can be duplicated.
            # So, from texts, find out how many we need to duplicate and do the operation.
            # Then, replace texts to get the final slides.
            template_slide_count = slide_range[1] - slide_range[0] + 1
            repeatable_slide_count = repeat_range[1] - repeat_range[0] + 1
            replace_texts = self.replace_texts
            if len(self.preprocessing_script):
                replace_texts = self.run_script(self.preprocessing_script, self.replace_texts)
                if replace_texts is None:
                    replace_texts = self.replace_texts
            else:
                replace_texts = self.replace_texts
            total_slide_count = len(replace_texts)

            duplicate_count = total_slide_count - template_slide_count
            duplicate_count = int(math.ceil(float(duplicate_count) / repeatable_slide_count))
            source_index = list(range(repeat_range[0], repeat_range[1] + 1))
            _slides = prs.duplicate_slides(source_index, repeat_range[0], duplicate_count)

            for i, text in enumerate(replace_texts):
                percent = 100 * i / total_slide_count
                cm.progress_message(percent, _("Replacing texts."))

                text = cm.replace_format_vars(text)
                text_dict = {self.find_text: text}
                prs.replace_one_slide_texts(slide_range[0] + i, text_dict)

        if self.archive_lyric_file:
            lyric_filelist = self.produce_as_lyric_file(cm)
            cm.add_lyric_file(lyric_filelist)

    def produce_as_lyric_file(self, cm):
        song = hymncore.Song()
        song.title = self.find_text.replace("{", "").replace("}", "")
        replace_texts = self.replace_texts
        if self.enable_wordwrap and self.wordwrap_font and self.wordwrap_pagewidth > 0:
            with WordWrap(self.wordwrap_font) as wwo:
                replace_texts = wwo.wordwrap(replace_texts, self.wordwrap_pagewidth)
        for vno, text in enumerate(replace_texts):
            text = cm.replace_format_vars(text)
            v = hymncore.Verse()
            v.no = "v" + str(vno + 1)
            sl = text.splitlines()
            len_sl_1 = len(sl) - 1
            for i, l in enumerate(sl):
                if self.optional_line_break != 0:
                    optional_break = i != len_sl_1 and (i % self.optional_line_break == 0)
                else:
                    optional_break = i != len_sl_1
                line = hymncore.Line(l, optional_break)
                v.lines.append(line)

            song.verses.append(v)

        xml_content = ""
        filename = ""
        with tempfile.NamedTemporaryFile(suffix=".xml", delete=False) as f:
            filename = f.name

        writer = OpenLyricsWriter()
        writer.write_song(filename, song)

        with open(filename, "r", encoding="utf-8") as file:
            xml_content = file.read().replace("\n", "")

        os.remove(filename)

        return {"song": song, "xml_content": xml_content}

    def run_script(self, script, text_list):
        try:
            gdict = {"input": text_list}
            exec(script, gdict)
            return gdict["output"]
        except Exception as e:
            print("Error: %s" % e)


class GenerateBibleVerse(Command):
    def __init__(self, bible_format: str):
        """GenerateBibleVerse is two operations combined.

        1. It will replace all main_verse_name1 to main_verses in all slides.
        2. verses_text will be populated from main_verses and additional_verses.
        Then, the repeat_range slides will be duplicated to match with len(verses_text) == repeated len(repeat_range)
        and replace each_verse_name1 to pair (no, text) in verses_text with bible verse variable substitution.

        bible_format is one of supported bible file format: See also bible.fileformat.
        bible_version1 and bible_version2 can be a valid Bible Version.
        main_verse_name1 is the variable name of the main bible verses.
        each_verse_name1 is the variable name of each verse in the slides.
        main_verses is a comma separated main verse titles for the service.
        additional_verses is an additional comma separated verse titles for the service.
        repeat_range is a slide range mark to duplicate the slides.
        """
        super().__init__()

        self.bible_format = bible_format
        self.bible_version1 = ""
        self.main_verse_name1 = ""
        self.each_verse_name1 = ""
        self.bible_version2 = ""
        self.main_verse_name2 = ""
        self.each_verse_name2 = ""
        self.main_verses = ""
        self.additional_verses = ""
        self.repeat_range = ""

        self._verses_text = []

    def populate_all_verses(self, cm: typing.Any):
        self._verses_text = []

        bibles = []
        valid_each_verse_names = []
        bible_indices = None
        versions = [self.bible_version1, self.bible_version2]
        each_verse_names = [self.each_verse_name1, self.each_verse_name2]
        for i in range(len(versions)):
            version = versions[i]
            each_verse_name = each_verse_names[i]
            if not version or not each_verse_name:
                continue

            bible = bibfileformat.read_version(self.bible_format, version)
            if bible is None:
                cm.error_message(
                    _("Cannot find the Bible format={bible_format} and version={version}.").format(
                        bible_format=self.bible_format, version=version
                    )
                )
                return

            bibles.append(bible)
            valid_each_verse_names.append(each_verse_name)

            if bible_indices is None:
                bible_indices = self.translate_to_bible_index(cm, bible, self.main_verses)
                bible_indices = bible_indices + self.translate_to_bible_index(cm, bible, self.additional_verses)

        all_verses_text = []
        for bible_index in bible_indices:
            startpos = len(all_verses_text)
            verses_text = {}
            for b in range(len(bibles)):
                i = startpos
                bible = bibles[b]
                bi, ct1, vs1, ct2, vs2 = bible_index
                texts = bible.extract_texts_from_bible_index(bi, ct1, vs1, ct2, vs2)
                if texts is None:
                    cm.error_message(_("Cannot extract Bible {bi=}, {ct1=}, {vs1=}, {ct2=}, {vs2=}."))

                for text in texts:
                    if i < len(all_verses_text):
                        all_verses_text[i][valid_each_verse_names[b]] = text
                    else:
                        verses_text = {valid_each_verse_names[b]: text}
                        all_verses_text.append(verses_text)
                    i += 1

        self._verses_text = all_verses_text

    def split_verses(self, verses: str) -> typing.List[str]:
        """Split verses by comma. If there is no book and chapter parts by checking ':',
        add it from the previous one."""
        splitted_verses = verses.split(",")
        book_chapter = ""
        for i, v in enumerate(splitted_verses):
            v = v.strip()
            index = v.find(":")
            if index != -1:
                book_chapter = v[: index + 1]  # include ':'
            else:
                v = book_chapter + v

            splitted_verses[i] = v

        return splitted_verses

    def translate_to_bible_index(self, cm: typing.Any, bible: typing.Any, verses: str) -> typing.List[typing.Any]:
        bible_index = []
        if verses is not None:
            splitted_verses = self.split_verses(verses)
            for verse in splitted_verses:
                verse = verse.strip()
                if not verse:
                    continue

                verses_text = None
                try:
                    result = bible.translate_to_bible_index(verse)
                    if result is None:
                        if verse == verses:
                            cm.error_message(_("Cannot find the Bible verse={verse}.").format(verse=verse))
                        else:
                            cm.error_message(_("Cannot find the Bible verse={verse} in {verses}.").format(verse=verse, verses=verses))

                    bi, ct1, vs1, ct2, vs2 = result
                    bible_index.append(result)
                except ValueError:
                    pass

                if verses_text is None:
                    continue

        return bible_index

    def execute(self, cm: typing.Any, prs: typing.Any):
        self.execute_on_slides(cm, prs)
        self.execute_on_notes(cm)

    def execute_on_slides(self, cm: typing.Any, prs: typing.Any):
        cm.progress_message(0, _("Processing Bible Verse."))

        self.populate_all_verses(cm)

        cm.set_bible_verse(self)

        text_dict = {self.main_verse_name1: self.main_verses}
        cm.add_global_variables(text_dict)
        cm.process_variable_substitution()

        repeat_range = evaluate_to_multiple_slide(prs, self.repeat_range)
        repeat_range = prs.slide_index_to_ID(repeat_range)

        repeat_count = len(repeat_range) * (len(self._verses_text) - 1)
        if repeat_count > 0:
            duplicate_format = ngettext(
                "Duplicating {repeat_count} slide for Bible Verse.", "Duplicating {repeat_count} slides for Bible Verse.", repeat_count
            )
            cm.progress_message(0, duplicate_format.format(repeat_count=repeat_count))

        repeat_count = len(repeat_range)
        for i, sid in enumerate(repeat_range):
            percent = 100 * i / repeat_count
            cm.progress_message(percent, _("Duplicating slides and replacing texts."))

            index = prs.slide_ID_to_index(sid)
            _slides = prs.duplicate_slides(index, index, len(self._verses_text) - 1)

            for i, v in enumerate(self._verses_text):
                slide_text = prs.get_text_in_slide(index + i, False)
                text_dict = self.get_verse_dict(slide_text, v)
                prs.replace_one_slide_texts(index + i, text_dict)

    def execute_on_notes(self, cm: typing.Any):
        if cm.notes:
            notes = cm.notes
            text_dict = {self.main_verse_name1: self.main_verses}
            _, notes = replace_all_notes_text(notes, text_dict)

            # check whether we need to duplicate the line by checking self.each_verse_name1 exists.
            if self.each_verse_name1 and self.each_verse_name1 in notes:
                lines = notes.split("\n")
                repeat_range = [i for i, l in enumerate(lines) if self.each_verse_name1 in l]
                for index in reversed(repeat_range):
                    for _ in range(len(self._verses_text) - 1):
                        lines.insert(index, lines[index])

                    for i, v in enumerate(self._verses_text):
                        text_dict = self.get_verse_dict(lines[index + i], v)
                        _, lines[index + i] = replace_all_notes_text(lines[index + i], text_dict)

                notes = "\n".join(lines)

            cm.notes = notes

    def get_verse_dict(self, slide_text: typing.Any, verses: typing.Dict[str, bibcore.Verse]):
        text_dict = BibleVerseFormat.build_fr_dict(slide_text, verses)
        return text_dict

    def get_flattened_dict(self):
        return {
            "bible_version1": self.bible_version1,
            "main_verse_name1": self.main_verse_name1,
            "each_verse_name1": self.each_verse_name1,
            "bible_version2": self.bible_version2,
            "main_verse_name2": self.main_verse_name2,
            "each_verse_name2": self.each_verse_name2,
            "main_verses": self.main_verses,
            "additional_verses": self.additional_verses,
            "repeat_range": self.repeat_range,
        }


def rename_filename_to_zeropadded(dirname, num_digits):
    r"""Rename Slide(\d+).PNG to Slide%03d.png so that the length is same
    and they can be sorted properly.
    Powerpoint generates filename as 1-9 and 10, etc. that make sorting difficult.
    """
    if num_digits <= 1:
        return

    def replace_format_3digits(m):
        if m.group(2).isdigit():
            num = int(m.group(2))
            fmt = r"%s%0" + str(num_digits) + r"d%s"

            ext = m.group(3)
            ext = ext.lower()
            if len(ext) == 5:
                if ext == ".jpeg":
                    ext = ".jpg"
                elif ext == ".tiff":
                    ext = ".tif"

            s = fmt % (m.group(1), num, ext)
        else:
            s = m.group(0)
        return s

    matching_fn_pattern = r"(.*?)(\d+)(\.(gif|jpg|jpeg|png|tif|tiff))"
    repl_func = replace_format_3digits

    fn_re = re.compile(matching_fn_pattern, re.IGNORECASE)

    files = os.listdir(dirname)
    for fn in files:
        new_fn = fn_re.sub(repl_func, fn)
        if new_fn == fn:
            continue

        old_fullname = os.path.join(dirname, fn)
        new_fullname = os.path.join(dirname, new_fn)
        os.rename(old_fullname, new_fullname)


def rmtree_except_self(dirname):
    try:
        files = os.listdir(dirname)
        for fn in files:
            fullname = os.path.join(dirname, fn)
            if os.path.isdir(fullname):
                shutil.rmtree(fullname)
            else:
                os.remove(fullname)
    except OSError as err:
        if err.errno == errno.ENOENT:
            return

        raise


def mkdir_if_not_exists(dirname):
    try:
        os.mkdir(dirname)
    except OSError as err:
        if err.errno == errno.EEXIST:
            return

        raise


class PromptCommand(Command):
    def __init__(self, message):
        super().__init__()

        self.message = message

    def execute(self, cm, prs):
        cm.progress_message(0, _("Processing PromptCommand."))

        print(f"{self.message}")
        input()

        prs.refresh_page_cache()


class ExportSlides(Command):
    def __init__(self, slide_range, out_dirname, image_type, flags, color=None):
        super().__init__()

        self.slide_range = slide_range
        self.out_dirname = out_dirname
        self.image_type = image_type
        self.flags = flags
        self.color = color

    def execute(self, cm, prs):
        cm.progress_message(0, _("Exporting slide shapes as images to '{dirname}'.").format(dirname=self.out_dirname))

        if self.flags & Export_CleanupFiles:
            rmtree_except_self(self.out_dirname)
            mkdir_if_not_exists(self.out_dirname)

        if not hasattr(prs, "export_slide_as"):
            self.export_all(cm, prs)
            return

        slide_range = evaluate_to_multiple_slide(prs, self.slide_range)
        if slide_range is None:
            return

        image_type = cm.normalize_image_type(self.image_type)
        prs.export_slides_as(slide_range, self.out_dirname, image_type)

        if (self.flags & Export_Transparent) and self.image_type == "png" and self.color:
            cm.progress_message(50, _("Converting to transparent images."))

            color = ImageColor.getrgb(self.color)
            for index in slide_range:
                f = cm.get_filename_from_slideno(self.image_type, index)
                filename = os.path.join(self.out_dirname, f)
                color_to_transparent(filename, filename, color)

    def export_all(self, cm, prs):
        src_dirname = cm.generate_image_files(self.image_type, self.color)

        slide_range = evaluate_to_multiple_slide(prs, self.slide_range)
        slide_range = get_contiguous_range(slide_range)
        if slide_range is None:
            return

        if (self.flags & Export_Transparent) and self.image_type == "png" and self.color:
            cm.progress_message(50, _("Converting to transparent images."))

            color = ImageColor.getrgb(self.color)
            for index in range(slide_range[0], slide_range[1] + 1):
                f = cm.get_filename_from_slideno(self.image_type, index)
                filename = os.path.join(src_dirname, f)
                color_to_transparent(filename, filename, color)

        for index in range(slide_range[0], slide_range[1] + 1):
            f = cm.get_filename_from_slideno(self.image_type, index)
            shutil.copy(os.path.join(src_dirname, f), self.out_dirname)


class ExportShapes(Command):
    def __init__(self, slide_range, out_dirname, image_type, flags):
        super().__init__()

        self.slide_range = slide_range
        self.out_dirname = out_dirname
        self.image_type = image_type
        self.flags = flags

    def execute(self, cm, prs):
        cm.progress_message(0, _("Exporting slide shapes as images to '{dirname}'.").format(dirname=self.out_dirname))

        if self.flags & Export_CleanupFiles:
            rmtree_except_self(self.out_dirname)
            mkdir_if_not_exists(self.out_dirname)

        slide_range = evaluate_to_multiple_slide(prs, self.slide_range)
        if slide_range is None:
            return

        prs.export_slide_shapes_as(slide_range, self.out_dirname, self.image_type)


class LyricManager:
    def __init__(self, cm):
        self.cm = cm
        self.reader = OpenLyricsReader()
        self.lyric_file_map = {}
        self.all_lyric_files = []
        self.lyric_search_path = None

    def reset_exec_vars(self):
        self.lyric_file_map = {}
        self.all_lyric_files = []

    def read_song(self, filename):
        if filename in self.lyric_file_map:
            return self.lyric_file_map[filename]

        if not os.path.exists(filename):
            self.cm.error_message(_("Cannot open a lyric file '{filename}'.").format(filename=filename))
            raise FileNotFoundError(errno.ENOENT, os.strerror(errno.ENOENT), filename)

        try:
            song = self.reader.read_song(filename)
        except Exception as e:
            self.cm.error_message(_("Cannot open a lyric file '{filename}'.").format(filename=filename))
            raise

        self.lyric_file_map[filename] = song
        return song

    def read_songs(self, filelist):
        songs = []
        for filename in filelist:
            song = self.read_song(filename)
            songs.append(song)

        return songs

    def search_lyric_file(self, filename: str) -> typing.Optional[str]:
        _dir, fn = os.path.split(filename)
        xml_pathname = os.path.splitext(filename)[0] + ".xml"
        file_exist = os.path.exists(xml_pathname)
        if file_exist:
            return xml_pathname

        # seach xml file from search path
        if self.lyric_search_path:
            xml_filename = os.path.splitext(fn)[0] + ".xml"
            searched_xml_pathname = os.path.join(self.lyric_search_path, xml_filename)
            file_exist = os.path.exists(searched_xml_pathname)
            if file_exist:
                return searched_xml_pathname

        return xml_pathname

    def add_lyric_file(self, filelist):
        if isinstance(filelist, list):
            _songs = self.read_songs(filelist)
            self.all_lyric_files.extend(filelist)
        elif isinstance(filelist, dict):
            self.all_lyric_files.append(filelist)
        else:
            song = self.read_song(filelist)
            self.all_lyric_files.append(filelist)


class CommandManager:
    def __init__(self):
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

        self.image_dir = {}
        self.num_digits = 0

        self.bible_verse = None

        self.lyric_manager = LyricManager(self)

    def __del__(self):
        self.close()

    def close(self):
        if self.prs:
            self.prs.close()
            self.prs = None

        if not self.running_already and self.powerpoint:
            self.powerpoint.quit()

        self.powerpoint = None

        self.remove_cache()

    def reset_exec_vars(self):
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

    def set_presentation(self, prs):
        self.prs = prs

    def add_global_variables(self, str_dict, format_dict=None):
        self.string_variables.update(str_dict)
        if format_dict:
            self.format_variables.update(format_dict)

        self.var_dict = {}
        for var in self.string_variables:
            new_var = var
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

    def process_format_var(self, var):
        if isinstance(var, list):
            for elem in var:
                self.process_format_var(elem)
        elif isinstance(var, str):
            if var in self.var_dict:
                return

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

    def replace_format_vars(self, custom_str):
        if self.varname_re:
            self.process_format_var(custom_str)
        _, custom_str = replace_all_notes_text(custom_str, self.var_dict)
        return custom_str

    def process_variable_substitution(self):
        if self.varname_re:
            self.process_format_var(self.prs.get_text_in_all_slides(False))
        self.prs.replace_all_slides_texts(self.var_dict)

        if self.notes:
            if self.varname_re:
                self.process_format_var(self.notes)
            _, self.notes = replace_all_notes_text(self.notes, self.var_dict)

    def get_notes(self):
        return self.notes

    def set_notes(self, notes):
        self.notes = notes

    def set_bible_verse(self, bible_verse):
        self.bible_verse = bible_verse

    def read_songs(self, filelist):
        songs = self.lyric_manager.read_songs(filelist)
        return songs

    def search_lyric_files(self, filelist):
        xml_filelist = [self.lyric_manager.search_lyric_file(file) for file in filelist]
        return xml_filelist

    def add_lyric_file(self, lyric_file):
        self.lyric_manager.add_lyric_file(lyric_file)

    def progress_message(self, progress, message):
        self._continue = self.monitor.progress_message(progress, message)

    def error_message(self, message):
        self.monitor.error_message(message)

    # slide image file caching functions.
    def remove_cache(self):
        for _, dirname in self.image_dir.items():
            shutil.rmtree(dirname)
        self.image_dir = {}

    def normalize_image_type(self, image_type):
        image_type = image_type.lower()
        if image_type == "jpeg":
            image_type = "jpg"
        elif image_type == "tiff":
            image_type = "tif"
        return image_type

    def generate_image_files(self, image_type, color=None):
        if self.prs:
            slide_count = self.prs.slide_count()
            self.num_digits = len(f"{slide_count+1}")
        else:
            self.num_digits = 0

        image_type = self.normalize_image_type(image_type)

        if image_type not in self.image_dir:
            outdir = tempfile.mkdtemp(prefix="slides")

            self.prs.saveas_format(outdir, image_type)

            rename_filename_to_zeropadded(outdir, self.num_digits)

            self.image_dir[image_type] = outdir

        return self.image_dir[image_type]

    def get_filename_from_slideno(self, image_type, slideno):
        image_type = self.normalize_image_type(image_type)

        num = slideno + 1
        fmt = r"%0" + str(self.num_digits) + r"d"
        filename = "Slide" + (fmt % num) + "." + image_type

        return filename

    def execute_commands(self, instructions, monitor):
        """execute_commands() is the main function calling each command's execute()
        to achieve the goal specificed in each command.
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
