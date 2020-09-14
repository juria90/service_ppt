'''
'''
import errno
import math
import os
import re
import shutil
import sys
import tempfile
import traceback

from PIL import ImageColor

sys.path.append(os.path.join(os.path.dirname(os.path.abspath(__file__)), 'bible'))
sys.path.append(os.path.join(os.path.dirname(os.path.abspath(__file__)), 'hymn'))
from bible import fileformat as bibfileformat
from bible import biblang

from hymn.openlyrics import OpenLyricsReader

if sys.platform.startswith('win32'):
    import powerpoint_win32 as PowerPoint
else:
    import powerpoint_osx as PowerPoint

from make_transparent import color_to_transparent


_ = lambda s: s
ngettext = lambda s1, s2, c: s1 if c == 1 else s2

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
    '''class EvalShape is to support slide matching logic that supports
    slide and note with contains_text method accepting string that returns True.
    '''

    def __init__(self, prs, slide_index, note_shapes):
        self.prs = prs
        self.slide_index = slide_index
        self.note_shapes = note_shapes

    def contains_text(self, text, ignore_case=False, whole_words=False):
        return self.prs.find_text_in_slide(self.slide_index, self.note_shapes, text, ignore_case, whole_words)


def populate_slide_dict(prs, slide_index):
    '''populate_slide_dict() construct dict that will be used in eval() function
    which matches to a slide.
    '''

    sdict = {'slide': EvalShape(prs, slide_index, False),
             'note': EvalShape(prs, slide_index, True)
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


class Command:
    def __init__(self):
        '''Command base class.
        '''
        self.enabled = True

    def get_enabled(self):
        return self.enabled

    def set_enabled(self, enabled):
        self.enabled = enabled


class OpenFile(Command):
    def __init__(self, filename, notes_filename=None):
        '''OpenFile opens the template ppt file to operate on.
        '''
        super().__init__()

        self.filename = filename
        self.notes_filename = notes_filename

    def execute(self, cm, prs):
        if not self.filename:
            cm.progress_message(0, _('Creating a new presentation.'))

            prs = cm.powerpoint.new_presentation()
        else:
            cm.progress_message(0, _('Opening a template presentation file \'{filename}\'.').format(filename=self.filename))

            if not os.path.exists(self.filename):
                cm.error_message(_('Cannot open a template presentation file \'{filename}\'.').format(filename=self.filename))
                raise FileNotFoundError(errno.ENOENT, os.strerror(errno.ENOENT), self.filename)

            prs = cm.powerpoint.open_presentation(self.filename)

        cm.set_presentation(prs)

        if self.notes_filename:
            cm.progress_message(90, _('Opening a template notes file \'{filename}\'.').format(filename=self.notes_filename))

            if not os.path.exists(self.notes_filename):
                cm.error_message(_('Cannot open a template notes file \'{filename}\'.').format(filename=self.notes_filename))
                raise FileNotFoundError(errno.ENOENT, os.strerror(errno.ENOENT), self.notes_filename)

            notes = ''
            with open(self.notes_filename, 'rt', encoding='utf-8') as f:
                notes = f.read()

            cm.set_notes(notes)


class SaveFiles(Command):
    def __init__(self, filename, notes_filename=None, verses_filename=None):
        '''SaveFiles saves the current processed presentation to a given filename.
        '''
        super().__init__()

        self.filename = filename
        self.notes_filename = notes_filename
        self.verses_filename = verses_filename

    def execute(self, cm, prs):
        cm.progress_message(0, _('Saving the presentation to the file \'{filename}\'.').format(filename=self.filename))
        prs.saveas(self.filename)

        if self.notes_filename:
            cm.progress_message(90, _('Saving the notes to the file \'{filename}\'.').format(filename=self.notes_filename))

            try:
                notes = cm.get_notes()
                with open(self.notes_filename, 'wt', encoding='utf-8') as f:
                    f.write(notes)
            except FileNotFoundError:
                cm.error_message(_('Cannot save the notes file \'{filename}\'.').format(filename=self.notes_filename))

        if self.verses_filename:
            cm.progress_message(95, _('Saving Bible verses to file \'{filename}\'.').format(filename=self.verses_filename))

            main_verses = cm.bible_verse.main_verses
            verses_text = cm.bible_verse._verses_text

            # Save bible verse to text file.
            try:
                with open(self.verses_filename, 'wt', encoding='utf-8') as f:
                    f.write(biblang.UNICODE_BOM)
                    print(f'{main_verses}', file=f)
                    for v in verses_text:
                        print(f'{v.no} {v.text}', file=f)
            except FileNotFoundError:
                cm.error_message(_('Cannot save bible verses to the file \'{filename}\'.').format(filename=self.verses_filename))


class InsertSlides(Command):
    def __init__(self, insert_location, separator_slides, filelist):
        '''InsertSlides inserts all slides from each file in filelist after insert_location.
        insert_location is the expression that the inserted slide will located.
        separator_slides is separator slides between each inserted slides, and last one will not be inserted.
        '''
        super().__init__()

        self.insert_location = insert_location
        self.separator_slides = separator_slides
        self.filelist = filelist

    def execute(self, cm, prs):
        file_count = len(self.filelist)
        insert_format = ngettext('Inserting slides from {file_count} file.',
                                 'Inserting slides from {file_count} files.', file_count)
        cm.progress_message(0, insert_format.format(file_count=file_count))

        insert_location = evaluate_to_single_slide(prs, self.insert_location)
        if insert_location is None:
            insert_location = prs.slide_count()

        separator_slides = None
        if self.separator_slides:
            separator_slides = evaluate_to_multiple_slide(prs, self.separator_slides)
            separator_slides = prs.slide_index_to_ID(separator_slides)

        for i, filename in enumerate(self.filelist):
            percent = 100 * i / file_count
            cm.progress_message(percent, _('Inserting slides from file \'{filename}\'.').format(filename=filename))

            added_count = prs.insert_file_slides(insert_location, filename)
            insert_location = insert_location + added_count

            # Add separator except last file.
            if separator_slides and i+1 < len(self.filelist):
                separator_slides2 = prs.slide_ID_to_index(separator_slides)
                added_count = prs.duplicate_slides(separator_slides2, insert_location + 1)
                insert_location = insert_location + added_count


class InsertLyrics(Command):
    INSERT_LYRIC_SLIDE = 1
    INSERT_LYRIC_TEXT = 2
    INSERT_LYRIC_BOTH = 3

    def __init__(self, slide_insert_location, slide_separator_slides,
                 lyric_insert_location, lyric_separator_slides, lyric_pattern, filelist, flags=0):
        '''InsertLyrics has two operations into one class.
        The reason for having two operations in one class is to manage one filelist that can handle both operations.

        It inserts score slides from each file in filelist after slide_insert_location.
        slide_insert_location is the expression that the inserted slide will located.
        slide_separator_slides is separator slides between each inserted slides, and last one will not be inserted.

        It inserts lyrics from each file in filelist after lyric_insert_location.
        lyric_insert_location is the expression to the location of existing slide that will be duplicated.
        lyric_separator_slides is separator slides between each inserted slides, and last one will not be inserted.

        filelist is the ppt filename and for lyric file, the extension .xml will be used to get the lyrics.

        flags is an option how to handle both score and lyric slides.
        '''
        super().__init__()

        self.slide_insert_location = slide_insert_location
        self.slide_separator_slides = slide_separator_slides

        self.lyric_insert_location = lyric_insert_location
        self.lyric_separator_slides = lyric_separator_slides
        self.lyric_pattern = lyric_pattern

        self.filelist = filelist
        self.flags = flags

    def get_filelist(self, filetype):
        filelist = []
        if filetype == self.INSERT_LYRIC_SLIDE:
            for fn in self.filelist:
                base, ext = os.path.splitext(fn)
                ext = ext.lower()
                if ext == '.pptx' or ext == '.ppt':
                    pass
                else:
                    fn = base + '.pptx'
                    if os.path.exists(fn):
                        pass
                    else:
                        # do not check, so the execute() function can throw exception.
                        fn = base + '.ppt'
                filelist.append(fn)
        elif filetype == self.INSERT_LYRIC_TEXT:
            filelist = [os.path.splitext(fn)[0] + '.xml' for fn in self.filelist]

        return filelist

    def execute(self, cm, prs):
        if (self.flags & self.INSERT_LYRIC_SLIDE):
            lyric_filelist = self.get_filelist(self.INSERT_LYRIC_SLIDE)
            slides = InsertSlides(self.slide_insert_location, self.slide_separator_slides, lyric_filelist)
            slides.execute(cm, prs)
            del slides

        if (self.flags & self.INSERT_LYRIC_TEXT):
            lyric_filelist = self.get_filelist(self.INSERT_LYRIC_TEXT)
            self.execute_lyric_files(cm, prs, lyric_filelist)

    def execute_lyric_files(self, cm, prs, filelist):
        file_count = len(filelist)
        insert_format = ngettext('Inserting lyrics from {file_count} file.',
                                 'Inserting lyrics from {file_count} files.', file_count)
        cm.progress_message(0, insert_format.format(file_count=file_count))

        lyric_insert_location = evaluate_to_single_slide(prs, self.lyric_insert_location)
        if lyric_insert_location is None:
            cm.progress_message(0, _('No repeatable slides are found. Aborting the command.'))
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
            cm.progress_message(percent, _('Inserting lyric from \'{filename}\'.').format(filename=filelist[i]))

            lines = list(song.get_lines_by_order())
            for j, l in enumerate(lines):
                text_dict = {self.lyric_pattern: l}
                prs.replace_one_slide_texts(lyric_insert_location + j, text_dict)

            added_count = len(lines)
            lyric_insert_location = lyric_insert_location + added_count

            # Skip separator
            if separator_slide_count != 0 and i < last_index:
                added_count = separator_slide_count
                lyric_insert_location = lyric_insert_location + added_count

    def duplicate_slides(self, prs, source_location, lyric_separator_slides, separator_slide_count, songs):
        '''Duplicate slide based on count_of_lyric1, separators, count_of_lyric2, separators, ...
        songs is a two dimensional string array.
        '''
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


class FindReplaceText(Command):
    def __init__(self, texts):
        '''FindReplaceText finds all occurrence of find_text and replace it with replace_text,
        which are (k, v) pair in dictionary texts.
        '''
        super().__init__()

        self.texts = texts

    def execute(self, cm, prs):
        text_count = len(self.texts)
        replace_format = ngettext('Replacing {text_count} text.',
                                  'Replacing {text_count} texts.', text_count)
        cm.progress_message(0, replace_format.format(text_count=text_count))

        prs.replace_all_slides_texts(self.texts)

        _, cm.notes = replace_all_notes_text(cm.notes, self.texts)


class DuplicateWithText(Command):
    def __init__(self, slide_range, repeat_range, find_text, replace_texts):
        '''DuplicateWithText finds all occurrence of find_text and replace it replace_texts,
        which are list of text.
        Because replace_texts are a list, slides in repeat_range will be duplicated
        to match with len(replace_texts) == len(slide_range) + repeated len(repeat_range).
        '''
        super().__init__()

        self.slide_range = slide_range
        self.repeat_range = repeat_range
        self.find_text = find_text
        self.replace_texts = replace_texts

    def execute(self, cm, prs):
        cm.progress_message(0, _('Duplicating slides and replacing texts.'))

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

            # THe input template pptx contains slide_range[1] - slide_range[0] + 1 slides for the operation.
            # From that, repeat_range[1] - repeat_range[0] + 1 slides are repeatable and can be duplicated.
            # So, from texts, find out how many we need to duplicate and do the operation.
            # Then, replace texts to get the final slides.
            template_slide_count = slide_range[1] - slide_range[0] + 1
            repeatable_slide_count = repeat_range[1] - repeat_range[0] + 1
            total_slide_count = len(self.replace_texts)

            duplicate_count = (total_slide_count - template_slide_count)
            duplicate_count = int(math.ceil(float(duplicate_count)/repeatable_slide_count))
            source_index = list(range(repeat_range[0], repeat_range[1]+1))
            _slides = prs.duplicate_slides(source_index, repeat_range[0], duplicate_count)

            for i, text in enumerate(self.replace_texts):
                percent = 100 * i / total_slide_count
                cm.progress_message(percent, _('Replacing texts.'))

                text_dict = {self.find_text: text}
                prs.replace_one_slide_texts(slide_range[0] + i, text_dict)


class GenerateBibleVerse(Command):
    INDEX_TITLE_PATTERN = 0
    INDEX_BOOK_PATTERN = 1
    INDEX_SHORT_BOOK_PATTERN = 2
    INDEX_CHAPTER_PATTERN = 3
    INDEX_VERSE_NO_PATTERN = 4
    INDEX_VERSE_TEXT_PATTERN = 5

    def __init__(self, bible_format, bible_version, verse_patterns, main_verses, additional_verses, repeat_range):
        '''GenerateBibleVerse is a two operations combined.

        1. It will replace all verse_patterns[0] to main_verses in all slides.
        2. verses_text will be populated from main_verses and additional_verses.
        Then, the repeat_range slides will be duplicated to match with len(verses_text) == repeated len(repeat_range)
        and replace verse_patterns[1], verse_patterns[2] to pair (no, text) in verses_text.

        bible_format is one of supported bible file format: See also bible.fileformat.
        bible_version can be a valid Bible Version.
        verse_patterns is a list containing main_verse_title_pattern, book_pattern, book_short_pattern, chapter_pattern, verse_no_pattern, verse_text_pattern.
        main_verses is a comma separated main verse titles for the service.
        additional_verses is an additional comma separated verse titles for the service.
        repeat_range is a slide range mark to duplicate the slides.
        '''
        super().__init__()

        assert isinstance(verse_patterns, list) and len(verse_patterns) == (self.INDEX_VERSE_TEXT_PATTERN + 1)

        self.bible_format = bible_format
        self.bible_version = bible_version
        self.main_verses = main_verses
        self.additional_verses = additional_verses
        self.verse_patterns = verse_patterns
        self.repeat_range = repeat_range

        self._verses_text = []

    def populate_verses(self, fmt=None):
        self._verses_text = []

        bible = bibfileformat.read_version(self.bible_format, self.bible_version)
        if bible is None:
            return

        fmt1 = None
        fmt2 = '%b %c:%v'
        if fmt and isinstance(fmt, list) and len(fmt) == 2:
            fmt1 = fmt[0]
            fmt2 = fmt[1]

        all_verses_text = self.populate_one_verses(bible, self.main_verses, fmt1)
        all_verses_text = all_verses_text + self.populate_one_verses(bible, self.additional_verses, fmt2)

        self._verses_text = all_verses_text

    def populate_one_verses(self, bible, verses, fmt=None):
        all_verses_text = []
        if verses is not None:
            splitted_verses = verses.split(',')
            for verse in splitted_verses:
                verse = verse.strip()
                if not verse:
                    continue

                book, chapter, verses_text = bible.extract_texts(verse)
                if fmt:
                    bc_fmt = fmt.replace('%B', book.name)
                    bc_fmt = bc_fmt.replace('%b', book.short_name)
                    bc_fmt = bc_fmt.replace('%c', str(chapter.no))
                    for v in verses_text:
                        new_fmt = bc_fmt.replace('%v', str(v.no))
                        v.no = new_fmt

                all_verses_text = all_verses_text + verses_text

        return all_verses_text

    def execute(self, cm, prs):
        self.execute_on_slides(cm, prs)
        self.execute_on_notes(cm)

    def execute_on_slides(self, cm, prs):
        cm.progress_message(0, _('Processing Bible Verse.'))

        self.populate_verses()

        cm.set_bible_verse(self)

        text_dict = {self.verse_patterns[self.INDEX_TITLE_PATTERN]: self.main_verses}
        prs.replace_all_slides_texts(text_dict)

        repeat_range = evaluate_to_multiple_slide(prs, self.repeat_range)
        repeat_range = prs.slide_index_to_ID(repeat_range)

        repeat_count = len(repeat_range) * (len(self._verses_text)-1)
        if repeat_count > 0:
            duplicate_format = ngettext('Duplicating {repeat_count} slide for Bible Verse.',
                                        'Duplicating {repeat_count} slides for Bible Verse.', repeat_count)
            cm.progress_message(0, duplicate_format.format(repeat_count=repeat_count))

        repeat_count = len(repeat_range)
        for i, sid in enumerate(repeat_range):
            percent = 100 * i / repeat_count
            cm.progress_message(percent, _('Duplicating slides and replacing texts.'))

            index = prs.slide_ID_to_index(sid)
            _slides = prs.duplicate_slides(index, index, len(self._verses_text)-1)

            for i, v in enumerate(self._verses_text):
                text_dict = self.get_verse_dict(v)
                prs.replace_one_slide_texts(index + i, text_dict)

    def execute_on_notes(self, cm):
        if cm.notes:
            notes = cm.notes
            text_dict = {self.verse_patterns[self.INDEX_TITLE_PATTERN]: self.main_verses}
            _, notes = replace_all_notes_text(notes, text_dict)

            # check whether we need to duplicate the line by self.verse_patterns[INDEX_VERSE_NO_PATTERN] or self.verse_patterns[INDEX_VERSE_TEXT_PATTERN] exists.
            if self.verse_patterns[self.INDEX_VERSE_NO_PATTERN] in notes or self.verse_patterns[self.INDEX_VERSE_TEXT_PATTERN] in notes:
                lines = notes.split('\n')
                repeat_range = [i for i, l in enumerate(lines) if self.verse_patterns[self.INDEX_VERSE_NO_PATTERN] in l or
                                                                  self.verse_patterns[self.INDEX_VERSE_TEXT_PATTERN] in l]
                for index in reversed(repeat_range):
                    for _ in range(len(self._verses_text)-1):
                        lines.insert(index, lines[index])

                    for i, v in enumerate(self._verses_text):
                        text_dict = self.get_verse_dict(v)
                        _, lines[index+i] = replace_all_notes_text(lines[index+i], text_dict)

                notes = '\n'.join(lines)

            cm.notes = notes

    def get_verse_dict(self, v):
        text_dict = {}
        if self.verse_patterns[self.INDEX_BOOK_PATTERN]:
            text_dict[self.verse_patterns[self.INDEX_BOOK_PATTERN]] = v.book.name
        if self.verse_patterns[self.INDEX_SHORT_BOOK_PATTERN] and v.book.short_name:
            text_dict[self.verse_patterns[self.INDEX_SHORT_BOOK_PATTERN]] = v.book.short_name
        if self.verse_patterns[self.INDEX_CHAPTER_PATTERN]:
            text_dict[self.verse_patterns[self.INDEX_CHAPTER_PATTERN]] = str(v.chapter.no)
        if self.verse_patterns[self.INDEX_VERSE_NO_PATTERN]:
            text_dict[self.verse_patterns[self.INDEX_VERSE_NO_PATTERN]] = v.no
        if self.verse_patterns[self.INDEX_VERSE_TEXT_PATTERN]:
            text_dict[self.verse_patterns[self.INDEX_VERSE_TEXT_PATTERN]] = v.text

        return text_dict

    def get_flattened_dict(self):
        return {'bible_version': self.bible_version,
                'main_verses': self.main_verses,
                'additional_verses': self.additional_verses,
                'verse_patterns': self.verse_patterns,
                'repeat_range': self.repeat_range
                }


def rename_filename_to_zeropadded(dirname, num_digits):
    r'''Rename Slide(\d+).PNG to Slide%03d.png so that the length is same
    and they can be sorted properly.
    Powerpoint generates filename as 1-9 and 10, etc. that make sorting difficult.
    '''
    if num_digits <= 1:
        return

    def replace_format_3digits(m):
        if m.group(2).isdigit():
            num = int(m.group(2))
            fmt = r'%s%0' + str(num_digits) + r'd%s'

            ext = m.group(3)
            ext = ext.lower()
            if len(ext) == 5:
                if ext == '.jpeg':
                    ext = '.jpg'
                elif ext == '.tiff':
                    ext = '.tif'

            s = fmt % (m.group(1), num, ext)
        else:
            s = m.group(0)
        return s

    matching_fn_pattern = r'(.*?)(\d+)(\.(gif|jpg|jpeg|png|tif|tiff))'
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
    files = os.listdir(dirname)
    for fn in files:
        fullname = os.path.join(dirname, fn)
        if os.path.isdir(fullname):
            shutil.rmtree(fullname)
        else:
            os.remove(fullname)


class PromptCommand(Command):
    def __init__(self, message):
        super().__init__()

        self.message = message

    def execute(self, cm, prs):
        cm.progress_message(0, _('Processing PromptCommand.'))

        print(f'{self.message}')
        input()

        prs.check_modified()


class ExportSlides(Command):
    def __init__(self, slide_range, out_dirname, image_type, flags, color=None):
        super().__init__()

        self.slide_range = slide_range
        self.out_dirname = out_dirname
        self.image_type = image_type
        self.flags = flags
        self.color = color

    def execute(self, cm, prs):
        cm.progress_message(0, _('Exporting slide shapes as images to \'{dirname}\'.').format(dirname=self.out_dirname))

        if self.flags & Export_CleanupFiles:
            rmtree_except_self(self.out_dirname)

        if not hasattr(prs, 'export_slide_as'):
            self.export_all(cm, prs)
            return

        slide_range = evaluate_to_multiple_slide(prs, self.slide_range)
        if slide_range is None:
            return

        image_type = cm.normalize_image_type(self.image_type)
        prs.export_slides_as(slide_range, self.out_dirname, image_type)

        if (self.flags & Export_Transparent) and self.image_type == 'png' and self.color:
            cm.progress_message(50, _('Converting to transparent images.'))

            color = ImageColor.getrgb(self.color)
            for index in slide_range:
                f = cm.get_filename_from_slideno(self.image_type, index)
                filename = os.path.join(src_dirname, f)
                color_to_transparent(filename, filename, color)

    def export_all(self, cm, prs):
        src_dirname = cm.generate_image_files(self.image_type, self.color)

        slide_range = evaluate_to_multiple_slide(prs, self.slide_range)
        slide_range = get_contiguous_range(slide_range)
        if slide_range is None:
            return

        if (self.flags & Export_Transparent) and self.image_type == 'png' and self.color:
            cm.progress_message(50, _('Converting to transparent images.'))

            color = ImageColor.getrgb(self.color)
            for index in range(slide_range[0], slide_range[1]+1):
                f = cm.get_filename_from_slideno(self.image_type, index)
                filename = os.path.join(src_dirname, f)
                color_to_transparent(filename, filename, color)

        for index in range(slide_range[0], slide_range[1]+1):
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
        cm.progress_message(0, _('Exporting slide shapes as images to \'{dirname}\'.').format(dirname=self.out_dirname))

        if self.flags & Export_CleanupFiles:
            rmtree_except_self(self.out_dirname)

        slide_range = evaluate_to_multiple_slide(prs, self.slide_range)
        if slide_range is None:
            return

        prs.export_slide_shapes_as(slide_range, self.out_dirname, self.image_type)


class CommandManager:
    def __init__(self):
        self.running_already = False
        self.powerpoint = None
        self.prs = None

        self.notes = ''

        self.monitor = None
        self._continue = True

        self.image_dir = {}
        self.num_digits = 0

        self.bible_verse = None

    def __del__(self):
        self.close()

    def close(self):
        if self.prs:
            self.prs.close()
            self.prs = None

        if not self.running_already and self.powerpoint:
            self.powerpoint.quit_if_empty()

        self.powerpoint = None

        self.remove_cache()

    def set_presentation(self, prs):
        self.prs = prs

    def get_notes(self):
        return self.notes

    def set_notes(self, notes):
        self.notes = notes

    def set_bible_verse(self, bible_verse):
        self.bible_verse = bible_verse

    def read_songs(self, filelist):
        reader = OpenLyricsReader()
        songs = []
        for filename in filelist:
            if not os.path.exists(filename):
                self.error_message(_('Cannot open a lyric file \'{filename}\'.').format(filename=filename))
                raise FileNotFoundError(errno.ENOENT, os.strerror(errno.ENOENT), filename)

            try:
                song = reader.read_song(filename)
                songs.append(song)
            except Exception as e:
                self.error_message(_('Cannot open a lyric file \'{filename}\'.').format(filename=filename))
                raise

        return songs

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
        if image_type == 'jpeg':
            image_type = 'jpg'
        elif image_type == 'tiff':
            image_type = 'tif'
        return image_type

    def generate_image_files(self, image_type, color=None):
        if self.prs:
            slide_count = self.prs.slide_count()
            self.num_digits = len(f'{slide_count+1}')
        else:
            self.num_digits = 0

        image_type = self.normalize_image_type(image_type)

        if image_type not in self.image_dir:
            outdir = tempfile.mkdtemp(prefix='slides')

            self.prs.saveas_format(outdir, image_type)

            rename_filename_to_zeropadded(outdir, self.num_digits)

            self.image_dir[image_type] = outdir

        return self.image_dir[image_type]

    def get_filename_from_slideno(self, image_type, slideno):
        image_type = self.normalize_image_type(image_type)

        num = slideno + 1
        fmt = r'%0' + str(self.num_digits) + r'd'
        filename = 'Slide' + (fmt % num) + '.' + image_type

        return filename

    def execute_commands(self, instructions, monitor):
        '''execute_commands() is the main function calling each command's execute()
        to achieve the goal specificed in each command.
        '''
        self.monitor = monitor
        self._continue = True
        self.progress_message(0, _('Start processing commands.'))

        self.running_already = PowerPoint.App.is_running()
        self.powerpoint = PowerPoint.App()

        self.prs = None
        count = len(instructions)
        for i, bi in enumerate(instructions):
            if not self._continue:
                break

            if not bi.enabled:
                continue

            monitor.set_subrange(i * 100 / count, (i+1) * 100 / count)
            try:
                bi.execute(self, self.prs)
            except Exception as e:
                traceback.print_exc()
                self.error_message(str(e))

        self.progress_message(100, _('Cleaning up after running all the commands.'))
        self.close()

        self.progress_message(100, _('Processing is done successfully.'))
