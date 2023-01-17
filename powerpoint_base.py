import errno
import os
import re


def find_text_in_text_list(text, text_list, match_case=True, whole_words=True):
    if whole_words or not match_case:
        esc_text = re.escape(text)
        pattern = ""
        if not match_case:
            pattern = "(?i)"
        if whole_words:
            pattern = pattern + r"\b" + esc_text + r"\b"
        else:
            pattern = pattern + esc_text
        for t in text_list:
            if re.search(pattern, t):
                return True
    else:
        for t in text_list:
            index = t.find(text)
            if index != -1:
                return True

    return False


class FindAnyText:
    def __init__(self, find_text_list, match_case=True, whole_words=True):
        esc_text = ""
        for text in find_text_list:
            if esc_text:
                esc_text = esc_text + "|"
            esc_text = esc_text + re.escape(text)

        esc_text = "(" + esc_text + ")"

        pattern = ""
        if not match_case:
            pattern = "(?i)"
        if whole_words:
            pattern = pattern + r"\b" + esc_text + r"\b"
        else:
            pattern = pattern + esc_text

        self.pattern = pattern

    def find(self, text_list):
        for t in text_list:
            if re.search(self.pattern, t):
                return True

        return False


class SlideCache:
    def __init__(self):
        self.valid_cache = False
        self.id = 0
        self.slide_text = []
        self.notes_text = []


class PresentationBase:
    """PresentationBase is a common base class for Presentation(Win32) and Presentation(OSX).
    It also supports caching text for find and replace.
    """

    def __init__(self, app, prs):
        self.app = app
        self.prs = prs

        self._slide_count = self.slide_count()

        self.valid_cache = False
        self.slide_caches = []
        self.id_to_index = {}

    def reset(self):
        self.app = None
        self.prs = None

        self._slide_count = 0

        self.valid_cache = False
        self.slide_caches = []
        self.id_to_index = {}

    def close(self):
        if self.prs:
            self.prs.Close()
            self.prs = None

    def __enter__(self) -> "PresentationBase":
        return self

    def __exit__(self, exc_type, exc_value, traceback) -> None:
        self.close()

    def slide_count(self):
        return 0

    # slide_index_to_ID is platform dependent.
    # But, slide_ID_to_index is implemented using SlideCache

    def _slides_inserted(self, index, count):
        if index < len(self.slide_caches):
            for _ in range(count):
                self.slide_caches.insert(index, None)

        self._slide_count = self._slide_count + count
        self.valid_cache = False

    def _slides_deleted(self, index, count):
        if index < len(self.slide_caches):
            del self.slide_caches[index : index + count]

        self._slide_count = self._slide_count - count
        self.id_to_index = {}
        self.valid_cache = False

    def check_modified(self):
        slide_count = self.slide_count()
        if self._slide_count != slide_count or len(self.slide_caches) != slide_count:
            self._slide_count = slide_count
            self.valid_cache = False
            self.slide_caches = []  # invalid all caches
            self.id_to_index = {}

        self._update_slide_ID_cache()

    def _update_slide_ID_cache(self):
        if not self.valid_cache:
            id_to_index = {}
            for index, sc in enumerate(self.slide_caches):
                if sc is None or not sc.valid_cache:
                    sc = self._fetch_slide_cache(index)
                    self.slide_caches[index] = sc

                id_to_index[sc.id] = index

            for index in range(len(self.slide_caches), self._slide_count):
                sc = self._fetch_slide_cache(index)
                self.slide_caches.append(sc)

                id_to_index[sc.id] = index

            self.id_to_index = id_to_index
            self.valid_cache = True

    def _fetch_slide_cache(self, _i):
        return SlideCache()

    def slide_ID_to_index(self, var):
        self._update_slide_ID_cache()

        if isinstance(var, int):
            slide_index = self.id_to_index[var]
            return slide_index
        elif isinstance(var, list):
            result = []
            for sid in var:
                slide_index = self.id_to_index[sid]
                result.append(slide_index)

            return result

    def get_text_in_slide(self, slide_index, is_note_shape):
        self._update_slide_ID_cache()

        sc = self.slide_caches[slide_index]
        if not is_note_shape:
            text_list = sc.slide_text
        else:
            text_list = sc.notes_text

        return text_list

    def get_text_in_all_slides(self, is_note_shape):
        text_list = []
        for slide_index in range(self.slide_count()):
            text_list.append(self.get_text_in_slide(slide_index, is_note_shape))

        return text_list

    def find_text_in_slide(self, slide_index, is_note_shape, text, ignore_case=False, whole_words=False):
        self._update_slide_ID_cache()

        sc = self.slide_caches[slide_index]
        if not is_note_shape:
            text_list = sc.slide_text
        else:
            text_list = sc.notes_text

        found = find_text_in_text_list(text, text_list, ignore_case, whole_words)
        return found

    def replace_all_slides_texts(self, find_replace_dict):
        has_replace_all_slides = False
        if callable(getattr(self, "_replace_all_slides_texts", None)):
            has_replace_all_slides = True

        self._update_slide_ID_cache()

        ft = FindAnyText([t for t in find_replace_dict], True, False)

        for slide_index in range(self.slide_count()):
            sc = self.slide_caches[slide_index]
            if ft.find(sc.slide_text):
                if not has_replace_all_slides:
                    self._replace_texts_in_slide_shapes(slide_index, find_replace_dict)
                sc.valid_cache = False
                self.valid_cache = False

        if has_replace_all_slides:
            self._replace_all_slides_texts(find_replace_dict)

    def replace_one_slide_texts(self, slide_index, find_replace_dict):
        self._update_slide_ID_cache()

        ft = FindAnyText([t for t in find_replace_dict], True, False)

        sc = self.slide_caches[slide_index]
        if ft.find(sc.slide_text):
            self._replace_texts_in_slide_shapes(slide_index, find_replace_dict)
            sc.valid_cache = False
            self.valid_cache = False

    def _replace_texts_in_slide_shapes(self, _slide_index, _find_replace_dict):
        pass

    def insert_blank_slide(self, slide_index):
        pass

    def _paste_keep_source_formatting(self, insert_location):
        pass

    def delete_slide(self, slide_index):
        pass

    def insert_file_slides(self, insert_location, src_ppt_filename):

        if insert_location is None:
            insert_location = self.slide_count()
        elif isinstance(insert_location, int):
            insert_location = insert_location + 1
        else:
            raise ValueError("Invalid insert location")

        if not os.path.exists(src_ppt_filename):
            raise FileNotFoundError(errno.ENOENT, os.strerror(errno.ENOENT), src_ppt_filename)

        # If the presentation is empty, Activate() will not work.
        # So, add a blank slide and delete it later.
        empty_slide_added = False
        slide_no = 0
        if self.slide_count() == 0:
            self.insert_blank_slide(slide_no)
            empty_slide_added = True

        src_prs = self.app.open_presentation(src_ppt_filename)
        added_count = src_prs.slide_count()
        src_prs.copy_all_and_close()
        src_prs = None

        self._paste_keep_source_formatting(insert_location)

        if empty_slide_added:
            self.delete_slide(slide_no)

        self._slides_inserted(insert_location - 1, added_count)

        return added_count

    def export_slides_as(self, slidenumbers, dirname, image_type="png"):
        slide_count = self._slide_count
        num_digits = len(f"{slide_count+1}")
        fmt = r"Slide%0" + str(num_digits) + r"d%s"

        ext = "." + image_type

        if slidenumbers is None:
            slidenumbers = range(self._slide_count)

        for slideno in slidenumbers:
            filename = fmt % (slideno + 1, ext)
            self.export_slide_as(slideno, os.path.join(dirname, filename), image_type)

    def export_slide_as(self, slideno, filename, image_type="png"):
        """Save all shapes in slide as image files."""

    def export_slide_shapes_as(self, slidenumbers, dirname, image_type="png"):
        slide_count = self._slide_count
        num_digits = len(f"{slide_count+1}")
        fmt = r"Slide%0" + str(num_digits) + r"d%s"

        ext = "." + image_type

        if slidenumbers is None:
            slidenumbers = range(self._slide_count)

        for slideno in slidenumbers:
            filename = fmt % (slideno + 1, ext)
            self.export_shapes_as(slideno, os.path.join(dirname, filename), image_type)

    def export_shapes_as(self, slideno, filename, image_type="png"):
        """Save all shapes in slide as image files."""


class PPTAppBase:
    def __init__(self):
        pass
