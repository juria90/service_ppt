"""HTML Bible format reader.

This module provides support for reading Bible text from HTML format files,
including parsing HTML tags and extracting verse content.
"""

import html
import os
import re

from service_ppt.bible import biblang


class HTMLWriter:
    lang = biblang.LANG_EN
    charset = "utf-8"

    def __init__(self, dynamic_page, process_bible_tags):
        super().__init__()
        self.dynamic_page = dynamic_page
        self.process_bible_tags = process_bible_tags

    def _get_extension(self):
        return ".html"

    def write_bible(self, dirname, bible, encoding="utf-8"):
        self.lang = bible.lang

        extension = self._get_extension()
        for i, b in enumerate(bible.books):
            bible.ensure_loaded(b)

            filename = f"book{i + 1}{extension}"
            with open(os.path.join(dirname, filename), "w", encoding=encoding) as file:
                self._write_book(file, b)

        filename = f"index{extension}"
        with open(os.path.join(dirname, filename), "w", encoding=encoding) as file:
            self._write_index(file, bible)

    def _write_html_full_header(self, file, title):
        print(
            f"""<html>
<head>
<meta http-equiv=\"Content-Language\" content=\"{self.lang}\">
<meta http-equiv=\"Content-Type\" content=\"text/html; charset={self.charset}\">
<title>{title}</title>
<link rel=\"stylesheet\" href=\"bible.css\" type=\"text/css\">
""",
            file=file,
            end="",
        )

        if self.dynamic_page:
            print(
                """<script type="text/javascript">
function changeBible(obj) {
  var bookno = obj.value;
  load_book(bookno);
}

function load_book(bookno) {
  filename = "book" + bookno + ".html";
  // console.log('here: ' + filename);
  document.getElementById("bibleContent").innerHTML='<object type="text/html" data="' + filename + '" ></object>';
}
</script>
""",
                file=file,
                end="",
            )

        onload = ""
        if self.dynamic_page:
            onload = ''' onload="load_book('1')"'''

        print(
            f"""</head>
<body bgcolor=\"#FFFFFF\" text=\"#000000\"{onload}>
""",
            file=file,
            end="",
        )

    def _write_html_footer(self, file):
        print(
            """</body>
</html>
""",
            file=file,
            end="",
        )

    def _write_html_simple_header(self, file):
        print(
            f"""<html>
<head>
<meta http-equiv=\"Content-Language\" content=\"{self.lang}\">
<meta http-equiv=\"Content-Type\" content=\"text/html; charset={self.charset}\">
</head>
<body bgcolor=\"#FFFFFF\" text=\"#000000\">
""",
            file=file,
            end="",
        )

    def _write_html_footer(self, file):
        print(
            """</body>
</html>
""",
            file=file,
            end="",
        )

    def _write_book(self, file, book):
        self._write_book_begin(file, book)

        for chapter in book.chapters:
            self._write_chapter(file, book, chapter)

        self._write_book_end(file, book)

    def _write_book_begin(self, file, book):
        if self.dynamic_page:
            self._write_html_simple_header(file)
        else:
            self._write_html_full_header(file, book.name)

        self._write_book_toc(file, book)

    def _write_book_end(self, file, _book):
        self._write_html_footer(file)

    def _write_book_toc(self, file, book):
        print(f'<h1 align="center">{book.name}</h1>', file=file)

        print(' <table border="0" width="100%%" cellpadding="0" cellspacing="0">', file=file)

        for i in range(0, len(book.chapters), 2):
            print("  <tr>", file=file)

            chapter = book.chapters[i]
            bc_name = html.escape(biblang.L18N.get_book_chapter_name(self.lang, book.name, chapter.no))
            url = f'<a href="#chapter{chapter.no}">{bc_name}</a>'
            print(f'   <td width="50%%">{url}</td>', file=file)

            if (i + 1) < len(book.chapters):
                chapter = book.chapters[i + 1]
                bc_name = html.escape(biblang.L18N.get_book_chapter_name(self.lang, book.name, chapter.no))
                url = f'<a href="#chapter{chapter.no}">{bc_name}</a>'
            else:
                url = ""
            print(f'   <td width="50%%">{url}</td>', file=file)

            print("  </tr>", file=file)

        print(" </table>", file=file)

    def _write_chapter(self, file, book, chapter):
        self._write_chapter_begin(file, book, chapter)

        for verse in chapter.verses:
            self._write_verse(file, verse)

        self._write_chapter_end(file, book, chapter)

    def _write_chapter_begin(self, file, book, chapter):
        bc_name = html.escape(biblang.L18N.get_book_chapter_name(self.lang, book.name, chapter.no))
        print(f'<h2 align="center"><a name="chapter{chapter.no}">{bc_name}</a></h2>', file=file)

    def _write_chapter_end(self, file, book, chapter):
        pass

    def _process_tags(self, text):
        if "<" not in text:
            return text

        # FO/Fo : OT Quote => Ignore.
        # PF#, PI# : First line indent, indent
        # RF/Rf : Translators' notes
        # TS/Ts : Title
        text = re.sub("<CI>", " ", text)
        text = re.sub("<RF>.*<Rf>", "", text)
        text = re.sub("<TS>.*<Ts>", "", text)
        text = re.sub("<(CM|FO|Fo|PF[0-7]|PI[0-7])>", "", text)

        # FI/Fi : Italic
        # FU/Fu : underlined words
        # FR/Fr : words of Jesus in Red
        text = re.sub("<FI>", "<i>", text)
        text = re.sub("<Fi>", "</i>", text)
        text = re.sub("<FU>", "<u>", text)
        text = re.sub("<Fu>", "</u>", text)
        text = re.sub("<FR>", '<span style="color:red;">', text)
        text = re.sub("<Fr>", "</span>", text)

        return text

    def _write_verse(self, file, verse):
        text = verse.text
        if self.process_bible_tags:
            text = self._process_tags(text)
        else:
            text = html.escape(verse.text)

        if verse.no:
            print(f"<sup><font color=red>{verse.no}&nbsp;</font></sup>{text}<br>", file=file)
        else:
            print(f"<font color=blue>{text}</font><br>", file=file)

    def _write_index(self, file, bible):
        """_write_index writes two column table for old and new testament."""

        self._write_html_full_header(file, bible.name)

        if self.dynamic_page:
            self._write_index_dynamic_form(file, bible)
        else:
            self._write_index_toc(file, bible)

        self._write_html_footer(file)

    def _write_index_toc(self, file, bible):
        print('<table border="0" width="100%%" cellpadding="0" cellspacing="0">', file=file)
        print("  <tr>", file=file)
        testament_name = biblang.L18N.get_testament_name(False, self.lang)
        print(f'    <td width="50%%">{testament_name}</td>', file=file)
        testament_name = biblang.L18N.get_testament_name(True, self.lang)
        print(f'    <td width="50%%">{testament_name}</td>', file=file)
        print("  </tr>", file=file)

        extension = self._get_extension()

        # assume old testament comes before new testament.
        old = [b for b in bible.books if b.new_testament == False]
        new = [b for b in bible.books if b.new_testament != False]
        max_row = max(len(old), len(new))
        for i in range(max_row):
            print("  <tr>", file=file)

            # left column
            if i < len(old):
                index = i + 1
                filename = f"book{index}{extension}"
                bookname = html.escape(old[i].name)
                print(f'    <td width="50%%"><a href="{filename}">{bookname}</a></td>', file=file)
            else:
                print('    <td width="50%%"></td>', file=file)

            if i < len(new):
                index = len(old) + i + 1
                filename = f"book{index}{extension}"
                bookname = html.escape(new[i].name)

                print(f'    <td width="50%%"><a href="{filename}">{bookname}</a></td>', file=file)
            else:
                print('    <td width="50%%"></td>', file=file)

            print("  </tr>", file=file)

        print("</table>", file=file)

    def _write_index_dynamic_form(self, file, bible):
        print("""<select name="book" id="book" onchange="changeBible(this)">""", file=file)

        for i, book in enumerate(bible.books):
            book_no = i + 1
            print(f"""<option value="{book_no}">{book.name}</option>""", file=file)

        print(
            """</select>
<br>

<div id='bibleContent'>
</div>

""",
            file=file,
            end="",
        )
