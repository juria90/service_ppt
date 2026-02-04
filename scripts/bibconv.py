#!/usr/bin/env python
"""Bible format conversion utility.

This module provides functionality to convert Bible text between different formats
supported by the service_ppt application.
"""

import argparse

from service_ppt.bible import bibleformat, csv_bible, html_bible, mybible, opensong_bible, zefania_bible


def write_bible(bible, out_format, filename, encoding, dynamic_page, remove_bible_tags):
    writer = None
    if out_format == "csv":
        writer = csv_bible.CSVWriter()
    elif out_format == "html":
        dynamic_page = False
        process_bible_tags = not remove_bible_tags
        writer = html_bible.HTMLWriter(dynamic_page, process_bible_tags)
    elif out_format == "MyBible":
        writer = mybible.MyBibleWriter()
    elif out_format == "opensong-xml":
        writer = opensong_bible.OpenSongXMLWriter()
    elif out_format == "zefania":
        writer = zefania_bible.ZefaniaWriter()

    if writer:
        writer.write_bible(filename, bible, encoding=encoding)


def str_to_bool(value):
    if value.lower() in {"false", "f", "0", "no", "n"}:
        return False
    elif value.lower() in {"true", "t", "1", "yes", "y"}:
        return True
    raise ValueError(f"{value} is not a valid boolean value")


def parse_cmdline():
    parser = argparse.ArgumentParser(description="Convert Bible text to various format.")
    parser.add_argument("--in-format", choices=["MyBible", "MySword", "Sword", "Zefania"], required=True, help="Valid input format.")
    parser.add_argument("in_version", nargs=1, help="Input Bible Version.")
    parser.add_argument(
        "--dynamic-page",
        type=str_to_bool,
        nargs="?",
        const=True,
        default=False,
        help="Generate dynamic index.html page.",
    )
    parser.add_argument("--remove-special-chars", action="store_true", help="Remove special chars in MyBible.")
    parser.add_argument("--remove-bible-tags", type=str_to_bool, nargs="?", const=True, help="Remove Bible Tags in MySword.")
    parser.add_argument(
        "--out-format",
        choices=["csv", "html", "MyBible", "opensong-xml", "zefania"],
        required=True,
        help="Valid output format.",
    )
    parser.add_argument("--out-encoding", help="Encoding for the output file.")
    parser.add_argument("out_filename", nargs=1, help="Output filename or directory name.")

    return parser


if __name__ == "__main__":
    parser = parse_cmdline()
    # args = parser.parse_args()
    args = parser.parse_args(
        [
            "--in-format",
            "MyBible",
            "개역개정",
            # "--in-format", "MyBible", "NIV",
            # '--in-format', 'MyBible', 'ESV',
            # '--in-format', 'Sword', 'GerLut1545',
            # '--in-format', 'Zefania', 'Spanish Reina-Valera',
            # '--remove-special-chars',
            "--remove-bible-tags",
            "false",
            # '--out-format', 'csv', '--out-encoding', 'utf-8', r'C:\Users\juria\Church\bible-output',
            # "--out-format", "html", "--out-encoding", "utf-8", r"C:\Users\juria\Church\Bible.html\NIV",
            # '--out-format', 'MyBible', '--out-encoding', 'utf-8', r'C:\Users\juria\Church\Bible.text\ESV',
            # '--out-format', 'opensong-xml', r'C:\Users\juria\Church\bible-output',
            "--out-format",
            "zefania",
            "--out-encoding",
            "utf-8",
            r"C:\Users\juria\Church\bible-output",
        ]
    )

    bibleformat.set_format_option(bibleformat.BibleFormat.MYBIBLE.value, "remove_special_chars", args.remove_special_chars)
    if args.remove_bible_tags:
        bibleformat.set_format_option(bibleformat.BibleFormat.MYSWORD.value, "remove_bible_tags", args.remove_bible_tags)
    elif args.out_format != "html":
        bibleformat.set_format_option(bibleformat.BibleFormat.MYSWORD.value, "remove_bible_tags", True)

    bible = bibleformat.read_version(args.in_format, args.in_version[0])
    write_bible(bible, args.out_format, args.out_filename[0], args.out_encoding, args.dynamic_page, args.remove_bible_tags)
