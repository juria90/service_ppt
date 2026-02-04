#!/usr/bin/env python
"""Core hymn and lyric data structures.

This module defines the core data structures (Line, Verse, Song) for representing
hymns and lyrics, and provides functionality for reading and processing lyric files.
"""

from typing import TYPE_CHECKING

if TYPE_CHECKING:
    from collections.abc import Iterator


class Line:
    """Represents a single line of text in a verse.

    A line can have an optional break marker indicating it can be split
    across slides or pages.
    """

    def __init__(self, text: str, optional_break: bool = False) -> None:
        """Initialize a Line object.

        :param text: The text content of the line
        :param optional_break: Whether this line has an optional break marker
        """
        self.text: str = text
        self.optional_break: bool = optional_break


class Verse:
    """Represents a verse in a song.

    A verse contains multiple lines of text and may have a verse number
    or identifier.
    """

    def __init__(self) -> None:
        """Initialize a Verse object."""
        self.no: str | int | None = None
        self.lines: list[Line] = []


class Song:
    """Represents a complete song or hymn.

    A song contains metadata (title, authors, etc.) and multiple verses.
    Verses can be accessed in their original order or in a custom order
    specified by verse_order.
    """

    def __init__(self) -> None:
        """Initialize a Song object."""
        self.title: str | None = None
        """Song title."""
        self.authors: list[tuple[str, str]] | None = None
        """List containing tuples of ('words', 'music', 'translation/lang') and name."""
        self.verse_order: str | None = None
        """String that describes the order of verses."""
        self.songbook: dict[str, str] | None = None
        """Dict containing name (mandatory) and entry (optional)."""
        self.released: str | None = None
        """Release date in format Y, Y-M, Y-M-d, or Y-M-dTh:m."""
        self.keywords: str | None = None
        """Keywords associated with the song."""
        self.verses: list[Verse] = []
        """List of verses in the song."""

        self.ordered_verses: list[Verse] | None = None
        """Cached list of verses in the order specified by verse_order."""

    def get_verses_by_order(self) -> list[Verse]:
        """Get verses in the order specified by verse_order.

        If verse_order is set, returns verses in that order. Otherwise,
        returns verses in their original order. Results are cached for
        subsequent calls.

        :returns: List of Verse objects in the specified order
        """
        if isinstance(self.verse_order, str) and len(self.verse_order) > 0:
            if self.ordered_verses is None:
                order = self.verse_order.split()
                verse_map: dict[str | int, Verse] = {}
                for v in self.verses:
                    if v.no:
                        verse_map[v.no] = v

                verses: list[Verse] = []
                for name in order:
                    if name in verse_map:
                        verses.append(verse_map[name])

                self.ordered_verses = verses

            return self.ordered_verses
        return self.verses

    def get_lines_by_order(self) -> "Iterator[Line]":
        """Get all lines from all verses in order.

        Iterates through verses (in their specified order) and yields
        all lines from each verse.

        :returns: Iterator yielding Line objects from all verses in order
        """
        verses = self.get_verses_by_order()
        for v in verses:
            yield from v.lines


class Book:
    """Represents a collection of songs.

    A book contains a name and a list of songs, typically representing
    a hymnbook or song collection.
    """

    def __init__(self) -> None:
        """Initialize a Book object."""
        self.name: str | None = None
        self.songs: list[Song] = []
