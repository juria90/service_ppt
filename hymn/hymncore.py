#!/usr/bin/env python
'''
'''

class Verse:
    def __init__(self):
        self.no = None
        self.lines = []


class Song:
    def __init__(self):
        self.title = None
        self.authors = None # list containing 'words', 'music', 'translation/lang' and name.
        self.verse_order = None # string that describes the order of verses.
        self.songbook = None # dict containing name(mandatory), entry(optional)
        self.released = None # Release Y, Y-M, Y-M-d, Y-M-dTh:m
        self.keywords = None
        self.verses = []

        self.ordered_verses = None

    def get_verses_by_order(self):
        if isinstance(self.verse_order, str) and len(self.verse_order) > 0:
            if self.ordered_verses is None:
                order = self.verse_order.split()
                verse_map = {}
                for v in self.verses:
                    if v.no:
                        verse_map[v.no] = v

                verses = []
                for name in order:
                    if name in verse_map:
                        verses.append(verse_map[name])

                self.ordered_verses = verses

            return self.ordered_verses
        else:
            return self.verses

    def get_lines_by_order(self):
        verses = self.get_verses_by_order()
        for v in verses:
            for l in v.lines:
                yield l


class Book:
    def __init__(self):
        self.name = None
        self.songs = []
