# coding: utf-8
from datum import Datum

class Row(object):
    def __init__(self, rowid, spans, data=None):
        self.id = int(rowid) if rowid else rowid
        self.spans = spans
        self.data = data if data else []

    def add(self, datum):
        if isinstance(datum, Datum):
            raise ValueError('Inalid data type: %s' % repr(datum))
        self.data.append(datum)

