# coding: utf-8
from row import Row

class Sheet(object):
    def __init__(self, sheetid, name, rows=None):
        self.id = int(sheetid) if sheetid else sheetid
        self.name = name
        self.rows = rows if rows else []

    def add(self, row):
        if not isinstance(row, Row):
            raise ValueError('Inalid data type: %s' % repr(row))
        self.rows.append(row)

