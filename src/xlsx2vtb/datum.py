# coding: utf-8

class Datum(object):
    def __init__(self, pos, fmt, typ, value):
        self.pos = pos
        self.fmt = int(fmt) if fmt else fmt
        self.typ = typ
        self.value = value

