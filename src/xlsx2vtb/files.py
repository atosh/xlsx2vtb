# coding: utf-8

import csv
import os


_pwd = os.getcwd().replace('\\', '/')

class CSV(object):
    def __init__(self, name, rows):
        self.name = name
        self.rows = [[datum.encode('utf-8') for datum in row] for row in rows]

    def write(self, out=None):
        if not out:
            out = open('%s/%s.csv' % (_pwd, self.name), 'wb')
        writer = csv.writer(out)
        writer.writerows(self.rows)

    def pprint(self):
        import sys; self.write(sys.stdout)


class VTB(object):
    def __init__(self, name, fields):
        self.name = name
        self.fields = [field.encode('utf-8') for field in fields]

    def write(self, out=None):
        from vtbtemplate import header_tmpl, field_tmpl
        if not out:
            out = open('%s/%s.vtb' % (_pwd, self.name), 'wb')
        out.write(header_tmpl % ('%s/%s.csv' % (_pwd, self.name.encode('utf-8'))))
        for field in self.fields:
            out.write(field_tmpl % field)

    def pprint(self):
        import sys; self.write(sys.stdout)

