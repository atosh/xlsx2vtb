# coding: utf-8
import sys

class Datum(object):
    def __init__(self, pos, fmt, typ, value):
        self.pos = pos
        self.fmt = int(fmt) if fmt else fmt
        self.typ = typ
        self.value = value


import csv
import os


_pwd = os.getcwd().replace('\\', '/')

header_tmpl = '''// 
// Virtual table file.
// Version 4.0.01
// 

// Type.
ConnectType = "CSV"

// Encoding.
Encoding = "UTF_8"

// DataSource.
DataSourceName = "%s"
UserID = ""
PassWord = ""

// Select statement.
SQL_SELECT = ""
SQL_FROM = ""
SQL_WHERE = ""
SQL_OTHER = ""

// Date-Time format.
DateType = "0"
DateDelimit = "/"
TimeType = "0"
TimeDelimit = ":"

// CSV options.
Begin CSVOption
	Delimit = ","
	ColNameHeader = "-1"
	CharacterEnclose = """
End CSVOption

// FetchSize.
FetchSize = "0"
// AutoCommit.
AutoCommit = "0"
'''

field_tmpl = '''
Begin VirtualField
	VirtualFieldName = "%s"
	RealFieldName = ""
	Type = "VARCHAR"
	Precision = "0"
	Scale = "0"
	Position = "%d"
	Expression = ""
End VirtualField
'''


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
        self.write(sys.stdout)


class VTB(object):
    def __init__(self, name, fields):
        self.name = name
        self.fields = [field.encode('utf-8') for field in fields]

    def write(self, out=None):
        if not out:
            out = open('%s/%s.vtb' % (_pwd, self.name), 'wb')
        out.write(header_tmpl % ('%s/%s.csv' % (_pwd, self.name.encode('utf-8'))))
        for i in range(len(self.fields)):
            field = self.fields[i]
            out.write(field_tmpl % (field, i + 1))

    def pprint(self):
        self.write(sys.stdout)


class Row(object):
    def __init__(self, rowid, spans, data=None):
        self.id = int(rowid) if rowid else rowid
        self.spans = spans
        self.data = data if data else []

    def add(self, datum):
        self.data.append(datum)


class Sheet(object):
    def __init__(self, sheetid, name, rows=None):
        self.id = int(sheetid) if sheetid else sheetid
        self.name = name
        self.rows = rows if rows else []

    def add(self, row):
        self.rows.append(row)


from datetime import datetime, timedelta


class Style(object):
    ORIGIN = datetime(1899, 12, 30)

    def __init__(self, applyNumFmt, numFmtId):
        self.applyNumFmt = bool(int(applyNumFmt)) if applyNumFmt else False
        self.numFmtId = int(numFmtId) if numFmtId else numFmtId

    @classmethod
    def format(self, value, numFmtId):
        _format = self.get_formatter(numFmtId)
        return _format(value)

    @classmethod
    def get_formatter(self, numFmtId):
        if 12 <= numFmtId <= 17: return self.format_as_date
        if 18 <= numFmtId <= 21: return self.format_as_time
        if numFmtId == 22: return self.format_as_datetime
        return lambda x: x

    @classmethod
    def format_as_date(self, value):
        d = self.ORIGIN + timedelta(int(value))
        return '%04d/%02d/%02d' % (d.year, d.month, d.day)

    @classmethod
    def format_as_time(self, value):
        from math import floor
        totalsec = float(value) * 86400.0
        hour = floor(totalsec / 3600.0)
        rem = totalsec % 3600.0
        minute = floor(rem / 60.0)
        rem = rem % 60.0
        second = floor(rem)
        return '%02d:%02d:%02d' % (hour, minute, second)

    @classmethod
    def format_as_datetime(self, value):
        dt = self.ORIGIN + timedelta(float(value))
        return '%04d/%02d/%02d %02d:%02d:%02d' % (
            dt.year, dt.month, dt.day, dt.hour, dt.minute, dt.second)


import re


def xml2dict(elem):
    ret = {
        'tag': re.sub(r'^{.+}', '', elem.tag),
        'text': elem.text,
        'tail': elem.tail,
        'children': [xml2dict(child) for child in elem.getchildren()],
    }
    ret.update({re.sub(r'^{.+}', '', k): v for k, v in elem.attrib.items()})
    return ret


from zipfile import ZipFile
from xml.etree import ElementTree


class WorkBook(object):
    namespaces = {
        'main': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main', 
        'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships', 
        'x15': 'http://schemas.microsoft.com/office/spreadsheetml/2010/11/main', 
        'mc': 'http://schemas.openxmlformats.org/markup-compatibility/2006',
    }

    def __init__(self, sheets, styles, sharedstrs):
        self.sheets = sheets
        self.styles = styles
        self.sharedstrs = sharedstrs

    @classmethod
    def fromfile(self, filename):
        return self._fromfile(filename, self.namespaces)

    @classmethod
    def _fromfile(self, filename, namespaces):
        xlsxfile = ZipFile(filename, 'r')
        sheets = self._sheets(xlsxfile, namespaces)
        sharedstrs = self._sharedstrs(xlsxfile, namespaces)
        styles = self._styles(xlsxfile, namespaces)
        return WorkBook(sheets, styles, sharedstrs)

    @classmethod
    def _sheets(self, xlsxfile, namespaces):
        sheet_dics = [xml2dict(sheet_elem)
            for sheet_elem in ElementTree.fromstring(
                xlsxfile.read('xl/workbook.xml')
            ).findall('.//main:sheet', namespaces)]
        ret_sheets = [Sheet(sheet.get('sheetId'), sheet.get('name')) for sheet in sheet_dics]

        rows_dics = [[xml2dict(elem) for elem in ElementTree.fromstring(
                xlsxfile.read('xl/worksheets/sheet%s.xml' % sheet.id)
            ).findall('.//main:row', namespaces)]
            for sheet in ret_sheets]
        rowsgroup = [[Row(row.get('r'), row.get('spans'), 
            [Datum(datum.get('r'), datum.get('s'), datum.get('t'), datum.get('children')[0].get('text'))
                for datum in row.get('children')]) for row in rows_dic]
            for rows_dic in rows_dics]

        for i in range(len(ret_sheets)):
            ret_sheets[i].rows = rowsgroup[i]

        return ret_sheets

    @classmethod
    def _sharedstrs(self, xlsxfile, namespaces):
        sharedstr_dics = [xml2dict(sharedstr_elem)
            for sharedstr_elem in ElementTree.fromstring(
                xlsxfile.read('xl/sharedStrings.xml')
            ).findall('.//main:t', namespaces)]
        return [sharedstr.get('text') for sharedstr in sharedstr_dics]

    @classmethod
    def _styles(self, xlsxfile, namespaces):
        style_dics = [xml2dict(style_elem)
            for style_elem in ElementTree.fromstring(
                xlsxfile.read('xl/styles.xml')
            ).findall('.//main:cellXfs/main:xf', namespaces)]
        return [Style(style_dic.get('applyNumberFormat'), style_dic.get('numFmtId'))
            for style_dic in style_dics]

    def csvlist(self):
        csvs = []
        for sheet in self.sheets:
            ret_name = sheet.name
            ret_rows = []
            for row in sheet.rows:
                ret_row = []
                for datum in row.data:
                    value = datum.value
                    if datum.typ == 's':
                        value = self.sharedstrs[int(value)]
                    if datum.fmt:
                        value = Style.format(value, self.styles[datum.fmt].numFmtId)
                    ret_row.append(value)
                ret_rows.append(ret_row)
            csvs.append(CSV(ret_name, ret_rows))
        return csvs

    def vtblist(self):
        vtbs = []
        for sheet in self.sheets:
            ret_name = sheet.name
            ret_fields = []
            header_row = sheet.rows[0]
            for datum in header_row.data:
                value = datum.value
                if datum.typ == 's':
                    value = self.sharedstrs[int(value)]
                if datum.fmt:
                    value = Style.format(value, self.styles[datum.fmt].numFmtId)
                ret_fields.append(value)
            vtbs.append(VTB(ret_name, ret_fields))
        return vtbs


def main(argv):
    if len(argv) == 1:
        msg = 'usage: python xlsx2vtb.py <xlsx_filename>\n'
        sys.stderr.write(msg)
        return 1

    filename = argv[1]
    workbook = WorkBook.fromfile(filename)
    csvs = workbook.csvlist()
    for csv in csvs:
        csv.write()
    vtbs = workbook.vtblist()
    for vtb in vtbs:
        vtb.write()


if __name__ == '__main__':
    argv = sys.argv
    exit = main(argv)
    if exit:
        sys.exit(exit)
