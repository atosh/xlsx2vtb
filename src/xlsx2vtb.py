# coding: utf-8
from zipfile import ZipFile
from xml.etree import ElementTree
import re


_namespaces = {
    'main': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main', 
    'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships', 
    'x15': 'http://schemas.microsoft.com/office/spreadsheetml/2010/11/main', 
    'mc': 'http://schemas.openxmlformats.org/markup-compatibility/2006',
}


def xml2dict(elem):
    ret = {
        'tag': re.sub(r'^{.+}', '', elem.tag),
        'text': elem.text,
        'tail': elem.tail,
        'children': [xml2dict(child) for child in elem.getchildren()],
    }
    ret.update({re.sub(r'^{.+}', '', k): v for k, v in elem.attrib.items()})
    return ret


class Datum(object):
    def __init__(self, pos, fmt, typ, value):
        self.pos = pos
        self.fmt = int(fmt) if fmt else fmt
        self.typ = typ
        self.value = value


class Row(object):
    def __init__(self, id, spans, data=[]):
        self.id = int(id) if id else id
        self.spans = spans
        self.data = data

    def add(self, datum):
        if isinstance(datum, Datum):
            raise ValueError('Inalid data type: %s' % repr(datum))
        self.data.append(datum)


class Sheet(object):
    def __init__(self, id, name, rows=[]):
        self.id = int(id) if id else id
        self.name = name
        self.rows = rows

    def add(self, row):
        if not isinstance(row, Row):
            raise ValueError('Inalid data type: %s' % repr(row))
        self.rows.append(row)


from datetime import datetime, date, time, timedelta
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


class WorkBook(object):
    def __init__(self, sheets, styles, sharedstrs):
        self.sheets = sheets
        self.styles = styles
        self.sharedstrs = sharedstrs

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

import csv, os
_pwd = os.getcwd().replace('\\', '/')
class CSV(object):
    def __init__(self, name, rows):
        self.name = name
        self.rows = [[datum.encode('utf-8') for datum in row] for row in rows]

    def write(self, out=None):
        if not out:
            out = open('%s/%s.csv' % (_pwd, self.name), 'w')
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
            out = open('%s/%s.vtb' % (_pwd, self.name), 'w')
        out.write(header_tmpl % ('%s/%s.csv' % (_pwd, self.name.encode('utf-8'))))
        for field in self.fields:
            out.write(field_tmpl % field)

    def pprint(self):
        import sys; self.write(sys.stdout)

from StringIO import StringIO
def main(filename):
    """
    xlsxファイル読み込み
    workbook読み取り
    sheet読み取り
    style読み取り
    sharedstring読み取り
    適当にjoin


    シートごとにcsv, vtbファイルを生成
    """
    namespaces = _namespaces
    xlsx = ZipFile(filename, 'r')
    sheet_dics = [xml2dict(sheet_elem)
        for sheet_elem in ElementTree.fromstring(
            xlsx.read('xl/workbook.xml')
        ).findall('.//main:sheet', namespaces)]
    sheets = [Sheet(sheet.get('sheetId'), sheet.get('name')) for sheet in sheet_dics]

    rows_dics = [[xml2dict(elem) for elem in ElementTree.fromstring(
            xlsx.read('xl/worksheets/sheet%s.xml' % sheet.id)
        ).findall('.//main:row', namespaces)]
        for sheet in sheets]
    rowsgroup = [[Row(row.get('r'), row.get('spans'), 
        [Datum(datum.get('r'), datum.get('s'), datum.get('t'), datum.get('children')[0].get('text'))
            for datum in row.get('children')]) for row in rows_dic]
        for rows_dic in rows_dics]

    for i in range(len(sheets)):
        sheets[i].rows = rowsgroup[i]

    sharedstr_dics = [xml2dict(sharedstr_elem)
        for sharedstr_elem in ElementTree.fromstring(
            xlsx.read('xl/sharedStrings.xml')
        ).findall('.//main:t', namespaces)]
    sharedstrs = [sharedstr.get('text') for sharedstr in sharedstr_dics]

    style_dics = [xml2dict(style_elem)
        for style_elem in ElementTree.fromstring(
            xlsx.read('xl/styles.xml')
        ).findall('.//main:cellXfs/main:xf', namespaces)]
    styles = [Style(style_dic.get('applyNumberFormat'), style_dic.get('numFmtId'))
        for style_dic in style_dics]

    workbook = WorkBook(sheets, styles, sharedstrs)
    csvs = workbook.csvlist()
    for csv in csvs:
        csv.write()
    vtbs = workbook.vtblist()
    for vtb in vtbs:
        vtb.write()

    return namespaces, workbook

if __name__ == '__main__':
    import sys; argv = sys.argv
    main(argv[1])