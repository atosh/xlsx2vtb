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


class Style(object):
    def __init__(self, applyNumFmt, numFmtId):
        self.applyNumFmt = bool(int(applyNumFmt)) if applyNumFmt else False
        self.numFmtId = int(numFmtId) if numFmtId else numFmtId


class WorkBook(object):
    def __init__(self, sheets, styles, sharedstrs):
        self.sheets = sheets
        self.styles = styles
        self.sharedstrs = sharedstrs

    def tostring(self):
        for sheet in self.sheets:
            print sheet.name, '-' * 10
            for row in sheet.rows:
                for datum in row.data:
                    # str判定
                    value = sharedstrs[int(datum.value)] if datum.typ == 's' else datum.value
                    # fmt判定
                    value = format(self.styles, datum.fmt, float(value)) if datum.fmt else value
                    print value,
                print ''


def format(styles, fmtId, value):
    numFmtId = styles[fmtId].numFmtId
    formatter = create_formatter(numFmtId)
    return formatter.format(value)


def main(argv):
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
    xlsx = ZipFile('Book1.xlsx', 'r')
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

    print_sheets(workbook)

    return namespaces, sheets, sharedstrs, styles

if __name__ == '__main__':
    import sys; argv = sys.argv
    main(argv)