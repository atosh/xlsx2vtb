# coding: utf-8
from zipfile import ZipFile
from xml.etree import ElementTree

from datum import Datum
from row import Row
from sheet import Sheet
from styles import Style
from files import CSV, VTB
from utils import xml2dict

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

