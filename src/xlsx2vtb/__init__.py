#!/bin/usr/env python
# coding: utf-8
import sys
from xlsx2vtb.workbook import WorkBook


def main(argv):
    """
    xlsxファイル読み込み
    workbook読み取り
    sheet読み取り
    style読み取り
    sharedstring読み取り
    適当にjoin
    ...
    シートごとにcsv, vtbファイルを生成
    """
    if len(argv) == 1:
        msg = 'usage: python xlsx2vtb.py <xlsx_filename>\n'
        sys.stderr.write(msg)
        exit(1)

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
