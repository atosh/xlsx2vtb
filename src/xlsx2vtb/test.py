# coding: utf-8
'''
Created on Oct 10, 2013

@author: toshi
'''
import unittest


class Test(unittest.TestCase):


    def setUp(self):
        pass


    def tearDown(self):
        pass


    def test_main(self):
        import xlsx2vtb
        workbook = xlsx2vtb.main('Book1.xlsx')
        csvs = workbook.csvlist()
        for csv in csvs:
            print csv.rows


if __name__ == "__main__":
    # import sys;sys.argv = ['', 'Test.testName']
    unittest.main()
