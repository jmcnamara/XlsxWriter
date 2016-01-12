###############################################################################
#
# Tests for XlsxWriter.
#
# Copyright (c), 2013-2016, John McNamara, jmcnamara@cpan.org
#

import unittest
from ...workbook import Workbook


class TestCheckSheetname(unittest.TestCase):
    """
    Test the Workbook _check_sheetname() method.

    """

    def setUp(self):
        self.workbook = Workbook()

    def test_check_sheetname(self):
        """Test the _check_sheetname() method"""

        got = self.workbook._check_sheetname('name')
        exp = 'name'
        self.assertEqual(got, exp)

        got = self.workbook._check_sheetname('Sheet1')
        exp = 'Sheet1'
        self.assertEqual(got, exp)

        got = self.workbook._check_sheetname(None)
        exp = 'Sheet3'
        self.assertEqual(got, exp)

    def test_check_sheetname_with_exception1(self):
        """Test the _check_sheetname() method with exception"""

        name = 'name_that_is_longer_than_thirty_one_characters'
        self.assertRaises(Exception, self.workbook._check_sheetname, name)

    def test_check_sheetname_with_exception2(self):
        """Test the _check_sheetname() method with exception"""

        name = 'name_with_special_character_?'
        self.assertRaises(Exception, self.workbook._check_sheetname, name)

    def test_check_sheetname_with_exception3(self):
        """Test the _check_sheetname() method with exception"""

        name1 = 'Duplicate_name'
        name2 = name1.lower()

        self.workbook.add_worksheet(name1)
        self.assertRaises(Exception, self.workbook.add_worksheet, name2)

    def tearDown(self):
        self.workbook.fileclosed = 1
