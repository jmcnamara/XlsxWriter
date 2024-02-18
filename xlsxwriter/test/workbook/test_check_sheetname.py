###############################################################################
#
# Tests for XlsxWriter.
#
# SPDX-License-Identifier: BSD-2-Clause
# Copyright (c), 2013-2024, John McNamara, jmcnamara@cpan.org
#

import unittest
from ...workbook import Workbook
from ...exceptions import DuplicateWorksheetName
from ...exceptions import InvalidWorksheetName


class TestCheckSheetname(unittest.TestCase):
    """
    Test the Workbook _check_sheetname() method.

    """

    def setUp(self):
        self.workbook = Workbook()

    def test_check_sheetname(self):
        """Test the _check_sheetname() method"""

        got = self.workbook._check_sheetname("name")
        exp = "name"
        self.assertEqual(got, exp)

        got = self.workbook._check_sheetname("Sheet1")
        exp = "Sheet1"
        self.assertEqual(got, exp)

        got = self.workbook._check_sheetname(None)
        exp = "Sheet3"
        self.assertEqual(got, exp)

        got = self.workbook._check_sheetname("")
        exp = "Sheet4"
        self.assertEqual(got, exp)

    def test_check_sheetname_with_long_name(self):
        """Test the _check_sheetname() method with exception"""

        name = "name_that_is_longer_than_thirty_one_characters"
        self.assertRaises(InvalidWorksheetName, self.workbook._check_sheetname, name)

    def test_check_sheetname_with_invalid_name(self):
        """Test the _check_sheetname() method with exception"""

        name = "name_with_special_character_?"
        self.assertRaises(InvalidWorksheetName, self.workbook._check_sheetname, name)

        name = "'start with apostrophe"
        self.assertRaises(InvalidWorksheetName, self.workbook._check_sheetname, name)

        name = "end with apostrophe'"
        self.assertRaises(InvalidWorksheetName, self.workbook._check_sheetname, name)

        name = "'start and end with apostrophe'"
        self.assertRaises(InvalidWorksheetName, self.workbook._check_sheetname, name)

    def test_check_sheetname_with_duplicate_name(self):
        """Test the _check_sheetname() method with exception"""

        name1 = "Duplicate_name"
        name2 = name1.lower()

        self.workbook.add_worksheet(name1)
        self.assertRaises(DuplicateWorksheetName, self.workbook.add_worksheet, name2)

    def tearDown(self):
        self.workbook.fileclosed = True
