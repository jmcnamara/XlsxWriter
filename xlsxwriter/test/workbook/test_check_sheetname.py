###############################################################################
#
# Tests for XlsxWriter.
#
# SPDX-License-Identifier: BSD-2-Clause
#
# Copyright (c), 2013-2026, John McNamara, jmcnamara@cpan.org
#

import unittest

from xlsxwriter.exceptions import DuplicateWorksheetName, InvalidWorksheetName
from xlsxwriter.workbook import Workbook


class TestCheckSheetname(unittest.TestCase):
    """
    Test the Workbook _check_sheetname() method.

    """

    def setUp(self):
        self.workbook = Workbook()

    def test_check_sheetname_valid_custom_names(self):
        self.assertEqual(self.workbook._check_sheetname("name"), "name")
        self.assertEqual(self.workbook._check_sheetname("Sheet1"), "Sheet1")

        # Excel's limit is in UTF-16 units, so these emojis count as 2.
        name = "\U0001f600" * 15 + "a"
        self.assertEqual(self.workbook._check_sheetname(name), name)

    def test_check_sheetname_none_generates_default(self):
        got = self.workbook._check_sheetname(None)
        self.assertTrue(got.startswith("Sheet"))

    def test_check_sheetname_empty_generates_default(self):
        got = self.workbook._check_sheetname("")
        self.assertTrue(got.startswith("Sheet"))

    def test_check_sheetname_with_long_name(self):
        name = "name_that_is_longer_than_thirty_one_characters"
        self.assertRaises(InvalidWorksheetName, self.workbook._check_sheetname, name)

        # Excel's limit is in UTF-16 units, so these emojis count as 2.
        name = "\U0001f600" * 16
        self.assertRaises(InvalidWorksheetName, self.workbook._check_sheetname, name)

    def test_check_sheetname_with_invalid_characters_or_formatting(self):
        invalid_names = [
            "name_with_special_character_?",
            "'start with apostrophe",
            "end with apostrophe'",
            "'start and end with apostrophe'",
        ]

        for name in invalid_names:
            with self.subTest(name=name):
                self.assertRaises(
                    InvalidWorksheetName, self.workbook._check_sheetname, name
                )

    def test_check_sheetname_with_duplicate_name(self):
        """Test the _check_sheetname() method with exception"""

        name1 = "Duplicate_name"
        name2 = name1.lower()

        self.workbook.add_worksheet(name1)
        self.assertRaises(DuplicateWorksheetName, self.workbook.add_worksheet, name2)

    def tearDown(self):
        self.workbook.fileclosed = True
