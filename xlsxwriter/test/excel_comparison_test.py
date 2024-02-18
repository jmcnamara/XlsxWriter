###############################################################################
#
# Tests for XlsxWriter.
#
# SPDX-License-Identifier: BSD-2-Clause
# Copyright (c), 2013-2024, John McNamara, jmcnamara@cpan.org
#

import unittest
import os
from .helperfunctions import _compare_xlsx_files


class ExcelComparisonTest(unittest.TestCase):
    """
    Test class for comparing a file created by XlsxWriter against a file
    created by Excel.

    """

    def set_filename(self, filename):
        # Set the filename and paths for the test xlsx files.
        self.maxDiff = None
        self.got_filename = ""
        self.exp_filename = ""
        self.ignore_files = []
        self.ignore_elements = {}
        self.test_dir = "xlsxwriter/test/comparison/"
        self.vba_dir = self.test_dir + "xlsx_files/"
        self.image_dir = self.test_dir + "images/"

        # The reference Excel generated file.
        self.exp_filename = self.test_dir + "xlsx_files/" + filename

        # The generated XlsxWriter file.
        self.got_filename = self.test_dir + "_test_" + filename

    def set_text_file(self, filename):
        # Set the filename and path for text files used in tests.
        self.txt_filename = self.test_dir + "xlsx_files/" + filename

    def assertExcelEqual(self):
        # Compare the generate file and the reference Excel file.
        got, exp = _compare_xlsx_files(
            self.got_filename,
            self.exp_filename,
            self.ignore_files,
            self.ignore_elements,
        )

        self.assertEqual(exp, got)

    def tearDown(self):
        # Cleanup by removing the temp excel file created for testing.
        if os.path.exists(self.got_filename):
            os.remove(self.got_filename)
