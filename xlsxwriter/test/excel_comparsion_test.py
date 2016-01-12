###############################################################################
#
# Tests for XlsxWriter.
#
# Copyright (c), 2013-2016, John McNamara, jmcnamara@cpan.org
#

import unittest
import os
from .helperfunctions import _compare_xlsx_files


class ExcelComparisonTest(unittest.TestCase):
    """
    Test class for comparing a file created by XlsxWriter against a file
    created by Excel.

    """

    def assertExcelEqual(self):

        got, exp = _compare_xlsx_files(self.got_filename,
                                       self.exp_filename,
                                       self.ignore_files,
                                       self.ignore_elements)

        self.assertEqual(got, exp)

    def tearDown(self):
        # Cleanup by removing the temp excel file created for testing.
        if os.path.exists(self.got_filename):
            os.remove(self.got_filename)
