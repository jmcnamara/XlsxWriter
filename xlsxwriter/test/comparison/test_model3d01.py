###############################################################################
#
# Tests for XlsxWriter.
#
# SPDX-License-Identifier: BSD-2-Clause
#
# Copyright (c), 2013-2025, John McNamara, jmcnamara@cpan.org
#

import unittest

from xlsxwriter.workbook import Workbook

from ..excel_comparison_test import ExcelComparisonTest


class TestCompareXLSXFiles(ExcelComparisonTest):
    """
    Test file created by XlsxWriter against a file created by Excel.

    """

    def setUp(self):
        self.set_filename("model3d01.xlsx")
        self.model_dir = self.test_dir + "models/"

    @unittest.skip("Reference Excel file model3d01.xlsx needs to be created")
    def test_create_file(self):
        """Test the creation of a simple XlsxWriter file with a 3D model."""

        workbook = Workbook(self.got_filename)

        worksheet = workbook.add_worksheet()

        worksheet.insert_3d_model("A1", self.model_dir + "Duck.glb")

        workbook.close()

        self.assertExcelEqual()

    @unittest.skip("Reference Excel file model3d01.xlsx needs to be created")
    def test_create_file_in_memory(self):
        """Test the creation of a simple XlsxWriter file with a 3D model."""

        workbook = Workbook(self.got_filename, {"in_memory": True})

        worksheet = workbook.add_worksheet()

        worksheet.insert_3d_model("A1", self.model_dir + "Duck.glb")

        workbook.close()

        self.assertExcelEqual()
