###############################################################################
#
# Tests for XlsxWriter.
#
# SPDX-License-Identifier: BSD-2-Clause
# Copyright (c), 2013-2024, John McNamara, jmcnamara@cpan.org
#

import unittest
from io import StringIO
from ...workbook import Workbook
from ...exceptions import OverlappingRange


class TestOverlapRanges(unittest.TestCase):
    """
    Test overlapping merge and table ranges.

    """

    def setUp(self):
        fh = StringIO()
        workbook = Workbook()
        workbook._set_filehandle(fh)
        self.workbook = workbook

    def test_overlaps01(self):
        """Test Worksheet range overlap exceptions"""
        worksheet = self.workbook.add_worksheet()

        worksheet.merge_range("A1:G10", "")

        with self.assertRaises(OverlappingRange):
            worksheet.merge_range("A1:G10", "")

    def test_overlaps02(self):
        """Test Worksheet range overlap exceptions"""
        worksheet = self.workbook.add_worksheet()

        worksheet.merge_range("A1:G10", "")

        with self.assertRaises(OverlappingRange):
            worksheet.merge_range("B3:C3", "")

    def test_overlaps03(self):
        """Test Worksheet range overlap exceptions"""
        worksheet = self.workbook.add_worksheet()

        worksheet.merge_range("A1:G10", "")

        with self.assertRaises(OverlappingRange):
            worksheet.merge_range("G10:G11", "")

    def test_overlaps04(self):
        """Test Worksheet range overlap exceptions"""
        worksheet = self.workbook.add_worksheet()

        worksheet.add_table("A1:G10")

        with self.assertRaises(OverlappingRange):
            worksheet.add_table("A1:G10")

    def test_overlaps05(self):
        """Test Worksheet range overlap exceptions"""
        worksheet = self.workbook.add_worksheet()

        worksheet.add_table("A1:G10")

        with self.assertRaises(OverlappingRange):
            worksheet.add_table("B3:C3")

    def test_overlaps06(self):
        """Test Worksheet range overlap exceptions"""
        worksheet = self.workbook.add_worksheet()

        worksheet.add_table("A1:G10")

        with self.assertRaises(OverlappingRange):
            worksheet.add_table("G1:G11")

    def test_overlaps07(self):
        """Test Worksheet range overlap exceptions"""
        worksheet = self.workbook.add_worksheet()

        worksheet.merge_range("A1:G10", "")

        with self.assertRaises(OverlappingRange):
            worksheet.add_table("B3:C3")

    def test_overlaps08(self):
        """Test Worksheet range overlap exceptions"""
        worksheet = self.workbook.add_worksheet()

        worksheet.add_table("A1:G10")

        with self.assertRaises(OverlappingRange):
            worksheet.merge_range("B3:C3", "")
