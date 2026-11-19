###############################################################################
#
# Tests for XlsxWriter.
#
# SPDX-License-Identifier: BSD-2-Clause
#
# Copyright (c), 2013-2026, John McNamara, jmcnamara@cpan.org
#

import unittest
from io import StringIO

from xlsxwriter.exceptions import OverlappingRange
from xlsxwriter.workbook import Workbook


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

    def test_overlaps09(self):
        """Test Worksheet range overlap exceptions"""
        worksheet = self.workbook.add_worksheet()

        worksheet.add_table("A1:C5", {"autofilter": True})

        with self.assertRaises(OverlappingRange):
            worksheet.autofilter("A1:C5")

    def test_overlaps10(self):
        """Test Worksheet range overlap exceptions"""
        worksheet = self.workbook.add_worksheet()

        worksheet.autofilter("A1:C5")

        with self.assertRaises(OverlappingRange):
            worksheet.add_table("A1:C5", {"autofilter": True})

    def test_overlaps11(self):
        """Test that a rejected merge range doesn't change the worksheet"""
        worksheet = self.workbook.add_worksheet()

        worksheet.merge_range("B1:B2", "Foo")

        with self.assertRaises(OverlappingRange):
            worksheet.merge_range("A1:B2", "Bar")

        self.assertNotIn(0, worksheet.table[0])
        self.assertEqual(len(worksheet.merged_cells), 2)

        # The cells that didn't overlap should still be available.
        self.assertEqual(worksheet.merge_range("A1:A2", "Bar"), 0)

    def test_overlaps12(self):
        """Test that a rejected table range doesn't change the worksheet"""
        worksheet = self.workbook.add_worksheet()

        worksheet.add_table("B1:B3")

        with self.assertRaises(OverlappingRange):
            worksheet.add_table(
                "A1:B3", {"columns": [{"header": "Foo"}, {"header": "Bar"}]}
            )

        self.assertNotIn(0, worksheet.table[0])
        self.assertEqual(len(worksheet.table_cells), 3)
        self.assertEqual(len(worksheet.tables), 1)

        # The cells that didn't overlap should still be available.
        self.assertEqual(worksheet.add_table("A1:A3"), 0)

    def test_overlaps13(self):
        """Test that a table rejected by an autofilter doesn't change the worksheet"""
        worksheet = self.workbook.add_worksheet()

        worksheet.autofilter("A1:A6")

        with self.assertRaises(OverlappingRange):
            worksheet.add_table("A1:B3")

        self.assertEqual(len(worksheet.table), 0)
        self.assertEqual(len(worksheet.table_cells), 0)
        self.assertEqual(len(worksheet.filter_cells), 1)
        self.assertEqual(len(worksheet.tables), 0)

        # A table outside the autofilter should still be allowed.
        self.assertEqual(worksheet.add_table("B2:B4"), 0)

    def test_overlaps14(self):
        """Test that an autofilter rejected by a table doesn't change the worksheet"""
        worksheet = self.workbook.add_worksheet()

        worksheet.add_table("B1:B3")

        with self.assertRaises(OverlappingRange):
            worksheet.autofilter("A1:B6")

        self.assertIsNone(worksheet.autofilter_ref)
        self.assertEqual(len(worksheet.filter_cells), 1)

    def test_overlaps15(self):
        """Test that a table rejected by validation doesn't reserve its range"""
        worksheet = self.workbook.add_worksheet()

        with self.assertWarns(UserWarning):
            result = worksheet.add_table(
                "A1:B2", {"columns": [{"header": "Foo"}, {"header": "foo"}]}
            )

        self.assertEqual(result, -2)
        self.assertEqual(worksheet.add_table("A1:B2"), 0)
