###############################################################################
#
# Tests for XlsxWriter.
#
# Copyright (c), 2013-2016, John McNamara, jmcnamara@cpan.org
#

import unittest
from ...compatibility import StringIO
from ...worksheet import Worksheet


class TestCalculateSpans(unittest.TestCase):
    """
    Test the _calculate_spans Worksheet method for different cell ranges.

    """

    def setUp(self):
        self.fh = StringIO()
        self.worksheet = Worksheet()
        self.worksheet._set_filehandle(self.fh)

    def test_calculate_spans_0(self):
        """Test Worksheet _calculate_spans()"""
        row = 0
        col = 0

        for i in range(row, row + 17):
            self.worksheet.write_number(i, col, 1)
            col = col + 1

        self.worksheet._calculate_spans()

        exp = {0: '1:16', 1: '17:17'}
        got = self.worksheet.row_spans

        self.assertEqual(got, exp)

    def test_calculate_spans_1(self):
        """Test Worksheet _calculate_spans()"""
        row = 0
        col = 0

        for i in range(row, row + 17):
            self.worksheet.write_number(i, col, 1)
            col = col + 1

        self.worksheet._calculate_spans()

        got = self.worksheet.row_spans
        exp = {0: '1:16', 1: '17:17'}

        self.assertEqual(got, exp)

    def test_calculate_spans_2(self):
        """Test Worksheet _calculate_spans()"""
        row = 1
        col = 0

        for i in range(row, row + 17):
            self.worksheet.write_number(i, col, 1)
            col = col + 1

        self.worksheet._calculate_spans()

        got = self.worksheet.row_spans
        exp = {0: '1:15', 1: '16:17'}

        self.assertEqual(got, exp)

    def test_calculate_spans_3(self):
        """Test Worksheet _calculate_spans()"""
        row = 2
        col = 0

        for i in range(row, row + 17):
            self.worksheet.write_number(i, col, 1)
            col = col + 1

        self.worksheet._calculate_spans()

        got = self.worksheet.row_spans
        exp = {0: '1:14', 1: '15:17'}

        self.assertEqual(got, exp)

    def test_calculate_spans_4(self):
        """Test Worksheet _calculate_spans()"""
        row = 3
        col = 0

        for i in range(row, row + 17):
            self.worksheet.write_number(i, col, 1)
            col = col + 1

        self.worksheet._calculate_spans()

        got = self.worksheet.row_spans
        exp = {0: '1:13', 1: '14:17'}

        self.assertEqual(got, exp)

    def test_calculate_spans_5(self):
        """Test Worksheet _calculate_spans()"""
        row = 4
        col = 0

        for i in range(row, row + 17):
            self.worksheet.write_number(i, col, 1)
            col = col + 1

        self.worksheet._calculate_spans()

        got = self.worksheet.row_spans
        exp = {0: '1:12', 1: '13:17'}

        self.assertEqual(got, exp)

    def test_calculate_spans_6(self):
        """Test Worksheet _calculate_spans()"""
        row = 5
        col = 0

        for i in range(row, row + 17):
            self.worksheet.write_number(i, col, 1)
            col = col + 1

        self.worksheet._calculate_spans()

        got = self.worksheet.row_spans
        exp = {0: '1:11', 1: '12:17'}

        self.assertEqual(got, exp)

    def test_calculate_spans_7(self):
        """Test Worksheet _calculate_spans()"""
        row = 6
        col = 0

        for i in range(row, row + 17):
            self.worksheet.write_number(i, col, 1)
            col = col + 1

        self.worksheet._calculate_spans()

        got = self.worksheet.row_spans
        exp = {0: '1:10', 1: '11:17'}

        self.assertEqual(got, exp)

    def test_calculate_spans_8(self):
        """Test Worksheet _calculate_spans()"""
        row = 7
        col = 0

        for i in range(row, row + 17):
            self.worksheet.write_number(i, col, 1)
            col = col + 1

        self.worksheet._calculate_spans()

        got = self.worksheet.row_spans
        exp = {0: '1:9', 1: '10:17'}

        self.assertEqual(got, exp)

    def test_calculate_spans_9(self):
        """Test Worksheet _calculate_spans()"""
        row = 8
        col = 0

        for i in range(row, row + 17):
            self.worksheet.write_number(i, col, 1)
            col = col + 1

        self.worksheet._calculate_spans()

        got = self.worksheet.row_spans
        exp = {0: '1:8', 1: '9:17'}

        self.assertEqual(got, exp)

    def test_calculate_spans_10(self):
        """Test Worksheet _calculate_spans()"""
        row = 9
        col = 0

        for i in range(row, row + 17):
            self.worksheet.write_number(i, col, 1)
            col = col + 1

        self.worksheet._calculate_spans()

        got = self.worksheet.row_spans
        exp = {0: '1:7', 1: '8:17'}

        self.assertEqual(got, exp)

    def test_calculate_spans_11(self):
        """Test Worksheet _calculate_spans()"""
        row = 10
        col = 0

        for i in range(row, row + 17):
            self.worksheet.write_number(i, col, 1)
            col = col + 1

        self.worksheet._calculate_spans()

        got = self.worksheet.row_spans
        exp = {0: '1:6', 1: '7:17'}

        self.assertEqual(got, exp)

    def test_calculate_spans_12(self):
        """Test Worksheet _calculate_spans()"""
        row = 11
        col = 0

        for i in range(row, row + 17):
            self.worksheet.write_number(i, col, 1)
            col = col + 1

        self.worksheet._calculate_spans()

        got = self.worksheet.row_spans
        exp = {0: '1:5', 1: '6:17'}

        self.assertEqual(got, exp)

    def test_calculate_spans_13(self):
        """Test Worksheet _calculate_spans()"""
        row = 12
        col = 0

        for i in range(row, row + 17):
            self.worksheet.write_number(i, col, 1)
            col = col + 1

        self.worksheet._calculate_spans()

        got = self.worksheet.row_spans
        exp = {0: '1:4', 1: '5:17'}

        self.assertEqual(got, exp)

    def test_calculate_spans_14(self):
        """Test Worksheet _calculate_spans()"""
        row = 13
        col = 0

        for i in range(row, row + 17):
            self.worksheet.write_number(i, col, 1)
            col = col + 1

        self.worksheet._calculate_spans()

        got = self.worksheet.row_spans
        exp = {0: '1:3', 1: '4:17'}

        self.assertEqual(got, exp)

    def test_calculate_spans_15(self):
        """Test Worksheet _calculate_spans()"""
        row = 14
        col = 0

        for i in range(row, row + 17):
            self.worksheet.write_number(i, col, 1)
            col = col + 1

        self.worksheet._calculate_spans()

        got = self.worksheet.row_spans
        exp = {0: '1:2', 1: '3:17'}

        self.assertEqual(got, exp)

    def test_calculate_spans_16(self):
        """Test Worksheet _calculate_spans()"""
        row = 15
        col = 0

        for i in range(row, row + 17):
            self.worksheet.write_number(i, col, 1)
            col = col + 1

        self.worksheet._calculate_spans()

        got = self.worksheet.row_spans
        exp = {0: '1:1', 1: '2:17'}

        self.assertEqual(got, exp)

    def test_calculate_spans_17(self):
        """Test Worksheet _calculate_spans()"""
        row = 16
        col = 0

        for i in range(row, row + 17):
            self.worksheet.write_number(i, col, 1)
            col = col + 1

        self.worksheet._calculate_spans()

        got = self.worksheet.row_spans
        exp = {1: '1:16', 2: '17:17'}

        self.assertEqual(got, exp)

    def test_calculate_spans_18(self):
        """Test Worksheet _calculate_spans()"""
        row = 16
        col = 1

        for i in range(row, row + 17):
            self.worksheet.write_number(i, col, 1)
            col = col + 1

        self.worksheet._calculate_spans()

        got = self.worksheet.row_spans
        exp = {1: '2:17', 2: '18:18'}

        self.assertEqual(got, exp)
