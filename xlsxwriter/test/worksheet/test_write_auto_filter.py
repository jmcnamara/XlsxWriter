###############################################################################
#
# Tests for XlsxWriter.
#
# SPDX-License-Identifier: BSD-2-Clause
# Copyright (c), 2013-2024, John McNamara, jmcnamara@cpan.org
#

import unittest
from io import StringIO
from ...worksheet import Worksheet


class TestWriteAutoFilter(unittest.TestCase):
    """
    Test the Worksheet _write_auto_filter() method.

    """

    def setUp(self):
        self.fh = StringIO()
        self.worksheet = Worksheet()
        self.worksheet._set_filehandle(self.fh)
        self.worksheet.name = "Sheet1"
        self.worksheet.autofilter("A1:D51")

    def test_write_auto_filter_1(self):
        """Test the _write_auto_filter() method"""

        self.worksheet._write_auto_filter()

        exp = """<autoFilter ref="A1:D51"/>"""
        got = self.fh.getvalue()

        self.assertEqual(got, exp)

    def test_write_auto_filter_2(self):
        """Test the _write_auto_filter() method"""

        filter_condition = "x == East"
        exp = """<autoFilter ref="A1:D51"><filterColumn colId="0"><filters><filter val="East"/></filters></filterColumn></autoFilter>"""

        self.worksheet.filter_column(0, filter_condition)
        self.worksheet._write_auto_filter()

        got = self.fh.getvalue()

        self.assertEqual(got, exp)

    def test_write_auto_filter_3(self):
        """Test the _write_auto_filter() method"""

        filter_condition = "x == East or  x == North"
        exp = """<autoFilter ref="A1:D51"><filterColumn colId="0"><filters><filter val="East"/><filter val="North"/></filters></filterColumn></autoFilter>"""

        self.worksheet.filter_column("A", filter_condition)
        self.worksheet._write_auto_filter()

        got = self.fh.getvalue()

        self.assertEqual(got, exp)

    def test_write_auto_filter_4(self):
        """Test the _write_auto_filter() method"""

        filter_condition = "x == East and x == North"
        exp = """<autoFilter ref="A1:D51"><filterColumn colId="0"><customFilters and="1"><customFilter val="East"/><customFilter val="North"/></customFilters></filterColumn></autoFilter>"""

        self.worksheet.filter_column("A", filter_condition)
        self.worksheet._write_auto_filter()

        got = self.fh.getvalue()

        self.assertEqual(got, exp)

    def test_write_auto_filter_5(self):
        """Test the _write_auto_filter() method"""

        filter_condition = "x != East"
        exp = '<autoFilter ref="A1:D51"><filterColumn colId="0"><customFilters><customFilter operator="notEqual" val="East"/></customFilters></filterColumn></autoFilter>'

        self.worksheet.filter_column("A", filter_condition)
        self.worksheet._write_auto_filter()

        got = self.fh.getvalue()

        self.assertEqual(got, exp)

    def test_write_auto_filter_6(self):
        """Test the _write_auto_filter() method"""

        filter_condition = "x == S*"
        exp = '<autoFilter ref="A1:D51"><filterColumn colId="0"><customFilters><customFilter val="S*"/></customFilters></filterColumn></autoFilter>'

        self.worksheet.filter_column("A", filter_condition)
        self.worksheet._write_auto_filter()

        got = self.fh.getvalue()

        self.assertEqual(got, exp)

    def test_write_auto_filter_7(self):
        """Test the _write_auto_filter() method"""

        filter_condition = "x != S*"
        exp = '<autoFilter ref="A1:D51"><filterColumn colId="0"><customFilters><customFilter operator="notEqual" val="S*"/></customFilters></filterColumn></autoFilter>'

        self.worksheet.filter_column("A", filter_condition)
        self.worksheet._write_auto_filter()

        got = self.fh.getvalue()

        self.assertEqual(got, exp)

    def test_write_auto_filter_8(self):
        """Test the _write_auto_filter() method"""

        filter_condition = "x == *h"
        exp = '<autoFilter ref="A1:D51"><filterColumn colId="0"><customFilters><customFilter val="*h"/></customFilters></filterColumn></autoFilter>'

        self.worksheet.filter_column("A", filter_condition)
        self.worksheet._write_auto_filter()

        got = self.fh.getvalue()

        self.assertEqual(got, exp)

    def test_write_auto_filter_9(self):
        """Test the _write_auto_filter() method"""

        filter_condition = "x != *h"
        exp = '<autoFilter ref="A1:D51"><filterColumn colId="0"><customFilters><customFilter operator="notEqual" val="*h"/></customFilters></filterColumn></autoFilter>'

        self.worksheet.filter_column("A", filter_condition)
        self.worksheet._write_auto_filter()

        got = self.fh.getvalue()

        self.assertEqual(got, exp)

    def test_write_auto_filter_10(self):
        """Test the _write_auto_filter() method"""

        filter_condition = "x =~ *o*"
        exp = '<autoFilter ref="A1:D51"><filterColumn colId="0"><customFilters><customFilter val="*o*"/></customFilters></filterColumn></autoFilter>'

        self.worksheet.filter_column("A", filter_condition)
        self.worksheet._write_auto_filter()

        got = self.fh.getvalue()

        self.assertEqual(got, exp)

    def test_write_auto_filter_11(self):
        """Test the _write_auto_filter() method"""

        filter_condition = "x !~ *r*"
        exp = '<autoFilter ref="A1:D51"><filterColumn colId="0"><customFilters><customFilter operator="notEqual" val="*r*"/></customFilters></filterColumn></autoFilter>'

        self.worksheet.filter_column("A", filter_condition)
        self.worksheet._write_auto_filter()

        got = self.fh.getvalue()

        self.assertEqual(got, exp)

    def test_write_auto_filter_12(self):
        """Test the _write_auto_filter() method"""

        filter_condition = "x == 1000"
        exp = '<autoFilter ref="A1:D51"><filterColumn colId="2"><filters><filter val="1000"/></filters></filterColumn></autoFilter>'

        self.worksheet.filter_column(2, filter_condition)
        self.worksheet._write_auto_filter()

        got = self.fh.getvalue()

        self.assertEqual(got, exp)

    def test_write_auto_filter_13(self):
        """Test the _write_auto_filter() method"""

        filter_condition = "x != 2000"
        exp = '<autoFilter ref="A1:D51"><filterColumn colId="2"><customFilters><customFilter operator="notEqual" val="2000"/></customFilters></filterColumn></autoFilter>'

        self.worksheet.filter_column("C", filter_condition)
        self.worksheet._write_auto_filter()

        got = self.fh.getvalue()

        self.assertEqual(got, exp)

    def test_write_auto_filter_14(self):
        """Test the _write_auto_filter() method"""

        filter_condition = "x > 3000"
        exp = '<autoFilter ref="A1:D51"><filterColumn colId="2"><customFilters><customFilter operator="greaterThan" val="3000"/></customFilters></filterColumn></autoFilter>'

        self.worksheet.filter_column("C", filter_condition)
        self.worksheet._write_auto_filter()

        got = self.fh.getvalue()

        self.assertEqual(got, exp)

    def test_write_auto_filter_15(self):
        """Test the _write_auto_filter() method"""

        filter_condition = "x >= 4000"
        exp = '<autoFilter ref="A1:D51"><filterColumn colId="2"><customFilters><customFilter operator="greaterThanOrEqual" val="4000"/></customFilters></filterColumn></autoFilter>'

        self.worksheet.filter_column("C", filter_condition)
        self.worksheet._write_auto_filter()

        got = self.fh.getvalue()

        self.assertEqual(got, exp)

    def test_write_auto_filter_16(self):
        """Test the _write_auto_filter() method"""

        filter_condition = "x < 5000"
        exp = '<autoFilter ref="A1:D51"><filterColumn colId="2"><customFilters><customFilter operator="lessThan" val="5000"/></customFilters></filterColumn></autoFilter>'

        self.worksheet.filter_column("C", filter_condition)
        self.worksheet._write_auto_filter()

        got = self.fh.getvalue()

        self.assertEqual(got, exp)

    def test_write_auto_filter_17(self):
        """Test the _write_auto_filter() method"""

        filter_condition = "x <= 6000"
        exp = '<autoFilter ref="A1:D51"><filterColumn colId="2"><customFilters><customFilter operator="lessThanOrEqual" val="6000"/></customFilters></filterColumn></autoFilter>'

        self.worksheet.filter_column("C", filter_condition)
        self.worksheet._write_auto_filter()

        got = self.fh.getvalue()

        self.assertEqual(got, exp)

    def test_write_auto_filter_18(self):
        """Test the _write_auto_filter() method"""

        filter_condition = "x >= 1000 and x <= 2000"
        exp = '<autoFilter ref="A1:D51"><filterColumn colId="2"><customFilters and="1"><customFilter operator="greaterThanOrEqual" val="1000"/><customFilter operator="lessThanOrEqual" val="2000"/></customFilters></filterColumn></autoFilter>'

        self.worksheet.filter_column("C", filter_condition)
        self.worksheet._write_auto_filter()

        got = self.fh.getvalue()

        self.assertEqual(got, exp)

    def test_write_auto_filter_19(self):
        """Test the _write_auto_filter() method"""

        matches = ["East"]
        exp = '<autoFilter ref="A1:D51"><filterColumn colId="0"><filters><filter val="East"/></filters></filterColumn></autoFilter>'

        self.worksheet.filter_column_list("A", matches)
        self.worksheet._write_auto_filter()

        got = self.fh.getvalue()

        self.assertEqual(got, exp)

    def test_write_auto_filter_20(self):
        """Test the _write_auto_filter() method"""

        matches = ["East", "North"]
        exp = '<autoFilter ref="A1:D51"><filterColumn colId="0"><filters><filter val="East"/><filter val="North"/></filters></filterColumn></autoFilter>'

        self.worksheet.filter_column_list("A", matches)
        self.worksheet._write_auto_filter()

        got = self.fh.getvalue()

        self.assertEqual(got, exp)

    def test_write_auto_filter_21(self):
        """Test the _write_auto_filter() method"""

        matches = ["February", "January", "July", "June"]
        exp = '<autoFilter ref="A1:D51"><filterColumn colId="3"><filters><filter val="February"/><filter val="January"/><filter val="July"/><filter val="June"/></filters></filterColumn></autoFilter>'

        self.worksheet.filter_column_list(3, matches)
        self.worksheet._write_auto_filter()

        got = self.fh.getvalue()

        self.assertEqual(got, exp)
