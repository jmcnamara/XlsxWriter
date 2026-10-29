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

from xlsxwriter.worksheet import Worksheet


class TestParseFilterExpression(unittest.TestCase):
    """
    Test the Worksheet _parse_filter_expression() method.

    """

    def setUp(self):
        self.fh = StringIO()
        self.worksheet = Worksheet()
        self.worksheet._set_filehandle(self.fh)

    def test_equal_1(self):
        expression = "x =  2000"
        tokens = self.worksheet._extract_filter_tokens(expression)
        got = self.worksheet._parse_filter_expression(expression, tokens)
        self.assertEqual([2, "2000"], got)

    def test_equal_2(self):
        expression = "x == 2000"
        tokens = self.worksheet._extract_filter_tokens(expression)
        got = self.worksheet._parse_filter_expression(expression, tokens)
        self.assertEqual([2, "2000"], got)

    def test_equal_3(self):
        expression = "x =~ 2000"
        tokens = self.worksheet._extract_filter_tokens(expression)
        got = self.worksheet._parse_filter_expression(expression, tokens)
        self.assertEqual([2, "2000"], got)

    def test_equal_4(self):
        expression = "x eq 2000"
        tokens = self.worksheet._extract_filter_tokens(expression)
        got = self.worksheet._parse_filter_expression(expression, tokens)
        self.assertEqual([2, "2000"], got)

    def test_not_equal_1(self):
        expression = "x <> 2000"
        tokens = self.worksheet._extract_filter_tokens(expression)
        got = self.worksheet._parse_filter_expression(expression, tokens)
        self.assertEqual([5, "2000"], got)

    def test_not_equal_2(self):
        expression = "x != 2000"
        tokens = self.worksheet._extract_filter_tokens(expression)
        got = self.worksheet._parse_filter_expression(expression, tokens)
        self.assertEqual([5, "2000"], got)

    def test_not_equal_3(self):
        expression = "x ne 2000"
        tokens = self.worksheet._extract_filter_tokens(expression)
        got = self.worksheet._parse_filter_expression(expression, tokens)
        self.assertEqual([5, "2000"], got)

    def test_not_equal_4(self):
        expression = "x !~ 2000"
        tokens = self.worksheet._extract_filter_tokens(expression)
        got = self.worksheet._parse_filter_expression(expression, tokens)
        self.assertEqual([5, "2000"], got)

    def test_greater_than(self):
        expression = "x >  2000"
        tokens = self.worksheet._extract_filter_tokens(expression)
        got = self.worksheet._parse_filter_expression(expression, tokens)
        self.assertEqual([4, "2000"], got)

    def test_less_than(self):
        expression = "x <  2000"
        tokens = self.worksheet._extract_filter_tokens(expression)
        got = self.worksheet._parse_filter_expression(expression, tokens)
        self.assertEqual([1, "2000"], got)

    def test_greater_or_equal(self):
        expression = "x >= 2000"
        tokens = self.worksheet._extract_filter_tokens(expression)
        got = self.worksheet._parse_filter_expression(expression, tokens)
        self.assertEqual([6, "2000"], got)

    def test_less_or_equal(self):
        expression = "x <= 2000"
        tokens = self.worksheet._extract_filter_tokens(expression)
        got = self.worksheet._parse_filter_expression(expression, tokens)
        self.assertEqual([3, "2000"], got)

    def test_and_1(self):
        expression = "x >  2000 and x <  5000"
        tokens = self.worksheet._extract_filter_tokens(expression)
        got = self.worksheet._parse_filter_expression(expression, tokens)
        self.assertEqual([4, "2000", 0, 1, "5000"], got)

    def test_and_2(self):
        expression = "x >  2000 &&  x <  5000"
        tokens = self.worksheet._extract_filter_tokens(expression)
        got = self.worksheet._parse_filter_expression(expression, tokens)
        self.assertEqual([4, "2000", 0, 1, "5000"], got)

    def test_or_1(self):
        expression = "x >  2000 or  x <  5000"
        tokens = self.worksheet._extract_filter_tokens(expression)
        got = self.worksheet._parse_filter_expression(expression, tokens)
        self.assertEqual([4, "2000", 1, 1, "5000"], got)

    def test_or_2(self):
        expression = "x >  2000 ||  x <  5000"
        tokens = self.worksheet._extract_filter_tokens(expression)
        got = self.worksheet._parse_filter_expression(expression, tokens)
        self.assertEqual([4, "2000", 1, 1, "5000"], got)

    def test_equal_blanks(self):
        expression = "x =  Blanks"
        tokens = self.worksheet._extract_filter_tokens(expression)
        got = self.worksheet._parse_filter_expression(expression, tokens)
        self.assertEqual([2, "blanks"], got)

    def test_equal_nonblanks(self):
        expression = "x =  NonBlanks"
        tokens = self.worksheet._extract_filter_tokens(expression)
        got = self.worksheet._parse_filter_expression(expression, tokens)
        self.assertEqual([5, " "], got)

    def test_not_equal_blanks(self):
        expression = "x <> Blanks"
        tokens = self.worksheet._extract_filter_tokens(expression)
        got = self.worksheet._parse_filter_expression(expression, tokens)
        self.assertEqual([5, " "], got)

    def test_not_equal_nonblanks(self):
        expression = "x <> NonBlanks"
        tokens = self.worksheet._extract_filter_tokens(expression)
        got = self.worksheet._parse_filter_expression(expression, tokens)
        self.assertEqual([2, "blanks"], got)

    def test_top_items(self):
        expression = "Top 10 Items"
        tokens = self.worksheet._extract_filter_tokens(expression)
        got = self.worksheet._parse_filter_expression(expression, tokens)
        self.assertEqual([30, "10"], got)

    def test_top_percent(self):
        expression = "Top 20 %"
        tokens = self.worksheet._extract_filter_tokens(expression)
        got = self.worksheet._parse_filter_expression(expression, tokens)
        self.assertEqual([31, "20"], got)

    def test_bottom_items(self):
        expression = "Bottom 5 Items"
        tokens = self.worksheet._extract_filter_tokens(expression)
        got = self.worksheet._parse_filter_expression(expression, tokens)
        self.assertEqual([32, "5"], got)

    def test_bottom_percent(self):
        expression = "Bottom 101 %"
        tokens = self.worksheet._extract_filter_tokens(expression)
        got = self.worksheet._parse_filter_expression(expression, tokens)
        self.assertEqual([33, "101"], got)
