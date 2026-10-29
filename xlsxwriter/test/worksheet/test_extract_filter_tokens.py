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


class TestExtractFilterTokens(unittest.TestCase):
    """
    Test the Worksheet _extract_filter_tokens() method.

    """

    def setUp(self):
        self.fh = StringIO()
        self.worksheet = Worksheet()
        self.worksheet._set_filehandle(self.fh)

    def test_none_returns_empty_list(self):
        expression = None

        exp = []
        got = self.worksheet._extract_filter_tokens(expression)

        self.assertEqual(exp, got)

    def test_empty_string_returns_empty_list(self):
        expression = ""

        exp = []
        got = self.worksheet._extract_filter_tokens(expression)

        self.assertEqual(exp, got)

    def test_numeric_left_operand(self):
        expression = "0 <  2001"

        exp = ["0", "<", "2001"]
        got = self.worksheet._extract_filter_tokens(expression)

        self.assertEqual(exp, got)

    def test_less_than(self):
        expression = "x <  2000"

        exp = ["x", "<", "2000"]
        got = self.worksheet._extract_filter_tokens(expression)

        self.assertEqual(exp, got)

    def test_greater_than(self):
        expression = "x >  2000"

        exp = ["x", ">", "2000"]
        got = self.worksheet._extract_filter_tokens(expression)

        self.assertEqual(exp, got)

    def test_double_equal_operator(self):
        expression = "x == 2000"

        exp = ["x", "==", "2000"]
        got = self.worksheet._extract_filter_tokens(expression)

        self.assertEqual(exp, got)

    def test_and_expression(self):
        expression = "x >  2000 and x <  5000"

        exp = ["x", ">", "2000", "and", "x", "<", "5000"]
        got = self.worksheet._extract_filter_tokens(expression)

        self.assertEqual(exp, got)

    def test_quoted_string(self):
        expression = 'x = "goo"'

        exp = ["x", "=", "goo"]
        got = self.worksheet._extract_filter_tokens(expression)

        self.assertEqual(exp, got)

    def test_unquoted_string(self):
        expression = "x = moo"

        exp = ["x", "=", "moo"]
        got = self.worksheet._extract_filter_tokens(expression)

        self.assertEqual(exp, got)

    def test_quoted_string_with_space(self):
        expression = 'x = "foo baz"'

        exp = ["x", "=", "foo baz"]
        got = self.worksheet._extract_filter_tokens(expression)

        self.assertEqual(exp, got)

    def test_quoted_string_with_escaped_quote(self):
        expression = 'x = "moo "" bar"'

        exp = ["x", "=", 'moo " bar']
        got = self.worksheet._extract_filter_tokens(expression)

        self.assertEqual(exp, got)

    def test_or_expression_with_quoted_strings(self):
        expression = 'x = "foo bar" or x = "bar foo"'

        exp = ["x", "=", "foo bar", "or", "x", "=", "bar foo"]
        got = self.worksheet._extract_filter_tokens(expression)

        self.assertEqual(exp, got)

    def test_or_expression_with_escaped_quotes(self):
        expression = 'x = "foo "" bar" or x = "bar "" foo"'

        exp = ["x", "=", 'foo " bar', "or", "x", "=", 'bar " foo']
        got = self.worksheet._extract_filter_tokens(expression)

        self.assertEqual(exp, got)

    def test_only_escaped_quotes(self):
        expression = "x = " + '"' * 8

        exp = ["x", "=", '"' * 3]
        got = self.worksheet._extract_filter_tokens(expression)

        self.assertEqual(exp, got)

    def test_blanks(self):
        expression = "x = Blanks"

        exp = ["x", "=", "Blanks"]
        got = self.worksheet._extract_filter_tokens(expression)

        self.assertEqual(exp, got)

    def test_nonblanks(self):
        expression = "x = NonBlanks"

        exp = ["x", "=", "NonBlanks"]
        got = self.worksheet._extract_filter_tokens(expression)

        self.assertEqual(exp, got)

    def test_top_percent(self):
        expression = "top 10 %"

        exp = ["top", "10", "%"]
        got = self.worksheet._extract_filter_tokens(expression)

        self.assertEqual(exp, got)

    def test_top_items(self):
        expression = "top 10 items"

        exp = ["top", "10", "items"]
        got = self.worksheet._extract_filter_tokens(expression)

        self.assertEqual(exp, got)
