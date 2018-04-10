###############################################################################
#
# Tests for XlsxWriter.
#
# Copyright (c), 2013-2018, John McNamara, jmcnamara@cpan.org
#

import unittest
from ...compatibility import StringIO
from ...worksheet import Worksheet


class TestParseFilterExpression(unittest.TestCase):
    """
    Test the Worksheet _parse_filter_expression() method.

    """

    def setUp(self):
        self.fh = StringIO()
        self.worksheet = Worksheet()
        self.worksheet._set_filehandle(self.fh)

    def test_parse_filter_expression(self):
        """Test the _parse_filter_expression() method"""

        testcases = [

            [
                'x =  2000',
                [2, '2000'],
            ],

            [
                'x == 2000',
                [2, '2000'],
            ],

            [
                'x =~ 2000',
                [2, '2000'],
            ],

            [
                'x eq 2000',
                [2, '2000'],
            ],

            [
                'x <> 2000',
                [5, '2000'],
            ],

            [
                'x != 2000',
                [5, '2000'],
            ],

            [
                'x ne 2000',
                [5, '2000'],
            ],

            [
                'x !~ 2000',
                [5, '2000'],
            ],

            [
                'x >  2000',
                [4, '2000'],
            ],

            [
                'x <  2000',
                [1, '2000'],
            ],

            [
                'x >= 2000',
                [6, '2000'],
            ],

            [
                'x <= 2000',
                [3, '2000'],
            ],

            [
                'x >  2000 and x <  5000',
                [4, '2000', 0, 1, '5000'],
            ],

            [
                'x >  2000 &&  x <  5000',
                [4, '2000', 0, 1, '5000'],
            ],

            [
                'x >  2000 or  x <  5000',
                [4, '2000', 1, 1, '5000'],
            ],

            [
                'x >  2000 ||  x <  5000',
                [4, '2000', 1, 1, '5000'],
            ],

            [
                'x =  Blanks',
                [2, 'blanks'],
            ],

            [
                'x =  NonBlanks',
                [5, ' '],
            ],

            [
                'x <> Blanks',
                [5, ' '],
            ],

            [
                'x <> NonBlanks',
                [2, 'blanks'],
            ],

            [
                'Top 10 Items',
                [30, '10'],
            ],

            [
                'Top 20 %',
                [31, '20'],
            ],

            [
                'Bottom 5 Items',
                [32, '5'],
            ],

            [
                'Bottom 101 %',
                [33, '101'],
            ],
        ]

        for testcase in testcases:
            expression = testcase[0]
            tokens = self.worksheet._extract_filter_tokens(expression)

            exp = testcase[1]
            got = self.worksheet._parse_filter_expression(expression, tokens)

            self.assertEqual(got, exp)
