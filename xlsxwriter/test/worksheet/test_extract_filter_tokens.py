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


class TestExtractFilterTokens(unittest.TestCase):
    """
    Test the Worksheet _extract_filter_tokens() method.

    """

    def setUp(self):
        self.fh = StringIO()
        self.worksheet = Worksheet()
        self.worksheet._set_filehandle(self.fh)

    def test_extract_filter_tokens(self):
        """Test the _extract_filter_tokens() method"""

        testcases = [
            [
                None,
                [],
            ],
            [
                "",
                [],
            ],
            [
                "0 <  2001",
                ["0", "<", "2001"],
            ],
            [
                "x <  2000",
                ["x", "<", "2000"],
            ],
            [
                "x >  2000",
                ["x", ">", "2000"],
            ],
            [
                "x == 2000",
                ["x", "==", "2000"],
            ],
            [
                "x >  2000 and x <  5000",
                ["x", ">", "2000", "and", "x", "<", "5000"],
            ],
            [
                'x = "goo"',
                ["x", "=", "goo"],
            ],
            [
                "x = moo",
                ["x", "=", "moo"],
            ],
            [
                'x = "foo baz"',
                ["x", "=", "foo baz"],
            ],
            [
                'x = "moo "" bar"',
                ["x", "=", 'moo " bar'],
            ],
            [
                'x = "foo bar" or x = "bar foo"',
                ["x", "=", "foo bar", "or", "x", "=", "bar foo"],
            ],
            [
                'x = "foo "" bar" or x = "bar "" foo"',
                ["x", "=", 'foo " bar', "or", "x", "=", 'bar " foo'],
            ],
            [
                'x = """"""""',
                ["x", "=", '"""'],
            ],
            [
                "x = Blanks",
                ["x", "=", "Blanks"],
            ],
            [
                "x = NonBlanks",
                ["x", "=", "NonBlanks"],
            ],
            [
                "top 10 %",
                ["top", "10", "%"],
            ],
            [
                "top 10 items",
                ["top", "10", "items"],
            ],
        ]

        for testcase in testcases:
            expression = testcase[0]

            exp = testcase[1]
            got = self.worksheet._extract_filter_tokens(expression)

            self.assertEqual(got, exp)
