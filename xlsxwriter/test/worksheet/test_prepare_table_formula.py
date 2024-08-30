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


class TestPrepareTableFormula(unittest.TestCase):
    """
    Test the _prepare_table_formula Worksheet method for different formula types.

    """

    def test_prepare_table_formula(self):
        self.fh = StringIO()
        self.worksheet = Worksheet()
        self.worksheet._set_filehandle(self.fh)

        testcases = [
            ["", ""],
            ["@", "[#This Row],"],
            ['"Comment @"', r'"Comment @"'],
            [
                "SUM(Table1[@[Column1]:[Column3]])",
                "SUM(Table1[[#This Row],[Column1]:[Column3]])",
            ],
            [
                """HYPERLINK(CONCAT("http://myweb.com:1677/'@md=d&path/to/sour/...'@/",[@CL],"?ac=10"),[@CL])""",
                """HYPERLINK(CONCAT("http://myweb.com:1677/'@md=d&path/to/sour/...'@/",[[#This Row],CL],"?ac=10"),[[#This Row],CL])""",
            ],
        ]

        for testcase in testcases:
            formula = testcase[0]
            exp = testcase[1]
            got = self.worksheet._prepare_table_formula(formula)

            self.assertEqual(exp, got)
