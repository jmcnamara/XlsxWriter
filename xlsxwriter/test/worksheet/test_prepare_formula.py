###############################################################################
#
# Tests for XlsxWriter.
#
# Copyright (c), 2013-2021, John McNamara, jmcnamara@cpan.org
#

import unittest
from ...compatibility import StringIO
from ...worksheet import Worksheet


class TestCalculateSpans(unittest.TestCase):
    """
    Test the _prepare_formula Worksheet method for different formula types.

    """

    def test_prepare_formula(self):
        self.fh = StringIO()
        self.worksheet = Worksheet()
        self.worksheet._set_filehandle(self.fh)

        testcases = [
            ['=foo()', 'foo()'],
            ['{foo()}', 'foo()'],
            ['{=foo()}', 'foo()'],
            ['SEQUENCE(10)', '_xlfn.SEQUENCE(10)'],
            ['UNIQUES(A1:A10)', 'UNIQUES(A1:A10)'],
            ['UNIQUE(A1:A10)', '_xlfn.UNIQUE(A1:A10)'],
            ['_xlfn.SEQUENCE(10)', '_xlfn.SEQUENCE(10)'],
            ['SORT(A1:A10)', '_xlfn._xlws.SORT(A1:A10)'],
            ['RANDARRAY(10,1)', '_xlfn.RANDARRAY(10,1)'],
            ['ANCHORARRAY(C1)', '_xlfn.ANCHORARRAY(C1)'],
            ['SORTBY(A1:A10,B1)', '_xlfn.SORTBY(A1:A10,B1)'],
            ['FILTER(A1:A10,1)', '_xlfn._xlws.FILTER(A1:A10,1)'],
            ['XMATCH(B1:B2,A1:A10)', '_xlfn.XMATCH(B1:B2,A1:A10)'],
            ['COUNTA(ANCHORARRAY(C1))', 'COUNTA(_xlfn.ANCHORARRAY(C1))'],
            ['XLOOKUP("India",A22:A23,B22:B23)', '_xlfn.XLOOKUP("India",A22:A23,B22:B23)'],
            ['XLOOKUP(B1,A1:A10,ANCHORARRAY(D1))', '_xlfn.XLOOKUP(B1,A1:A10,_xlfn.ANCHORARRAY(D1))'],
            ['LAMBDA(_xlpm.number, _xlpm.number + 1)(1)', '_xlfn.LAMBDA(_xlpm.number, _xlpm.number + 1)(1)'],
        ]

        for testcase in testcases:
            formula = testcase[0]
            exp = testcase[1]
            got = self.worksheet._prepare_formula(formula)

            self.assertEqual(got, exp)
