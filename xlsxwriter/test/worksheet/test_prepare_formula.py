###############################################################################
#
# Tests for XlsxWriter.
#
# SPDX-License-Identifier: BSD-2-Clause
# Copyright (c), 2013-2022, John McNamara, jmcnamara@cpan.org
#

import unittest
from io import StringIO
from ...worksheet import Worksheet


class TestCalculateSpans(unittest.TestCase):
    """
    Test the _prepare_formula Worksheet method for different formula types.

    """

    def test_prepare_formula(self):
        self.fh = StringIO()
        self.worksheet = Worksheet()
        self.worksheet._set_filehandle(self.fh)

        self.worksheet.use_future_functions = True

        testcases = [
            ['=foo()', 'foo()'],
            ['{foo()}', 'foo()'],
            ['{=foo()}', 'foo()'],

            # Dynamic functions.
            ['LET()', '_xlfn.LET()'],
            ['SEQUENCE(10)', '_xlfn.SEQUENCE(10)'],
            ['UNIQUES(A1:A10)', 'UNIQUES(A1:A10)'],
            ['UUNIQUE(A1:A10)', 'UUNIQUE(A1:A10)'],
            ['SINGLE(A1:A3)', '_xlfn.SINGLE(A1:A3)'],
            ['UNIQUE(A1:A10)', '_xlfn.UNIQUE(A1:A10)'],
            ['_xlfn.SEQUENCE(10)', '_xlfn.SEQUENCE(10)'],
            ['SORT(A1:A10)', '_xlfn._xlws.SORT(A1:A10)'],
            ['RANDARRAY(10,1)', '_xlfn.RANDARRAY(10,1)'],
            ['ANCHORARRAY(C1)', '_xlfn.ANCHORARRAY(C1)'],
            ['SORTBY(A1:A10,B1)', '_xlfn.SORTBY(A1:A10,B1)'],
            ['FILTER(A1:A10,1)', '_xlfn._xlws.FILTER(A1:A10,1)'],
            ['XMATCH(B1:B2,A1:A10)', '_xlfn.XMATCH(B1:B2,A1:A10)'],
            ['COUNTA(ANCHORARRAY(C1))', 'COUNTA(_xlfn.ANCHORARRAY(C1))'],
            ['SEQUENCE(10)*SEQUENCE(10)', '_xlfn.SEQUENCE(10)*_xlfn.SEQUENCE(10)'],
            ['XLOOKUP("India",A22:A23,B22:B23)', '_xlfn.XLOOKUP("India",A22:A23,B22:B23)'],
            ['XLOOKUP(B1,A1:A10,ANCHORARRAY(D1))', '_xlfn.XLOOKUP(B1,A1:A10,_xlfn.ANCHORARRAY(D1))'],
            ['LAMBDA(_xlpm.number, _xlpm.number + 1)(1)', '_xlfn.LAMBDA(_xlpm.number, _xlpm.number + 1)(1)'],

            # Future functions.
            ['COT()', '_xlfn.COT()'],
            ['CSC()', '_xlfn.CSC()'],
            ['IFS()', '_xlfn.IFS()'],
            ['PHI()', '_xlfn.PHI()'],
            ['RRI()', '_xlfn.RRI()'],
            ['SEC()', '_xlfn.SEC()'],
            ['XOR()', '_xlfn.XOR()'],
            ['ACOT()', '_xlfn.ACOT()'],
            ['BASE()', '_xlfn.BASE()'],
            ['COTH()', '_xlfn.COTH()'],
            ['CSCH()', '_xlfn.CSCH()'],
            ['DAYS()', '_xlfn.DAYS()'],
            ['IFNA()', '_xlfn.IFNA()'],
            ['SECH()', '_xlfn.SECH()'],
            ['ACOTH()', '_xlfn.ACOTH()'],
            ['BITOR()', '_xlfn.BITOR()'],
            ['F.INV()', '_xlfn.F.INV()'],
            ['GAMMA()', '_xlfn.GAMMA()'],
            ['GAUSS()', '_xlfn.GAUSS()'],
            ['IMCOT()', '_xlfn.IMCOT()'],
            ['IMCSC()', '_xlfn.IMCSC()'],
            ['IMSEC()', '_xlfn.IMSEC()'],
            ['IMTAN()', '_xlfn.IMTAN()'],
            ['MUNIT()', '_xlfn.MUNIT()'],
            ['SHEET()', '_xlfn.SHEET()'],
            ['T.INV()', '_xlfn.T.INV()'],
            ['VAR.P()', '_xlfn.VAR.P()'],
            ['VAR.S()', '_xlfn.VAR.S()'],
            ['ARABIC()', '_xlfn.ARABIC()'],
            ['BITAND()', '_xlfn.BITAND()'],
            ['BITXOR()', '_xlfn.BITXOR()'],
            ['CONCAT()', '_xlfn.CONCAT()'],
            ['F.DIST()', '_xlfn.F.DIST()'],
            ['F.TEST()', '_xlfn.F.TEST()'],
            ['IMCOSH()', '_xlfn.IMCOSH()'],
            ['IMCSCH()', '_xlfn.IMCSCH()'],
            ['IMSECH()', '_xlfn.IMSECH()'],
            ['IMSINH()', '_xlfn.IMSINH()'],
            ['MAXIFS()', '_xlfn.MAXIFS()'],
            ['MINIFS()', '_xlfn.MINIFS()'],
            ['SHEETS()', '_xlfn.SHEETS()'],
            ['SKEW.P()', '_xlfn.SKEW.P()'],
            ['SWITCH()', '_xlfn.SWITCH()'],
            ['T.DIST()', '_xlfn.T.DIST()'],
            ['T.TEST()', '_xlfn.T.TEST()'],
            ['Z.TEST()', '_xlfn.Z.TEST()'],
            ['COMBINA()', '_xlfn.COMBINA()'],
            ['DECIMAL()', '_xlfn.DECIMAL()'],
            ['RANK.EQ()', '_xlfn.RANK.EQ()'],
            ['STDEV.P()', '_xlfn.STDEV.P()'],
            ['STDEV.S()', '_xlfn.STDEV.S()'],
            ['UNICHAR()', '_xlfn.UNICHAR()'],
            ['UNICODE()', '_xlfn.UNICODE()'],
            ['BETA.INV()', '_xlfn.BETA.INV()'],
            ['F.INV.RT()', '_xlfn.F.INV.RT()'],
            ['ISO.CEILING()', 'ISO.CEILING()'],
            ['NORM.INV()', '_xlfn.NORM.INV()'],
            ['RANK.AVG()', '_xlfn.RANK.AVG()'],
            ['T.INV.2T()', '_xlfn.T.INV.2T()'],
            ['TEXTJOIN()', '_xlfn.TEXTJOIN()'],
            ['AGGREGATE()', '_xlfn.AGGREGATE()'],
            ['BETA.DIST()', '_xlfn.BETA.DIST()'],
            ['BINOM.INV()', '_xlfn.BINOM.INV()'],
            ['BITLSHIFT()', '_xlfn.BITLSHIFT()'],
            ['BITRSHIFT()', '_xlfn.BITRSHIFT()'],
            ['CHISQ.INV()', '_xlfn.CHISQ.INV()'],
            ['ECMA.CEILING()', 'ECMA.CEILING()'],
            ['F.DIST.RT()', '_xlfn.F.DIST.RT()'],
            ['FILTERXML()', '_xlfn.FILTERXML()'],
            ['GAMMA.INV()', '_xlfn.GAMMA.INV()'],
            ['ISFORMULA()', '_xlfn.ISFORMULA()'],
            ['MODE.MULT()', '_xlfn.MODE.MULT()'],
            ['MODE.SNGL()', '_xlfn.MODE.SNGL()'],
            ['NORM.DIST()', '_xlfn.NORM.DIST()'],
            ['PDURATION()', '_xlfn.PDURATION()'],
            ['T.DIST.2T()', '_xlfn.T.DIST.2T()'],
            ['T.DIST.RT()', '_xlfn.T.DIST.RT()'],
            ['WORKDAY.INTL()', 'WORKDAY.INTL()'],
            ['BINOM.DIST()', '_xlfn.BINOM.DIST()'],
            ['CHISQ.DIST()', '_xlfn.CHISQ.DIST()'],
            ['CHISQ.TEST()', '_xlfn.CHISQ.TEST()'],
            ['EXPON.DIST()', '_xlfn.EXPON.DIST()'],
            ['FLOOR.MATH()', '_xlfn.FLOOR.MATH()'],
            ['GAMMA.DIST()', '_xlfn.GAMMA.DIST()'],
            ['ISOWEEKNUM()', '_xlfn.ISOWEEKNUM()'],
            ['NORM.S.INV()', '_xlfn.NORM.S.INV()'],
            ['WEBSERVICE()', '_xlfn.WEBSERVICE()'],
            ['ERF.PRECISE()', '_xlfn.ERF.PRECISE()'],
            ['FORMULATEXT()', '_xlfn.FORMULATEXT()'],
            ['LOGNORM.INV()', '_xlfn.LOGNORM.INV()'],
            ['NORM.S.DIST()', '_xlfn.NORM.S.DIST()'],
            ['NUMBERVALUE()', '_xlfn.NUMBERVALUE()'],
            ['QUERYSTRING()', '_xlfn.QUERYSTRING()'],
            ['CEILING.MATH()', '_xlfn.CEILING.MATH()'],
            ['CHISQ.INV.RT()', '_xlfn.CHISQ.INV.RT()'],
            ['CONFIDENCE.T()', '_xlfn.CONFIDENCE.T()'],
            ['COVARIANCE.P()', '_xlfn.COVARIANCE.P()'],
            ['COVARIANCE.S()', '_xlfn.COVARIANCE.S()'],
            ['ERFC.PRECISE()', '_xlfn.ERFC.PRECISE()'],
            ['FORECAST.ETS()', '_xlfn.FORECAST.ETS()'],
            ['HYPGEOM.DIST()', '_xlfn.HYPGEOM.DIST()'],
            ['LOGNORM.DIST()', '_xlfn.LOGNORM.DIST()'],
            ['PERMUTATIONA()', '_xlfn.PERMUTATIONA()'],
            ['POISSON.DIST()', '_xlfn.POISSON.DIST()'],
            ['QUARTILE.EXC()', '_xlfn.QUARTILE.EXC()'],
            ['QUARTILE.INC()', '_xlfn.QUARTILE.INC()'],
            ['WEIBULL.DIST()', '_xlfn.WEIBULL.DIST()'],
            ['CHISQ.DIST.RT()', '_xlfn.CHISQ.DIST.RT()'],
            ['FLOOR.PRECISE()', '_xlfn.FLOOR.PRECISE()'],
            ['NEGBINOM.DIST()', '_xlfn.NEGBINOM.DIST()'],
            ['NETWORKDAYS.INTL()', 'NETWORKDAYS.INTL()'],
            ['PERCENTILE.EXC()', '_xlfn.PERCENTILE.EXC()'],
            ['PERCENTILE.INC()', '_xlfn.PERCENTILE.INC()'],
            ['CEILING.PRECISE()', '_xlfn.CEILING.PRECISE()'],
            ['CONFIDENCE.NORM()', '_xlfn.CONFIDENCE.NORM()'],
            ['FORECAST.LINEAR()', '_xlfn.FORECAST.LINEAR()'],
            ['GAMMALN.PRECISE()', '_xlfn.GAMMALN.PRECISE()'],
            ['PERCENTRANK.EXC()', '_xlfn.PERCENTRANK.EXC()'],
            ['PERCENTRANK.INC()', '_xlfn.PERCENTRANK.INC()'],
            ['BINOM.DIST.RANGE()', '_xlfn.BINOM.DIST.RANGE()'],
            ['FORECAST.ETS.STAT()', '_xlfn.FORECAST.ETS.STAT()'],
            ['FORECAST.ETS.CONFINT()', '_xlfn.FORECAST.ETS.CONFINT()'],
            ['FORECAST.ETS.SEASONALITY()', '_xlfn.FORECAST.ETS.SEASONALITY()'],

            ['Z.TEST(Z.TEST(Z.TEST()))', '_xlfn.Z.TEST(_xlfn.Z.TEST(_xlfn.Z.TEST()))'],
        ]

        for testcase in testcases:
            formula = testcase[0]
            exp = testcase[1]
            got = self.worksheet._prepare_formula(formula)

            self.assertEqual(got, exp)
