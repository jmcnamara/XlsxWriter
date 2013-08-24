###############################################################################
#
# Tests for XlsxWriter.
#
# Copyright (c), 2013, John McNamara, jmcnamara@cpan.org
#

import unittest
from ..compatibility import StringIO
from ...workbook import Workbook


class TestWriteCalcPr(unittest.TestCase):
    """
    Test the Workbook _write_calc_pr() method.

    """

    def setUp(self):
        self.fh = StringIO()
        self.workbook = Workbook()
        self.workbook._set_filehandle(self.fh)

    def test_write_calc_pr(self):
        """Test the _write_calc_pr() method"""

        self.workbook._write_calc_pr()

        exp = """<calcPr calcId="124519" fullCalcOnLoad="1"/>"""
        got = self.fh.getvalue()

        self.assertEqual(got, exp)

    def tearDown(self):
        self.workbook.fileclosed = 1

if __name__ == '__main__':
    unittest.main()
