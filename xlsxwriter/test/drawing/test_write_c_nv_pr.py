###############################################################################
#
# Tests for XlsxWriter.
#
# Copyright (c), 2013, John McNamara, jmcnamara@cpan.org
#

import unittest
from ...compatibility import StringIO
from ...drawing import Drawing


class TestWriteXdrcNvPr(unittest.TestCase):
    """
    Test the Drawing _write_c_nv_pr() method.

    """

    def setUp(self):
        self.fh = StringIO()
        self.drawing = Drawing()
        self.drawing._set_filehandle(self.fh)

    def test_write_c_nv_pr(self):
        """Test the _write_c_nv_pr() method"""

        self.drawing._write_c_nv_pr(2, 'Chart 1')

        exp = """<xdr:cNvPr id="2" name="Chart 1"/>"""
        got = self.fh.getvalue()

        self.assertEqual(got, exp)


if __name__ == '__main__':
    unittest.main()
