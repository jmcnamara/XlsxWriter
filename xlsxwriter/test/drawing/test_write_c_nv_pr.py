###############################################################################
#
# Tests for XlsxWriter.
#
# Copyright (c), 2013-2018, John McNamara, jmcnamara@cpan.org
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

        options = {'url': 'https://www.github.com', 'tip': 'tip'}
        self.drawing._write_c_nv_pr(2, 'Chart 1', options)

        exp = """<xdr:cNvPr id="2" name="Chart 1"/><xdr:cNvPr id="2" name="Chart 1"><a:hlinkClick xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rId1" tooltip="tip"/></xdr:cNvPr>"""
        got = self.fh.getvalue()

        self.assertEqual(got, exp)
