###############################################################################
#
# Tests for XlsxWriter.
#
# SPDX-License-Identifier: BSD-2-Clause
#
# Copyright (c), 2013-2025, John McNamara, jmcnamara@cpan.org
#

import unittest
from io import StringIO

from xlsxwriter.drawing import Drawing, DrawingInfo
from xlsxwriter.url import Url


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

        drawing_info = DrawingInfo()

        self.drawing._write_c_nv_pr(2, drawing_info, "Chart 1")

        exp = """<xdr:cNvPr id="2" name="Chart 1"/>"""
        got = self.fh.getvalue()

        self.assertEqual(exp, got)

    def test_write_c_nv_pr_with_hyperlink(self):
        """Test the _write_c_nv_pr() method with a hyperlink"""

        url = Url("https://test")
        url.tip = "tip"
        url._rel_index = 1

        drawing_info = DrawingInfo()
        drawing_info._tip = "tip"
        drawing_info._rel_index = 1
        drawing_info._url = url

        self.drawing._write_c_nv_pr(2, drawing_info, "Chart 1")

        exp = """<xdr:cNvPr id="2" name="Chart 1"><a:hlinkClick xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rId1" tooltip="tip"/></xdr:cNvPr>"""
        got = self.fh.getvalue()

        self.assertEqual(exp, got)
