###############################################################################
#
# Tests for XlsxWriter.
#
# SPDX-License-Identifier: BSD-2-Clause
# Copyright (c), 2013-2024, John McNamara, jmcnamara@cpan.org
#

import unittest
from io import StringIO
from ...drawing import Drawing


class TestWriteXdrcNvGraphicFramePr(unittest.TestCase):
    """
    Test the Drawing _write_c_nv_graphic_frame_pr() method.

    """

    def setUp(self):
        self.fh = StringIO()
        self.drawing = Drawing()
        self.drawing._set_filehandle(self.fh)

    def test_write_c_nv_graphic_frame_pr(self):
        """Test the _write_c_nv_graphic_frame_pr() method"""

        self.drawing._write_c_nv_graphic_frame_pr()

        exp = """<xdr:cNvGraphicFramePr><a:graphicFrameLocks noGrp="1"/></xdr:cNvGraphicFramePr>"""
        got = self.fh.getvalue()

        self.assertEqual(got, exp)
