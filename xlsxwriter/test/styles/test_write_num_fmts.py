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

from xlsxwriter.format import Format
from xlsxwriter.styles import Styles


class TestWriteNumFmts(unittest.TestCase):
    """
    Test the Styles _write_num_fmts() method.

    """

    def setUp(self):
        self.fh = StringIO()
        self.styles = Styles()
        self.styles._set_filehandle(self.fh)

    def test_write_num_fmts(self):
        """Test the _write_num_fmts() method"""

        xf_format = Format()
        xf_format.num_format_index = 164
        xf_format.set_num_format("#,##0.0")

        self.styles._set_style_properties(
            [[xf_format], None, 0, ["#,##0.0"], 0, 0, [], [], 0]
        )

        self.styles._write_num_fmts()

        exp = """<numFmts count="1"><numFmt numFmtId="164" formatCode="#,##0.0"/></numFmts>"""
        got = self.fh.getvalue()

        self.assertEqual(exp, got)
