###############################################################################
#
# Tests for XlsxWriter.
#
# SPDX-License-Identifier: BSD-2-Clause
# Copyright (c), 2013-2024, John McNamara, jmcnamara@cpan.org
#

import unittest
from io import StringIO
from ...styles import Styles


class TestWriteNumFmt(unittest.TestCase):
    """
    Test the Styles _write_num_fmt() method.

    """

    def setUp(self):
        self.fh = StringIO()
        self.styles = Styles()
        self.styles._set_filehandle(self.fh)

    def test_write_num_fmt(self):
        """Test the _write_num_fmt() method"""

        self.styles._write_num_fmt(164, "#,##0.0")

        exp = """<numFmt numFmtId="164" formatCode="#,##0.0"/>"""
        got = self.fh.getvalue()

        self.assertEqual(got, exp)
