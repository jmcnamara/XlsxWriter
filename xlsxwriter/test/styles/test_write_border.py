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


class TestWriteBorder(unittest.TestCase):
    """
    Test the Styles _write_border() method.

    """

    def setUp(self):
        self.fh = StringIO()
        self.styles = Styles()
        self.styles._set_filehandle(self.fh)

    def test_write_border(self):
        """Test the _write_border() method"""

        xf_format = Format()
        xf_format.has_border = True

        self.styles._write_border(xf_format)

        exp = """<border><left/><right/><top/><bottom/><diagonal/></border>"""
        got = self.fh.getvalue()

        self.assertEqual(exp, got)
