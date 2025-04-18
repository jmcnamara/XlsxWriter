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


class TestWriteBorders(unittest.TestCase):
    """
    Test the Styles _write_borders() method.

    """

    def setUp(self):
        self.fh = StringIO()
        self.styles = Styles()
        self.styles._set_filehandle(self.fh)

    def test_write_borders(self):
        """Test the _write_borders() method"""

        xf_format = Format()
        xf_format.has_border = True

        self.styles._set_style_properties([[xf_format], None, 0, 0, 1, 0, [], [], 0])

        self.styles._write_borders()

        exp = """<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>"""
        got = self.fh.getvalue()

        self.assertEqual(exp, got)
