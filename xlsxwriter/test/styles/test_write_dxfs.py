###############################################################################
#
# Tests for XlsxWriter.
#
# SPDX-License-Identifier: BSD-2-Clause
# Copyright (c), 2013-2022, John McNamara, jmcnamara@cpan.org
#

import unittest
from io import StringIO
from ...styles import Styles


class TestWriteDxfs(unittest.TestCase):
    """
    Test the Styles _write_dxfs() method.

    """

    def setUp(self):
        self.fh = StringIO()
        self.styles = Styles()
        self.styles._set_filehandle(self.fh)

    def test_write_dxfs(self):
        """Test the _write_dxfs() method"""

        self.styles._write_dxfs()

        exp = """<dxfs count="0"/>"""
        got = self.fh.getvalue()

        self.assertEqual(got, exp)
