###############################################################################
#
# Tests for XlsxWriter.
#
# SPDX-License-Identifier: BSD-2-Clause
# Copyright (c), 2013-2024, John McNamara, jmcnamara@cpan.org
#

import unittest
from io import StringIO
from ...vml import Vml


class TestWriteVstroke(unittest.TestCase):
    """
    Test the Vml _write_stroke() method.

    """

    def setUp(self):
        self.fh = StringIO()
        self.vml = Vml()
        self.vml._set_filehandle(self.fh)

    def test_write_stroke(self):
        """Test the _write_stroke() method"""

        self.vml._write_stroke()

        exp = """<v:stroke joinstyle="miter"/>"""
        got = self.fh.getvalue()

        self.assertEqual(got, exp)
