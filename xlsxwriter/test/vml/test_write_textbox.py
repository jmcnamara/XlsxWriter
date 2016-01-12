###############################################################################
#
# Tests for XlsxWriter.
#
# Copyright (c), 2013-2016, John McNamara, jmcnamara@cpan.org
#

import unittest
from ...compatibility import StringIO
from ...vml import Vml


class TestWriteVtextbox(unittest.TestCase):
    """
    Test the Vml _write_textbox() method.

    """

    def setUp(self):
        self.fh = StringIO()
        self.vml = Vml()
        self.vml._set_filehandle(self.fh)

    def test_write_comment_textbox(self):
        """Test the _write_comment_textbox() method"""

        self.vml._write_comment_textbox()

        exp = """<v:textbox style="mso-direction-alt:auto"><div style="text-align:left"></div></v:textbox>"""
        got = self.fh.getvalue()

        self.assertEqual(got, exp)
