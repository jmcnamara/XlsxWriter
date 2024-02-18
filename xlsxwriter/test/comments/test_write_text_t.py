###############################################################################
#
# Tests for XlsxWriter.
#
# SPDX-License-Identifier: BSD-2-Clause
# Copyright (c), 2013-2024, John McNamara, jmcnamara@cpan.org
#

import unittest
from io import StringIO
from ...comments import Comments


class TestWriteText(unittest.TestCase):
    """
    Test the Comments _write_text_t() method.

    """

    def setUp(self):
        self.fh = StringIO()
        self.comments = Comments()
        self.comments._set_filehandle(self.fh)

    def test_write_text_t_1(self):
        """Test the _write_text_t() method"""

        self.comments._write_text_t("Some text")

        exp = """<t>Some text</t>"""
        got = self.fh.getvalue()

        self.assertEqual(got, exp)

    def test_write_text_t_2(self):
        """Test the _write_text_t() method"""

        self.comments._write_text_t(" Some text")

        exp = """<t xml:space="preserve"> Some text</t>"""
        got = self.fh.getvalue()

        self.assertEqual(got, exp)

    def test_write_text_t_3(self):
        """Test the _write_text_t() method"""

        self.comments._write_text_t("Some text ")

        exp = """<t xml:space="preserve">Some text </t>"""
        got = self.fh.getvalue()

        self.assertEqual(got, exp)

    def test_write_text_t_4(self):
        """Test the _write_text_t() method"""

        self.comments._write_text_t(" Some text ")

        exp = """<t xml:space="preserve"> Some text </t>"""
        got = self.fh.getvalue()

        self.assertEqual(got, exp)

    def test_write_text_t_5(self):
        """Test the _write_text_t() method"""

        self.comments._write_text_t("Some text\n")

        exp = """<t xml:space="preserve">Some text\n</t>"""
        got = self.fh.getvalue()

        self.assertEqual(got, exp)
