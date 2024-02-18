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
from ...format import Format


class TestWriteFont(unittest.TestCase):
    """
    Test the Styles _write_font() method.

    """

    def setUp(self):
        self.fh = StringIO()
        self.styles = Styles()
        self.styles._set_filehandle(self.fh)

    def test_write_font_1(self):
        """Test the _write_font() method. Default properties."""
        properties = {}

        xf_format = Format(properties)

        self.styles._write_font(xf_format)

        exp = """<font><sz val="11"/><color theme="1"/><name val="Calibri"/><family val="2"/><scheme val="minor"/></font>"""
        got = self.fh.getvalue()

        self.assertEqual(got, exp)

    def test_write_font_2(self):
        """Test the _write_font() method. Bold."""
        properties = {"bold": 1}

        xf_format = Format(properties)

        self.styles._write_font(xf_format)

        exp = """<font><b/><sz val="11"/><color theme="1"/><name val="Calibri"/><family val="2"/><scheme val="minor"/></font>"""
        got = self.fh.getvalue()

        self.assertEqual(got, exp)

    def test_write_font_3(self):
        """Test the _write_font() method. Italic."""
        properties = {"italic": 1}

        xf_format = Format(properties)

        self.styles._write_font(xf_format)

        exp = """<font><i/><sz val="11"/><color theme="1"/><name val="Calibri"/><family val="2"/><scheme val="minor"/></font>"""
        got = self.fh.getvalue()

        self.assertEqual(got, exp)

    def test_write_font_4(self):
        """Test the _write_font() method. Underline."""
        properties = {"underline": 1}

        xf_format = Format(properties)

        self.styles._write_font(xf_format)

        exp = """<font><u/><sz val="11"/><color theme="1"/><name val="Calibri"/><family val="2"/><scheme val="minor"/></font>"""
        got = self.fh.getvalue()

        self.assertEqual(got, exp)

    def test_write_font_5(self):
        """Test the _write_font() method. Strikeout."""
        properties = {"font_strikeout": 1}

        xf_format = Format(properties)

        self.styles._write_font(xf_format)

        exp = """<font><strike/><sz val="11"/><color theme="1"/><name val="Calibri"/><family val="2"/><scheme val="minor"/></font>"""
        got = self.fh.getvalue()

        self.assertEqual(got, exp)

    def test_write_font_6(self):
        """Test the _write_font() method. Superscript."""
        properties = {"font_script": 1}

        xf_format = Format(properties)

        self.styles._write_font(xf_format)

        exp = """<font><vertAlign val="superscript"/><sz val="11"/><color theme="1"/><name val="Calibri"/><family val="2"/><scheme val="minor"/></font>"""
        got = self.fh.getvalue()

        self.assertEqual(got, exp)

    def test_write_font_7(self):
        """Test the _write_font() method. Subscript."""
        properties = {"font_script": 2}

        xf_format = Format(properties)

        self.styles._write_font(xf_format)

        exp = """<font><vertAlign val="subscript"/><sz val="11"/><color theme="1"/><name val="Calibri"/><family val="2"/><scheme val="minor"/></font>"""
        got = self.fh.getvalue()

        self.assertEqual(got, exp)

    def test_write_font_8(self):
        """Test the _write_font() method. Font name."""
        properties = {"font_name": "Arial"}

        xf_format = Format(properties)

        self.styles._write_font(xf_format)

        exp = """<font><sz val="11"/><color theme="1"/><name val="Arial"/><family val="2"/></font>"""
        got = self.fh.getvalue()

        self.assertEqual(got, exp)

    def test_write_font_9(self):
        """Test the _write_font() method. Font size."""
        properties = {"size": 12}

        xf_format = Format(properties)

        self.styles._write_font(xf_format)

        exp = """<font><sz val="12"/><color theme="1"/><name val="Calibri"/><family val="2"/><scheme val="minor"/></font>"""
        got = self.fh.getvalue()

        self.assertEqual(got, exp)

    def test_write_font_10(self):
        """Test the _write_font() method. Outline."""
        properties = {"font_outline": 1}

        xf_format = Format(properties)

        self.styles._write_font(xf_format)

        exp = """<font><outline/><sz val="11"/><color theme="1"/><name val="Calibri"/><family val="2"/><scheme val="minor"/></font>"""
        got = self.fh.getvalue()

        self.assertEqual(got, exp)

    def test_write_font_11(self):
        """Test the _write_font() method. Shadow."""
        properties = {"font_shadow": 1}

        xf_format = Format(properties)

        self.styles._write_font(xf_format)

        exp = """<font><shadow/><sz val="11"/><color theme="1"/><name val="Calibri"/><family val="2"/><scheme val="minor"/></font>"""
        got = self.fh.getvalue()

        self.assertEqual(got, exp)

    def test_write_font_12(self):
        """Test the _write_font() method. Colour = red."""
        properties = {"color": "#FF0000"}

        xf_format = Format(properties)

        self.styles._write_font(xf_format)

        exp = """<font><sz val="11"/><color rgb="FFFF0000"/><name val="Calibri"/><family val="2"/><scheme val="minor"/></font>"""
        got = self.fh.getvalue()

        self.assertEqual(got, exp)

    def test_write_font_13(self):
        """Test the _write_font() method. All font attributes to check order."""
        properties = {
            "bold": 1,
            "color": "#FF0000",
            "font_outline": 1,
            "font_script": 1,
            "font_shadow": 1,
            "font_strikeout": 1,
            "italic": 1,
            "size": 12,
            "underline": 1,
        }

        xf_format = Format(properties)

        self.styles._write_font(xf_format)

        exp = """<font><b/><i/><strike/><outline/><shadow/><u/><vertAlign val="superscript"/><sz val="12"/><color rgb="FFFF0000"/><name val="Calibri"/><family val="2"/><scheme val="minor"/></font>"""
        got = self.fh.getvalue()

        self.assertEqual(got, exp)

    def test_write_font_14(self):
        """Test the _write_font() method. Double underline."""
        properties = {"underline": 2}

        xf_format = Format(properties)

        self.styles._write_font(xf_format)

        exp = """<font><u val="double"/><sz val="11"/><color theme="1"/><name val="Calibri"/><family val="2"/><scheme val="minor"/></font>"""
        got = self.fh.getvalue()

        self.assertEqual(got, exp)

    def test_write_font_15(self):
        """Test the _write_font() method. Double underline."""
        properties = {"underline": 33}

        xf_format = Format(properties)

        self.styles._write_font(xf_format)

        exp = """<font><u val="singleAccounting"/><sz val="11"/><color theme="1"/><name val="Calibri"/><family val="2"/><scheme val="minor"/></font>"""
        got = self.fh.getvalue()

        self.assertEqual(got, exp)

    def test_write_font_16(self):
        """Test the _write_font() method. Double underline."""
        properties = {"underline": 34}

        xf_format = Format(properties)

        self.styles._write_font(xf_format)

        exp = """<font><u val="doubleAccounting"/><sz val="11"/><color theme="1"/><name val="Calibri"/><family val="2"/><scheme val="minor"/></font>"""
        got = self.fh.getvalue()

        self.assertEqual(got, exp)

    def test_write_font_17(self):
        """Test the _write_font() method. Hyperlink."""
        properties = {"hyperlink": 1}

        xf_format = Format(properties)

        self.styles._write_font(xf_format)

        exp = """<font><u/><sz val="11"/><color theme="10"/><name val="Calibri"/><family val="2"/></font>"""
        got = self.fh.getvalue()

        self.assertEqual(got, exp)
