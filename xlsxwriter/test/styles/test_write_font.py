###############################################################################
#
# Tests for XlsxWriter.
#
# Copyright (c), 2013, John McNamara, jmcnamara@cpan.org
#

import unittest
from StringIO import StringIO
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

    def test_write_font1(self):
        """Test the _write_font() method"""

        xf_format = Format()

        self.styles._write_font(xf_format)

        exp = """<font><sz val="11"/><color theme="1"/><name val="Calibri"/><family val="2"/><scheme val="minor"/></font>"""
        got = self.fh.getvalue()

        self.assertEqual(got, exp)

    def test_write_font2(self):
        """Test the _write_font() method"""

        xf_format = Format()
        xf_format.bold = 1

        self.styles._write_font(xf_format)

        exp = """<font><b/><sz val="11"/><color theme="1"/><name val="Calibri"/><family val="2"/><scheme val="minor"/></font>"""
        got = self.fh.getvalue()

        self.assertEqual(got, exp)

    def test_write_font3(self):
        """Test the _write_font() method"""

        xf_format = Format()
        xf_format.italic = 1

        self.styles._write_font(xf_format)

        exp = """<font><i/><sz val="11"/><color theme="1"/><name val="Calibri"/><family val="2"/><scheme val="minor"/></font>"""
        got = self.fh.getvalue()

        self.assertEqual(got, exp)

    def test_write_font4(self):
        """Test the _write_font() method"""

        xf_format = Format()
        xf_format.underline = 1

        self.styles._write_font(xf_format)

        exp = """<font><u/><sz val="11"/><color theme="1"/><name val="Calibri"/><family val="2"/><scheme val="minor"/></font>"""
        got = self.fh.getvalue()

        self.assertEqual(got, exp)

    def test_write_font5(self):
        """Test the _write_font() method"""

        xf_format = Format()
        xf_format.font_strikeout = 1

        self.styles._write_font(xf_format)

        exp = """<font><strike/><sz val="11"/><color theme="1"/><name val="Calibri"/><family val="2"/><scheme val="minor"/></font>"""
        got = self.fh.getvalue()

        self.assertEqual(got, exp)

    def test_write_font6(self):
        """Test the _write_font() method"""

        xf_format = Format()
        xf_format.font_script = 1

        self.styles._write_font(xf_format)

        exp = """<font><vertAlign val="superscript"/><sz val="11"/><color theme="1"/><name val="Calibri"/><family val="2"/><scheme val="minor"/></font>"""
        got = self.fh.getvalue()

        self.assertEqual(got, exp)

    def test_write_font7(self):
        """Test the _write_font() method"""

        xf_format = Format()
        xf_format.font_script = 2

        self.styles._write_font(xf_format)

        exp = """<font><vertAlign val="subscript"/><sz val="11"/><color theme="1"/><name val="Calibri"/><family val="2"/><scheme val="minor"/></font>"""
        got = self.fh.getvalue()

        self.assertEqual(got, exp)

    def test_write_font8(self):
        """Test the _write_font() method"""

        xf_format = Format()
        xf_format.font = 'Arial'

        self.styles._write_font(xf_format)

        exp = """<font><sz val="11"/><color theme="1"/><name val="Arial"/><family val="2"/></font>"""
        got = self.fh.getvalue()

        self.assertEqual(got, exp)

    def test_write_font9(self):
        """Test the _write_font() method"""

        xf_format = Format()
        xf_format.size = 12

        self.styles._write_font(xf_format)

        exp = """<font><sz val="12"/><color theme="1"/><name val="Calibri"/><family val="2"/><scheme val="minor"/></font>"""
        got = self.fh.getvalue()

        self.assertEqual(got, exp)

    def test_write_font10(self):
        """Test the _write_font() method"""

        xf_format = Format()
        xf_format.font_outline = 1

        self.styles._write_font(xf_format)

        exp = """<font><outline/><sz val="11"/><color theme="1"/><name val="Calibri"/><family val="2"/><scheme val="minor"/></font>"""
        got = self.fh.getvalue()

        self.assertEqual(got, exp)

    def test_write_font11(self):
        """Test the _write_font() method"""

        xf_format = Format()
        xf_format.font_shadow = 1

        self.styles._write_font(xf_format)

        exp = """<font><shadow/><sz val="11"/><color theme="1"/><name val="Calibri"/><family val="2"/><scheme val="minor"/></font>"""
        got = self.fh.getvalue()

        self.assertEqual(got, exp)

    def test_write_font12(self):
        """Test the _write_font() method"""

        xf_format = Format()
        xf_format.color = '#FF0000'

        self.styles._write_font(xf_format)

        exp = """<font><sz val="11"/><color rgb="FFFF0000"/><name val="Calibri"/><family val="2"/><scheme val="minor"/></font>"""
        got = self.fh.getvalue()

        self.assertEqual(got, exp)

    def test_write_font13(self):
        """Test the _write_font() method"""

        xf_format = Format()
        xf_format.bold = 1
        xf_format.color = '#FF0000'
        xf_format.font_outline = 1
        xf_format.font_script = 1
        xf_format.font_shadow = 1
        xf_format.font_strikeout = 1
        xf_format.italic = 1
        xf_format.size = 12
        xf_format.underline = 1

        self.styles._write_font(xf_format)

        exp = """<font><b/><i/><strike/><outline/><shadow/><u/><vertAlign val="superscript"/><sz val="12"/><color rgb="FFFF0000"/><name val="Calibri"/><family val="2"/><scheme val="minor"/></font>"""
        got = self.fh.getvalue()

        self.assertEqual(got, exp)

    def test_write_font14(self):
        """Test the _write_font() method"""

        xf_format = Format()
        xf_format.underline = 2

        self.styles._write_font(xf_format)

        exp = """<font><u val="double"/><sz val="11"/><color theme="1"/><name val="Calibri"/><family val="2"/><scheme val="minor"/></font>"""
        got = self.fh.getvalue()

        self.assertEqual(got, exp)

    def test_write_font15(self):
        """Test the _write_font() method"""

        xf_format = Format()
        xf_format.underline = 33

        self.styles._write_font(xf_format)

        exp = """<font><u val="singleAccounting"/><sz val="11"/><color theme="1"/><name val="Calibri"/><family val="2"/><scheme val="minor"/></font>"""
        got = self.fh.getvalue()

        self.assertEqual(got, exp)

    def test_write_font16(self):
        """Test the _write_font() method"""

        xf_format = Format()
        xf_format.underline = 34

        self.styles._write_font(xf_format)

        exp = """<font><u val="doubleAccounting"/><sz val="11"/><color theme="1"/><name val="Calibri"/><family val="2"/><scheme val="minor"/></font>"""
        got = self.fh.getvalue()

        self.assertEqual(got, exp)

    def test_write_font17(self):
        """Test the _write_font() method"""

        xf_format = Format()
        xf_format.hyperlink = 1
        xf_format.underline = 1
        xf_format.theme = 10

        self.styles._write_font(xf_format)

        exp = """<font><u/><sz val="11"/><color theme="10"/><name val="Calibri"/><family val="2"/></font>"""
        got = self.fh.getvalue()

        self.assertEqual(got, exp)


if __name__ == '__main__':
    unittest.main()
