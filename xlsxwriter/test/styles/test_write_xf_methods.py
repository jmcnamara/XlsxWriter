###############################################################################
#
# Tests for XlsxWriter.
#
# Copyright (c), 2013-2016, John McNamara, jmcnamara@cpan.org
#

import unittest
from ...compatibility import StringIO
from ...styles import Styles
from ...format import Format


class TestWriteXf(unittest.TestCase):
    """
    Test the Styles _write_xf() method. This test case is similar to
    test_write_xf.py but with methods calls instead of properties.

    """

    def setUp(self):
        self.fh = StringIO()
        self.styles = Styles()
        self.styles._set_filehandle(self.fh)

    def test_write_xf_1(self):
        """Test the _write_xf() method. Default properties."""

        xf_format = Format()

        self.styles._write_xf(xf_format)

        exp = """<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>"""
        got = self.fh.getvalue()

        self.assertEqual(got, exp)

    def test_write_xf_2(self):
        """Test the _write_xf() method. Has font but is first XF."""

        xf_format = Format()
        xf_format.set_has_font()

        self.styles._write_xf(xf_format)

        exp = """<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>"""
        got = self.fh.getvalue()

        self.assertEqual(got, exp)

    def test_write_xf_3(self):
        """Test the _write_xf() method. Has font but isn't first XF."""

        xf_format = Format()
        xf_format.set_has_font()
        xf_format.set_font_index(1)

        self.styles._write_xf(xf_format)

        exp = """<xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/>"""
        got = self.fh.getvalue()

        self.assertEqual(got, exp)

    def test_write_xf_4(self):
        """Test the _write_xf() method. Uses built-in number format."""

        xf_format = Format()
        xf_format.set_num_format_index(2)

        self.styles._write_xf(xf_format)

        exp = """<xf numFmtId="2" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>"""
        got = self.fh.getvalue()

        self.assertEqual(got, exp)

    def test_write_xf_5(self):
        """Test the _write_xf() method. Uses built-in number format + font."""

        xf_format = Format()
        xf_format.set_num_format_index(2)
        xf_format.set_has_font()
        xf_format.set_font_index(1)

        self.styles._write_xf(xf_format)

        exp = """<xf numFmtId="2" fontId="1" fillId="0" borderId="0" xfId="0" applyNumberFormat="1" applyFont="1"/>"""
        got = self.fh.getvalue()

        self.assertEqual(got, exp)

    def test_write_xf_6(self):
        """Test the _write_xf() method. Vertical alignment = top."""

        xf_format = Format()
        xf_format.set_align('top')

        self.styles._write_xf(xf_format)

        exp = """<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1"><alignment vertical="top"/></xf>"""
        got = self.fh.getvalue()

        self.assertEqual(got, exp)

    def test_write_xf_7(self):
        """Test the _write_xf() method. Vertical alignment = centre."""

        xf_format = Format()
        xf_format.set_align('vcenter')

        self.styles._write_xf(xf_format)

        exp = """<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1"><alignment vertical="center"/></xf>"""
        got = self.fh.getvalue()

        self.assertEqual(got, exp)

    def test_write_xf_8(self):
        """Test the _write_xf() method. Vertical alignment = bottom."""

        xf_format = Format()
        xf_format.set_align('bottom')

        self.styles._write_xf(xf_format)

        exp = """<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1"/>"""
        got = self.fh.getvalue()

        self.assertEqual(got, exp)

    def test_write_xf_9(self):
        """Test the _write_xf() method. Vertical alignment = justify."""

        xf_format = Format()
        xf_format.set_align('vjustify')

        self.styles._write_xf(xf_format)

        exp = """<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1"><alignment vertical="justify"/></xf>"""
        got = self.fh.getvalue()

        self.assertEqual(got, exp)

    def test_write_xf_10(self):
        """Test the _write_xf() method. Vertical alignment = distributed."""

        xf_format = Format()
        xf_format.set_align('vdistributed')

        self.styles._write_xf(xf_format)

        exp = """<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1"><alignment vertical="distributed"/></xf>"""
        got = self.fh.getvalue()

        self.assertEqual(got, exp)

    def test_write_xf_11(self):
        """Test the _write_xf() method. Horizontal alignment = left."""

        xf_format = Format()
        xf_format.set_align('left')

        self.styles._write_xf(xf_format)

        exp = """<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1"><alignment horizontal="left"/></xf>"""
        got = self.fh.getvalue()

        self.assertEqual(got, exp)

    def test_write_xf_12(self):
        """Test the _write_xf() method. Horizontal alignment = center."""

        xf_format = Format()
        xf_format.set_align('center')

        self.styles._write_xf(xf_format)

        exp = """<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1"><alignment horizontal="center"/></xf>"""
        got = self.fh.getvalue()

        self.assertEqual(got, exp)

    def test_write_xf_13(self):
        """Test the _write_xf() method. Horizontal alignment = right."""

        xf_format = Format()
        xf_format.set_align('right')

        self.styles._write_xf(xf_format)

        exp = """<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1"><alignment horizontal="right"/></xf>"""
        got = self.fh.getvalue()

        self.assertEqual(got, exp)

    def test_write_xf_14(self):
        """Test the _write_xf() method. Horizontal alignment = left + indent."""

        xf_format = Format()
        xf_format.set_align('left')
        xf_format.set_indent()

        self.styles._write_xf(xf_format)

        exp = """<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1"><alignment horizontal="left" indent="1"/></xf>"""
        got = self.fh.getvalue()

        self.assertEqual(got, exp)

    def test_write_xf_15(self):
        """Test the _write_xf() method. Horizontal alignment = right + indent."""

        xf_format = Format()
        xf_format.set_align('right')
        xf_format.set_indent()

        self.styles._write_xf(xf_format)

        exp = """<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1"><alignment horizontal="right" indent="1"/></xf>"""
        got = self.fh.getvalue()

        self.assertEqual(got, exp)

    def test_write_xf_16(self):
        """Test the _write_xf() method. Horizontal alignment = fill."""

        xf_format = Format()
        xf_format.set_align('fill')

        self.styles._write_xf(xf_format)

        exp = """<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1"><alignment horizontal="fill"/></xf>"""
        got = self.fh.getvalue()

        self.assertEqual(got, exp)

    def test_write_xf_17(self):
        """Test the _write_xf() method. Horizontal alignment = justify."""

        xf_format = Format()
        xf_format.set_align('justify')

        self.styles._write_xf(xf_format)

        exp = """<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1"><alignment horizontal="justify"/></xf>"""
        got = self.fh.getvalue()

        self.assertEqual(got, exp)

    def test_write_xf_18(self):
        """Test the _write_xf() method. Horizontal alignment = center across."""

        xf_format = Format()
        xf_format.set_align('center_across')

        self.styles._write_xf(xf_format)

        exp = """<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1"><alignment horizontal="centerContinuous"/></xf>"""
        got = self.fh.getvalue()

        self.assertEqual(got, exp)

    def test_write_xf_19(self):
        """Test the _write_xf() method. Horizontal alignment = distributed."""

        xf_format = Format()
        xf_format.set_align('distributed')

        self.styles._write_xf(xf_format)

        exp = """<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1"><alignment horizontal="distributed"/></xf>"""
        got = self.fh.getvalue()

        self.assertEqual(got, exp)

    def test_write_xf_20(self):
        """Test the _write_xf() method. Horizontal alignment = distributed + indent."""

        xf_format = Format()
        xf_format.set_align('distributed')
        xf_format.set_indent()

        self.styles._write_xf(xf_format)

        exp = """<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1"><alignment horizontal="distributed" indent="1"/></xf>"""
        got = self.fh.getvalue()

        self.assertEqual(got, exp)

    def test_write_xf_21(self):
        """Test the _write_xf() method. Horizontal alignment = justify distributed."""

        xf_format = Format()
        xf_format.set_align('justify_distributed')

        self.styles._write_xf(xf_format)

        exp = """<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1"><alignment horizontal="distributed" justifyLastLine="1"/></xf>"""
        got = self.fh.getvalue()

        self.assertEqual(got, exp)

    def test_write_xf_22(self):
        """Test the _write_xf() method. Horizontal alignment = indent only."""

        xf_format = Format()
        xf_format.set_indent()

        self.styles._write_xf(xf_format)

        exp = """<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1"><alignment horizontal="left" indent="1"/></xf>"""
        got = self.fh.getvalue()

        self.assertEqual(got, exp)

    def test_write_xf_23(self):
        """Test the _write_xf() method. Horizontal alignment = distributed + indent."""

        xf_format = Format()
        xf_format.set_align('justify_distributed')
        xf_format.set_indent()

        self.styles._write_xf(xf_format)

        exp = """<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1"><alignment horizontal="distributed" indent="1"/></xf>"""
        got = self.fh.getvalue()

        self.assertEqual(got, exp)

    def test_write_xf_24(self):
        """Test the _write_xf() method. Alignment = text wrap"""

        xf_format = Format()
        xf_format.set_text_wrap()

        self.styles._write_xf(xf_format)

        exp = """<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1"><alignment wrapText="1"/></xf>"""
        got = self.fh.getvalue()

        self.assertEqual(got, exp)

    def test_write_xf_25(self):
        """Test the _write_xf() method. Alignment = shrink to fit"""

        xf_format = Format()
        xf_format.set_shrink()

        self.styles._write_xf(xf_format)

        exp = """<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1"><alignment shrinkToFit="1"/></xf>"""
        got = self.fh.getvalue()

        self.assertEqual(got, exp)

    def test_write_xf_26(self):
        """Test the _write_xf() method. Alignment = reading order"""

        xf_format = Format()
        xf_format.set_reading_order()

        self.styles._write_xf(xf_format)

        exp = """<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1"><alignment readingOrder="1"/></xf>"""
        got = self.fh.getvalue()

        self.assertEqual(got, exp)

    def test_write_xf_27(self):
        """Test the _write_xf() method. Alignment = reading order"""

        xf_format = Format()
        xf_format.set_reading_order(2)

        self.styles._write_xf(xf_format)

        exp = """<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1"><alignment readingOrder="2"/></xf>"""
        got = self.fh.getvalue()

        self.assertEqual(got, exp)

    def test_write_xf_28(self):
        """Test the _write_xf() method. Alignment = rotation"""

        xf_format = Format()
        xf_format.set_rotation(45)

        self.styles._write_xf(xf_format)

        exp = """<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1"><alignment textRotation="45"/></xf>"""
        got = self.fh.getvalue()

        self.assertEqual(got, exp)

    def test_write_xf_29(self):
        """Test the _write_xf() method. Alignment = rotation"""

        xf_format = Format()
        xf_format.set_rotation(-45)

        self.styles._write_xf(xf_format)

        exp = """<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1"><alignment textRotation="135"/></xf>"""
        got = self.fh.getvalue()

        self.assertEqual(got, exp)

    def test_write_xf_30(self):
        """Test the _write_xf() method. Alignment = rotation"""

        xf_format = Format()
        xf_format.set_rotation(270)

        self.styles._write_xf(xf_format)

        exp = """<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1"><alignment textRotation="255"/></xf>"""
        got = self.fh.getvalue()

        self.assertEqual(got, exp)

    def test_write_xf_31(self):
        """Test the _write_xf() method. Alignment = rotation"""

        xf_format = Format()
        xf_format.set_rotation(90)

        self.styles._write_xf(xf_format)

        exp = """<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1"><alignment textRotation="90"/></xf>"""
        got = self.fh.getvalue()

        self.assertEqual(got, exp)

    def test_write_xf_32(self):
        """Test the _write_xf() method. Alignment = rotation"""

        xf_format = Format()
        xf_format.set_rotation(-90)

        self.styles._write_xf(xf_format)

        exp = """<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1"><alignment textRotation="180"/></xf>"""
        got = self.fh.getvalue()

        self.assertEqual(got, exp)

    def test_write_xf_33(self):
        """Test the _write_xf() method. With cell protection."""

        xf_format = Format()
        xf_format.set_locked(0)

        self.styles._write_xf(xf_format)

        exp = """<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyProtection="1"><protection locked="0"/></xf>"""
        got = self.fh.getvalue()

        self.assertEqual(got, exp)

    def test_write_xf_34(self):
        """Test the _write_xf() method. With cell protection."""

        xf_format = Format()
        xf_format.set_hidden()

        self.styles._write_xf(xf_format)

        exp = """<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyProtection="1"><protection hidden="1"/></xf>"""
        got = self.fh.getvalue()

        self.assertEqual(got, exp)

    def test_write_xf_35(self):
        """Test the _write_xf() method. With cell protection."""

        xf_format = Format()
        xf_format.set_locked(0)
        xf_format.set_hidden()

        self.styles._write_xf(xf_format)

        exp = """<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyProtection="1"><protection locked="0" hidden="1"/></xf>"""
        got = self.fh.getvalue()

        self.assertEqual(got, exp)

    def test_write_xf_36(self):
        """Test the _write_xf() method. With cell protection + align."""

        xf_format = Format()
        xf_format.set_align('right')
        xf_format.set_locked(0)
        xf_format.set_hidden()

        self.styles._write_xf(xf_format)

        exp = """<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1" applyProtection="1"><alignment horizontal="right"/><protection locked="0" hidden="1"/></xf>"""
        got = self.fh.getvalue()

        self.assertEqual(got, exp)
