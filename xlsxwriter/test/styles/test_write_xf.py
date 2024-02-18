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


class TestWriteXf(unittest.TestCase):
    """
    Test the Styles _write_xf() method.

    """

    def setUp(self):
        self.fh = StringIO()
        self.styles = Styles()
        self.styles._set_filehandle(self.fh)

    def test_write_xf_1(self):
        """Test the _write_xf() method. Default properties."""
        properties = {}

        xf_format = Format(properties)

        self.styles._write_xf(xf_format)

        exp = """<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>"""
        got = self.fh.getvalue()

        self.assertEqual(got, exp)

    def test_write_xf_2(self):
        """Test the _write_xf() method. Has font but is first XF."""
        properties = {"has_font": 1}

        xf_format = Format(properties)

        self.styles._write_xf(xf_format)

        exp = """<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>"""
        got = self.fh.getvalue()

        self.assertEqual(got, exp)

    def test_write_xf_3(self):
        """Test the _write_xf() method. Has font but isn't first XF."""
        properties = {"has_font": 1, "font_index": 1}

        xf_format = Format(properties)

        self.styles._write_xf(xf_format)

        exp = """<xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/>"""
        got = self.fh.getvalue()

        self.assertEqual(got, exp)

    def test_write_xf_4(self):
        """Test the _write_xf() method. Uses built-in number format."""
        properties = {"num_format_index": 2}

        xf_format = Format(properties)

        self.styles._write_xf(xf_format)

        exp = """<xf numFmtId="2" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>"""
        got = self.fh.getvalue()

        self.assertEqual(got, exp)

    def test_write_xf_5(self):
        """Test the _write_xf() method. Uses built-in number format + font."""
        properties = {"num_format_index": 2, "has_font": 1, "font_index": 1}

        xf_format = Format(properties)

        self.styles._write_xf(xf_format)

        exp = """<xf numFmtId="2" fontId="1" fillId="0" borderId="0" xfId="0" applyNumberFormat="1" applyFont="1"/>"""
        got = self.fh.getvalue()

        self.assertEqual(got, exp)

    def test_write_xf_6(self):
        """Test the _write_xf() method. Vertical alignment = top."""
        properties = {"align": "top"}

        xf_format = Format(properties)

        self.styles._write_xf(xf_format)

        exp = """<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1"><alignment vertical="top"/></xf>"""
        got = self.fh.getvalue()

        self.assertEqual(got, exp)

    def test_write_xf_7(self):
        """Test the _write_xf() method. Vertical alignment = centre."""
        properties = {"align": "vcenter"}

        xf_format = Format(properties)

        self.styles._write_xf(xf_format)

        exp = """<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1"><alignment vertical="center"/></xf>"""
        got = self.fh.getvalue()

        self.assertEqual(got, exp)

    def test_write_xf_8(self):
        """Test the _write_xf() method. Vertical alignment = bottom."""
        properties = {"align": "bottom"}

        xf_format = Format(properties)

        self.styles._write_xf(xf_format)

        exp = """<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1"/>"""
        got = self.fh.getvalue()

        self.assertEqual(got, exp)

    def test_write_xf_9(self):
        """Test the _write_xf() method. Vertical alignment = justify."""
        properties = {"align": "vjustify"}

        xf_format = Format(properties)

        self.styles._write_xf(xf_format)

        exp = """<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1"><alignment vertical="justify"/></xf>"""
        got = self.fh.getvalue()

        self.assertEqual(got, exp)

    def test_write_xf_10(self):
        """Test the _write_xf() method. Vertical alignment = distributed."""
        properties = {"align": "vdistributed"}

        xf_format = Format(properties)

        self.styles._write_xf(xf_format)

        exp = """<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1"><alignment vertical="distributed"/></xf>"""
        got = self.fh.getvalue()

        self.assertEqual(got, exp)

    def test_write_xf_11(self):
        """Test the _write_xf() method. Horizontal alignment = left."""
        properties = {"align": "left"}

        xf_format = Format(properties)

        self.styles._write_xf(xf_format)

        exp = """<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1"><alignment horizontal="left"/></xf>"""
        got = self.fh.getvalue()

        self.assertEqual(got, exp)

    def test_write_xf_12(self):
        """Test the _write_xf() method. Horizontal alignment = center."""
        properties = {"align": "center"}

        xf_format = Format(properties)

        self.styles._write_xf(xf_format)

        exp = """<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1"><alignment horizontal="center"/></xf>"""
        got = self.fh.getvalue()

        self.assertEqual(got, exp)

    def test_write_xf_13(self):
        """Test the _write_xf() method. Horizontal alignment = right."""
        properties = {"align": "right"}

        xf_format = Format(properties)

        self.styles._write_xf(xf_format)

        exp = """<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1"><alignment horizontal="right"/></xf>"""
        got = self.fh.getvalue()

        self.assertEqual(got, exp)

    def test_write_xf_14(self):
        """Test the _write_xf() method. Horizontal alignment = left + indent."""
        properties = {"align": "left", "indent": 1}

        xf_format = Format(properties)

        self.styles._write_xf(xf_format)

        exp = """<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1"><alignment horizontal="left" indent="1"/></xf>"""
        got = self.fh.getvalue()

        self.assertEqual(got, exp)

    def test_write_xf_15(self):
        """Test the _write_xf() method. Horizontal alignment = right + indent."""
        properties = {"align": "right", "indent": 1}

        xf_format = Format(properties)

        self.styles._write_xf(xf_format)

        exp = """<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1"><alignment horizontal="right" indent="1"/></xf>"""
        got = self.fh.getvalue()

        self.assertEqual(got, exp)

    def test_write_xf_16(self):
        """Test the _write_xf() method. Horizontal alignment = fill."""
        properties = {"align": "fill"}

        xf_format = Format(properties)

        self.styles._write_xf(xf_format)

        exp = """<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1"><alignment horizontal="fill"/></xf>"""
        got = self.fh.getvalue()

        self.assertEqual(got, exp)

    def test_write_xf_17(self):
        """Test the _write_xf() method. Horizontal alignment = justify."""
        properties = {"align": "justify"}

        xf_format = Format(properties)

        self.styles._write_xf(xf_format)

        exp = """<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1"><alignment horizontal="justify"/></xf>"""
        got = self.fh.getvalue()

        self.assertEqual(got, exp)

    def test_write_xf_18(self):
        """Test the _write_xf() method. Horizontal alignment = center across."""
        properties = {"align": "center_across"}

        xf_format = Format(properties)

        self.styles._write_xf(xf_format)

        exp = """<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1"><alignment horizontal="centerContinuous"/></xf>"""
        got = self.fh.getvalue()

        self.assertEqual(got, exp)

    def test_write_xf_19(self):
        """Test the _write_xf() method. Horizontal alignment = distributed."""
        properties = {"align": "distributed"}

        xf_format = Format(properties)

        self.styles._write_xf(xf_format)

        exp = """<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1"><alignment horizontal="distributed"/></xf>"""
        got = self.fh.getvalue()

        self.assertEqual(got, exp)

    def test_write_xf_20(self):
        """Test the _write_xf() method. Horizontal alignment = distributed + indent."""
        properties = {"align": "distributed", "indent": 1}

        xf_format = Format(properties)

        self.styles._write_xf(xf_format)

        exp = """<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1"><alignment horizontal="distributed" indent="1"/></xf>"""
        got = self.fh.getvalue()

        self.assertEqual(got, exp)

    def test_write_xf_21(self):
        """Test the _write_xf() method. Horizontal alignment = justify distributed."""
        properties = {"align": "justify_distributed"}

        xf_format = Format(properties)

        self.styles._write_xf(xf_format)

        exp = """<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1"><alignment horizontal="distributed" justifyLastLine="1"/></xf>"""
        got = self.fh.getvalue()

        self.assertEqual(got, exp)

    def test_write_xf_22(self):
        """Test the _write_xf() method. Horizontal alignment = indent only."""
        properties = {"indent": 1}

        xf_format = Format(properties)

        self.styles._write_xf(xf_format)

        exp = """<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1"><alignment horizontal="left" indent="1"/></xf>"""
        got = self.fh.getvalue()

        self.assertEqual(got, exp)

    def test_write_xf_23(self):
        """Test the _write_xf() method. Horizontal alignment = distributed + indent."""
        properties = {"align": "justify_distributed", "indent": 1}

        xf_format = Format(properties)

        self.styles._write_xf(xf_format)

        exp = """<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1"><alignment horizontal="distributed" indent="1"/></xf>"""
        got = self.fh.getvalue()

        self.assertEqual(got, exp)

    def test_write_xf_24(self):
        """Test the _write_xf() method. Alignment = text wrap"""
        properties = {"text_wrap": 1}

        xf_format = Format(properties)

        self.styles._write_xf(xf_format)

        exp = """<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1"><alignment wrapText="1"/></xf>"""
        got = self.fh.getvalue()

        self.assertEqual(got, exp)

    def test_write_xf_25(self):
        """Test the _write_xf() method. Alignment = shrink to fit"""
        properties = {"shrink": 1}

        xf_format = Format(properties)

        self.styles._write_xf(xf_format)

        exp = """<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1"><alignment shrinkToFit="1"/></xf>"""
        got = self.fh.getvalue()

        self.assertEqual(got, exp)

    def test_write_xf_26(self):
        """Test the _write_xf() method. Alignment = reading order"""
        properties = {"reading_order": 1}

        xf_format = Format(properties)

        self.styles._write_xf(xf_format)

        exp = """<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1"><alignment readingOrder="1"/></xf>"""
        got = self.fh.getvalue()

        self.assertEqual(got, exp)

    def test_write_xf_27(self):
        """Test the _write_xf() method. Alignment = reading order"""
        properties = {"reading_order": 2}

        xf_format = Format(properties)

        self.styles._write_xf(xf_format)

        exp = """<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1"><alignment readingOrder="2"/></xf>"""
        got = self.fh.getvalue()

        self.assertEqual(got, exp)

    def test_write_xf_28(self):
        """Test the _write_xf() method. Alignment = rotation"""
        properties = {"rotation": 45}

        xf_format = Format(properties)

        self.styles._write_xf(xf_format)

        exp = """<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1"><alignment textRotation="45"/></xf>"""
        got = self.fh.getvalue()

        self.assertEqual(got, exp)

    def test_write_xf_29(self):
        """Test the _write_xf() method. Alignment = rotation"""
        properties = {"rotation": -45}

        xf_format = Format(properties)

        self.styles._write_xf(xf_format)

        exp = """<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1"><alignment textRotation="135"/></xf>"""
        got = self.fh.getvalue()

        self.assertEqual(got, exp)

    def test_write_xf_30(self):
        """Test the _write_xf() method. Alignment = rotation"""
        properties = {"rotation": 270}

        xf_format = Format(properties)

        self.styles._write_xf(xf_format)

        exp = """<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1"><alignment textRotation="255"/></xf>"""
        got = self.fh.getvalue()

        self.assertEqual(got, exp)

    def test_write_xf_31(self):
        """Test the _write_xf() method. Alignment = rotation"""
        properties = {"rotation": 90}

        xf_format = Format(properties)

        self.styles._write_xf(xf_format)

        exp = """<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1"><alignment textRotation="90"/></xf>"""
        got = self.fh.getvalue()

        self.assertEqual(got, exp)

    def test_write_xf_32(self):
        """Test the _write_xf() method. Alignment = rotation"""
        properties = {"rotation": -90}

        xf_format = Format(properties)

        self.styles._write_xf(xf_format)

        exp = """<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1"><alignment textRotation="180"/></xf>"""
        got = self.fh.getvalue()

        self.assertEqual(got, exp)

    def test_write_xf_33(self):
        """Test the _write_xf() method. With cell protection."""
        properties = {"locked": 0}

        xf_format = Format(properties)

        self.styles._write_xf(xf_format)

        exp = """<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyProtection="1"><protection locked="0"/></xf>"""
        got = self.fh.getvalue()

        self.assertEqual(got, exp)

    def test_write_xf_34(self):
        """Test the _write_xf() method. With cell protection."""
        properties = {"hidden": 1}

        xf_format = Format(properties)

        self.styles._write_xf(xf_format)

        exp = """<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyProtection="1"><protection hidden="1"/></xf>"""
        got = self.fh.getvalue()

        self.assertEqual(got, exp)

    def test_write_xf_35(self):
        """Test the _write_xf() method. With cell protection."""
        properties = {"locked": 0, "hidden": 1}

        xf_format = Format(properties)

        self.styles._write_xf(xf_format)

        exp = """<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyProtection="1"><protection locked="0" hidden="1"/></xf>"""
        got = self.fh.getvalue()

        self.assertEqual(got, exp)

    def test_write_xf_36(self):
        """Test the _write_xf() method. With cell protection + align."""
        properties = {"align": "right", "locked": 0, "hidden": 1}

        xf_format = Format(properties)

        self.styles._write_xf(xf_format)

        exp = """<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1" applyProtection="1"><alignment horizontal="right"/><protection locked="0" hidden="1"/></xf>"""
        got = self.fh.getvalue()

        self.assertEqual(got, exp)
