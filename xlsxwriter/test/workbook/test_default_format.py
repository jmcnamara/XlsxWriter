###############################################################################
#
# Tests for XlsxWriter.
#
# SPDX-License-Identifier: BSD-2-Clause
#
# Copyright (c), 2013-2026, John McNamara, jmcnamara@cpan.org
#

import unittest

from xlsxwriter.workbook import Workbook


class TestDefaultFormat(unittest.TestCase):
    """
    Test the Workbook default format options.

    """

    def test_user_font_scheme(self):
        """Test that a user font isn't tied to the default theme font scheme."""

        workbook = Workbook(None, {"excel_2023_theme": True})

        format1 = workbook.add_format({"font_name": "Calibri"})
        format2 = workbook.add_format({"bold": True})

        workbook.fileclosed = 1

        self.assertEqual(format1.font_scheme, "none")
        self.assertEqual(format2.font_scheme, "minor")

    def test_font_scheme_key(self):
        """Test that fonts that only differ by scheme aren't merged."""

        workbook = Workbook()

        format1 = workbook.add_format({"bold": True})
        format2 = workbook.add_format({"bold": True, "font_scheme": "none"})

        # Get (and set) the XF format indices.
        format1._get_xf_index()
        format2._get_xf_index()

        workbook._prepare_format_properties()

        workbook.fileclosed = 1

        self.assertEqual(workbook.font_count, 3)

    def test_options_unchanged(self):
        """Test that the user default_format_properties aren't modified."""

        properties = {"font_name": "Arial", "font_size": 11}
        workbook = Workbook(None, {"default_format_properties": properties})

        workbook.fileclosed = 1

        self.assertEqual(properties, {"font_name": "Arial", "font_size": 11})

    def test_unsupported_column_width(self):
        """Test the fallback for an unsupported default_column_width."""

        with self.assertWarns(UserWarning):
            workbook = Workbook(None, {"default_column_width": 100})

        workbook.fileclosed = 1

        self.assertEqual(workbook.default_col_width, 64)

    def test_rich_string_default_font(self):
        """Test that rich string fragments use the default font."""

        workbook = Workbook(None, {"excel_2023_theme": True})
        worksheet = workbook.add_worksheet()
        bold = workbook.add_format({"bold": True})

        worksheet.write_rich_string("A1", "a", bold, "b", "c")

        workbook.fileclosed = 1

        exp = (
            '<r><t>a</t></r><r><rPr><b/><sz val="11"/><color theme="1"/>'
            '<rFont val="Aptos Narrow"/><family val="2"/><scheme val="minor"/>'
            '</rPr><t>b</t></r><r><rPr><sz val="11"/><color theme="1"/>'
            '<rFont val="Aptos Narrow"/><family val="2"/><scheme val="minor"/>'
            "</rPr><t>c</t></r>"
        )

        self.assertIn(exp, worksheet.str_table.string_table)
