###############################################################################
#
# Tests for XlsxWriter.
#
# SPDX-License-Identifier: BSD-2-Clause
#
# Copyright (c), 2013-2026, John McNamara, jmcnamara@cpan.org
#

import unittest
from io import StringIO

from xlsxwriter.worksheet import Worksheet

from ..helperfunctions import _xml_to_list


class TestAssembleWorksheet(unittest.TestCase):
    """
    Test assembling a complete Worksheet file.

    """

    def test_assemble_xml_file(self):
        """Test that stop_if_true is ignored for types that don't support it."""
        self.maxDiff = None

        fh = StringIO()
        worksheet = Worksheet()
        worksheet._set_filehandle(fh)
        worksheet.select()

        worksheet.conditional_format("A1", {"type": "data_bar", "stop_if_true": True})
        worksheet.conditional_format(
            "A2", {"type": "2_color_scale", "stop_if_true": True}
        )
        worksheet.conditional_format(
            "A3", {"type": "3_color_scale", "stop_if_true": True}
        )
        worksheet.conditional_format(
            "A4",
            {"type": "icon_set", "icon_style": "3_arrows", "stop_if_true": True},
        )

        worksheet._assemble_xml_file()

        exp = _xml_to_list("""
                <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
                <dimension ref="A1"/>
                <sheetViews>
                    <sheetView tabSelected="1" workbookViewId="0"/>
                </sheetViews>
                <sheetFormatPr defaultRowHeight="15"/>
                <sheetData/>
                <conditionalFormatting sqref="A1">
                    <cfRule type="dataBar" priority="1">
                    <dataBar>
                        <cfvo type="min" val="0"/>
                        <cfvo type="max" val="0"/>
                        <color rgb="FF638EC6"/>
                    </dataBar>
                    </cfRule>
                </conditionalFormatting>
                <conditionalFormatting sqref="A2">
                    <cfRule type="colorScale" priority="2">
                    <colorScale>
                        <cfvo type="min" val="0"/>
                        <cfvo type="max" val="0"/>
                        <color rgb="FFFF7128"/>
                        <color rgb="FFFFEF9C"/>
                    </colorScale>
                    </cfRule>
                </conditionalFormatting>
                <conditionalFormatting sqref="A3">
                    <cfRule type="colorScale" priority="3">
                    <colorScale>
                        <cfvo type="min" val="0"/>
                        <cfvo type="percentile" val="50"/>
                        <cfvo type="max" val="0"/>
                        <color rgb="FFF8696B"/>
                        <color rgb="FFFFEB84"/>
                        <color rgb="FF63BE7B"/>
                    </colorScale>
                    </cfRule>
                </conditionalFormatting>
                <conditionalFormatting sqref="A4">
                    <cfRule type="iconSet" priority="4">
                    <iconSet iconSet="3Arrows">
                        <cfvo type="percent" val="0"/>
                        <cfvo type="percent" val="33"/>
                        <cfvo type="percent" val="67"/>
                    </iconSet>
                    </cfRule>
                </conditionalFormatting>
                <pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>
                </worksheet>
            """)

        got = _xml_to_list(fh.getvalue())

        self.assertEqual(exp, got)
