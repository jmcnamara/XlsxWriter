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
        """Test writing a worksheet with an Excel 2010 icon set."""
        self.maxDiff = None

        fh = StringIO()
        worksheet = Worksheet()
        worksheet._set_filehandle(fh)
        worksheet.select()
        worksheet.index = 0

        worksheet.conditional_format(
            "A1", {"type": "icon_set", "icon_style": "3_triangles"}
        )

        worksheet._assemble_xml_file()

        exp = _xml_to_list("""
                <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac" mc:Ignorable="x14ac">
                    <dimension ref="A1"/>
                    <sheetViews>
                        <sheetView tabSelected="1" workbookViewId="0"/>
                    </sheetViews>
                    <sheetFormatPr defaultRowHeight="15" x14ac:dyDescent="0.25"/>
                    <sheetData/>
                    <pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>
                    <extLst>
                        <ext xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" uri="{78C0D931-6437-407d-A8EE-F0AAD7539E65}">
                            <x14:conditionalFormattings>
                                <x14:conditionalFormatting xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main">
                                    <x14:cfRule type="iconSet" priority="1" id="{DA7ABA51-AAAA-BBBB-0001-000000000001}">
                                        <x14:iconSet iconSet="3Triangles">
                                            <x14:cfvo type="percent">
                                                <xm:f>0</xm:f>
                                            </x14:cfvo>
                                            <x14:cfvo type="percent">
                                                <xm:f>33</xm:f>
                                            </x14:cfvo>
                                            <x14:cfvo type="percent">
                                                <xm:f>67</xm:f>
                                            </x14:cfvo>
                                        </x14:iconSet>
                                    </x14:cfRule>
                                    <xm:sqref>A1</xm:sqref>
                                </x14:conditionalFormatting>
                            </x14:conditionalFormattings>
                        </ext>
                    </extLst>
                </worksheet>
            """)

        got = _xml_to_list(fh.getvalue())

        self.assertEqual(exp, got)
