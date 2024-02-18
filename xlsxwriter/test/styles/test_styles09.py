###############################################################################
#
# Tests for XlsxWriter.
#
# SPDX-License-Identifier: BSD-2-Clause
# Copyright (c), 2013-2024, John McNamara, jmcnamara@cpan.org
#

import unittest
from io import StringIO
from ..helperfunctions import _xml_to_list
from ...styles import Styles
from ...workbook import Workbook


class TestAssembleStyles(unittest.TestCase):
    """
    Test assembling a complete Styles file.

    """

    def test_assemble_xml_file(self):
        """Test for simple font styles."""
        self.maxDiff = None

        fh = StringIO()
        style = Styles()
        style._set_filehandle(fh)

        workbook = Workbook()

        format1 = workbook.add_format(
            {
                "font_color": "#9C0006",
                "bg_color": "#FFC7CE",
                "font_condense": 1,
                "font_extend": 1,
                "has_fill": 1,
                "has_font": 1,
            }
        )

        # Get (and set) the DXF format index.
        format1._get_dxf_index()

        workbook._prepare_format_properties()

        style._set_style_properties(
            [
                workbook.xf_formats,
                workbook.palette,
                workbook.font_count,
                workbook.num_formats,
                workbook.border_count,
                workbook.fill_count,
                workbook.custom_colors,
                workbook.dxf_formats,
                workbook.has_comments,
            ]
        )

        style._assemble_xml_file()
        workbook.fileclosed = 1

        exp = _xml_to_list(
            """
                <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                <styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
                  <fonts count="1">
                    <font>
                      <sz val="11"/>
                      <color theme="1"/>
                      <name val="Calibri"/>
                      <family val="2"/>
                      <scheme val="minor"/>
                    </font>
                  </fonts>
                  <fills count="2">
                    <fill>
                      <patternFill patternType="none"/>
                    </fill>
                    <fill>
                      <patternFill patternType="gray125"/>
                    </fill>
                  </fills>
                  <borders count="1">
                    <border>
                      <left/>
                      <right/>
                      <top/>
                      <bottom/>
                      <diagonal/>
                    </border>
                  </borders>
                  <cellStyleXfs count="1">
                    <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
                  </cellStyleXfs>
                  <cellXfs count="1">
                    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
                  </cellXfs>
                  <cellStyles count="1">
                    <cellStyle name="Normal" xfId="0" builtinId="0"/>
                  </cellStyles>
                  <dxfs count="1">
                    <dxf>
                      <font>
                        <condense val="0"/>
                        <extend val="0"/>
                        <color rgb="FF9C0006"/>
                      </font>
                      <fill>
                        <patternFill>
                          <bgColor rgb="FFFFC7CE"/>
                        </patternFill>
                      </fill>
                    </dxf>
                  </dxfs>
                  <tableStyles count="0" defaultTableStyle="TableStyleMedium9" defaultPivotStyle="PivotStyleLight16"/>
                </styleSheet>
                """
        )

        got = _xml_to_list(fh.getvalue())

        self.assertEqual(got, exp)
