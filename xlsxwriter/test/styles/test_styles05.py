###############################################################################
#
# Tests for XlsxWriter.
#
# Copyright (c), 2013-2016, John McNamara, jmcnamara@cpan.org
#

import unittest
from ...compatibility import StringIO
from ..helperfunctions import _xml_to_list
from ...styles import Styles
from ...workbook import Workbook


class TestAssembleStyles(unittest.TestCase):
    """
    Test assembling a complete Styles file.

    """
    def test_assemble_xml_file(self):
        """Tests for diagonal border styles."""
        self.maxDiff = None

        fh = StringIO()
        style = Styles()
        style._set_filehandle(fh)

        workbook = Workbook()

        workbook.add_format({'left': 1})
        workbook.add_format({'right': 1})
        workbook.add_format({'top': 1})
        workbook.add_format({'bottom': 1})
        workbook.add_format({'diag_type': 1, 'diag_border': 1})
        workbook.add_format({'diag_type': 2, 'diag_border': 1})
        workbook.add_format({'diag_type': 3})  # Test default border.

        workbook._set_default_xf_indices()
        workbook._prepare_format_properties()

        style._set_style_properties([
            workbook.xf_formats,
            workbook.palette,
            workbook.font_count,
            workbook.num_format_count,
            workbook.border_count,
            workbook.fill_count,
            workbook.custom_colors,
            workbook.dxf_formats,
        ])

        style._assemble_xml_file()
        workbook.fileclosed = 1

        exp = _xml_to_list("""
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
                  <borders count="8">
                    <border>
                      <left/>
                      <right/>
                      <top/>
                      <bottom/>
                      <diagonal/>
                    </border>
                    <border>
                      <left style="thin">
                        <color auto="1"/>
                      </left>
                      <right/>
                      <top/>
                      <bottom/>
                      <diagonal/>
                    </border>
                    <border>
                      <left/>
                      <right style="thin">
                        <color auto="1"/>
                      </right>
                      <top/>
                      <bottom/>
                      <diagonal/>
                    </border>
                    <border>
                      <left/>
                      <right/>
                      <top style="thin">
                        <color auto="1"/>
                      </top>
                      <bottom/>
                      <diagonal/>
                    </border>
                    <border>
                      <left/>
                      <right/>
                      <top/>
                      <bottom style="thin">
                        <color auto="1"/>
                      </bottom>
                      <diagonal/>
                    </border>
                    <border diagonalUp="1">
                      <left/>
                      <right/>
                      <top/>
                      <bottom/>
                      <diagonal style="thin">
                        <color auto="1"/>
                      </diagonal>
                    </border>
                    <border diagonalDown="1">
                      <left/>
                      <right/>
                      <top/>
                      <bottom/>
                      <diagonal style="thin">
                        <color auto="1"/>
                      </diagonal>
                    </border>
                    <border diagonalUp="1" diagonalDown="1">
                      <left/>
                      <right/>
                      <top/>
                      <bottom/>
                      <diagonal style="thin">
                        <color auto="1"/>
                      </diagonal>
                    </border>
                  </borders>
                  <cellStyleXfs count="1">
                    <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
                  </cellStyleXfs>
                  <cellXfs count="8">
                    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
                    <xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyBorder="1"/>
                    <xf numFmtId="0" fontId="0" fillId="0" borderId="2" xfId="0" applyBorder="1"/>
                    <xf numFmtId="0" fontId="0" fillId="0" borderId="3" xfId="0" applyBorder="1"/>
                    <xf numFmtId="0" fontId="0" fillId="0" borderId="4" xfId="0" applyBorder="1"/>
                    <xf numFmtId="0" fontId="0" fillId="0" borderId="5" xfId="0" applyBorder="1"/>
                    <xf numFmtId="0" fontId="0" fillId="0" borderId="6" xfId="0" applyBorder="1"/>
                    <xf numFmtId="0" fontId="0" fillId="0" borderId="7" xfId="0" applyBorder="1"/>
                  </cellXfs>
                  <cellStyles count="1">
                    <cellStyle name="Normal" xfId="0" builtinId="0"/>
                  </cellStyles>
                  <dxfs count="0"/>
                  <tableStyles count="0" defaultTableStyle="TableStyleMedium9" defaultPivotStyle="PivotStyleLight16"/>
                </styleSheet>
                """)

        got = _xml_to_list(fh.getvalue())

        self.assertEqual(got, exp)
