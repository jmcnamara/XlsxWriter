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
        """Test for border styles."""
        self.maxDiff = None

        fh = StringIO()
        style = Styles()
        style._set_filehandle(fh)

        workbook = Workbook()
        workbook.fileclosed = 1

        workbook.add_format({'top': 7})
        workbook.add_format({'top': 4})
        workbook.add_format({'top': 11})
        workbook.add_format({'top': 9})
        workbook.add_format({'top': 3})
        workbook.add_format({'top': 1})
        workbook.add_format({'top': 12})
        workbook.add_format({'top': 13})
        workbook.add_format({'top': 10})
        workbook.add_format({'top': 8})
        workbook.add_format({'top': 2})
        workbook.add_format({'top': 5})
        workbook.add_format({'top': 6})

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
                  <borders count="14">
                    <border>
                      <left/>
                      <right/>
                      <top/>
                      <bottom/>
                      <diagonal/>
                    </border>
                    <border>
                      <left/>
                      <right/>
                      <top style="hair">
                        <color auto="1"/>
                      </top>
                      <bottom/>
                      <diagonal/>
                    </border>
                    <border>
                      <left/>
                      <right/>
                      <top style="dotted">
                        <color auto="1"/>
                      </top>
                      <bottom/>
                      <diagonal/>
                    </border>
                    <border>
                      <left/>
                      <right/>
                      <top style="dashDotDot">
                        <color auto="1"/>
                      </top>
                      <bottom/>
                      <diagonal/>
                    </border>
                    <border>
                      <left/>
                      <right/>
                      <top style="dashDot">
                        <color auto="1"/>
                      </top>
                      <bottom/>
                      <diagonal/>
                    </border>
                    <border>
                      <left/>
                      <right/>
                      <top style="dashed">
                        <color auto="1"/>
                      </top>
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
                      <top style="mediumDashDotDot">
                        <color auto="1"/>
                      </top>
                      <bottom/>
                      <diagonal/>
                    </border>
                    <border>
                      <left/>
                      <right/>
                      <top style="slantDashDot">
                        <color auto="1"/>
                      </top>
                      <bottom/>
                      <diagonal/>
                    </border>
                    <border>
                      <left/>
                      <right/>
                      <top style="mediumDashDot">
                        <color auto="1"/>
                      </top>
                      <bottom/>
                      <diagonal/>
                    </border>
                    <border>
                      <left/>
                      <right/>
                      <top style="mediumDashed">
                        <color auto="1"/>
                      </top>
                      <bottom/>
                      <diagonal/>
                    </border>
                    <border>
                      <left/>
                      <right/>
                      <top style="medium">
                        <color auto="1"/>
                      </top>
                      <bottom/>
                      <diagonal/>
                    </border>
                    <border>
                      <left/>
                      <right/>
                      <top style="thick">
                        <color auto="1"/>
                      </top>
                      <bottom/>
                      <diagonal/>
                    </border>
                    <border>
                      <left/>
                      <right/>
                      <top style="double">
                        <color auto="1"/>
                      </top>
                      <bottom/>
                      <diagonal/>
                    </border>
                  </borders>
                  <cellStyleXfs count="1">
                    <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
                  </cellStyleXfs>
                  <cellXfs count="14">
                    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
                    <xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyBorder="1"/>
                    <xf numFmtId="0" fontId="0" fillId="0" borderId="2" xfId="0" applyBorder="1"/>
                    <xf numFmtId="0" fontId="0" fillId="0" borderId="3" xfId="0" applyBorder="1"/>
                    <xf numFmtId="0" fontId="0" fillId="0" borderId="4" xfId="0" applyBorder="1"/>
                    <xf numFmtId="0" fontId="0" fillId="0" borderId="5" xfId="0" applyBorder="1"/>
                    <xf numFmtId="0" fontId="0" fillId="0" borderId="6" xfId="0" applyBorder="1"/>
                    <xf numFmtId="0" fontId="0" fillId="0" borderId="7" xfId="0" applyBorder="1"/>
                    <xf numFmtId="0" fontId="0" fillId="0" borderId="8" xfId="0" applyBorder="1"/>
                    <xf numFmtId="0" fontId="0" fillId="0" borderId="9" xfId="0" applyBorder="1"/>
                    <xf numFmtId="0" fontId="0" fillId="0" borderId="10" xfId="0" applyBorder="1"/>
                    <xf numFmtId="0" fontId="0" fillId="0" borderId="11" xfId="0" applyBorder="1"/>
                    <xf numFmtId="0" fontId="0" fillId="0" borderId="12" xfId="0" applyBorder="1"/>
                    <xf numFmtId="0" fontId="0" fillId="0" borderId="13" xfId="0" applyBorder="1"/>
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
