###############################################################################
#
# Tests for XlsxWriter.
#
# SPDX-License-Identifier: BSD-2-Clause
#
# Copyright (c), 2013-2025, John McNamara, jmcnamara@cpan.org
#

import unittest
from io import StringIO

from xlsxwriter.drawing import Drawing, DrawingInfo, DrawingTypes

from ..helperfunctions import _xml_to_list


class TestAssembleDrawing(unittest.TestCase):
    """
    Test assembling a complete Drawing file.

    """

    def test_assemble_xml_file(self):
        """Test writing a drawing with no cell data."""
        self.maxDiff = None

        fh = StringIO()
        drawing = Drawing()
        drawing._set_filehandle(fh)

        dimensions = [4, 8, 457200, 104775, 12, 22, 152400, 180975, 0, 0]
        drawing_object = DrawingInfo()
        drawing_object._drawing_type = DrawingTypes.CHART
        drawing_object._dimensions = dimensions
        drawing_object._width = 0
        drawing_object._height = 0
        drawing_object._description = None
        drawing_object._shape = None
        drawing_object._anchor = 1
        drawing_object._rel_index = 1
        drawing_object._url_rel_index = 0
        drawing_object._tip = None

        drawing._add_drawing_object(drawing_object)

        drawing.embedded = 1

        drawing._assemble_xml_file()

        exp = _xml_to_list(
            """
                <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                <xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                  <xdr:twoCellAnchor>
                    <xdr:from>
                      <xdr:col>4</xdr:col>
                      <xdr:colOff>457200</xdr:colOff>
                      <xdr:row>8</xdr:row>
                      <xdr:rowOff>104775</xdr:rowOff>
                    </xdr:from>
                    <xdr:to>
                      <xdr:col>12</xdr:col>
                      <xdr:colOff>152400</xdr:colOff>
                      <xdr:row>22</xdr:row>
                      <xdr:rowOff>180975</xdr:rowOff>
                    </xdr:to>
                    <xdr:graphicFrame macro="">
                      <xdr:nvGraphicFramePr>
                        <xdr:cNvPr id="2" name="Chart 1"/>
                        <xdr:cNvGraphicFramePr/>
                      </xdr:nvGraphicFramePr>
                      <xdr:xfrm>
                        <a:off x="0" y="0"/>
                        <a:ext cx="0" cy="0"/>
                      </xdr:xfrm>
                      <a:graphic>
                        <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">
                          <c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rId1"/>
                        </a:graphicData>
                      </a:graphic>
                    </xdr:graphicFrame>
                    <xdr:clientData/>
                  </xdr:twoCellAnchor>
                </xdr:wsDr>
                """
        )

        got = _xml_to_list(fh.getvalue())

        self.assertEqual(exp, got)
