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

from xlsxwriter.chart import Chart
from xlsxwriter.chart_drawing import ChartDrawing


class TestAssembleChartDrawing(unittest.TestCase):
    """
    Test assembling a complete ChartDrawing (c:userShapes) file.

    """

    def test_assemble_xml_file_string(self):
        """Test writing a chart drawing with a plain string textbox"""

        chart = Chart()
        chart.add_textbox("Hello", {"x": 0.5, "y": 0.1})

        drawing = ChartDrawing()
        drawing.textboxes = chart.user_shapes

        fh = StringIO()
        drawing._set_filehandle(fh)
        drawing._assemble_xml_file()

        exp = (
            """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n"""
            """<c:userShapes xmlns:cdr="http://schemas.openxmlformats.org/drawingml/2006/chartDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">"""
            """<cdr:relSizeAnchor>"""
            """<cdr:from><cdr:x>0.5</cdr:x><cdr:y>0.1</cdr:y></cdr:from>"""
            """<cdr:to><cdr:x>0.7</cdr:x><cdr:y>0.2</cdr:y></cdr:to>"""
            """<cdr:sp macro="" textlink="">"""
            """<cdr:nvSpPr><cdr:cNvPr id="2" name="TextBox 2"/><cdr:cNvSpPr txBox="1"/></cdr:nvSpPr>"""
            """<cdr:spPr><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/><a:ln w="9525" cmpd="sng"><a:noFill/></a:ln></cdr:spPr>"""
            """<cdr:style><a:lnRef idx="0"><a:scrgbClr r="0" g="0" b="0"/></a:lnRef><a:fillRef idx="0"><a:scrgbClr r="0" g="0" b="0"/></a:fillRef><a:effectRef idx="0"><a:scrgbClr r="0" g="0" b="0"/></a:effectRef><a:fontRef idx="minor"><a:schemeClr val="dk1"/></a:fontRef></cdr:style>"""
            """<cdr:txBody><a:bodyPr wrap="square" rtlCol="0" anchor="t"/><a:lstStyle/>"""
            """<a:p><a:r><a:rPr lang="en-US" sz="1100"/><a:t>Hello</a:t></a:r></a:p>"""
            """</cdr:txBody></cdr:sp></cdr:relSizeAnchor></c:userShapes>"""
        )
        got = fh.getvalue()

        self.assertEqual(exp, got)

    def test_assemble_xml_file_rich(self):
        """Test writing a chart drawing with rich multi-run text"""

        chart = Chart()
        chart.add_textbox(
            [
                {
                    "align": "center",
                    "runs": [
                        {
                            "text": "C",
                            "font": {"italic": True, "name": "Times New Roman"},
                        },
                        {"text": "max", "font": {"italic": True, "baseline": -25000}},
                        {"text": "=161.3"},
                    ],
                }
            ],
            {"x": 0.5, "y": 0.3, "width": 0.3, "height": 0.2},
        )

        drawing = ChartDrawing()
        drawing.textboxes = chart.user_shapes

        fh = StringIO()
        drawing._set_filehandle(fh)
        drawing._assemble_xml_file()

        exp = (
            """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n"""
            """<c:userShapes xmlns:cdr="http://schemas.openxmlformats.org/drawingml/2006/chartDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">"""
            """<cdr:relSizeAnchor>"""
            """<cdr:from><cdr:x>0.5</cdr:x><cdr:y>0.3</cdr:y></cdr:from>"""
            """<cdr:to><cdr:x>0.8</cdr:x><cdr:y>0.5</cdr:y></cdr:to>"""
            """<cdr:sp macro="" textlink="">"""
            """<cdr:nvSpPr><cdr:cNvPr id="2" name="TextBox 2"/><cdr:cNvSpPr txBox="1"/></cdr:nvSpPr>"""
            """<cdr:spPr><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/><a:ln w="9525" cmpd="sng"><a:noFill/></a:ln></cdr:spPr>"""
            """<cdr:style><a:lnRef idx="0"><a:scrgbClr r="0" g="0" b="0"/></a:lnRef><a:fillRef idx="0"><a:scrgbClr r="0" g="0" b="0"/></a:fillRef><a:effectRef idx="0"><a:scrgbClr r="0" g="0" b="0"/></a:effectRef><a:fontRef idx="minor"><a:schemeClr val="dk1"/></a:fontRef></cdr:style>"""
            """<cdr:txBody><a:bodyPr wrap="square" rtlCol="0" anchor="t"/><a:lstStyle/>"""
            """<a:p><a:pPr algn="ctr"/>"""
            """<a:r><a:rPr lang="en-US" sz="1100" i="1"><a:latin typeface="Times New Roman"/><a:cs typeface="Times New Roman"/></a:rPr><a:t>C</a:t></a:r>"""
            """<a:r><a:rPr lang="en-US" sz="1100" i="1" baseline="-25000"/><a:t>max</a:t></a:r>"""
            """<a:r><a:rPr lang="en-US" sz="1100"/><a:t>=161.3</a:t></a:r>"""
            """</a:p></cdr:txBody></cdr:sp></cdr:relSizeAnchor></c:userShapes>"""
        )
        got = fh.getvalue()

        self.assertEqual(exp, got)


if __name__ == "__main__":
    unittest.main()
