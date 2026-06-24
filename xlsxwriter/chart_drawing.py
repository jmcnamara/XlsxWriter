###############################################################################
#
# ChartDrawing - A class for writing the Excel XLSX chart user-shapes file.
#
# SPDX-License-Identifier: BSD-2-Clause
#
# Copyright (c) 2013-2025, John McNamara, jmcnamara@cpan.org
#

from xlsxwriter import xmlwriter
from xlsxwriter.shape import Shape


class ChartDrawing(xmlwriter.XMLwriter):
    """
    A class for writing the Excel XLSX chart user-shapes (c:userShapes) file.


    """

    ###########################################################################
    #
    # Public API.
    #
    ###########################################################################

    def __init__(self) -> None:
        """
        Constructor.

        """

        super().__init__()

        self.textboxes = []

    ###########################################################################
    #
    # Private API.
    #
    ###########################################################################

    def _assemble_xml_file(self) -> None:
        # Assemble and write the XML file.

        # Write the XML declaration.
        self._xml_declaration()

        # Write the c:userShapes element.
        self._write_user_shapes()

        self._xml_end_tag("c:userShapes")

        # Close the file.
        self._xml_close()

    ###########################################################################
    #
    # XML methods.
    #
    ###########################################################################

    def _write_user_shapes(self) -> None:
        # Write the <c:userShapes> element.
        schema = "http://schemas.openxmlformats.org/drawingml/2006/"
        attributes = [
            ("xmlns:cdr", schema + "chartDrawing"),
            ("xmlns:a", schema + "main"),
            ("xmlns:c", schema + "chart"),
        ]

        self._xml_start_tag("c:userShapes", attributes)

        index = 1
        for textbox in self.textboxes:
            self._write_rel_size_anchor(index, textbox)
            index += 1

    def _write_rel_size_anchor(self, index: int, textbox) -> None:
        # Write the <cdr:relSizeAnchor> element.
        self._xml_start_tag("cdr:relSizeAnchor")

        # Write the cdr:from element.
        self._write_from(textbox["from"])

        # Write the cdr:to element.
        self._write_to(textbox["to"])

        # Write the cdr:sp element.
        self._write_sp(index, textbox)

        self._xml_end_tag("cdr:relSizeAnchor")

    def _write_from(self, vertex) -> None:
        # Write the <cdr:from> element.
        self._xml_start_tag("cdr:from")

        self._xml_data_element("cdr:x", vertex[0])
        self._xml_data_element("cdr:y", vertex[1])

        self._xml_end_tag("cdr:from")

    def _write_to(self, vertex) -> None:
        # Write the <cdr:to> element.
        self._xml_start_tag("cdr:to")

        self._xml_data_element("cdr:x", vertex[0])
        self._xml_data_element("cdr:y", vertex[1])

        self._xml_end_tag("cdr:to")

    def _write_sp(self, index: int, textbox) -> None:
        # Write the <cdr:sp> element.
        attributes = [("macro", ""), ("textlink", "")]

        self._xml_start_tag("cdr:sp", attributes)

        # Write the cdr:nvSpPr element.
        self._write_nv_sp_pr(index, textbox)

        # Write the cdr:spPr element.
        self._write_sp_pr()

        # Write the cdr:style element.
        self._write_style()

        # Write the cdr:txBody element.
        self._write_tx_body(textbox)

        self._xml_end_tag("cdr:sp")

    def _write_nv_sp_pr(self, index: int, textbox) -> None:
        # Write the <cdr:nvSpPr> element.
        self._xml_start_tag("cdr:nvSpPr")

        shape_id = index + 1
        name = textbox.get("name") or f"TextBox {shape_id}"

        # Write the cdr:cNvPr element.
        self._xml_empty_tag("cdr:cNvPr", [("id", shape_id), ("name", name)])

        # Write the cdr:cNvSpPr element.
        self._xml_empty_tag("cdr:cNvSpPr", [("txBox", "1")])

        self._xml_end_tag("cdr:nvSpPr")

    def _write_sp_pr(self) -> None:
        # Write the <cdr:spPr> element.
        self._xml_start_tag("cdr:spPr")

        # Write the a:prstGeom element.
        self._xml_start_tag("a:prstGeom", [("prst", "rect")])
        self._xml_empty_tag("a:avLst")
        self._xml_end_tag("a:prstGeom")

        # Write the a:noFill element.
        self._xml_empty_tag("a:noFill")

        # Write the a:ln element.
        self._xml_start_tag("a:ln", [("w", "9525"), ("cmpd", "sng")])
        self._xml_empty_tag("a:noFill")
        self._xml_end_tag("a:ln")

        self._xml_end_tag("cdr:spPr")

    def _write_style(self) -> None:
        # Write the <cdr:style> element.
        self._xml_start_tag("cdr:style")

        self._write_a_ln_ref()
        self._write_a_fill_ref()
        self._write_a_effect_ref()
        self._write_a_font_ref()

        self._xml_end_tag("cdr:style")

    def _write_a_ln_ref(self) -> None:
        # Write the <a:lnRef> element.
        self._xml_start_tag("a:lnRef", [("idx", "0")])
        self._write_a_scrgb_clr()
        self._xml_end_tag("a:lnRef")

    def _write_a_fill_ref(self) -> None:
        # Write the <a:fillRef> element.
        self._xml_start_tag("a:fillRef", [("idx", "0")])
        self._write_a_scrgb_clr()
        self._xml_end_tag("a:fillRef")

    def _write_a_effect_ref(self) -> None:
        # Write the <a:effectRef> element.
        self._xml_start_tag("a:effectRef", [("idx", "0")])
        self._write_a_scrgb_clr()
        self._xml_end_tag("a:effectRef")

    def _write_a_scrgb_clr(self) -> None:
        # Write the <a:scrgbClr> element.
        self._xml_empty_tag("a:scrgbClr", [("r", "0"), ("g", "0"), ("b", "0")])

    def _write_a_font_ref(self) -> None:
        # Write the <a:fontRef> element.
        self._xml_start_tag("a:fontRef", [("idx", "minor")])
        self._write_a_scheme_clr("dk1")
        self._xml_end_tag("a:fontRef")

    def _write_a_scheme_clr(self, val) -> None:
        # Write the <a:schemeClr> element.
        self._xml_empty_tag("a:schemeClr", [("val", val)])

    def _write_tx_body(self, textbox) -> None:
        # Write the <cdr:txBody> element.
        self._xml_start_tag("cdr:txBody")

        # Write the a:bodyPr element.
        attributes = [("wrap", "square"), ("rtlCol", "0"), ("anchor", "t")]
        self._xml_empty_tag("a:bodyPr", attributes)

        # Write the a:lstStyle element.
        self._xml_empty_tag("a:lstStyle")

        for paragraph in textbox["paragraphs"]:
            self._write_paragraph(paragraph)

        self._xml_end_tag("cdr:txBody")

    def _write_paragraph(self, paragraph) -> None:
        # Write the <a:p> element.
        self._xml_start_tag("a:p")

        align = paragraph.get("align")
        if align == "left":
            self._xml_empty_tag("a:pPr", [("algn", "l")])
        elif align == "center":
            self._xml_empty_tag("a:pPr", [("algn", "ctr")])
        elif align == "right":
            self._xml_empty_tag("a:pPr", [("algn", "r")])

        for run in paragraph["runs"]:
            self._write_run(run)

        self._xml_end_tag("a:p")

    def _write_run(self, run) -> None:
        # Write the <a:r> element.
        self._xml_start_tag("a:r")

        # pylint: disable=protected-access
        font = run["font"]
        style_attrs = Shape._get_font_style_attributes(font)
        latin_attrs = Shape._get_font_latin_attributes(font)
        style_attrs.insert(0, ("lang", font["lang"]))

        self._write_run_properties(font, style_attrs, latin_attrs)

        self._xml_data_element("a:t", run["text"])

        self._xml_end_tag("a:r")

    def _write_run_properties(self, font, style_attrs, latin_attrs) -> None:
        # Write the <a:rPr> element.
        has_color = font.get("color") is not None

        if latin_attrs or has_color:
            self._xml_start_tag("a:rPr", style_attrs)

            if has_color:
                self._write_a_solid_fill(font["color"])

            if latin_attrs:
                self._write_a_latin(latin_attrs)
                self._write_a_cs(latin_attrs)

            self._xml_end_tag("a:rPr")
        else:
            self._xml_empty_tag("a:rPr", style_attrs)

    def _write_a_solid_fill(self, color) -> None:
        # Write the <a:solidFill> element.
        self._xml_start_tag("a:solidFill")
        self._xml_empty_tag("a:srgbClr", [("val", color._rgb_hex_value())])
        self._xml_end_tag("a:solidFill")

    def _write_a_latin(self, attributes) -> None:
        # Write the <a:latin> element.
        self._xml_empty_tag("a:latin", attributes)

    def _write_a_cs(self, attributes) -> None:
        # Write the <a:cs> element.
        self._xml_empty_tag("a:cs", attributes)
