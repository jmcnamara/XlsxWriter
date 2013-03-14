###############################################################################
#
# Drawing - A class for writing the Excel XLSX Drawing file.
#
# Copyright 2013, John McNamara, jmcnamara@cpan.org
#

import xmlwriter


class Drawing(xmlwriter.XMLwriter):
    """
    A class for writing the Excel XLSX Drawing file.


    """

    ###########################################################################
    #
    # Public API.
    #
    ###########################################################################

    def __init__(self):
        """
        Constructor.

        """

        super(Drawing, self).__init__()

        self.drawings = []
        self.embedded = 0
        self.orientation = 0

    ###########################################################################
    #
    # Private API.
    #
    ###########################################################################

    def _assemble_xml_file(self):
        # Assemble and write the XML file.

        # Write the XML declaration.
        self._xml_declaration()

        # Write the xdr:wsDr element.
        self._write_drawing_workspace()

        if self.embedded:
            index = 1
            for dimensions in self.drawings:
                # Write the xdr:twoCellAnchor element.
                self._write_two_cell_anchor(index, dimensions)
                index += 1
        else:
            # Write the xdr:absoluteAnchor element.
            self._write_absolute_anchor(1)

        self._xml_end_tag('xdr:wsDr')

        # Close the file.
        self._xml_close()

    def _add_drawing_object(self, drawing_object):
        # Add a chart, image or shape sub object to the drawing.
        self.drawings.append(drawing_object)

    ###########################################################################
    #
    # XML methods.
    #
    ###########################################################################

    # Write the <xdr:wsDr> element.
    #
    def _write_drawing_workspace(self):
        schema = 'http://schemas.openxmlformats.org/drawingml/'
        xmlns_xdr = schema + '2006/spreadsheetDrawing'
        xmlns_a = schema + '2006/main'

        attributes = [
            ('xmlns:xdr', xmlns_xdr),
            ('xmlns:a', xmlns_a),
        ]

        self._xml_start_tag('xdr:wsDr', attributes)

    #
    #
    # Write the <xdr:twoCellAnchor> element.
    #
    def _write_two_cell_anchor(self, index, dimensions):
        anchor_type = dimensions[0]
        col_from = dimensions[1]
        row_from = dimensions[2]
        col_from_offset = dimensions[3]
        row_from_offset = dimensions[4]
        col_to = dimensions[5]
        row_to = dimensions[6]
        col_to_offset = dimensions[7]
        row_to_offset = dimensions[8]
        col_absolute = dimensions[9]
        row_absolute = dimensions[10]
        width = dimensions[11]
        height = dimensions[12]
        description = dimensions[13]
        shape = dimensions[14]

        attributes = []

        # Add attribute for images.
        if anchor_type == 2:
            attributes.append(('editAs', 'oneCell'))

        # Add editAs attribute for shapes.
        if shape and shape.editAs:
            attributes.append(('editAs', shape.editAs))

        self._xml_start_tag('xdr:twoCellAnchor', attributes)

        # Write the xdr:from element.
        self._write_from(
            col_from,
            row_from,
            col_from_offset,
            row_from_offset)

        # Write the xdr:from element.
        self._write_to(
            col_to,
            row_to,
            col_to_offset,
            row_to_offset)

        if anchor_type == 1:
            # Graphic frame.
            # Write the xdr:graphicFrame element for charts.
            self._write_graphic_frame(index, description)
        elif anchor_type == 2:
            # Write the xdr:pic element.
            self._write_pic(index, col_absolute, row_absolute, width,
                height, description)
        else:
            # Write the xdr:sp element for shapes.
            self._write_sp(index, col_absolute, row_absolute, width, height,
                shape)

        # Write the xdr:clientData element.
        self._write_client_data()

        self._xml_end_tag('xdr:twoCellAnchor')

    #
    #
    # Write the <xdr:absoluteAnchor> element.
    #
    def _write_absolute_anchor(self, index):
        self._xml_start_tag('xdr:absoluteAnchor')

        # Different co-ordinates for horizontal (= 0) and vertical (= 1).
        if self.orientation == 0:
            # Write the xdr:pos element.
            self._write_pos(0, 0)

            # Write the xdr:ext element.
            self._write_ext(9308969, 6078325)

        else:
            # Write the xdr:pos element.
            self._write_pos(0, -47625)

            # Write the xdr:ext element.
            self._write_ext(6162675, 6124575)

        # Write the xdr:graphicFrame element.
        self._write_graphic_frame(index)

        # Write the xdr:clientData element.
        self._write_client_data()

        self._xml_end_tag('xdr:absoluteAnchor')

    #
    #
    # Write the <xdr:from> element.
    #
    def _write_from(self, col, row, col_offset, row_offset):
        self._xml_start_tag('xdr:from')

        # Write the xdr:col element.
        self._write_col(col)

        # Write the xdr:colOff element.
        self._write_col_off(col_offset)

        # Write the xdr:row element.
        self._write_row(row)

        # Write the xdr:rowOff element.
        self._write_row_off(row_offset)

        self._xml_end_tag('xdr:from')

    #
    #
    # Write the <xdr:to> element.
    #
    def _write_to(self, col, row, col_offset, row_offset):
        self._xml_start_tag('xdr:to')

        # Write the xdr:col element.
        self._write_col(col)

        # Write the xdr:colOff element.
        self._write_col_off(col_offset)

        # Write the xdr:row element.
        self._write_row(row)

        # Write the xdr:rowOff element.
        self._write_row_off(row_offset)

        self._xml_end_tag('xdr:to')

    #
    #
    # Write the <xdr:col> element.
    #
    def _write_col(self, data):
        self._xml_data_element('xdr:col', data)

    #
    #
    # Write the <xdr:colOff> element.
    #
    def _write_col_off(self, data):
        self._xml_data_element('xdr:colOff', data)

    #
    #
    # Write the <xdr:row> element.
    #
    def _write_row(self, data):
        self._xml_data_element('xdr:row', data)

    #
    #
    # Write the <xdr:rowOff> element.
    #
    def _write_row_off(self, data):
        self._xml_data_element('xdr:rowOff', data)

    #
    #
    # Write the <xdr:pos> element.
    #
    def _write_pos(self, x, y):

        attributes = [('x', x), ('y', y)]

        self._xml_empty_tag('xdr:pos', attributes)

    #
    #
    # Write the <xdr:ext> element.
    #
    def _write_ext(self, cx, cy):

        attributes = [('cx', cx), ('cy', cy)]

        self._xml_empty_tag('xdr:ext', attributes)

    #
    #
    # Write the <xdr:graphicFrame> element.
    #
    def _write_graphic_frame(self, index, name):
        attributes = [('macro', '')]

        self._xml_start_tag('xdr:graphicFrame', attributes)

        # Write the xdr:nvGraphicFramePr element.
        self._write_nv_graphic_frame_pr(index, name)

        # Write the xdr:xfrm element.
        self._write_xfrm()

        # Write the a:graphic element.
        self._write_atag_graphic(index)

        self._xml_end_tag('xdr:graphicFrame')

    #
    #
    # Write the <xdr:nvGraphicFramePr> element.
    #
    def _write_nv_graphic_frame_pr(self, index, name):

        if not name:
            name = 'Chart ' + str(index)

        self._xml_start_tag('xdr:nvGraphicFramePr')

        # Write the xdr:cNvPr element.
        self._write_c_nv_pr(index + 1, name)

        # Write the xdr:cNvGraphicFramePr element.
        self._write_c_nv_graphic_frame_pr()

        self._xml_end_tag('xdr:nvGraphicFramePr')

    #
    #
    # Write the <xdr:cNvPr> element.
    #
    def _write_c_nv_pr(self, index, name, descr=None):

        attributes = [('id', index), ('name', name)]

        # Add description attribute for images.
        if descr is not None:
            attributes.append(('descr', descr))

        self._xml_empty_tag('xdr:cNvPr', attributes)

    #
    #
    # Write the <xdr:cNvGraphicFramePr> element.
    #
    def _write_c_nv_graphic_frame_pr(self):
        if self.embedded:
            self._xml_empty_tag('xdr:cNvGraphicFramePr')
        else:
            self._xml_start_tag('xdr:cNvGraphicFramePr')

            # Write the a:graphicFrameLocks element.
            self._write_a_graphic_frame_locks()

            self._xml_end_tag('xdr:cNvGraphicFramePr')

    #
    #
    # Write the <a:graphicFrameLocks> element.
    #
    def _write_a_graphic_frame_locks(self):
        attributes = [('noGrp', 1)]

        self._xml_empty_tag('a:graphicFrameLocks', attributes)

    #
    #
    # Write the <xdr:xfrm> element.
    #
    def _write_xfrm(self):
        self._xml_start_tag('xdr:xfrm')

        # Write the xfrmOffset element.
        self._write_xfrm_offset()

        # Write the xfrmOffset element.
        self._write_xfrm_extension()

        self._xml_end_tag('xdr:xfrm')

    #
    #
    # Write the <a:off> xfrm sub-element.
    #
    def _write_xfrm_offset(self):

        attributes = [
            ('x', 0),
            ('y', 0),
        ]

        self._xml_empty_tag('a:off', attributes)

    #
    #
    # Write the <a:ext> xfrm sub-element.
    #
    def _write_xfrm_extension(self):

        attributes = [
            ('cx', 0),
            ('cy', 0),
        ]

        self._xml_empty_tag('a:ext', attributes)

    #
    #
    # Write the <a:graphic> element.
    #
    def _write_atag_graphic(self, index):
        self._xml_start_tag('a:graphic')

        # Write the a:graphicData element.
        self._write_atag_graphic_data(index)

        self._xml_end_tag('a:graphic')

    #
    #
    # Write the <a:graphicData> element.
    #
    def _write_atag_graphic_data(self, index):
        uri = 'http://schemas.openxmlformats.org/drawingml/2006/chart'

        attributes = [('uri', uri,)]

        self._xml_start_tag('a:graphicData', attributes)

        # Write the c:chart element.
        self._write_c_chart('rId' + str(index))

        self._xml_end_tag('a:graphicData')

    #
    #
    # Write the <c:chart> element.
    #
    def _write_c_chart(self, r_id):

        schema = 'http://schemas.openxmlformats.org/'
        xmlns_c = schema + 'drawingml/2006/chart'
        xmlns_r = schema + 'officeDocument/2006/relationships'

        attributes = [
            ('xmlns:c', xmlns_c),
            ('xmlns:r', xmlns_r),
            ('r:id', r_id),
        ]

        self._xml_empty_tag('c:chart', attributes)

    #
    #
    # Write the <xdr:clientData> element.
    #
    def _write_client_data(self):
        self._xml_empty_tag('xdr:clientData')

    #
    #
    # Write the <xdr:sp> element.
    #
    def _write_sp(self, index, col_absolute, row_absolute,
                  width, height, shape):

        if shape and shape.connect:
            attributes = [('macro', '')]
            self._xml_start_tag('xdr:cxnSp', attributes)

            # Write the xdr:nvCxnSpPr element.
            self._write_nv_cxn_sp_pr(index, shape)

            # Write the xdr:spPr element.
            self._write_xdr_sp_pr(index, col_absolute, row_absolute, width,
                height, shape)

            self._xml_end_tag('xdr:cxnSp')
        else:
            # Add attribute for shapes.
            attributes = [('macro', ''), ('textlink', '')]
            self._xml_start_tag('xdr:sp', attributes)

            # Write the xdr:nvSpPr element.
            self._write_nv_sp_pr(index, shape)

            # Write the xdr:spPr element.
            self._write_xdr_sp_pr(index, col_absolute, row_absolute, width,
                height, shape)

            # Write the xdr:txBody element.
            if shape.text:
                self._write_txBody(col_absolute, row_absolute, width, height,
                    shape)

            self._xml_end_tag('xdr:sp')

    #
    #
    # Write the <xdr:nvCxnSpPr> element.
    #
    def _write_nv_cxn_sp_pr(self, index, shape):

        self._xml_start_tag('xdr:nvCxnSpPr')

        shape.name = shape.type + ' ' + index
        if shape.name is not None:
            self._write_c_nv_pr(shape.id, shape.name)

        self._xml_start_tag('xdr:cNvCxnSpPr')

        attributes = [('noChangeShapeType', '1')]
        self._xml_empty_tag('a:cxnSpLocks', attributes)

        if shape.start:
            attributes = [('id', shape.start), ('idx', shape.start_index)]
            self._xml_empty_tag('a:stCxn', attributes)

        if shape.end:
            attributes = [('id', shape.end), ('idx', shape.end_index)]
            self._xml_empty_tag('a:endCxn', attributes)
        self._xml_end_tag('xdr:cNvCxnSpPr')
        self._xml_end_tag('xdr:nvCxnSpPr')

    #
    #
    # Write the <xdr:NvSpPr> element.
    #
    def _write_nv_sp_pr(self, index, shape):

        attributes = []

        self._xml_start_tag('xdr:nvSpPr')

        shape_name = shape.type + ' ' + index

        self._write_c_nv_pr(shape.id, shape_name)

        if shape.txBox:
            attributes = [('txBox', 1)]

        self._xml_start_tag('xdr:cNvSpPr', attributes)

        attributes = [('noChangeArrowheads', '1')]

        self._xml_empty_tag('a:spLocks', attributes)

        self._xml_end_tag('xdr:cNvSpPr')
        self._xml_end_tag('xdr:nvSpPr')

    #
    #
    # Write the <xdr:pic> element.
    #
    def _write_pic(self, index, col_absolute, row_absolute,
                   width, height, description):

        self._xml_start_tag('xdr:pic')

        # Write the xdr:nvPicPr element.
        self._write_nv_pic_pr(index, description)

        # Write the xdr:blipFill element.
        self._write_blip_fill(index)

        # Pictures are rectangle shapes by default.
        shape = {'type': 'rect'}

        # Write the xdr:spPr element.
        self._write_sp_pr(col_absolute, row_absolute, width, height,
            shape)

        self._xml_end_tag('xdr:pic')

    #
    #
    # Write the <xdr:nvPicPr> element.
    #
    def _write_nv_pic_pr(self, index, description):

        self._xml_start_tag('xdr:nvPicPr')

        # Write the xdr:cNvPr element.
        self._write_c_nv_pr(index + 1, 'Picture ' + str(index), description)

        # Write the xdr:cNvPicPr element.
        self._write_c_nv_pic_pr()

        self._xml_end_tag('xdr:nvPicPr')

    #
    #
    # Write the <xdr:cNvPicPr> element.
    #
    def _write_c_nv_pic_pr(self):
        self._xml_start_tag('xdr:cNvPicPr')

        # Write the a:picLocks element.
        self._write_a_pic_locks()

        self._xml_end_tag('xdr:cNvPicPr')

    #
    #
    # Write the <a:picLocks> element.
    #
    def _write_a_pic_locks(self):
        attributes = [('noChangeAspect', 1)]

        self._xml_empty_tag('a:picLocks', attributes)

    #
    #
    # Write the <xdr:blipFill> element.
    #
    def _write_blip_fill(self, index):
        self._xml_start_tag('xdr:blipFill')

        # Write the a:blip element.
        self._write_a_blip(index)

        # Write the a:stretch element.
        self._write_a_stretch()

        self._xml_end_tag('xdr:blipFill')

    #
    #
    # Write the <a:blip> element.
    #
    def _write_a_blip(self, index):
        schema = 'http://schemas.openxmlformats.org/officeDocument/'
        xmlns_r = schema + '2006/relationships'
        r_embed = 'rId' + str(index)

        attributes = [
            ('xmlns:r', xmlns_r),
            ('r:embed', r_embed)]

        self._xml_empty_tag('a:blip', attributes)

    #
    #
    # Write the <a:stretch> element.
    #
    def _write_a_stretch(self):
        self._xml_start_tag('a:stretch')

        # Write the a:fillRect element.
        self._write_a_fill_rect()

        self._xml_end_tag('a:stretch')

    #
    #
    # Write the <a:fillRect> element.
    #
    def _write_a_fill_rect(self):
        self._xml_empty_tag('a:fillRect')

    #
    #
    # Write the <xdr:spPr> element, for charts.
    #
    def _write_sp_pr(self, col_absolute, row_absolute, width, height,
                     shape={}):

        self._xml_start_tag('xdr:spPr')

        # Write the a:xfrm element.
        self._write_a_xfrm(col_absolute, row_absolute, width, height)

        # Write the a:prstGeom element.
        self._write_a_prst_geom(shape)

        self._xml_end_tag('xdr:spPr')

    #
    #
    # Write the <xdr:spPr> element for shapes.
    #
    def _write_xdr_sp_pr(self, index, col_absolute, row_absolute, width,
                         height, shape={}):

        attributes = [('bwMode', 'auto')]

        self._xml_start_tag('xdr:spPr', attributes)

        # Write the a:xfrm element.
        self._write_a_xfrm(col_absolute, row_absolute, width, height,
            shape)

        # Write the a:prstGeom element.
        self._write_a_prst_geom(shape)

        fill = shape.fill

        if len(fill) > 1:

            # Write the a:solidFill element.
            self._write_a_solid_fill(fill)
        else:
            self._xml_empty_tag('a:noFill')

        # Write the a:ln element.
        self._write_a_ln(shape)

        self._xml_end_tag('xdr:spPr')

    #
    #
    # Write the <a:xfrm> element.
    #
    def _write_a_xfrm(self, col_absolute, row_absolute, width, height,
                      shape={}):
        attributes = []

        if "rotation" in shape:
            rotation = shape.rotation
            rotation *= 60000
            attributes.append(('rot', rotation))

        if 'flip_h' in shape:
            attributes.append(('flipH', 1))
        if 'flip_v' in shape:
            attributes.append(('flipV', 1))

        self._xml_start_tag('a:xfrm', attributes)

        # Write the a:off element.
        self._write_a_off(col_absolute, row_absolute)

        # Write the a:ext element.
        self._write_a_ext(width, height)

        self._xml_end_tag('a:xfrm')

    #
    #
    # Write the <a:off> element.
    #
    def _write_a_off(self, x, y):

        attributes = [
            ('x', x),
            ('y', y),
        ]

        self._xml_empty_tag('a:off', attributes)

    #
    #
    # Write the <a:ext> element.
    #
    def _write_a_ext(self, cx, cy):

        attributes = [
            ('cx', cx),
            ('cy', cy),
        ]

        self._xml_empty_tag('a:ext', attributes)

    #
    #
    # Write the <a:prstGeom> element.
    #
    def _write_a_prst_geom(self, shape={}):
        attributes = []

        if 'type' in shape:
            attributes = [('prst', shape['type'])]

        self._xml_start_tag('a:prstGeom', attributes)

        # Write the a:avLst element.
        self._write_a_av_lst(shape)

        self._xml_end_tag('a:prstGeom')

    #
    #
    # Write the <a:avLst> element.
    #
    def _write_a_av_lst(self, shape={}):
        adjustments = []

        if 'adjustments' in shape:
            adjustments = shape['adjustments']

        if adjustments:
            self._xml_start_tag('a:avLst')

            i = 0
            for adj in adjustments:
                i += 1
                # Only connectors have multiple adjustments.
                suffix = shape.connect  # TODO

                # Scale Adjustments: 100,000 = 100%.
                adj_int = int(adj * 1000)

                attributes = [('name', 'adj' + suffix),
                              ('fmla', 'val' + adj_int)]

                self._xml_empty_tag('a:gd', attributes)

            self._xml_end_tag('a:avLst')
        else:
            self._xml_empty_tag('a:avLst')

    #
    #
    # Write the <a:solidFill> element.
    #
    def _write_a_solid_fill(self, rgb):
        if not rgb is not None:
            rgb = '000000'

        attributes = [('val', rgb)]

        self._xml_start_tag('a:solidFill')

        self._xml_empty_tag('a:srgbClr', attributes)

        self._xml_end_tag('a:solidFill')

    #
    #
    # Write the <a:ln> element.
    #
    def _write_a_ln(self, shape={}):
        weight = shape.line_weight

        attributes = [('w', weight * 9525)]

        self._xml_start_tag('a:ln', attributes)

        line = shape.line

        if len(line) > 1:

            # Write the a:solidFill element.
            self._write_a_solid_fill(line)
        else:
            self._xml_empty_tag('a:noFill')

        if shape.line_type:

            attributes = [('val', shape.line_type)]
            self._xml_empty_tag('a:prstDash', attributes)

        if shape.connect:
            self._xml_empty_tag('a:round')
        else:
            attributes = [('lim', 800000)]
            self._xml_empty_tag('a:miter', attributes)

        self._xml_empty_tag('a:headEnd')
        self._xml_empty_tag('a:tailEnd')

        self._xml_end_tag('a:ln')

    #
    # _write_txBody
    #
    # Write the <xdr:txBody> element.
    #
    def _write_txBody(self, col_absolute, row_absolute, width, height, shape):

        attributes = [
            ('vertOverflow', "clip"),
            ('wrap', "square"),
            ('lIns', "27432"),
            ('tIns', "22860"),
            ('rIns', "27432"),
            ('bIns', "22860"),
            ('anchor', shape.valign),
            ('upright', "1"),
        ]

        self._xml_start_tag('xdr:txBody')
        self._xml_empty_tag('a:bodyPr', attributes)
        self._xml_empty_tag('a:lstStyle')

        self._xml_start_tag('a:p')

        rotation = shape.format.rotation
        if not rotation is not None:
            rotation = 0
        rotation *= 60000

        attributes = [('algn', shape.align), ('rtl', rotation)]
        self._xml_start_tag('a:pPr', attributes)

        attributes = [('sz', "1000")]
        self._xml_empty_tag('a:defRPr', attributes)

        self._xml_end_tag('a:pPr')
        self._xml_start_tag('a:r')

        size = shape.format.size
        if not size is not None:
            size = 8
        size *= 100

        bold = shape.format.bold
        if not bold is not None:
            bold = 0

        italic = shape.format.italic
        if not italic is not None:
            italic = 0

        underline = shape.format.underline
        underline = underline  # TODO ? 'sng': 'none'

        strike = shape.format.font_strikeout
        strike = strike  # TODO? 'Strike': 'noStrike'

        attributes = [
            ('lang', 'en-US'),
            ('sz', size),
            ('b', bold),
            ('i', italic),
            ('u', underline),
            ('strike', strike),
            ('baseline', 0),
        ]

        self._xml_start_tag('a:rPr', attributes)

        color = shape.format.color
        if color is not None:
            color = shape._get_palette_color(color)
            # color =~ s/^FF//; # Remove leading FF from rgb for shape color.
        else:
            color = '000000'

        self._write_a_solid_fill(color)

        font = shape.format.font
        if font is not None:
            font = 'Calibri'

        attributes = [('typeface', font)]
        self._xml_empty_tag('a:latin', attributes)

        self._xml_empty_tag('a:cs', attributes)

        self._xml_end_tag('a:rPr')

        self._xml_data_element('a:t', shape.text)

        self._xml_end_tag('a:r')
        self._xml_end_tag('a:p')
        self._xml_end_tag('xdr:txBody')
