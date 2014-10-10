###############################################################################
#
# Drawing - A class for writing the Excel XLSX Drawing file.
#
# Copyright 2013-2014, John McNamara, jmcnamara@cpan.org
#

from . import xmlwriter


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
            for drawing in self.drawings:
                # Write the xdr:twoCellAnchor element.
                self._write_two_cell_anchor(index, drawing)
                index += 1

                if drawing['url']:
                    index += 1

        else:
            # Write the xdr:absoluteAnchor element.
            self._write_absolute_anchor(1)

        self._xml_end_tag('xdr:wsDr')

        # Close the file.
        self._xml_close()

    def _add_drawing_object(self, drawing_object):
        # Add a chart, image or shape sub object to the drawing.
        obj = {
            'anchor_type': drawing_object[0],
            'col_from': drawing_object[1],
            'row_from': drawing_object[2],
            'col_from_offset': drawing_object[3],
            'row_from_offset': drawing_object[4],
            'col_to': drawing_object[5],
            'row_to': drawing_object[6],
            'col_to_offset': drawing_object[7],
            'row_to_offset': drawing_object[8],
            'col_absolute': drawing_object[9],
            'row_absolute': drawing_object[10],
            'width': drawing_object[11],
            'height': drawing_object[12],
            'description': drawing_object[13],
            'shape': drawing_object[14],
            'url': None,
            'tip': None,
            'anchor': None
        }

        if len(drawing_object) > 15:
            obj['url'] = drawing_object[15]
            obj['tip'] = drawing_object[16]
            obj['anchor'] = drawing_object[17]

        self.drawings.append(obj)

    ###########################################################################
    #
    # XML methods.
    #
    ###########################################################################

    def _write_drawing_workspace(self):
        # Write the <xdr:wsDr> element.
        schema = 'http://schemas.openxmlformats.org/drawingml/'
        xmlns_xdr = schema + '2006/spreadsheetDrawing'
        xmlns_a = schema + '2006/main'

        attributes = [
            ('xmlns:xdr', xmlns_xdr),
            ('xmlns:a', xmlns_a),
        ]

        self._xml_start_tag('xdr:wsDr', attributes)

    def _write_two_cell_anchor(self, index, drawing):
        # Write the <xdr:twoCellAnchor> element.
        shape = drawing['shape']

        options = {
            'description': drawing['description'],
            'url': drawing['url'],
            'tip': drawing['tip']
        }

        attributes = []

        # Add attribute for images.
        if drawing['anchor_type'] == 2:
            if drawing['anchor'] == 3:
                attributes.append(('editAs', 'absolute'))
            elif drawing['anchor'] == 1:
                pass
            else:
                attributes.append(('editAs', 'oneCell'))

        # Add editAs attribute for shapes.
        if shape and shape.editAs:
            attributes.append(('editAs', shape.editAs))

        self._xml_start_tag('xdr:twoCellAnchor', attributes)

        # Write the xdr:from element.
        self._write_from(
            drawing['col_from'],
            drawing['row_from'],
            drawing['col_from_offset'],
            drawing['row_from_offset'])

        # Write the xdr:from element.
        self._write_to(
            drawing['col_to'],
            drawing['row_to'],
            drawing['col_to_offset'],
            drawing['row_to_offset'])

        if drawing['anchor_type'] == 1:
            # Graphic frame.
            # Write the xdr:graphicFrame element for charts.
            self._write_graphic_frame(index, drawing['description'])
        elif drawing['anchor_type'] == 2:
            # Write the xdr:pic element.
            self._write_pic(index,
                            drawing['col_absolute'],
                            drawing['row_absolute'],
                            drawing['width'],
                            drawing['height'],
                            options)
        else:
            # Write the xdr:sp element for shapes.
            self._write_sp(index,
                           drawing['col_absolute'],
                           drawing['row_absolute'],
                           drawing['width'],
                           drawing['height'],
                           shape)

        # Write the xdr:clientData element.
        self._write_client_data()

        self._xml_end_tag('xdr:twoCellAnchor')

    def _write_absolute_anchor(self, frame_index):
        self._xml_start_tag('xdr:absoluteAnchor')
        # Write the <xdr:absoluteAnchor> element.

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
        self._write_graphic_frame(frame_index)

        # Write the xdr:clientData element.
        self._write_client_data()

        self._xml_end_tag('xdr:absoluteAnchor')

    def _write_from(self, col, row, col_offset, row_offset):
        # Write the <xdr:from> element.
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

    def _write_to(self, col, row, col_offset, row_offset):
        # Write the <xdr:to> element.
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

    def _write_col(self, data):
        # Write the <xdr:col> element.
        self._xml_data_element('xdr:col', data)

    def _write_col_off(self, data):
        # Write the <xdr:colOff> element.
        self._xml_data_element('xdr:colOff', data)

    def _write_row(self, data):
        # Write the <xdr:row> element.
        self._xml_data_element('xdr:row', data)

    def _write_row_off(self, data):
        # Write the <xdr:rowOff> element.
        self._xml_data_element('xdr:rowOff', data)

    def _write_pos(self, x, y):
        # Write the <xdr:pos> element.

        attributes = [('x', x), ('y', y)]

        self._xml_empty_tag('xdr:pos', attributes)

    def _write_ext(self, cx, cy):
        # Write the <xdr:ext> element.

        attributes = [('cx', cx), ('cy', cy)]

        self._xml_empty_tag('xdr:ext', attributes)

    def _write_graphic_frame(self, index, name=None):
        # Write the <xdr:graphicFrame> element.
        attributes = [('macro', '')]

        self._xml_start_tag('xdr:graphicFrame', attributes)

        # Write the xdr:nvGraphicFramePr element.
        self._write_nv_graphic_frame_pr(index, name)

        # Write the xdr:xfrm element.
        self._write_xfrm()

        # Write the a:graphic element.
        self._write_atag_graphic(index)

        self._xml_end_tag('xdr:graphicFrame')

    def _write_nv_graphic_frame_pr(self, index, name):
        # Write the <xdr:nvGraphicFramePr> element.

        if not name:
            name = 'Chart ' + str(index)

        self._xml_start_tag('xdr:nvGraphicFramePr')

        # Write the xdr:cNvPr element.
        self._write_c_nv_pr(index + 1, name)

        # Write the xdr:cNvGraphicFramePr element.
        self._write_c_nv_graphic_frame_pr()

        self._xml_end_tag('xdr:nvGraphicFramePr')

    def _write_c_nv_pr(self, index, name, options={}):
        # Write the <xdr:cNvPr> element.
        descr = options.get('description', None)
        url = options.get('url', None)
        tip = options.get('tip', None)

        attributes = [('id', index), ('name', name)]

        # Add description attribute for images.
        if descr is not None:
            attributes.append(('descr', descr))

        if url:
            self._xml_start_tag('xdr:cNvPr', attributes)
            schema = "http://schemas.openxmlformats.org"
            att = [
                ('xmlns:r', schema + "/officeDocument/2006/relationships"),
                ('r:id', "rId" + str(index - 1))
            ]

            if tip:
                att.append(('tooltip', tip))

            self._xml_empty_tag('a:hlinkClick', att)
            self._xml_end_tag('xdr:cNvPr')
        else:
            self._xml_empty_tag('xdr:cNvPr', attributes)

    def _write_c_nv_graphic_frame_pr(self):
        # Write the <xdr:cNvGraphicFramePr> element.
        if self.embedded:
            self._xml_empty_tag('xdr:cNvGraphicFramePr')
        else:
            self._xml_start_tag('xdr:cNvGraphicFramePr')

            # Write the a:graphicFrameLocks element.
            self._write_a_graphic_frame_locks()

            self._xml_end_tag('xdr:cNvGraphicFramePr')

    def _write_a_graphic_frame_locks(self):
        # Write the <a:graphicFrameLocks> element.
        attributes = [('noGrp', 1)]

        self._xml_empty_tag('a:graphicFrameLocks', attributes)

    def _write_xfrm(self):
        # Write the <xdr:xfrm> element.
        self._xml_start_tag('xdr:xfrm')

        # Write the xfrmOffset element.
        self._write_xfrm_offset()

        # Write the xfrmOffset element.
        self._write_xfrm_extension()

        self._xml_end_tag('xdr:xfrm')

    def _write_xfrm_offset(self):
        # Write the <a:off> xfrm sub-element.

        attributes = [
            ('x', 0),
            ('y', 0),
        ]

        self._xml_empty_tag('a:off', attributes)

    def _write_xfrm_extension(self):
        # Write the <a:ext> xfrm sub-element.

        attributes = [
            ('cx', 0),
            ('cy', 0),
        ]

        self._xml_empty_tag('a:ext', attributes)

    def _write_atag_graphic(self, index):
        # Write the <a:graphic> element.
        self._xml_start_tag('a:graphic')

        # Write the a:graphicData element.
        self._write_atag_graphic_data(index)

        self._xml_end_tag('a:graphic')

    def _write_atag_graphic_data(self, index):
        # Write the <a:graphicData> element.
        uri = 'http://schemas.openxmlformats.org/drawingml/2006/chart'

        attributes = [('uri', uri,)]

        self._xml_start_tag('a:graphicData', attributes)

        # Write the c:chart element.
        self._write_c_chart('rId' + str(index))

        self._xml_end_tag('a:graphicData')

    def _write_c_chart(self, r_id):
        # Write the <c:chart> element.

        schema = 'http://schemas.openxmlformats.org/'
        xmlns_c = schema + 'drawingml/2006/chart'
        xmlns_r = schema + 'officeDocument/2006/relationships'

        attributes = [
            ('xmlns:c', xmlns_c),
            ('xmlns:r', xmlns_r),
            ('r:id', r_id),
        ]

        self._xml_empty_tag('c:chart', attributes)

    def _write_client_data(self):
        # Write the <xdr:clientData> element.
        self._xml_empty_tag('xdr:clientData')

    def _write_sp(self, index, col_absolute, row_absolute,
                  width, height, shape):
        # Write the <xdr:sp> element.

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
                self._write_tx_body(col_absolute, row_absolute, width, height,
                                    shape)

            self._xml_end_tag('xdr:sp')

    def _write_nv_cxn_sp_pr(self, index, shape):
        # Write the <xdr:nvCxnSpPr> element.
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

    def _write_nv_sp_pr(self, index, shape):
        # Write the <xdr:NvSpPr> element.
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

    def _write_pic(self, index, col_absolute, row_absolute,
                   width, height, options):
        # Write the <xdr:pic> element.
        self._xml_start_tag('xdr:pic')

        # Write the xdr:nvPicPr element.
        self._write_nv_pic_pr(index, options)

        # Write the xdr:blipFill element.
        if options.get('url', None):
            index = index + 1

        self._write_blip_fill(index)

        # Pictures are rectangle shapes by default.
        shape = {'type': 'rect'}

        # Write the xdr:spPr element.
        self._write_sp_pr(col_absolute, row_absolute, width, height,
                          shape)

        self._xml_end_tag('xdr:pic')

    def _write_nv_pic_pr(self, index, options):
        # Write the <xdr:nvPicPr> element.
        self._xml_start_tag('xdr:nvPicPr')

        # Write the xdr:cNvPr element.
        self._write_c_nv_pr(index + 1, 'Picture ' + str(index), options)

        # Write the xdr:cNvPicPr element.
        self._write_c_nv_pic_pr()

        self._xml_end_tag('xdr:nvPicPr')

    def _write_c_nv_pic_pr(self):
        # Write the <xdr:cNvPicPr> element.
        self._xml_start_tag('xdr:cNvPicPr')

        # Write the a:picLocks element.
        self._write_a_pic_locks()

        self._xml_end_tag('xdr:cNvPicPr')

    def _write_a_pic_locks(self):
        # Write the <a:picLocks> element.
        attributes = [('noChangeAspect', 1)]

        self._xml_empty_tag('a:picLocks', attributes)

    def _write_blip_fill(self, index):
        # Write the <xdr:blipFill> element.
        self._xml_start_tag('xdr:blipFill')

        # Write the a:blip element.
        self._write_a_blip(index)

        # Write the a:stretch element.
        self._write_a_stretch()

        self._xml_end_tag('xdr:blipFill')

    def _write_a_blip(self, index):
        # Write the <a:blip> element.
        schema = 'http://schemas.openxmlformats.org/officeDocument/'
        xmlns_r = schema + '2006/relationships'
        r_embed = 'rId' + str(index)

        attributes = [
            ('xmlns:r', xmlns_r),
            ('r:embed', r_embed)]

        self._xml_empty_tag('a:blip', attributes)

    def _write_a_stretch(self):
        # Write the <a:stretch> element.
        self._xml_start_tag('a:stretch')

        # Write the a:fillRect element.
        self._write_a_fill_rect()

        self._xml_end_tag('a:stretch')

    def _write_a_fill_rect(self):
        # Write the <a:fillRect> element.
        self._xml_empty_tag('a:fillRect')

    def _write_sp_pr(self, col_absolute, row_absolute, width, height,
                     shape={}):
        # Write the <xdr:spPr> element, for charts.

        self._xml_start_tag('xdr:spPr')

        # Write the a:xfrm element.
        self._write_a_xfrm(col_absolute, row_absolute, width, height)

        # Write the a:prstGeom element.
        self._write_a_prst_geom(shape)

        self._xml_end_tag('xdr:spPr')

    def _write_xdr_sp_pr(self, index, col_absolute, row_absolute, width,
                         height, shape={}):
        # Write the <xdr:spPr> element for shapes.

        attributes = [('bwMode', 'auto')]

        self._xml_start_tag('xdr:spPr', attributes)

        # Write the a:xfrm element.
        self._write_a_xfrm(col_absolute, row_absolute, width, height, shape)

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

    def _write_a_xfrm(self, col_absolute, row_absolute, width, height,
                      shape={}):
        # Write the <a:xfrm> element.
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

    def _write_a_off(self, x, y):
        # Write the <a:off> element.
        attributes = [
            ('x', x),
            ('y', y),
        ]

        self._xml_empty_tag('a:off', attributes)

    def _write_a_ext(self, cx, cy):
        # Write the <a:ext> element.
        attributes = [
            ('cx', cx),
            ('cy', cy),
        ]

        self._xml_empty_tag('a:ext', attributes)

    def _write_a_prst_geom(self, shape={}):
        # Write the <a:prstGeom> element.
        attributes = []

        if 'type' in shape:
            attributes = [('prst', shape['type'])]

        self._xml_start_tag('a:prstGeom', attributes)

        # Write the a:avLst element.
        self._write_a_av_lst(shape)

        self._xml_end_tag('a:prstGeom')

    def _write_a_av_lst(self, shape={}):
        # Write the <a:avLst> element.
        adjustments = []

        if 'adjustments' in shape:
            adjustments = shape['adjustments']

        if adjustments:
            self._xml_start_tag('a:avLst')

            i = 0
            for adj in adjustments:
                i += 1
                # Only connectors have multiple adjustments.
                if 'connect' in shape:
                    suffix = i
                else:
                    suffix = ''

                # Scale Adjustments: 100,000 = 100%.
                adj_int = str(int(adj * 1000))

                attributes = [('name', 'adj' + suffix),
                              ('fmla', 'val' + adj_int)]

                self._xml_empty_tag('a:gd', attributes)

            self._xml_end_tag('a:avLst')
        else:
            self._xml_empty_tag('a:avLst')

    def _write_a_solid_fill(self, rgb):
        # Write the <a:solidFill> element.
        if rgb is not None:
            rgb = '000000'

        attributes = [('val', rgb)]

        self._xml_start_tag('a:solidFill')

        self._xml_empty_tag('a:srgbClr', attributes)

        self._xml_end_tag('a:solidFill')

    def _write_a_ln(self, shape={}):
        # Write the <a:ln> element.
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

    def _write_tx_body(self, col_absolute, row_absolute, width, height, shape):
        # Write the <xdr:txBody> element.
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
        if rotation is not None:
            rotation = 0
        rotation *= 60000

        attributes = [('algn', shape.align), ('rtl', rotation)]
        self._xml_start_tag('a:pPr', attributes)

        attributes = [('sz', "1000")]
        self._xml_empty_tag('a:defRPr', attributes)

        self._xml_end_tag('a:pPr')
        self._xml_start_tag('a:r')

        size = shape.format.size
        if size is not None:
            size = 8
        size *= 100

        bold = shape.format.bold
        if bold is not None:
            bold = 0

        italic = shape.format.italic
        if italic is not None:
            italic = 0

        underline = shape['format']['underline']
        if underline:
            underline = 'sng'
        else:
            underline = 'none'

        strike = shape['format']['font_strikeout']
        if strike:
            strike = 'Strike'
        else:
            strike = 'noStrike'

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
