###############################################################################
#
# Worksheet - A class for writing Excel Worksheets.
#
# Copyright 2013, John McNamara, jmcnamara@cpan.org
#

import xmlwriter


class Worksheet(xmlwriter.XMLwriter):
    """
    A class for writing Excel Worksheets.

    """

    ###########################################################################
    #
    # Public API.
    #

    ###########################################################################
    #
    # Private API.
    #

    def _assemble_xml_file(self):
        # Assemble and write the XML file.

        # Write the XML declaration.
        self._xml_declaration()

        # Write the worksheet element.
        self._write_worksheet()

        # Write the dimension element.
        self._write_dimension()

        # Write the sheetViews element.
        self._write_sheet_views()

        # Write the sheetFormatPr element.
        self._write_sheet_format_pr()

        # Write the sheetData element.
        self._write_sheet_data()

        # Write the pageMargins element.
        self._write_page_margins()

        # Close the worksheet tag.
        self._xml_end_tag('worksheet')

    ###########################################################################
    #
    # XML methods.
    #
    def _write_worksheet(self):
        # Write the <worksheet> element.
        schema = 'http://schemas.openxmlformats.org/'
        xmlns = schema + 'spreadsheetml/2006/main'
        xmlns_r = schema + 'officeDocument/2006/relationships'

        attributes = [
            ('xmlns', xmlns),
            ('xmlns:r', xmlns_r),
        ]

        self._xml_start_tag('worksheet', attributes)

    def _write_dimension(self):
        # Write the <dimension> element.
        ref = 'A1'

        attributes = [('ref', ref)]

        self._xml_empty_tag('dimension', attributes)

    def _write_sheet_views(self):
        # Write the <sheetViews> element.
        self._xml_start_tag('sheetViews')

        # Write the sheetView element.
        self._write_sheet_view()

        self._xml_end_tag('sheetViews')

    def _write_sheet_view(self):
        # Write the <sheetView> element.
        tab_selected = '1'
        workbook_view_id = '0'

        attributes = [
            ('tabSelected', tab_selected),
            ('workbookViewId', workbook_view_id),
        ]

        self._xml_empty_tag('sheetView', attributes)

    def _write_sheet_format_pr(self):
        # Write the <sheetFormatPr> element.
        default_row_height = '15'

        attributes = [('defaultRowHeight', default_row_height)]

        self._xml_empty_tag('sheetFormatPr', attributes)

    def _write_sheet_data(self):
        # Write the <sheetData> element.
        self._xml_empty_tag('sheetData')

    def _write_page_margins(self):
        # Write the <pageMargins> element.
        left = '0.7'
        right = '0.7'
        top = '0.75'
        bottom = '0.75'
        header = '0.3'
        footer = '0.3'

        attributes = [
            ('left',   left),
            ('right',  right),
            ('top',    top),
            ('bottom', bottom),
            ('header', header),
            ('footer', footer),
        ]

        self._xml_empty_tag('pageMargins', attributes)
