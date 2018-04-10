###############################################################################
#
# Chartsheet - A class for writing the Excel XLSX Worksheet file.
#
# Copyright 2013-2018, John McNamara, jmcnamara@cpan.org
#

from . import worksheet
from .drawing import Drawing


class Chartsheet(worksheet.Worksheet):
    """
    A class for writing the Excel XLSX Chartsheet file.


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

        super(Chartsheet, self).__init__()

        self.is_chartsheet = True
        self.drawing = None
        self.chart = None
        self.charts = []
        self.zoom_scale_normal = 0
        self.orientation = 0
        self.protection = False

    def set_chart(self, chart):
        """
        Set the chart object for the chartsheet.
        Args:
            chart:  Chart object.
        Returns:
            chart:  A reference to the chart object.
        """
        chart.embedded = False
        chart.protection = self.protection
        self.chart = chart
        self.charts.append([0, 0, chart, 0, 0, 1, 1])
        return chart

    def protect(self, password='', options=None):
        """
        Set the password and protection options of the worksheet.

        Args:
            password: An optional password string.
            options:  A dictionary of worksheet objects to protect.

        Returns:
            Nothing.

        """
        # Overridden from parent worksheet class.
        if self.chart:
            self.chart.protection = True
        else:
            self.protection = True

        if not options:
            options = {}

        options = options.copy()

        options['sheet'] = False
        options['content'] = True
        options['scenarios'] = True

        # Call the parent method.
        super(Chartsheet, self).protect(password, options)

    ###########################################################################
    #
    # Private API.
    #
    ###########################################################################
    def _assemble_xml_file(self):
        # Assemble and write the XML file.

        # Write the XML declaration.
        self._xml_declaration()

        # Write the root worksheet element.
        self._write_chartsheet()

        # Write the worksheet properties.
        self._write_sheet_pr()

        # Write the sheet view properties.
        self._write_sheet_views()

        # Write the sheetProtection element.
        self._write_sheet_protection()

        # Write the printOptions element.
        self._write_print_options()

        # Write the worksheet page_margins.
        self._write_page_margins()

        # Write the worksheet page setup.
        self._write_page_setup()

        # Write the headerFooter element.
        self._write_header_footer()

        # Write the drawing element.
        self._write_drawings()

        # Close the worksheet tag.
        self._xml_end_tag('chartsheet')

        # Close the file.
        self._xml_close()

    def _prepare_chart(self, index, chart_id, drawing_id):
        # Set up chart/drawings.

        self.chart.id = chart_id - 1

        self.drawing = Drawing()
        self.drawing.orientation = self.orientation

        self.external_drawing_links.append(['/drawing',
                                            '../drawings/drawing'
                                            + str(drawing_id)
                                            + '.xml'])

        self.drawing_links.append(['/chart',
                                   '../charts/chart'
                                   + str(chart_id)
                                   + '.xml'])

    ###########################################################################
    #
    # XML methods.
    #
    ###########################################################################

    def _write_chartsheet(self):
        # Write the <worksheet> element. This is the root element.

        schema = 'http://schemas.openxmlformats.org/'
        xmlns = schema + 'spreadsheetml/2006/main'
        xmlns_r = schema + 'officeDocument/2006/relationships'

        attributes = [
            ('xmlns', xmlns),
            ('xmlns:r', xmlns_r)]

        self._xml_start_tag('chartsheet', attributes)

    def _write_sheet_pr(self):
        # Write the <sheetPr> element for Sheet level properties.
        attributes = []

        if self.filter_on:
            attributes.append(('filterMode', 1))

        if (self.fit_page or self.tab_color):
            self._xml_start_tag('sheetPr', attributes)
            self._write_tab_color()
            self._write_page_set_up_pr()
            self._xml_end_tag('sheetPr')
        else:
            self._xml_empty_tag('sheetPr', attributes)
