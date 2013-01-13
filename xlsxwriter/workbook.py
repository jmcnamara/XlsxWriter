###############################################################################
#
# Workbook - A class for writing the Excel XLSX Workbook file.
#
# Copyright 2013, John McNamara, jmcnamara@cpan.org
#

import xmlwriter
from worksheet import Worksheet


class Workbook(xmlwriter.XMLwriter):
    """
    A class for writing the Excel XLSX Workbook file.


    """

    ###########################################################################
    #
    # Public API.
    #
    ###########################################################################

    def __init__(self):
        """
        Construtor.

        """

        super(Workbook, self).__init__()

        self.filename = ''
        self.tempdir = None
        self.date_1904 = 0
        self.activesheet = 0
        self.firstsheet = 0
        self.selected = 0
        self.fileclosed = 0
        self.filehandle = None
        self.internal_fh = 0
        self.sheet_name = 'Sheet'
        self.chart_name = 'Chart'
        self.sheetname_count = 0
        self.chartname_count = 0
        self.worksheets = []
        self.charts = []
        self.drawings = []
        self.sheetnames = []
        self.formats = []
        self.xf_formats = []
        self.xf_format_indices = {}
        self.dxf_formats = []
        self.dxf_format_indices = {}
        self.palette = []
        self.font_count = 0
        self.num_format_count = 0
        self.defined_names = []
        self.named_ranges = []
        self.custom_colors = []
        self.doc_properties = {}
        self.localtime = []
        self.num_vml_files = 0
        self.num_comment_files = 0
        self.optimization = 0
        self.x_window = 240
        self.y_window = 15
        self.window_width = 16095
        self.window_height = 9660
        self.tab_ratio = 500
        self.table_count = 0
        self.str_table = {}
        self.vba_project = None
        self.vba_codename = None

    def add_worksheet(self, name):
        """
        Add a new worksheet to the Excel workbook.

        Args:
            name: The worksheet name. Defaults to 'Sheet1', etc.

        Returns:
            Reference to a worksheet object.

        """
        """
        TODO

        """

        index = len(self.worksheets)
        # TODO: name = self._check_sheetname(name)

        # fh = None
        #
        #        init_data = (
        #            fh,
        #            name,
        #            index,
        #
        #            self.activesheet,
        #            self.firstsheet,
        #
        #            self.str_total,
        #            self.str_unique,
        #            self.str_table,
        #
        #            self.table_count,
        #
        #            self.date_1904,
        #            self.palette,
        #            self.optimization,
        #            self.tempdir,
        #
        #            )

        init_data = {
            'name': name,
        }

        worksheet = Worksheet()
        worksheet._initialize(init_data)

        self.worksheets.append(worksheet)
        self.sheetnames.append(name)

        return worksheet

    ###########################################################################
    #
    # Private API.
    #
    ###########################################################################

    def _assemble_xml_file(self):
        # Assemble and write the XML file.

        # Write the XML declaration.
        self._xml_declaration()

        # Write the workbook element.
        self._write_workbook()

        # Write the fileVersion element.
        self._write_file_version()

        # Write the workbookPr element.
        self._write_workbook_pr()

        # Write the bookViews element.
        self._write_book_views()

        # Write the sheets element.
        self._write_sheets()

        # Write the calcPr element.
        self._write_calc_pr()

        # Close the workbook tag.
        self._xml_end_tag('workbook')

    ###########################################################################
    #
    # XML methods.
    #
    ###########################################################################

    def _write_workbook(self):
        # Write <workbook> element.

        schema = 'http://schemas.openxmlformats.org'
        xmlns = schema + '/spreadsheetml/2006/main'
        xmlns_r = schema + '/officeDocument/2006/relationships'

        attributes = [
            ('xmlns', xmlns),
            ('xmlns:r', xmlns_r),
        ]

        self._xml_start_tag('workbook', attributes)

    def _write_file_version(self):
        # Write the <fileVersion> element.

        app_name = 'xl'
        last_edited = 4
        lowest_edited = 4
        rup_build = 4505

        attributes = [
            ('appName', app_name),
            ('lastEdited', last_edited),
            ('lowestEdited', lowest_edited),
            ('rupBuild', rup_build),
        ]

        if self.vba_project:
            attributes.append(
                ('codeName', '{37E998C4-C9E5-D4B9-71C8-EB1FF731991C}'))

        self._xml_empty_tag('fileVersion', attributes)

    def _write_workbook_pr(self):
        # Write <workbookPr> element.
        default_theme_version = 124226
        attributes = []

        if self.vba_codename:
            attributes.append(('codeName', self.vba_codename))
        if self.date_1904:
            attributes.append(('date1904', 1))

        attributes.append(('defaultThemeVersion', default_theme_version))

        self._xml_empty_tag('workbookPr', attributes)

    def _write_book_views(self):
        # Write <bookViews> element.
        self._xml_start_tag('bookViews')
        self._write_workbook_view()
        self._xml_end_tag('bookViews')

    def _write_workbook_view(self):
        # Write <workbookView> element.
        attributes = [
            ('xWindow', self.x_window),
            ('yWindow', self.y_window),
            ('windowWidth', self.window_width),
            ('windowHeight', self.window_height),
        ]

        # Store the tabRatio attribute when it isn't the default.
        if self.tab_ratio != 500:
            attributes.append(('tabRatio', self.tab_ratio))

        # Store the firstSheet attribute when it isn't the default.
        if self.firstsheet > 0:
            attributes.append(('firstSheet', self.firstsheet))

        # Store the activeTab attribute when it isn't the first sheet.
        if self.activesheet > 0:
            attributes.append(('activeTab', self.activesheet))

        self._xml_empty_tag('workbookView', attributes)

    def _write_sheets(self):
        # Write <sheets> element.
        self._xml_start_tag('sheets')

        id_num = 1
        for worksheet in (self.worksheets):
            self._write_sheet(worksheet.name, id_num, worksheet.hidden)
            id_num += 1

        self._xml_end_tag('sheets')

    def _write_sheet(self, name, sheet_id, hidden):
        # Write <sheet> element.
        attributes = [
            ('name', name),
            ('sheetId', sheet_id),
        ]

        if hidden:
            attributes.append(('state', 'hidden'))

        attributes.append(('r:id', 'rId' + str(sheet_id)))

        self._xml_empty_tag('sheet', attributes)

    def _write_calc_pr(self):
        # Write the <calcPr> element.

        calc_id = '124519'

        attributes = [('calcId', calc_id)]

        self._xml_empty_tag('calcPr', attributes)
