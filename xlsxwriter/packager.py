###############################################################################
#
# Packager - A class for writing the Excel XLSX Worksheet file.
#
# Copyright 2013, John McNamara, jmcnamara@cpan.org
#

# Standard packages.
import os
from shutil import copy

# Package imports.
from xlsxwriter.app import App
from xlsxwriter.contenttypes import ContentTypes
from xlsxwriter.core import Core
from xlsxwriter.relationships import Relationships
from xlsxwriter.sharedstrings import SharedStrings
from xlsxwriter.styles import Styles
from xlsxwriter.theme import Theme
from xlsxwriter.vml import Vml
from xlsxwriter.comments import Comments


class Packager(object):
    """
    A class for writing the Excel XLSX Packager file.

    This module is used in conjunction with XlsxWriter to create an
    Excel XLSX container file.

    From Wikipedia: The Open Packaging Conventions (OPC) is a
    container-file technology initially created by Microsoft to store
    a combination of XML and non-XML files that together form a single
    entity such as an Open XML Paper Specification (OpenXPS)
    document. http://en.wikipedia.org/wiki/Open_Packaging_Conventions.

    At its simplest an Excel XLSX file contains the following elements:

         ____ [Content_Types].xml
        |
        |____ docProps
        | |____ app.xml
        | |____ core.xml
        |
        |____ xl
        | |____ workbook.xml
        | |____ worksheets
        | | |____ sheet1.xml
        | |
        | |____ styles.xml
        | |
        | |____ theme
        | | |____ theme1.xml
        | |
        | |_____rels
        | |____ workbook.xml.rels
        |
        |_____rels
          |____ .rels

    The Packager class coordinates the classes that represent the
    elements of the package and writes them into the XLSX file.

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

        super(Packager, self).__init__()

        self.package_dir = ''
        self.workbook = None
        self.sheet_names = []
        self.worksheet_count = 0
        self.chartsheet_count = 0
        self.chart_count = 0
        self.drawing_count = 0
        self.table_count = 0
        self.named_ranges = []

    ###########################################################################
    #
    # Private API.
    #
    ###########################################################################

    def _set_package_dir(self, package_dir):
        # Set the XLSX OPC package directory.
        self.package_dir = package_dir

    def _add_workbook(self, workbook):
        # Add the Excel::Writer::XLSX::Workbook object to the package.
        self.workbook = workbook
        self.sheet_names = workbook.sheetnames
        self.chart_count = len(workbook.charts)
        self.drawing_count = len(workbook.drawings)
        self.num_vml_files = workbook.num_vml_files
        self.num_comment_files = workbook.num_comment_files
        self.named_ranges = workbook.named_ranges

        for worksheet in self.workbook.worksheets():
            if worksheet.is_chartsheet:
                self.chartsheet_count += 1
            else:
                self.worksheet_count += 1

    def _create_package(self):
        # Write the xml files that make up the XLSX OPC package.
        self._write_worksheet_files()
        self._write_chartsheet_files()
        self._write_workbook_file()
        self._write_chart_files()
        self._write_drawing_files()
        self._write_vml_files()
        self._write_comment_files()
        self._write_table_files()
        self._write_shared_strings_file()
        self._write_app_file()
        self._write_core_file()
        self._write_content_types_file()
        self._write_styles_file()
        self._write_theme_file()
        self._write_root_rels_file()
        self._write_workbook_rels_file()
        self._write_worksheet_rels_files()
        self._write_chartsheet_rels_files()
        self._write_drawing_rels_files()
        self._add_image_files()
        self._add_vba_project()

    def _write_workbook_file(self):
        # Write the workbook.xml file.
        xlsx_dir = self.package_dir
        workbook = self.workbook

        self._mkdir(xlsx_dir + '/xl')

        workbook._set_xml_writer(xlsx_dir + '/xl/workbook.xml')
        workbook._assemble_xml_file()

    def _write_worksheet_files(self):
        # Write the worksheet files.

        xlsx_dir = self.package_dir

        self._mkdir(xlsx_dir + '/xl')
        self._mkdir(xlsx_dir + '/xl/worksheets')

        index = 1
        for worksheet in self.workbook.worksheets():
            if worksheet.is_chartsheet:
                continue

            worksheet._set_xml_writer(
                xlsx_dir + '/xl/worksheets/sheet' + str(index) + '.xml')
            worksheet._assemble_xml_file()
            index += 1

    def _write_chartsheet_files(self):
        # Write the chartsheet files.

        xlsx_dir = self.package_dir

        index = 1
        for worksheet in self.workbook.worksheets():
            if not worksheet.is_chartsheet:
                continue

            self._mkdir(xlsx_dir + '/xl')
            self._mkdir(xlsx_dir + '/xl/chartsheets')

            worksheet._set_xml_writer(
                xlsx_dir + '/xl/chartsheets/sheet' + str(index) + '.xml')
            worksheet._assemble_xml_file()
            index += 1

    def _write_chart_files(self):
        # Write the chart files.

        xlsx_dir = self.package_dir

        if not self.workbook.charts:
            return

        self._mkdir(xlsx_dir + '/xl')
        self._mkdir(xlsx_dir + '/xl/charts')

        index = 1
        for chart in self.workbook.charts:
            chart._set_xml_writer(
                xlsx_dir + '/xl/charts/chart' + str(index) + '.xml')
            chart._assemble_xml_file()
            index += 1

    def _write_drawing_files(self):
        # Write the drawing files.

        xlsx_dir = self.package_dir

        if not self.drawing_count:
            return

        self._mkdir(xlsx_dir + '/xl')
        self._mkdir(xlsx_dir + '/xl/drawings')

        index = 1
        for drawing in self.workbook.drawings:
            drawing._set_xml_writer(
                xlsx_dir + '/xl/drawings/drawing' + str(index) + '.xml')
            drawing._assemble_xml_file()
            index += 1

    def _write_vml_files(self):
        # Write the comment VML files.
        xlsx_dir = self.package_dir

        index = 1
        for worksheet in self.workbook.worksheets():
            if not worksheet.has_vml:
                continue

            vml = Vml()

            self._mkdir(xlsx_dir + '/xl')
            self._mkdir(xlsx_dir + '/xl/drawings')

            vml._set_xml_writer(
               xlsx_dir + '/xl/drawings/vmlDrawing' + str(index) + '.vml')
            vml._assemble_xml_file(
               worksheet.vml_data_id,
               worksheet.vml_shape_id,
               worksheet.comments_array,
               worksheet.buttons_array)
            index += 1

    def _write_comment_files(self):
        # Write the comment files.
        xlsx_dir = self.package_dir

        index = 1
        for worksheet in self.workbook.worksheets():
            if not worksheet.has_comments:
                continue

            comment = Comments()

            self._mkdir(xlsx_dir + '/xl')
            self._mkdir(xlsx_dir + '/xl/drawings')

            comment._set_xml_writer(xlsx_dir + '/xl/comments'
                                    + str(index) + '.xml')
            comment._assemble_xml_file(worksheet.comments_array)
            index += 1

    def _write_shared_strings_file(self):
        # Write the sharedStrings.xml file.
        xlsx_dir = self.package_dir
        sst = SharedStrings()

        sst.string_table = self.workbook.str_table

        if not self.workbook.str_table.count:
            return

        self._mkdir(xlsx_dir + '/xl')

        sst._set_xml_writer(xlsx_dir + '/xl/sharedStrings.xml')
        sst._assemble_xml_file()

    def _write_app_file(self):
        # Write the app.xml file.
        xlsx_dir = self.package_dir
        properties = self.workbook.doc_properties
        app = App()

        self._mkdir(xlsx_dir + '/docProps')

        # Add the Worksheet heading pairs.
        app._add_heading_pair(['Worksheets', self.worksheet_count])

        # Add the Chartsheet heading pairs.
        app._add_heading_pair(['Charts', self.chartsheet_count])

        # Add the Worksheet parts.
        for worksheet in self.workbook.worksheets():
            if worksheet.is_chartsheet:
                continue
            app._add_part_name(worksheet.name)

        # Add the Chartsheet parts.
        for worksheet in self.workbook.worksheets():
            if not worksheet.is_chartsheet:
                continue
            app._add_part_name(worksheet.name)

        # Add the Named Range heading pairs.
        if self.named_ranges:
            app._add_heading_pair(['Named Ranges', len(self.named_ranges)])

        # Add the Named Ranges parts.
        for named_range in self.named_ranges:
            app._add_part_name(named_range)

        app._set_properties(properties)

        app._set_xml_writer(xlsx_dir + '/docProps/app.xml')
        app._assemble_xml_file()

    def _write_core_file(self):
        # Write the core.xml file.
        xlsx_dir = self.package_dir
        properties = self.workbook.doc_properties
        core = Core()

        self._mkdir(xlsx_dir + '/docProps')

        core._set_properties(properties)
        core._set_xml_writer(xlsx_dir + '/docProps/core.xml')
        core._assemble_xml_file()

    def _write_content_types_file(self):
        # Write the ContentTypes.xml file.
        xlsx_dir = self.package_dir
        content = ContentTypes()

        content._add_image_types(self.workbook.image_types)

        worksheet_index = 1
        chartsheet_index = 1
        for worksheet in self.workbook.worksheets():
            if worksheet.is_chartsheet:
                content._add_chartsheet_name('sheet' + str(chartsheet_index))
                chartsheet_index += 1
            else:
                content._add_worksheet_name('sheet' + str(worksheet_index))
                worksheet_index += 1

        for i in range(1, self.chart_count + 1):
            content._add_chart_name('chart' + str(i))

        for i in range(1, self.drawing_count + 1):
            content._add_drawing_name('drawing' + str(i))

        if self.num_vml_files:
            content._add_vml_name()

        for i in range(1, self.table_count + 1):
            content._add_table_name('table' + str(i))

        for i in range(1, self.num_comment_files + 1):
            content._add_comment_name('comments' + str(i))

        # Add the sharedString rel if there is string data in the workbook.
        if self.workbook.str_table.count:
            content._add_shared_strings()

        # Add vbaProject if present.
        if self.workbook.vba_project:
            content._add_vba_project()

        content._set_xml_writer(xlsx_dir + '/[Content_Types].xml')
        content._assemble_xml_file()

    def _write_styles_file(self):
        # Write the style xml file.
        xlsx_dir = self.package_dir
        xf_formats = self.workbook.xf_formats
        palette = self.workbook.palette
        font_count = self.workbook.font_count
        num_format_count = self.workbook.num_format_count
        border_count = self.workbook.border_count
        fill_count = self.workbook.fill_count
        custom_colors = self.workbook.custom_colors
        dxf_formats = self.workbook.dxf_formats

        styles = Styles()

        self._mkdir(xlsx_dir + '/xl')

        styles._set_style_properties([
            xf_formats,
            palette,
            font_count,
            num_format_count,
            border_count,
            fill_count,
            custom_colors,
            dxf_formats])

        styles._set_xml_writer(xlsx_dir + '/xl/styles.xml')
        styles._assemble_xml_file()

    def _write_theme_file(self):
        # Write the theme xml file.
        xlsx_dir = self.package_dir
        theme = Theme()

        self._mkdir(xlsx_dir + '/xl')
        self._mkdir(xlsx_dir + '/xl/theme')

        theme._set_xml_writer(xlsx_dir + '/xl/theme/theme1.xml')
        theme._assemble_xml_file()

    def _write_table_files(self):
        # Write the table files.
        # xlsx_dir = self.package_dir

        index = 1
        for worksheet in self.workbook.worksheets():
            table_props = worksheet.tables

            if not table_props:
                continue

            # self._mkdir(xlsx_dir + '/xl')
            # self._mkdir(xlsx_dir + '/xl/tables')

            # for table_props in table_props:
            #    table = Table()
            #    table._set_xml_writer(
            #        xlsx_dir + '/xl/tables/table' + str(index) + '.xml')
            #    table._set_properties(table_props)
            #    table._assemble_xml_file()
            #    self.table_count += 1
                index += 1

    def _write_root_rels_file(self):
        # Write the _rels/.rels xml file.

        xlsx_dir = self.package_dir
        rels = Relationships()

        self._mkdir(xlsx_dir + '/_rels')

        rels._add_document_relationship('/officeDocument', 'xl/workbook.xml')
        rels._add_package_relationship('/metadata/core-properties',
                                       'docProps/core.xml')
        rels._add_document_relationship('/extended-properties',
                                        'docProps/app.xml')

        rels._set_xml_writer(xlsx_dir + '/_rels/.rels')
        rels._assemble_xml_file()

    def _write_workbook_rels_file(self):
        # Write the _rels/.rels xml file.

        xlsx_dir = self.package_dir
        rels = Relationships()

        self._mkdir(xlsx_dir + '/xl')
        self._mkdir(xlsx_dir + '/xl/_rels')

        worksheet_index = 1
        chartsheet_index = 1

        for worksheet in self.workbook.worksheets():
            if worksheet.is_chartsheet:
                rels._add_document_relationship('/chartsheet',
                                                'chartsheets/sheet'
                                                + str(chartsheet_index)
                                                + '.xml')
                chartsheet_index += 1
            else:
                rels._add_document_relationship('/worksheet',
                                                'worksheets/sheet'
                                                + str(worksheet_index)
                                                + '.xml')
                worksheet_index += 1

        rels._add_document_relationship('/theme', 'theme/theme1.xml')
        rels._add_document_relationship('/styles', 'styles.xml')

        # Add the sharedString rel if there is string data in the workbook.
        if self.workbook.str_table.count:
            rels._add_document_relationship('/sharedStrings',
                                            'sharedStrings.xml')

        # Add vbaProject if present.
        if self.workbook.vba_project:
            rels._add_ms_package_relationship('/vbaProject', 'vbaProject.bin')

        rels._set_xml_writer(xlsx_dir + '/xl/_rels/workbook.xml.rels')
        rels._assemble_xml_file()

    def _write_worksheet_rels_files(self):
        # data such as hyperlinks or drawings.
        xlsx_dir = self.package_dir

        index = 0
        for worksheet in self.workbook.worksheets():

            if worksheet.is_chartsheet:
                continue

            index += 1

            external_links = (worksheet.external_hyper_links +
                              worksheet.external_drawing_links +
                              worksheet.external_vml_links +
                              worksheet.external_table_links +
                              worksheet.external_comment_links)

            if not external_links:
                continue

            # Create the worksheet .rels dirs.
            self._mkdir(xlsx_dir + '/xl')
            self._mkdir(xlsx_dir + '/xl/worksheets')
            self._mkdir(xlsx_dir + '/xl/worksheets/_rels')

            rels = Relationships()

            for link_data in external_links:
                rels._add_worksheet_relationship(*link_data)

            # Create .rels file such as /xl/worksheets/_rels/sheet1.xml.rels.
            rels._set_xml_writer(xlsx_dir + '/xl/worksheets/_rels/sheet'
                                 + str(index) + '.xml.rels')
            rels._assemble_xml_file()

    def _write_chartsheet_rels_files(self):
        # Write the chartsheet .rels files for links to drawing files.

        xlsx_dir = self.package_dir

        index = 0
        for worksheet in self.workbook.worksheets():

            if not worksheet.is_chartsheet:
                continue

            index += 1

            external_links = worksheet.external_drawing_links

            if not external_links:
                continue

            # Create the chartsheet .rels xlsx_dir.
            self._mkdir(xlsx_dir + '/xl')
            self._mkdir(xlsx_dir + '/xl/chartsheets')
            self._mkdir(xlsx_dir + '/xl/chartsheets/_rels')

            rels = Relationships()

            for link_data in external_links:
                rels._add_worksheet_relationship(link_data)

            # Create .rels file such as /xl/chartsheets/_rels/sheet1.xml.rels.
            rels._set_xml_writer(xlsx_dir + '/xl/chartsheets/_rels/sheet'
                                 + str(index) + '.xml.rels')
            rels._assemble_xml_file()

    def _write_drawing_rels_files(self):
        # Write the drawing .rels files for worksheets with charts or drawings.

        xlsx_dir = self.package_dir

        index = 0
        for worksheet in self.workbook.worksheets():
            if not worksheet.drawing_links:
                continue
            index += 1

            # Create the drawing .rels xlsx_dir.
            self._mkdir(xlsx_dir + '/xl')
            self._mkdir(xlsx_dir + '/xl/drawings')
            self._mkdir(xlsx_dir + '/xl/drawings/_rels')

            rels = Relationships()

            for drawing_data in worksheet.drawing_links:
                rels._add_document_relationship(*drawing_data)

            # Create .rels file such as /xl/drawings/_rels/sheet1.xml.rels.
            rels._set_xml_writer(xlsx_dir + '/xl/drawings/_rels/drawing'
                                 + str(index) + '.xml.rels')
            rels._assemble_xml_file()

    def _add_image_files(self):
        # Write the /xl/media/image?.xml files.
        xlsx_dir = self.package_dir
        workbook = self.workbook
        index = 1

        for image in workbook.images:
            filename = image[0]
            extension = '.' + image[1]

            self._mkdir(xlsx_dir + '/xl')
            self._mkdir(xlsx_dir + '/xl/media')

            copy(filename,
                 xlsx_dir + '/xl/media/image' + str(index) + extension)
            index += 1

    def _add_vba_project(self):
        # Write the vbaProject.bin file.
        # xlsx_dir = self.package_dir
        vba_project = self.workbook.vba_project

        if not vba_project:
            return

        # self._mkdir(xlsx_dir + '/xl')
        # copy(vba_project, xlsx_dir + '/xl/vbaProject.bin')

    def _mkdir(self, xlsx_dir):
        # TODO Wrap in try/catch.
        if not os.path.exists(xlsx_dir):
            os.makedirs(xlsx_dir)
