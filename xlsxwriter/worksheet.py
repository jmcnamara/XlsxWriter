###############################################################################
#
# Worksheet - A class for writing the Excel XLSX Worksheet file.
#
# Copyright 2013, John McNamara, jmcnamara@cpan.org
#

# Standard packages.
import re
from collections import defaultdict
from collections import namedtuple

# Package imports.
import xmlwriter
from utility import xl_rowcol_to_cell
from utility import xl_cell_to_rowcol


class Worksheet(xmlwriter.XMLwriter):
    """
    A class for writing the Excel XLSX Worksheet file.

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

        super(Worksheet, self).__init__()

        self.name = None
        self.index = None
        self.activesheet = 0
        self.firstsheet = 0
        self.str_table = None
        self.date_1904 = None
        self.palette = None
        self.optimization = 0
        self.tempdir = None

        self.ext_sheets = []
        self.fileclosed = 0
        self.excel_version = 2007

        self.xls_rowmax = 1048576
        self.xls_colmax = 16384
        self.xls_strmax = 32767
        self.dim_rowmin = None
        self.dim_rowmax = None
        self.dim_colmin = None
        self.dim_colmax = None

        self.colinfo = []
        self.selections = []
        self.hidden = 0
        self.active = 0
        self.tab_color = 0

        self.panes = []
        self.active_pane = 3
        self.selected = 0

        self.page_setup_changed = 0
        self.paper_size = 0
        self.orientation = 1

        self.print_options_changed = 0
        self.hcenter = 0
        self.vcenter = 0
        self.print_gridlines = 0
        self.screen_gridlines = 1
        self.print_headers = 0

        self.header_footer_changed = 0
        self.header = ''
        self.footer = ''

        self.margin_left = 0.7
        self.margin_right = 0.7
        self.margin_top = 0.75
        self.margin_bottom = 0.75
        self.margin_header = 0.3
        self.margin_footer = 0.3

        self.repeat_rows = ''
        self.repeat_cols = ''
        self.print_area = ''

        self.page_order = 0
        self.black_white = 0
        self.draft_quality = 0
        self.print_comments = 0
        self.page_start = 0

        self.fit_page = 0
        self.fit_width = 0
        self.fit_height = 0

        self.hbreaks = []
        self.vbreaks = []

        self.protect = 0
        self.password = None

        self.set_cols = {}
        self.set_rows = {}

        self.zoom = 100
        self.zoom_scale_normal = 1
        self.print_scale = 100
        self.right_to_left = 0
        self.show_zeros = 1
        self.leading_zeros = 0

        self.outline_row_level = 0
        self.outline_col_level = 0
        self.outline_style = 0
        self.outline_below = 1
        self.outline_right = 1
        self.outline_on = 1
        self.outline_changed = 0

        self.default_row_height = 15
        self.default_row_zeroed = 0

        self.names = {}
        self.write_match = []
        self.table = defaultdict(dict)
        self.merge = []
        self.row_spans = {}

        self.has_vml = 0
        self.has_comments = 0
        self.comments = {}
        self.comments_array = []
        self.comments_author = ''
        self.comments_visible = 0
        self.vml_shape_id = 1024
        self.buttons_array = []

        self.autofilter = ''
        self.filter_on = 0
        self.filter_range = []
        self.filter_cols = {}

        self.col_sizes = {}
        self.row_sizes = {}
        self.col_formats = {}
        self.col_size_changed = 0
        self.row_size_changed = 0

        self.last_shape_id = 1
        self.rel_count = 0
        self.hlink_count = 0
        self.hlink_refs = []
        self.external_hyper_links = []
        self.external_drawing_links = []
        self.external_comment_links = []
        self.external_vml_links = []
        self.external_table_links = []
        self.drawing_links = []
        self.charts = []
        self.images = []
        self.tables = []
        self.sparklines = []
        self.shapes = []
        self.shape_hash = {}
        self.drawing = 0

        self.rstring = ''
        self.previous_row = 0

        self.validations = []
        self.cond_formats = {}
        self.dxf_priority = 1
        self.is_chartsheet = 0
        self.page_view = 0

    def write(self, *args):
        """TODO Temp code."""
        is_number = re.compile(r'^[+-]?(?=\d|\.\d)\d*(\.\d*)?([Ee][+-]?\d+)?$')

        if not args[0].isdigit():
            (row, col, _, _) = xl_cell_to_rowcol(args[0])
            token = args[1] if len(args) > 1 else None
            xf_format = args[2] if len(args) > 2 else None
            # options = args[3] if len(args) > 3 else None
        else:
            row = args[0]
            col = args[1]
            token = args[2] if len(args) > 2 else None
            xf_format = args[3] if len(args) > 3 else None
            # options = args[4] if len(args) > 4 else None

        if is_number.match(str(token)):
            self.write_number(row, col, token, xf_format)
        else:
            self.write_string(row, col, token, xf_format)

    def write_string(self, row, col, string, cell_format=None):
        """
        Write a string to a worksheet cell.

        Args:
            row:    The cell row (zero indexed).
            col:    The cell column (zero indexed).
            string: Cell data. Str.
            format: An optional cell Format object.

        Returns:
            0:  Success.
            -1: Column number is out of worksheet bounds.
            -2: String truncated to 32k characters.

        """
        str_error = 0

        # Check that row and col are valid and store max and min values.
        if self._check_dimensions(row, col):
            return -1

        # Check that the string is < 32767 chars.
        if len(string) > self.xls_strmax:
            string = string[:self.xls_strmax]
            str_error = -2

        # Write a shared string or an in-line string in optimisation mode.
        if self.optimization == 0:
            string_index = self.str_table._get_shared_string_index(string)
        else:
            string_index = string

        # Write previous row if in in-line string optimization mode.
        if self.optimization and row > self.previous_row:
            self._write_single_row(row)

        # Store the cell data in the worksheet data table.
        cell_tuple = namedtuple('String', 'string, format')
        self.table[row][col] = cell_tuple(string_index, cell_format)

        return str_error

    def write_number(self, row, col, number, cell_format=None):
        """
        Write a number to a worksheet cell.

        Args:
            row:         The cell row (zero indexed).
            col:         The cell column (zero indexed).
            number:      Cell data. Int or float.
            cell_format: An optional cell Format object.

        Returns:
            0:  Success.
            -1: Column number is out of worksheet bounds.

        """
        # Check that row and col are valid and store max and min values.
        if self._check_dimensions(row, col):
            return -1

        # Write previous row if in in-line string optimization mode.
        if self.optimization and row > self.previous_row:
            self._write_single_row(row)

        # Store the cell data in the worksheet data table.
        cell_tuple = namedtuple('Number', 'number, format')
        self.table[row][col] = cell_tuple(number, cell_format)

        return 0

    def write_blank(self, row, col, cell_format=None):
        """
        Write a number to a worksheet cell.

        Args:
            row:         The cell row (zero indexed).
            col:         The cell column (zero indexed).
            cell_format: An optional cell Format object.

        Returns:
            0:  Success.
            -1: Column number is out of worksheet bounds.

        """
        # Don't write a blank cell unless it has a format.
        if cell_format is None:
            return 0

        # Check that row and col are valid and store max and min values.
        if self._check_dimensions(row, col):
            return -1

        # Write previous row if in in-line string optimization mode.
        if self.optimization and row > self.previous_row:
            self._write_single_row(row)

        # Store the cell data in the worksheet data table.
        cell_tuple = namedtuple('Blank', 'format')
        self.table[row][col] = cell_tuple(cell_format)

        return 0

    def write_formula(self, row, col, formula, cell_format=None, value=0):
        """
        Write a formula to a worksheet cell.

        Args:
            row:         The cell row (zero indexed).
            col:         The cell column (zero indexed).
            formual:     Cell formula.
            cell_format: An optional cell Format object.
            value:       An optional value for the formula. Default is 0.

        Returns:
            0:  Success.
            -1: Column number is out of worksheet bounds.

        """
        # Check that row and col are valid and store max and min values.
        if self._check_dimensions(row, col):
            return -1

        # Remove the formula '=' sign if it exists.
        if formula[0] == '=':
            formula = formula[1:]

        # Write previous row if in in-line string optimization mode.
        if self.optimization and row > self.previous_row:
            self._write_single_row(row)

        # Store the cell data in the worksheet data table.
        cell_tuple = namedtuple('Formula', 'formula, format, value')
        self.table[row][col] = cell_tuple(formula, cell_format, value)

        return 0

    def write_array_formula(self, firstrow, firstcol, lastrow, lastcol,
                            formula, cell_format=None, value=0):
        """
        Write a formula to a worksheet cell.

        Args:
            firstrow:    The first row of the cell range. (zero indexed).
            firstcol:    The first column of the cell range.
            lastrow:     The last row of the cell range. (zero indexed).
            lastcol:     The last column of the cell range.
            formuala:    Cell formula.
            cell_format: An optional cell Format object.
            value:       An optional value for the formula. Default is 0.

        Returns:
            0:  Success.
            -1: Column number is out of worksheet bounds.

        """

        # Swap last row/col with first row/col as necessary.
        if firstrow > lastrow:
            firstrow, lastrow = lastrow, firstrow
        if firstcol > lastcol:
            firstcol, lastcol = lastcol, firstcol

        # Check that row and col are valid and store max and min values
        if self._check_dimensions(lastrow, lastcol):
            return -1

        # Define array range
        if firstrow == lastrow and firstcol == lastcol:
            cell_range = xl_rowcol_to_cell(firstrow, firstcol)
        else:
            cell_range = (xl_rowcol_to_cell(firstrow, firstcol) + ':'
                          + xl_rowcol_to_cell(lastrow, lastcol))

        # Remove array formula braces and the leading =.
        if formula[0] == '{':
            formula = formula[1:]
        if formula[0] == '=':
            formula = formula[1:]
        if formula[-1] == '}':
            formula = formula[:-1]

        # Write previous row if in in-line string optimization mode.
        if self.optimization and firstrow > self.previous_row:
            self._write_single_row(firstrow)

        # Store the cell data in the worksheet data table.
        cell_tuple = namedtuple('ArrayFormula',
                                'formula, format, value, range')
        self.table[firstrow][firstcol] = cell_tuple(formula, cell_format,
                                                    value, cell_range)

        # Pad out the rest of the area with formatted zeroes.
        if not self.optimization:
            for row in range(firstrow, lastrow + 1):
                for col in range(firstcol, lastcol + 1):
                    if row != firstrow or col != firstcol:
                        self.write_number(row, col, 0, cell_format)

        return 0

    def select(self):
        """
        Set this worksheet as a selected worksheet, i.e. the worksheet
        has its tab highlighted.

        Note: A selected worksheet cannot be hidden.

        Args:
            None.

        Returns:
            Nothing.

        """
        self.selected = 1
        self.hidden = 0

    def set_column(self, firstcol, lastcol, width,
                   cell_format=None, hidden=False, level=0):
        """
        Set the width, and other properties of a single column or a
        range of columns.

        Args:
            firstcol: First column (zero-indexed).
            lastcol:  Last column (zero-indexed). Can be the same as firstcol.
            width:    Column width.
            cell_format:   Column cell_format. (optional).
            hidden:   Column is hidden. (default 0).
            level:    Column outline level. (default 0).

        Returns:
            0:  Success.
            -1: Column number is out of worksheet bounds.

        """
        # Ensure 2nd col is larger than first.
        if firstcol > lastcol:
            (firstcol, lastcol) = (lastcol, firstcol)

        # Don't modify the row dimensions when checking the columns.
        ignore_row = 1

        # Store the column dimension only in some conditions.
        if cell_format or (width and hidden):
            ignore_col = 0
        else:
            ignore_col = 1

        # Check that each column is valid and store the max and min values.
        if self._check_dimensions(0, lastcol, ignore_row, ignore_col):
            return -1
        if self._check_dimensions(0, firstcol, ignore_row, ignore_col):
            return -1

        # Set the limits for the outline levels (0 <= x <= 7).
        if level < 0:
            level = 0
        if level > 7:
            level = 7

        if level > self.outline_col_level:
            self.outline_col_level = level

        # Store the column data.
        self.colinfo.append([firstcol, lastcol, width, cell_format, hidden,
                             level])

        # Store the column change to allow optimisations.
        self.col_size_changed = 1

        # Store the col sizes for use when calculating image vertices taking
        # hidden columns into account. Also store the column formats.

        # Set width to zero if col is hidden
        if hidden:
            width = 0

        for col in range(firstcol, lastcol + 1):
            self.col_sizes[col] = width
            if cell_format:
                self.col_formats[col] = cell_format

        return 0

    def set_row(self, row, height, cell_format=None, hidden=False, level=0,
                collapsed=0):
        """
        Set the width, and other properties of a row.
        range of columns.

        Args:
            row:         Row number (zero-indexed).
            height:      Row width.
            cell_format: Row cell_format. (optional).
            hidden:      Row is hidden. (default 0).
            level:       Row outline level. (default 0).
            collapsed:   Row outline levels are collapsed. (default 0).
        Returns:
            0:  Success.
            -1: Row number is out of worksheet bounds.

        """
        # Use minimum col in _check_dimensions().
        if self.dim_colmin is not None:
            min_col = self.dim_colmin
        else:
            min_col = 0

        # Check that row is valid.
        if self._check_dimensions(row, min_col):
            return -1

        if height is None:
            height = self.default_row_height

        # If the height is 0 the row is hidden and the height is the default.
        if height == 0:
            hidden = 1
            height = self.default_row_height

        # Set the limits for the outline levels (0 <= x <= 7).
        if level < 0:
            level = 0
        if level > 7:
            level = 7

        if level > self.outline_row_level:
            self.outline_row_level = level

        # Store the row properties.
        self.set_rows[row] = [height, cell_format, hidden, level, collapsed]

        # Store the row change to allow optimisations.
        self.row_size_changed = 1

        # Store the row sizes for use when calculating image vertices.
        self.row_sizes[row] = height

    ###########################################################################
    #
    # Private API.
    #
    ###########################################################################

    def _initialize(self, init_data):
        self.name = init_data['name']
        self.index = init_data['index']
        self.str_table = init_data['str_table']

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

        # Write the cols element.
        self._write_cols()

        # Write the sheetData element.
        self._write_sheet_data()

        # Write the pageMargins element.
        self._write_page_margins()

        # Close the worksheet tag.
        self._xml_end_tag('worksheet')

        # Close the file.
        self._xml_close()

    def _check_dimensions(self, row, col, ignore_row=False, ignore_col=False):
        # Check that row and col are valid and store the max and min
        # values for use in other methods/elements. The ignore_row /
        # ignore_col flags is used to indicate that we wish to perform
        # the dimension check without storing the value. The ignore
        # flags are use by set_row() and data_validate.

        # Check that the row/col are within the worksheet bounds.
        if row >= self.xls_rowmax or col >= self.xls_colmax:
            return -1

        # In optimization mode we don't change dimensions for rows
        # that are already written.
        if not ignore_row and not ignore_col and self.optimization == 1:
            if row < self.previous_row:
                return -1

        if not ignore_row:
            if self.dim_rowmin is None or row < self.dim_rowmin:
                self.dim_rowmin = row
            if self.dim_rowmax is None or row > self.dim_rowmax:
                self.dim_rowmax = row

        if not ignore_col:
            if self.dim_colmin is None or col < self.dim_colmin:
                self.dim_colmin = col
            if self.dim_colmax is None or col > self.dim_colmax:
                self.dim_colmax = col

        return 0

    ###########################################################################
    #
    # XML methods.
    #
    ###########################################################################

    def _write_worksheet(self):
        # Write the <worksheet> element. This is the root element.

        schema = 'http://schemas.openxmlformats.org/'
        xmlns = schema + 'spreadsheetml/2006/main'
        xmlns_r = schema + 'officeDocument/2006/relationships'
        xmlns_mc = schema + 'markup-compatibility/2006'
        ms_schema = 'http://schemas.microsoft.com/'
        xmlns_x14ac = ms_schema + 'office/spreadsheetml/2009/9/ac'

        attributes = [
            ('xmlns', xmlns),
            ('xmlns:r', xmlns_r),
        ]

        # Add some extra attributes for Excel 2010. Mainly for sparklines.
        if self.excel_version == 2010:
            attributes.append(('xmlns:mc', xmlns_mc))
            attributes.append(('xmlns:x14ac', xmlns_x14ac))
            attributes.append(('mc:Ignorable', 'x14ac'))

        self._xml_start_tag('worksheet', attributes)

    def _write_dimension(self):
        # Write the <dimension> element. This specifies the range of
        # cells in the worksheet. As a special case, empty
        # spreadsheets use 'A1' as a range.

        if self.dim_rowmin is None and self.dim_colmin is None:
            # If the min dimensions are not defined then no dimensions
            # have been set and we use the default 'A1'.
            ref = 'A1'

        elif self.dim_rowmin is None and self.dim_colmin is not None:
            # If the row dimensions aren't set but the column
            # dimensions are set then they have been changed via
            # set_column().

            if self.dim_colmin == self.dim_colmax:
                # The dimensions are a single cell and not a range.
                ref = xl_rowcol_to_cell(0, self.dim_colmin)
            else:
                # The dimensions are a cell range.
                cell_1 = xl_rowcol_to_cell(0, self.dim_colmin)
                cell_2 = xl_rowcol_to_cell(0, self.dim_colmax)
                ref = cell_1 + ':' + cell_2

        elif (self.dim_rowmin == self.dim_rowmax and
              self.dim_colmin == self.dim_colmax):
            # The dimensions are a single cell and not a range.
            ref = xl_rowcol_to_cell(self.dim_rowmin, self.dim_colmin)
        else:
            # The dimensions are a cell range.
            cell_1 = xl_rowcol_to_cell(self.dim_rowmin, self.dim_colmin)
            cell_2 = xl_rowcol_to_cell(self.dim_rowmax, self.dim_colmax)
            ref = cell_1 + ':' + cell_2

        self._xml_empty_tag('dimension', [('ref', ref)])

    def _write_sheet_views(self):
        # Write the <sheetViews> element.
        self._xml_start_tag('sheetViews')

        # Write the sheetView element.
        self._write_sheet_view()

        self._xml_end_tag('sheetViews')

    def _write_sheet_view(self):
        # Write the <sheetViews> element.
        attributes = []

        # Hide screen gridlines if required
        if not self.screen_gridlines:
            attributes.append(('showGridLines', 0))

        # Hide zeroes in cells.
        if not self.show_zeros:
            attributes.append(('showZeros', 0))

        # Display worksheet right to left for Hebrew, Arabic and others.
        if self.right_to_left:
            attributes.append(('rightToLeft', 1))

        # Show that the sheet tab is selected.
        if self.selected:
            attributes.append(('tabSelected', 1))

        # Turn outlines off. Also required in the outlinePr element.
        if not self.outline_on:
            attributes.append(("showOutlineSymbols", 0))

        # Set the page view/layout mode if required.
        # TODO. Add pageBreakPreview mode when requested.
        if self.page_view:
            attributes.append(('view', 'pageLayout'))

        # Set the zoom level.
        if self.zoom != 100:
            if not self.page_view:
                attributes.append(('zoomScale', self.zoom))
                if self.zoom_scale_normal:
                    attributes.append(('zoomScaleNormal', self.zoom))

        attributes.append(('workbookViewId', 0))

        if self.panes or len(self.selections):
            self._xml_start_tag('sheetView', attributes)
            # self._write_panes()
            # self._write_selections()
            self._xml_end_tag('sheetView')
        else:
            self._xml_empty_tag('sheetView', attributes)

    def _write_sheet_format_pr(self):
        # Write the <sheetFormatPr> element.
        default_row_height = self.default_row_height

        attributes = [('defaultRowHeight', default_row_height)]

        self._xml_empty_tag('sheetFormatPr', attributes)

    def _write_cols(self):
        # Write the <cols> element and <col> sub elements.

        # Exit unless some column have been formatted.
        if not self.colinfo:
            return

        self._xml_start_tag('cols')

        for col_info in self.colinfo:
            self._write_col_info(col_info)

        self._xml_end_tag('cols')

    def _write_col_info(self, col_info):
        # Write the <col> element.

        col_min, col_max, width, cell_format, hidden, level = col_info
        collapsed = 0
        custom_width = 1
        xf_index = 0

        # Get the cell_format index.
        if cell_format:
            xf_index = cell_format._get_xf_index()

        # Set the Excel default column width.
        if width is None:
            if not hidden:
                width = 8.43
                custom_width = 0
            else:
                width = 0
        elif width == 8.43:
            # Width is defined but same as default.
            custom_width = 0

        # Convert column width from user units to character width.
        if width > 0:
            # For Calabri 11.
            max_digit_width = 7
            padding = 5
            width = int((float(width) * max_digit_width + padding)
                        / max_digit_width * 256.0) / 256.0

        attributes = [
            ('min', col_min + 1),
            ('max', col_max + 1),
            ('width', width),
        ]

        if xf_index:
            attributes.append(('style', xf_index))
        if hidden:
            attributes.append(('hidden', '1'))
        if custom_width:
            attributes.append(('customWidth', '1'))
        if level:
            attributes.append(('outlineLevel', level))
        if collapsed:
            attributes.append(('collapsed', '1'))

        self._xml_empty_tag('col', attributes)

    def _write_sheet_data(self):
        # Write the <sheetData> element.

        if self.dim_rowmin is None:
            # If the dimensions aren't defined there is no data to write.
            self._xml_empty_tag('sheetData')
        else:
            self._xml_start_tag('sheetData')
            self._write_rows()
            self._xml_end_tag('sheetData')

    def _write_page_margins(self):
        # Write the <pageMargins> element.
        left = '0.7'
        right = '0.7'
        top = '0.75'
        bottom = '0.75'
        header = '0.3'
        footer = '0.3'

        attributes = [
            ('left', left),
            ('right', right),
            ('top', top),
            ('bottom', bottom),
            ('header', header),
            ('footer', footer),
        ]

        self._xml_empty_tag('pageMargins', attributes)

    def _write_rows(self):
        # Write out the worksheet data as a series of rows and cells.
        self._calculate_spans()

        for row_num in range(self.dim_rowmin, self.dim_rowmax + 1):

            if (row_num in self.set_rows or row_num in self.comments
                    or self.table[row_num]):
                # Only process rows with formatting, cell data and/or comments.

                span_index = int(row_num / 16)

                if span_index in self.row_spans:
                    span = self.row_spans[span_index]
                else:
                    span = None

                if self.table[row_num]:
                    # Write the cells if the row contains data.
                    if row_num not in self.set_rows:
                        self._write_row(row_num, span)
                    else:
                        self._write_row(row_num, span, self.set_rows[row_num])

                    for col_num in range(self.dim_colmin, self.dim_colmax + 1):
                        if col_num in self.table[row_num]:
                            col_ref = self.table[row_num][col_num]
                            self._write_cell(row_num, col_num, col_ref)

                    self._xml_end_tag('row')

                elif row_num in self.comments:
                    # Row with comments in cells.
                    self._write_empty_row(row_num, span,
                                          self.set_rows[row_num])
                else:
                    # Blank row with attributes only.
                    self._write_empty_row(row_num, span,
                                          self.set_rows[row_num])

    def _write_single_row(self, current_row_num):
        # Write out the worksheet data as a single row with cells.
        # This method is used when memory optimisation is on. A single
        # row is written and the data table is reset. That way only
        # one row of data is kept in memory at any one time. We don't
        # write span data in the optimised case since it is optional.

        # Set the new previous row as the current row.
        row_num = self.previous_row
        self.previous_row = current_row_num

        if (row_num in self.set_rows or row_num in self.comments
                or self.table[row_num]):
            # Only process rows with formatting, cell data and/or comments.

            # No span data in optimised mode.
            span = None

            if self.table[row_num]:
                # Write the cells if the row contains data.
                if row_num not in self.set_rows:
                    self._write_row(row_num, span)
                else:
                    self._write_row(row_num, span, self.set_rows[row_num])

                for col_num in range(self.dim_colmin, self.dim_colmax + 1):
                    if col_num in self.table[row_num]:
                        col_ref = self.table[row_num][col_num]
                        self._write_cell(row_num, col_num, col_ref)

                self._xml_end_tag('row')
            else:
                # Row attributes or comments only.
                self._write_empty_row(row_num, span, self.set_rows[row_num])

        # Reset table.
        self.table.clear()

    def _calculate_spans(self):
        # Calculate the "spans" attribute of the <row> tag. This is an
        # XLSX optimisation and isn't strictly required. However, it
        # makes comparing files easier. The span is the same for each
        # block of 16 rows.
        spans = {}
        span_min = None
        span_max = None

        for row_num in range(self.dim_rowmin, self.dim_rowmax + 1):

            if row_num in self.table:
                # Calculate spans for cell data.
                for col_num in range(self.dim_colmin, self.dim_colmax + 1):
                    if col_num in self.table[row_num]:
                        if span_min is None:
                            span_min = col_num
                            span_max = col_num
                        else:
                            if col_num < span_min:
                                span_min = col_num
                            if col_num > span_max:
                                span_max = col_num

            if row_num in self.comments:
                # Calculate spans for comments.
                for col_num in range(self.dim_colmin, self.dim_colmax + 1):
                    if self.comments[row_num][col_num] is not None:

                        if span_min is None:
                            span_min = col_num
                            span_max = col_num
                        else:
                            if col_num < span_min:
                                span_min = col_num
                            if col_num > span_max:
                                span_max = col_num

            if ((row_num + 1) % 16 == 0) or row_num == self.dim_rowmax:
                span_index = int(row_num / 16)

                if span_min is not None:
                    span_min += 1
                    span_max += 1
                    spans[span_index] = "%s:%s" % (span_min, span_max)
                    span_min = None

        self.row_spans = spans

    def _write_row(self, row, spans, properties=None, empty_row=False):
        # Write the <row> element.
        xf_index = 0

        if properties:
            height, cell_format, hidden, level, collapsed = properties
        else:
            height, cell_format, hidden, level, collapsed = 15, None, 0, 0, 0

        if height is None:
            height = self.default_row_height

        attributes = [('r', row + 1)]

        # Get the cell_format index.
        if cell_format:
            xf_index = cell_format._get_xf_index()

        # Add row attributes where applicable.
        if spans:
            attributes.append(('spans', spans))
        if xf_index:
            attributes.append(('s', xf_index))
        if cell_format:
            attributes.append(('customFormat', 1))
        if height != 15:
            attributes.append(('ht', height))
        if hidden:
            attributes.append(('hidden', 1))
        if height != 15:
            attributes.append(('customHeight', 1))
        if level:
            attributes.append(('outlineLevel', level))
        if collapsed:
            attributes.append(('collapsed', 1))
        if self.excel_version == 2010:
            attributes.append(('x14ac:dyDescent', '0.25'))

        if empty_row:
            self._xml_empty_tag_unencoded('row', attributes)
        else:
            self._xml_start_tag_unencoded('row', attributes)

    def _write_empty_row(self, *args):
        # Write and empty <row> element.
        self._write_row(*args, empty_row=True)

    def _write_cell(self, row, col, cell):
        # Write the <cell> element.
        #
        # Note. This is the innermost loop so efficiency is important.
        cell_range = xl_rowcol_to_cell(row, col)

        attributes = [('r', cell_range)]

        if cell.format:
            # Add the cell format index.
            xf_index = cell.format._get_xf_index()
            attributes.append(('s', xf_index))
        elif row in self.set_rows and self.set_rows[row][1]:
            # Add the row format.
            row_xf = self.set_rows[row][1]
            attributes.append(('s', row_xf._get_xf_index()))
        elif col in self.col_formats:
            # Add the column format.
            col_xf = self.col_formats[col]
            attributes.append(('s', col_xf._get_xf_index()))

        # Write the various cell types.
        if type(cell).__name__ == 'Number':
            # Write a number.
            self._xml_number_element(cell.number, attributes)

        elif type(cell).__name__ == 'String':
            # Write a string.
            string = cell.string

            if not self.optimization:
                # Write a shared string.
                self._xml_string_element(string, attributes)
            else:
                # Write an optimised in-line string.

                # TODO: Fix control char encoding when unit test is ported.
                # Escape control characters. See SharedString.pm for details.
                # string =~ s/(_x[0-9a-fA-F]{4}_)/_x005F1/g
                # string =~s/([\x00-\x08\x0B-\x1F])/sprintf "_x04X_", ord(1)/eg

                # Write any rich strings without further tags.
                if re.search('^<r>', string) and re.search('</r>$', string):
                    self._xml_rich_inline_string(string, attributes)
                else:
                    # Add attribute to preserve leading or trailing whitespace.
                    preserve = 0
                    if re.search('^\s', string) or re.search('\s$', string):
                        preserve = 1

                    self._xml_inline_string(string, preserve, attributes)

        elif type(cell).__name__ == 'Formula':
            # Write a formula. First check if the formula value is a string.
            try:
                float(cell.value)
            except ValueError:
                attributes.append(('t', 'str'))

            self._xml_formula_element(cell.formula, cell.value, attributes)

        elif type(cell).__name__ == 'ArrayFormula':
            # Write a array formula.

            # First check if the formula value is a string.
            try:
                float(cell.value)
            except ValueError:
                attributes.append(('t', 'str'))

            # Write an array formula.
            self._xml_start_tag('c', attributes)
            self._write_cell_array_formula(cell.formula, cell.range)
            self._write_cell_value(cell.value)
            self._xml_end_tag('c')

        elif type(cell).__name__ == 'Blank':
            # Write a empty cell.
            self._xml_empty_tag('c', attributes)

    def _write_cell_value(self, value):
        # Write the cell value <v> element.
        if value is None:
            value = ''

        self._xml_data_element('v', value)

    def _write_cell_array_formula(self, formula, cell_range):
        # Write the cell array formula <f> element.
        attributes = [
            ('t', 'array'),
            ('ref', cell_range)
        ]

        self._xml_data_element('f', formula, attributes)
