###############################################################################
#
# Worksheet - A class for writing the Excel XLSX Worksheet file.
#
# Copyright 2013, John McNamara, jmcnamara@cpan.org
#

# Standard packages.
import re
import datetime
from warnings import warn
from collections import defaultdict
from collections import namedtuple

# Package imports.
from . import xmlwriter
from .utility import xl_rowcol_to_cell
from .utility import xl_cell_to_rowcol
from .utility import xl_col_to_name
from .utility import xl_range
from .utility import xl_color


###############################################################################
#
# Decorator functions.
#
###############################################################################
def convert_cell_args(method):
    """
    Decorator function to convert A1 notation in cell method calls
    to the default row/col notation.

    """
    def cell_wrapper(self, *args):

        try:
            # First arg is an int, default to row/col notation.
            int(args[0])
            return method(self, *args)
        except ValueError:
            # First arg isn't an int, convert to A1 notation.
            new_args = list(xl_cell_to_rowcol(args[0]))
            new_args.extend(args[1:])
            return method(self, *new_args)

    return cell_wrapper


def convert_range_args(method):
    """
    Decorator function to convert A1 notation in range method calls
    to the default row/col notation.

    """
    def cell_wrapper(self, *args):

        try:
            # First arg is an int, default to row/col notation.
            int(args[0])
            return method(self, *args)
        except ValueError:
            # First arg isn't an int, convert to A1 notation.
            cell_1, cell_2 = args[0].split(':')
            row_1, col_1 = xl_cell_to_rowcol(cell_1)
            row_2, col_2 = xl_cell_to_rowcol(cell_2)
            new_args = [row_1, col_1, row_2, col_2]
            new_args.extend(args[1:])
            return method(self, *new_args)

    return cell_wrapper


def convert_column_args(method):
    """
    Decorator function to convert A1 notation in columns method calls
    to the default row/col notation.

    """
    def column_wrapper(self, *args):

        try:
            # First arg is an int, default to row/col notation.
            int(args[0])
            return method(self, *args)
        except ValueError:
            # First arg isn't an int, convert to A1 notation.
            cell_1, cell_2 = [col + '1' for col in args[0].split(':')]
            _, col_1 = xl_cell_to_rowcol(cell_1)
            _, col_2 = xl_cell_to_rowcol(cell_2)
            new_args = [col_1, col_2]
            new_args.extend(args[1:])
            return method(self, *new_args)

    return column_wrapper


###############################################################################
#
# Worksheet Class definition.
#
###############################################################################
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
        self.str_table = None
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

        self.repeat_row_range = ''
        self.repeat_col_range = ''
        self.print_area_range = ''

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

        self.protect_options = {}
        self.set_cols = {}
        self.set_rows = {}

        self.zoom = 100
        self.zoom_scale_normal = 1
        self.print_scale = 100
        self.is_right_to_left = 0
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

        self.autofilter_area = ''
        self.autofilter_ref = None
        self.filter_range = []
        self.filter_on = 0
        self.filter_range = []
        self.filter_cols = {}
        self.filter_type = {}

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

        self.vba_codename = None

        self.date_1904 = False
        self.epoch = datetime.datetime(1899, 12, 31)
        self.hyperlinks = defaultdict(dict)

    @convert_cell_args
    def write(self, row, col, *args):
        """
        Write data to a worksheet cell by calling the appropriate write_*()
        method based on the type of data being passed.

        Args:
            row:     The cell row (zero indexed).
            col:     The cell column (zero indexed).
            token:   Cell data.
            format:  An optional cell Format object.
            options: Any options to pass to sub function.

        Returns:
             0:    Success.
            -1:    Row or column is out of worksheet bounds.
            other: Return value of called method.

        """
        # Check the number of args passed.
        if not len(args):
            raise TypeError("write() takes at least 4 arguments (3 given)")

        # The first arg should be the token for all write calls.
        token = args[0]

        # Convert None to an empty string and thus a blank cell.
        if token is None:
            token = ''

        # Check for a datetime object.
        if isinstance(token, datetime.datetime):
            return self.write_datetime(row, col, *args)

        # Then check if the token to write is a number.
        try:
            float(token)
            return self.write_number(row, col, *args)
        except ValueError:
            # Not a number. Continue to the checks below.
            pass

        # Map the data to the appropriate write_*() method.
        if token == '':
            return self.write_blank(row, col, *args)
        elif token.startswith('='):
            return self.write_formula(row, col, *args)
        elif token.startswith('{') and token.endswith('}'):
            return self.write_formula(row, col, *args)
        elif re.match('[fh]tt?ps?://', token):
            return self.write_url(row, col, *args)
        elif re.match('mailto:', token):
            return self.write_url(row, col, *args)
        elif re.match('(in|ex)ternal:', token):
            return self.write_url(row, col, *args)
        else:
            return self.write_string(row, col, *args)

    @convert_cell_args
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
            -1: Row or column is out of worksheet bounds.
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

    @convert_cell_args
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
            -1: Row or column is out of worksheet bounds.

        """
        # TODO catch and re-raise exception if token isn't a number.
        number = float(number)

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

    @convert_cell_args
    def write_blank(self, row, col, blank, cell_format=None):
        """
        Write a blank cell with formatting to a worksheet cell. The blank
        token is ignored and the format only is written to the cell.

        Args:
            row:         The cell row (zero indexed).
            col:         The cell column (zero indexed).
            blank:       Any value. It is ignored.
            cell_format: An optional cell Format object.

        Returns:
            0:  Success.
            -1: Row or column is out of worksheet bounds.

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

    @convert_cell_args
    def write_formula(self, row, col, formula, cell_format=None, value=0):
        """
        Write a formula to a worksheet cell.

        Args:
            row:         The cell row (zero indexed).
            col:         The cell column (zero indexed).
            formula:     Cell formula.
            cell_format: An optional cell Format object.
            value:       An optional value for the formula. Default is 0.

        Returns:
            0:  Success.
            -1: Row or column is out of worksheet bounds.

        """
        # Check that row and col are valid and store max and min values.
        if self._check_dimensions(row, col):
            return -1

        # Hand off array formulas.
        if formula.startswith('{') and formula.endswith('}'):
            return self.write_array_formula(row, col, row, col, formula,
                                            cell_format, value)

        # Remove the formula '=' sign if it exists.
        if formula.startswith('='):
            formula = formula.lstrip('=')

        # Write previous row if in in-line string optimization mode.
        if self.optimization and row > self.previous_row:
            self._write_single_row(row)

        # Store the cell data in the worksheet data table.
        cell_tuple = namedtuple('Formula', 'formula, format, value')
        self.table[row][col] = cell_tuple(formula, cell_format, value)

        return 0

    @convert_range_args
    def write_array_formula(self, first_row, first_col, last_row, last_col,
                            formula, cell_format=None, value=0):
        """
        Write a formula to a worksheet cell.

        Args:
            first_row:    The first row of the cell range. (zero indexed).
            first_col:    The first column of the cell range.
            last_row:     The last row of the cell range. (zero indexed).
            last_col:     The last column of the cell range.
            formula:      Cell formula.
            cell_format:  An optional cell Format object.
            value:        An optional value for the formula. Default is 0.

        Returns:
            0:  Success.
            -1: Row or column is out of worksheet bounds.

        """

        # Swap last row/col with first row/col as necessary.
        if first_row > last_row:
            first_row, last_row = last_row, first_row
        if first_col > last_col:
            first_col, last_col = last_col, first_col

        # Check that row and col are valid and store max and min values
        if self._check_dimensions(last_row, last_col):
            return -1

        # Define array range
        if first_row == last_row and first_col == last_col:
            cell_range = xl_rowcol_to_cell(first_row, first_col)
        else:
            cell_range = (xl_rowcol_to_cell(first_row, first_col) + ':'
                          + xl_rowcol_to_cell(last_row, last_col))

        # Remove array formula braces and the leading =.
        if formula[0] == '{':
            formula = formula[1:]
        if formula[0] == '=':
            formula = formula[1:]
        if formula[-1] == '}':
            formula = formula[:-1]

        # Write previous row if in in-line string optimization mode.
        if self.optimization and first_row > self.previous_row:
            self._write_single_row(first_row)

        # Store the cell data in the worksheet data table.
        cell_tuple = namedtuple('ArrayFormula',
                                'formula, format, value, range')
        self.table[first_row][first_col] = cell_tuple(formula, cell_format,
                                                      value, cell_range)

        # Pad out the rest of the area with formatted zeroes.
        if not self.optimization:
            for row in range(first_row, last_row + 1):
                for col in range(first_col, last_col + 1):
                    if row != first_row or col != first_col:
                        self.write_number(row, col, 0, cell_format)

        return 0

    @convert_cell_args
    def write_datetime(self, row, col, date, cell_format):
        """
        Write a date to a worksheet cell.

        Args:
            row:         The cell row (zero indexed).
            col:         The cell column (zero indexed).
            date:        Date and/or time as a datetime object.
            cell_format: A cell Format object.

        Returns:
            0:  Success.
            -1: Row or column is out of worksheet bounds.

        """
        # Check that row and col are valid and store max and min values.
        if self._check_dimensions(row, col):
            return -1

        # Write previous row if in in-line string optimization mode.
        if self.optimization and row > self.previous_row:
            self._write_single_row(row)

        # Convert datetime to an Excel date.
        number = self._convert_date_time(date)

        # Store the cell data in the worksheet data table.
        cell_tuple = namedtuple('Number', 'number, format')
        self.table[row][col] = cell_tuple(number, cell_format)

        return 0

    # Write a hyperlink. This is comprised of two elements: the displayed
    # string and the non-displayed link. The displayed string is the same as
    # the link unless an alternative string is specified. The display string
    # is written using the write_string() method. Therefore the max characters
    # string limit applies.
    #
    # The hyperlink can be to a http, ftp, mail, internal sheet, or external
    # directory urls.
    #
    # Returns  0 : normal termination
    #         -1 : insufficient number of arguments
    #         -2 : row or column out of range
    #         -3 : long string truncated to 32767 chars
    #         -4 : URL longer than 255 characters
    #         -5 : Exceeds limit of 65_530 urls per worksheet
    #
    @convert_cell_args
    def write_url(self, row, col, url, cell_format=None,
                  string=None, tip=None):
        """
        Write a hyperlink to a worksheet cell.

        Args:
            row:    The cell row (zero indexed).
            col:    The cell column (zero indexed).
            url:    Hyperlink url.
            format: An optional cell Format object.
            string: An optional display string for the hyperlink.
            tip:    An optional tooltip.
        Returns:
            0:  Success.
            -1: Row or column is out of worksheet bounds.
            -2: String longer than 32767 characters.
            -3: URL longer than Excel limit of 255 characters
            -4: Exceeds Excel limit of 65,530 urls per worksheet
        """
        # Default link type such as http://.
        link_type = 1

        # Remove the URI scheme from internal links.
        if re.match("internal:", url):
            url = url.replace('internal:', '')
            link_type = 2

        # Remove the URI scheme from external links.
        if re.match("external:", url):
            url = url.replace('external:', '')
            link_type = 3

        # Set the displayed string to the URL unless defined by the user.
        if string is None:
            string = url

        # For external links change the directory separator from Unix to Dos.
        if link_type == 3:
            url = url.replace('/', '\\')
            string = string.replace('/', '\\')

        # Strip the mailto header.
        string = string.replace('mailto:', '')

        # Check that row and col are valid and store max and min values
        if self._check_dimensions(row, col):
            return -1

        # Check that the string is < 32767 chars
        str_error = 0
        if len(string) > self.xls_strmax:
            warn("Ignoring URL since it exceeds Excel's string limit of "
                 "32767 characters")
            return -2

        # Copy string for use in hyperlink elements.
        url_str = string

        # External links to URLs and to other Excel workbooks have slightly
        # different characteristics that we have to account for.
        if link_type == 1:
            # Escape URL unless it looks already escaped.
            if not re.search('%[0-9a-fA-F]{2}', url):
                # Can't use url.quote() here because it doesn't match Excel.
                url = url.replace('%', '%25')
                url = url.replace('"', '%22')
                url = url.replace(' ', '%20')
                url = url.replace('<', '%3c')
                url = url.replace('>', '%3e')
                url = url.replace('[', '%5b')
                url = url.replace(']', '%5d')
                url = url.replace('^', '%5e')
                url = url.replace('`', '%60')
                url = url.replace('{', '%7b')
                url = url.replace('}', '%7d')

            # Ordinary URL style external links don't have a "location" string.
            url_str = None

        elif link_type == 3:

            # External Workbook links need to be modified into correct format.
            # The URL will look something like 'c:\temp\file.xlsx#Sheet!A1'.
            # We need the part to the left of the # as the URL and the part to
            # the right as the "location" string (if it exists).
            if re.search('#', url):
                url, url_str = url.split('#')
            else:
                url_str = None

            # Add the file:/// URI to the url if non-local.
            # Windows style "C:/" link. # Network share.
            if (re.match('\w:', url) or re.match(r'\\', url)):
                url = 'file:///' + url

            # Convert a .\dir\file.xlsx link to dir\file.xlsx.
            url = re.sub(r'^\.\\', '', url)

            # Treat as a default external link now the data has been modified.
            link_type = 1

        # Excel limits escaped URL to 255 characters.
        if len(url) > 255:
            warn("Ignoring URL '%s' > 255 characters since it exceeds "
                 "Excel's limit for URLS" % url)
            return -3

        # Check the limit of URLS per worksheet.
        self.hlink_count += 1

        if self.hlink_count > 65530:
            warn("Ignoring URL '%s' since it exceeds Excel's limit of "
                 "65,530 URLS per worksheet." % url)
            return -5

        # Write previous row if in in-line string optimization mode.
        if self.optimization == 1 and row > self.previous_row:
            self._write_single_row(row)

        # Write the hyperlink string.
        self.write_string(row, col, string, cell_format)

        # Store the hyperlink data in a separate structure.
        self.hyperlinks[row][col] = {
            'link_type': link_type,
            'url': url,
            'str': url_str,
            'tip': tip}

        return str_error

    @convert_cell_args
    def write_row(self, row, col, data, cell_format=None):
        """
        Write a row of data starting from (row, col).

        Args:
            row:    The cell row (zero indexed).
            col:    The cell column (zero indexed).
            data:   A list of tokens to be written with write().
            format: An optional cell Format object.
        Returns:
            0:  Success.
            other: Return value of write() method.

        """
        for token in (data):
            error = self.write(row, col, token, cell_format)
            if error:
                return error
            col += 1

        return 0

    @convert_cell_args
    def write_column(self, row, col, data, cell_format=None):
        """
        Write a column of data starting from (row, col).

        Args:
            row:    The cell row (zero indexed).
            col:    The cell column (zero indexed).
            data:   A list of tokens to be written with write().
            format: An optional cell Format object.
        Returns:
            0:  Success.
            other: Return value of write() method.

        """
        for token in (data):
            error = self.write(row, col, token, cell_format)
            if error:
                return error
            row += 1

        return 0

    def get_name(self):
        """
        Retrieve the worksheet name.

        Args:
            None.

        Returns:
            Nothing.

        """
        # There is no set_name() method. Name must be set in add_worksheet().
        return self.name

    def activate(self):
        """
        Set this worksheet as the active worksheet, i.e. the worksheet that is
        displayed when the workbook is opened. Also set it as selected.

        Note: An active worksheet cannot be hidden.

        Args:
            None.

        Returns:
            Nothing.

        """
        self.hidden = 0
        self.selected = 1
        self.worksheet_meta.activesheet = self.index

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

    def set_first_sheet(self):
        """
        Set current worksheet as the first visible sheet. This is necessary
        when there are a large number of worksheets and the activated
        worksheet is not visible on the screen.

        Note: A selected worksheet cannot be hidden.

        Args:
            None.

        Returns:
            Nothing.

        """
        self.hidden = 0  # Active worksheet can't be hidden.
        self.firstsheet = self.index

    @convert_column_args
    def set_column(self, firstcol, lastcol, width=None, cell_format=None,
                   options={}):
        """
        Set the width, and other properties of a single column or a
        range of columns.

        Args:
            firstcol:    First column (zero-indexed).
            lastcol:     Last column (zero-indexed). Can be same as firstcol.
            width:       Column width. (optional).
            cell_format: Column cell_format. (optional).
            options:     Dict of options such as hidden and level.

        Returns:
            0:  Success.
            -1: Column number is out of worksheet bounds.

        """
        # Ensure 2nd col is larger than first.
        if firstcol > lastcol:
            (firstcol, lastcol) = (lastcol, firstcol)

        # Don't modify the row dimensions when checking the columns.
        ignore_row = 1

        # Set optional column values.
        hidden = options.get('hidden', False)
        level = options.get('level', 0)

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

    def set_row(self, row, height=None, cell_format=None, options={}):
        """
        Set the width, and other properties of a row.
        range of columns.

        Args:
            row:         Row number (zero-indexed).
            height:      Row width. (optional).
            cell_format: Row cell_format. (optional).
            options:     Dict of options such as hidden, level and collapsed.

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

        # Set optional row values.
        hidden = options.get('hidden', False)
        collapsed = options.get('collapsed', False)
        level = options.get('level', 0)

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

    @convert_range_args
    def merge_range(self, first_row, first_col, last_row, last_col,
                    data, cell_format=None):
        """
        Merge a range of cells.

        Args:
            first_row:    The first row of the cell range. (zero indexed).
            first_col:    The first column of the cell range.
            last_row:     The last row of the cell range. (zero indexed).
            last_col:     The last column of the cell range.
            data:         Cell data.
            cell_format:  Cell Format object.

        Returns:
             0:    Success.
            -1:    Row or column is out of worksheet bounds.
            other: Return value of write().

        """
        # Merge a range of cells. The first cell should contain the data and
        # the others should be blank. All cells should have the same format.

        # Excel doesn't allow a single cell to be merged
        if first_row == last_row and first_col == last_col:
            warn("Can't merge single cell")
            return

        # Swap last row/col with first row/col as necessary
        if first_row > last_row:
            (first_row, last_row) = (last_row, first_row)
        if first_col > last_col:
            (first_col, last_col) = (last_col, first_col)

        # Check that column number is valid and store the max value
        if self._check_dimensions(last_row, first_col):
            return

        # Store the merge range.
        self.merge.append([first_row, first_col, last_row, last_col])

        # Write the first cell
        self.write(first_row, first_col, data, cell_format)

        # Pad out the rest of the area with formatted blank cells.
        for row in range(first_row, last_row + 1):
            for col in range(first_col, last_col + 1):
                if row == first_row and col == first_col:
                    continue
                self.write_blank(row, col, '', cell_format)

    @convert_range_args
    def autofilter(self, first_row, first_col, last_row, last_col):
        """
        Set the autofilter area in the worksheet.

        Args:
            first_row:    The first row of the cell range. (zero indexed).
            first_col:    The first column of the cell range.
            last_row:     The last row of the cell range. (zero indexed).
            last_col:     The last column of the cell range.

        Returns:
             Nothing.

        """
        # Reverse max and min values if necessary.
        if last_row < first_row:
            (first_row, last_row) = (last_row, first_row)
        if last_col < first_col:
            (first_col, last_col) = (last_col, first_col)

        # Build up the print area range "Sheet1!$A$1:$C$13".
        area = self._convert_name_area(first_row, first_col,
                                       last_row, last_col)
        ref = xl_range(first_row, first_col, last_row, last_col)

        self.autofilter_area = area
        self.autofilter_ref = ref
        self.filter_range = [first_col, last_col]

    def filter_column(self, col, criteria):
        """
        Set the column filter criteria.

        Args:
            col:       Filter column (zero-indexed).
            criteria:  Filter criteria.

        Returns:
             Nothing.

        """
        if not self.autofilter_area:
            warn("Must call autofilter() before filter_column()")
            return

        # Check for a column reference in A1 notation and substitute.
        try:
            int(col)
        except ValueError:
            # Convert col ref to a cell ref and then to a col number.
            col_letter = col
            (_, col) = xl_cell_to_rowcol(col + '1')

            if col >= self.xls_colmax:
                warn("Invalid column '%d'" % col_letter)
                return

        (col_first, col_last) = self.filter_range

        # Reject column if it is outside filter range.
        if col < col_first or col > col_last:
            warn("Column '%d' outside autofilter() column range (%d, %d)"
                 % (col, col_first, col_last))
            return

        tokens = self._extract_filter_tokens(criteria)

        if not (len(tokens) == 3 or len(tokens) == 7):
            warn("Incorrect number of tokens in criteria '%s'" % criteria)

        tokens = self._parse_filter_expression(criteria, tokens)

        # Excel handles single or double custom filters as default filters.
        #  We need to check for them and handle them accordingly.
        if len(tokens) == 2 and tokens[0] == 2:
            # Single equality.
            self.filter_column_list(col, [tokens[1]])
        elif (len(tokens) == 5 and tokens[0] == 2 and tokens[2] == 1
              and tokens[3] == 2):
            # Double equality with "or" operator.
            self.filter_column_list(col, [tokens[1], tokens[4]])
        else:
            # Non default custom filter.
            self.filter_cols[col] = tokens
            self.filter_type[col] = 0

        self.filter_on = 1

    def filter_column_list(self, col, filters):
        """
        Set the column filter criteria in Excel 2007 list style.

        Args:
            col:      Filter column (zero-indexed).
            filters:  List of filter criteria to match.

        Returns:
             Nothing.

        """
        if not self.autofilter_area:
            warn("Must call autofilter() before filter_column()")
            return

        # Check for a column reference in A1 notation and substitute.
        try:
            int(col)
        except ValueError:
            # Convert col ref to a cell ref and then to a col number.
            col_letter = col
            (_, col) = xl_cell_to_rowcol(col + '1')

            if col >= self.xls_colmax:
                warn("Invalid column '%d'" % col_letter)
                return

        (col_first, col_last) = self.filter_range

        # Reject column if it is outside filter range.
        if col < col_first or col > col_last:
            warn("Column '%d' outside autofilter() column range "
                 "(%d,%d)" % (col, col_first, col_last))
            return

        self.filter_cols[col] = filters
        self.filter_type[col] = 1
        self.filter_on = 1

    def set_zoom(self, zoom=100):
        """
        Set the worksheet zoom factor.

        Args:
            zoom: Scale factor: 10 <= zoom <= 400.

        Returns:
            Nothing.

        """
        # Ensure the zoom scale is in Excel's range.
        if zoom < 10 or zoom > 400:
            warn("Zoom factor %d outside range: 10 <= zoom <= 400" % zoom)
            zoom = 100

        self.zoom = int(zoom)

    def right_to_left(self):
        """
        Display the worksheet right to left for some versions of Excel.

        Args:
            None.

        Returns:
            Nothing.

        """
        self.is_right_to_left = 1

    def hide_zero(self):
        """
        Hide zero values in worksheet cells.

        Args:
            None.

        Returns:
            Nothing.

        """
        self.show_zeros = 0

    def set_tab_color(self, color):
        """
        Set the colour of the worksheet tab.

        Args:
            color: A #RGB color index.

        Returns:
            Nothing.

        """
        self.tab_color = xl_color(color)

    def protect(self, password='', options=None):
        """
        Set the colour of the worksheet tab.

        Args:
            password: An optional password string.
            options:  A dictionary of worksheet objects to protect.

        Returns:
            Nothing.

        """
        if password != '':
            password = self._encode_password(password)

        if not options:
            options = {}

        # Default values for objects that can be protected.
        defaults = {
            'sheet': 1,
            'content': 0,
            'objects': 0,
            'scenarios': 0,
            'format_cells': 0,
            'format_columns': 0,
            'format_rows': 0,
            'insert_columns': 0,
            'insert_rows': 0,
            'insert_hyperlinks': 0,
            'delete_columns': 0,
            'delete_rows': 0,
            'select_locked_cells': 1,
            'sort': 0,
            'autofilter': 0,
            'pivot_tables': 0,
            'select_unlocked_cells': 1}

        # Overwrite the defaults with user specified values.
        for key in (options.keys()):

            if key in defaults:
                defaults[key] = options[key]
            else:
                warn("Unknown protection object: '%s'\n" % key)

        # Set the password after the user defined values.
        defaults['password'] = password

        self.protect_options = defaults

    ###########################################################################
    #
    # Public API. Page Setup methods.
    #
    ###########################################################################
    def set_landscape(self):
        """
        Set the page orientation as landscape.

        Args:
            None.

        Returns:
            Nothing.

        """
        self.orientation = 0
        self.page_setup_changed = 1

    def set_portrait(self):
        """
        Set the page orientation as portrait.

        Args:
            None.

        Returns:
            Nothing.

        """
        self.orientation = 1
        self.page_setup_changed = 1

    def set_page_view(self):
        """
        Set the page view mode.

        Args:
            None.

        Returns:
            Nothing.

        """
        self.page_view = 1

    def set_paper(self, paper_size):
        """
        Set the paper type. US Letter = 1, A4 = 9.

        Args:
            paper_size: Paper index.

        Returns:
            Nothing.

        """
        if paper_size:
            self.paper_size = paper_size
            self.page_setup_changed = 1

    def center_horizontally(self):
        """
        Center the page horizontally.

        Args:
            None.

        Returns:
            Nothing.

        """
        self.print_options_changed = 1
        self.hcenter = 1

    def center_vertically(self):
        """
        Center the page vertically.

        Args:
            None.

        Returns:
            Nothing.

        """
        self.print_options_changed = 1
        self.vcenter = 1

    def set_margins(self, left=0.7, right=0.7, top=0.75, bottom=0.75):
        """
        Set all the page margins in inches.

        Args:
            left:   Left margin.
            right:  Right margin.
            top:    Top margin.
            bottom: Bottom margin.

        Returns:
            Nothing.

        """
        self.margin_left = left
        self.margin_right = right
        self.margin_top = top
        self.margin_bottom = bottom

    def set_header(self, header='', margin=0.3):
        """
        Set the page header caption and optional margin.

        Args:
            header: Header string.
            margin: Header margin.

        Returns:
            Nothing.

        """
        if len(header) >= 255:
            warn('Header string must be less than 255 characters')
            return

        self.header = header
        self.margin_header = margin
        self.header_footer_changed = 1

    def set_footer(self, footer='', margin=0.3):
        """
        Set the page footer caption and optional margin.

        Args:
            footer: Footer string.
            margin: Footer margin.

        Returns:
            Nothing.

        """
        if len(footer) >= 255:
            warn('Footer string must be less than 255 characters')
            return

        self.footer = footer
        self.margin_footer = margin
        self.header_footer_changed = 1

    def repeat_rows(self, first_row, last_row=None):
        """
        Set the rows to repeat at the top of each printed page.

        Args:
            first_row: Start row for range.
            last_row: End row for range.

        Returns:
            Nothing.

        """
        if last_row is None:
            last_row = first_row

        # Convert rows to 1 based.
        first_row += 1
        last_row += 1

        # Create the row range area like: $1:$2.
        area = '${}:${}'.format(first_row, last_row)

        # Build up the print titles area "Sheet1!$1:$2"
        sheetname = self._quote_sheetname(self.name)
        self.repeat_row_range = sheetname + '!' + area

    @convert_column_args
    def repeat_columns(self, first_col, last_col=None):
        """
        Set the columns to repeat at the left hand side of each printed page.

        Args:
            first_col: Start column for range.
            last_col: End column for range.

        Returns:
            Nothing.

        """
        if last_col is None:
            last_col = first_col

        # Convert to A notation.
        first_col = xl_col_to_name(first_col, 1)
        last_col = xl_col_to_name(last_col, 1)

        # Create a column range like $C:$D.
        area = first_col + ':' + last_col

        # Build up the print area range "=Sheet2!$C:$D"
        sheetname = self._quote_sheetname(self.name)
        self.repeat_col_range = sheetname + "!" + area

    def hide_gridlines(self, option=1):
        """
        Set the option to hide gridlines on the screen and the printed page.

        Args:
            option:    0 : Don't hide gridlines
                       1 : Hide printed gridlines only
                       2 : Hide screen and printed gridlines

        Returns:
            Nothing.

        """
        if option == 0:
            self.print_gridlines = 1
            self.screen_gridlines = 1
            self.print_options_changed = 1
        elif option == 1:
            self.print_gridlines = 0
            self.screen_gridlines = 1
        else:
            self.print_gridlines = 0
            self.screen_gridlines = 0

    def print_row_col_headers(self):
        """
        Set the option to print the row and column headers on the printed page.

        Args:
            None.

        Returns:
            Nothing.

        """
        self.print_headers = 1
        self.print_options_changed = 1

    @convert_range_args
    def print_area(self, first_row, first_col, last_row, last_col):
        """
        Set the print area in the current worksheet.

        Args:
            first_row:    The first row of the cell range. (zero indexed).
            first_col:    The first column of the cell range.
            last_row:     The last row of the cell range. (zero indexed).
            last_col:     The last column of the cell range.

        Returns:
            0:  Success.
            -1: Row or column is out of worksheet bounds.

        """
        # Set the print area in the current worksheet.

        # Ignore max print area since it is the same as no  area for Excel.
        if (first_row == 0 and first_col == 0
                and last_row == self.xls_rowmax - 1
                and last_col == self.xls_colmax - 1):
            return

        # Build up the print area range "Sheet1!$A$1:$C$13".
        area = self._convert_name_area(first_row, first_col,
                                       last_row, last_col)
        self.print_area_range = area

    def print_across(self):
        """
        Set the order in which pages are printed.

        Args:
            None.

        Returns:
            Nothing.

        """
        self.page_order = 1
        self.page_setup_changed = 1

    def fit_to_pages(self, width, height):
        """
        Fit the printed area to a specific number of pages both vertically and
        horizontally.

        Args:
            width:  Number of pages horizontally.
            height: Number of pages vertically.

        Returns:
            Nothing.

        """
        self.fit_page = 1
        self.fit_width = width
        self.fit_height = height
        self.page_setup_changed = 1

    def set_start_page(self, start_page):
        """
        Set the start page number when printing.

        Args:
            start_page: Start page number.

        Returns:
            Nothing.

        """
        self.page_start = start_page
        self.custom_start = 1

    def set_print_scale(self, scale):
        """
        Set the scale factor for the printed page.

        Args:
            scale: Print scale. 10 <= scale <= 400.

        Returns:
            Nothing.

        """
        # Confine the scale to Excel's range.
        if scale < 10 or scale > 400:
            warn("Print scale '%d' outside range: 10 <= scale <= 400" % scale)
            return

        # Turn off "fit to page" option when print scale is on.
        self.fit_page = 0

        self.print_scale = int(scale)
        self.page_setup_changed = 1

    def set_h_pagebreaks(self, breaks):
        """
        Set the horizontal page breaks on a worksheet.

        Args:
            breaks: List of rows where the page breaks should be added.

        Returns:
            Nothing.

        """
        self.hbreaks = breaks

    #
    # set_v_pagebreaks(@breaks)
    #
    # Store the vertical page breaks on a worksheet.
    #
    def set_v_pagebreaks(self, breaks):
        """
        Set the horizontal page breaks on a worksheet.

        Args:
            breaks: List of columns where the page breaks should be added.

        Returns:
            Nothing.

        """
        self.vbreaks = breaks

    ###########################################################################
    #
    # Private API.
    #
    ###########################################################################
    def _initialize(self, init_data):
        self.name = init_data['name']
        self.index = init_data['index']
        self.str_table = init_data['str_table']
        self.worksheet_meta = init_data['worksheet_meta']

    def _assemble_xml_file(self):
        # Assemble and write the XML file.

        # Write the XML declaration.
        self._xml_declaration()

        # Write the root worksheet element.
        self._write_worksheet()

        # Write the worksheet properties.
        self._write_sheet_pr()

        # Write the worksheet dimensions.
        self._write_dimension()

        # Write the sheet view properties.
        self._write_sheet_views()

        # Write the sheet format properties.
        self._write_sheet_format_pr()

        # Write the sheet column info.
        self._write_cols()

        # Write the worksheet data such as rows columns and cells.
        # if self.optimization == 0:
        #    self._write_sheet_data()
        # else:
        #    self._write_optimized_sheet_data()
        self._write_sheet_data()

        # Write the sheetProtection element.
        self._write_sheet_protection()

        # Write the worksheet calculation properties.
        # self._write_sheet_calc_pr()

        # Write the worksheet phonetic properties.
        # self._write_phonetic_pr()

        # Write the autoFilter element.
        self._write_auto_filter()

        # Write the mergeCells element.
        self._write_merge_cells()

        # Write the conditional formats.
        # self._write_conditional_formats()

        # Write the dataValidations element.
        # self._write_data_validations()

        # Write the hyperlink element.
        self._write_hyperlinks()

        # Write the printOptions element.
        self._write_print_options()

        # Write the worksheet page_margins.
        self._write_page_margins()

        # Write the worksheet page setup.
        self._write_page_setup()

        # Write the headerFooter element.
        self._write_header_footer()

        # Write the rowBreaks element.
        self._write_row_breaks()

        # Write the colBreaks element.
        self._write_col_breaks()

        # Write the drawing element.
        # self._write_drawings()

        # Write the legacyDrawing element.
        # self._write_legacy_drawing()

        # Write the tableParts element.
        # self._write_table_parts()

        # Write the extLst and sparklines.
        # self._write_ext_sparklines()

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

    def _convert_date_time(self, date):
        # Convert a Python datetime.datetime value to an Excel date number.
        delta = date - self.epoch
        excel_time = (delta.days
                      + (float(delta.seconds)
                      + float(delta.microseconds) / 1E6)
                      / (60 * 60 * 24))

        # Special case for datetime where time only has been specified and
        # the default date of 1900-01-01 is used.
        if date.isocalendar() == (1900, 1, 1):
            excel_time -= 1

        # Account for Excel erroneously treating 1900 as a leap year.
        if not self.date_1904 and delta.days > 59:
            excel_time += 1

        return excel_time

    def _options_changed(self):
        # Check to see if any of the worksheet options have changed.
        options_changed = 0
        print_changed = 0
        setup_changed = 0

        if (self.orientation == 0
                or self.hcenter == 1
                or self.vcenter == 1
                or self.header != ''
                or self.footer != ''
                or self.margin_header != 0.50
                or self.margin_footer != 0.50
                or self.margin_left != 0.75
                or self.margin_right != 0.75
                or self.margin_top != 1.00
                or self.margin_bottom != 1.00):
            setup_changed = 1

        # Special case for 1x1 page fit.
        if self.fit_width == 1 and self.fit_height == 1:
            options_changed = 1
            self.fit_width = 0
            self.fit_height = 0

        if (self.fit_width > 1
                or self.fit_height > 1
                or self.page_order == 1
                or self.black_white == 1
                or self.draft_quality == 1
                or self.print_comments == 1
                or self.paper_size != 0
                or self.print_scale != 100
                or self.print_gridlines == 1
                or self.print_headers == 1
                or self.hbreaks > 0
                or self.vbreaks > 0):
            print_changed = 1

        if (print_changed or setup_changed):
            options_changed = 1

        if self.screen_gridlines == 0:
            options_changed = 1
        if self.filter_on:
            options_changed = 1

        return (options_changed, print_changed, setup_changed)

    def _quote_sheetname(self, sheetname):
        # Sheetnames used in references should be quoted if they
        # contain any spaces, special characters or if the look like
        # something that isn't a sheet name.  TODO. We need to handle
        # more special cases.
        if re.match(r'Sheet\d+', sheetname):
            return sheetname
        else:
            return "'{}'".format(sheetname)

    def _convert_name_area(self, row_num_1, col_num_1, row_num_2, col_num_2):
        # Convert zero indexed rows and columns to the format required by
        # worksheet named ranges, eg, "Sheet1!$A$1:$C$13".

        range1 = ''
        range2 = ''
        area = ''
        row_col_only = 0

        # Convert to A1 notation.
        col_char_1 = xl_col_to_name(col_num_1, 1)
        col_char_2 = xl_col_to_name(col_num_2, 1)
        row_char_1 = '$' + str(row_num_1 + 1)
        row_char_2 = '$' + str(row_num_2 + 1)

        # We need to handle special cases that refer to rows or columns only.
        if row_num_1 == 0 and row_num_2 == self.xls_rowmax - 1:
            range1 = col_char_1
            range2 = col_char_2
            row_col_only = 1
        elif col_num_1 == 0 and col_num_2 == self.xls_colmax - 1:
            range1 = row_char_1
            range2 = row_char_2
            row_col_only = 1
        else:
            range1 = col_char_1 + row_char_1
            range2 = col_char_2 + row_char_2

        # A repeated range is only written once (if it isn't a special case).
        if range1 == range2 and not row_col_only:
            area = range1
        else:
            area = range1 + ':' + range2

        # Build up the print area range "Sheet1!$A$1:$C$13".
        sheetname = self._quote_sheetname(self.name)
        area = sheetname + "!" + area

        return area

    def _sort_pagebreaks(self, breaks):
        # This is an internal method used to filter elements of a list of
        # pagebreaks used in the _store_hbreak() and _store_vbreak() methods.
        # It:
        #   1. Removes duplicate entries from the list.
        #   2. Sorts the list.
        #   3. Removes 0 from the list if present.
        if not breaks:
            return

        breaks_set = set(breaks)

        if 0 in breaks_set:
            breaks_set.remove(0)

        breaks_list = list(breaks_set)
        breaks_list.sort()

        # The Excel 2007 specification says that the maximum number of page
        # breaks is 1026. However, in practice it is actually 1023.
        max_num_breaks = 1023
        if len(breaks_list) > max_num_breaks:
            breaks_list = breaks_list[:max_num_breaks]

        return breaks_list

    def _extract_filter_tokens(self, expression):
        # Extract the tokens from the filter expression. The tokens are mainly
        # non-whitespace groups. The only tricky part is to extract string
        # tokens that contain whitespace and/or quoted double quotes (Excel's
        # escaped quotes).
        #
        # Examples: 'x <  2000'
        #           'x >  2000 and x <  5000'
        #           'x = "foo"'
        #           'x = "foo bar"'
        #           'x = "foo "" bar"'
        #
        if not expression:
            return []

        token_re = re.compile(r'"(?:[^"]|"")*"|\S+')
        tokens = token_re.findall(expression)

        new_tokens = []
        # Remove single leading and trailing quotes and un-escape other quotes.
        for token in tokens:
            if token.startswith('"'):
                token = token[1:]

            if token.endswith('"'):
                token = token[:-1]

            token = token.replace('""', '"')

            new_tokens.append(token)

        return new_tokens

    def _parse_filter_expression(self, expression, tokens):
        # Converts the tokens of a possibly conditional expression into 1 or 2
        # sub expressions for further parsing.
        #
        # Examples:
        #          ('x', '==', 2000) -> exp1
        #          ('x', '>',  2000, 'and', 'x', '<', 5000) -> exp1 and exp2

        if len(tokens) == 7:
            # The number of tokens will be either 3 (for 1 expression)
            # or 7 (for 2  expressions).
            conditional = tokens[3]

            if re.match('(and|&&)', conditional):
                conditional = 0
            elif re.match('(or|\|\|)', conditional):
                conditional = 1
            else:
                warn("Token '%s' is not a valid conditional "
                     "in filter expression '%s'" % (conditional, expression))

            expression_1 = self._parse_filter_tokens(expression, tokens[0:3])
            expression_2 = self._parse_filter_tokens(expression, tokens[4:7])

            return expression_1 + [conditional] + expression_2
        else:
            return self._parse_filter_tokens(expression, tokens)

    def _parse_filter_tokens(self, expression, tokens):
        # Parse the 3 tokens of a filter expression and return the operator
        # and token. The use of numbers instead of operators is a legacy of
        # Spreadsheet::WriteExcel.
        operators = {
            '==': 2,
            '=': 2,
            '=~': 2,
            'eq': 2,

            '!=': 5,
            '!~': 5,
            'ne': 5,
            '<>': 5,

            '<': 1,
            '<=': 3,
            '>': 4,
            '>=': 6,
        }

        operator = operators.get(tokens[1], None)
        token = tokens[2]

        # Special handling of "Top" filter expressions.
        if re.match('top|bottom', tokens[0].lower()):
            value = int(tokens[1])

            if (value < 1 or value > 500):
                warn("The value '%d' in expression '%s' "
                     "must be in the range 1 to 500" % (value, expression))

            token = token.lower()

            if token != 'items' and token != '%':
                warn("The type '%s' in expression '%s' "
                     "must be either 'items' or '%'" % (token, expression))

            if tokens[0].lower() == 'top':
                operator = 30
            else:
                operator = 32

            if tokens[2] == '%':
                operator += 1

            token = str(value)

        if not operator and tokens[0]:
            warn("Token '%s' is not a valid operator "
                 "in filter expression '%s'" % (token[0], expression))

        # Special handling for Blanks/NonBlanks.
        if re.match('blanks|nonblanks', token.lower()):
            # Only allow Equals or NotEqual in this context.
            if operator != 2 and operator != 5:
                warn("The operator '%s' in expression '%s' "
                     "is not valid in relation to Blanks/NonBlanks'"
                     % (tokens[1], expression))

            token = token.lower()

            # The operator should always be 2 (=) to flag a "simple" equality
            # in the binary record. Therefore we convert <> to =.
            if token == 'blanks':
                if operator == 5:
                    token = ' '
            else:
                if operator == 5:
                    operator = 2
                    token = 'blanks'
                else:
                    operator = 5
                    token = ' '

        # if the string token contains an Excel match character then change the
        # operator type to indicate a non "simple" equality.
        if operator == 2 and re.search('[*?]', token):
            operator = 22

        return [operator, token]

    def _encode_password(self, plaintext):
        # Encode the worksheet protection "password" as a simple hash.
        # Based on the algorithm by Daniel Rentz of OpenOffice.
        i = 0
        count = len(plaintext)
        digits = []

        for char in (plaintext):
            i += 1
            char = ord(char) << i
            low_15 = char & 0x7fff
            high_15 = char & 0x7fff << 15
            high_15 = high_15 >> 15
            char = low_15 | high_15
            digits.append(char)

        password_hash = 0x0000

        for digit in digits:
            password_hash ^= digit

        password_hash ^= count
        password_hash ^= 0xCE4B

        return "%X" % password_hash

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
            ('xmlns:r', xmlns_r)]

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
        if self.is_right_to_left:
            attributes.append(('rightToLeft', 1))

        # Show that the sheet tab is selected.
        if self.selected:
            attributes.append(('tabSelected', 1))

        # Turn outlines off. Also required in the outlinePr element.
        if not self.outline_on:
            attributes.append(("showOutlineSymbols", 0))

        # Set the page view/layout mode if required.
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
            ('width', width)]

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
        attributes = [
            ('left', self.margin_left),
            ('right', self.margin_right),
            ('top', self.margin_top),
            ('bottom', self.margin_bottom),
            ('header', self.margin_header),
            ('footer', self.margin_footer)]

        self._xml_empty_tag('pageMargins', attributes)

    def _write_page_setup(self):
        # Write the <pageSetup> element.
        #
        # The following is an example taken from Excel.
        #
        # <pageSetup
        #     paperSize="9"
        #     scale="110"
        #     fitToWidth="2"
        #     fitToHeight="2"
        #     pageOrder="overThenDown"
        #     orientation="portrait"
        #     blackAndWhite="1"
        #     draft="1"
        #     horizontalDpi="200"
        #     verticalDpi="200"
        #     r:id="rId1"
        # />
        #
        attributes = []

        # Skip this element if no page setup has changed.
        if not self.page_setup_changed:
            return

        # Set paper size.
        if self.paper_size:
            attributes.append(('paperSize', self.paper_size))

        # Set the print_scale.
        if self.print_scale != 100:
            attributes.append(('scale', self.print_scale))

        # Set the "Fit to page" properties.
        if self.fit_page and self.fit_width != 1:
            attributes.append(('fitToWidth', self.fit_width))

        if self.fit_page and self.fit_height != 1:
            attributes.append(('fitToHeight', self.fit_height))

        # Set the page print direction.
        if self.page_order:
            attributes.append(('pageOrder', "overThenDown"))

        # Set page orientation.
        if self.orientation:
            attributes.append(('orientation', 'portrait'))
        else:
            attributes.append(('orientation', 'landscape'))

        # Set start page for printing.
        if self.page_start != 0:
            attributes.append(('useFirstPageNumber', self.page_start))

        self._xml_empty_tag('pageSetup', attributes)

    def _write_print_options(self):
        # Write the <printOptions> element.
        attributes = []

        if not self.print_options_changed:
            return

        # Set horizontal centering.
        if self.hcenter:
            attributes.append(('horizontalCentered', 1))

        # Set vertical centering.
        if self.vcenter:
            attributes.append(('verticalCentered', 1))

        # Enable row and column headers.
        if self.print_headers:
            attributes.append(('headings', 1))

        # Set printed gridlines.
        if self.print_gridlines:
            attributes.append(('gridLines', 1))

        self._xml_empty_tag('printOptions', attributes)

    def _write_header_footer(self):
        # Write the <headerFooter> element.

        if not self.header_footer_changed:
            return

        self._xml_start_tag('headerFooter')

        if self.header:
            self._write_odd_header()
        if self.footer:
            self._write_odd_footer()

        self._xml_end_tag('headerFooter')

    def _write_odd_header(self):
        # Write the <headerFooter> element.
        self._xml_data_element('oddHeader', self.header)

    def _write_odd_footer(self):
        # Write the <headerFooter> element.
        self._xml_data_element('oddFooter', self.footer)

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

    def _write_sheet_pr(self):
        # Write the <sheetPr> element for Sheet level properties.
        attributes = []

        if (not self.fit_page
                and not self.filter_on
                and not self.tab_color
                and not self.outline_changed
                and not self.vba_codename):
            return

        if self.vba_codename:
            attributes.append(('codeName', self.vba_codename))

        if self.filter_on:
            attributes.append(('filterMode', 1))

        if (self.fit_page
                or self.tab_color
                or self.outline_changed):
            self._xml_start_tag('sheetPr', attributes)
            self._write_tab_color()
            self._write_outline_pr()
            self._write_page_set_up_pr()
            self._xml_end_tag('sheetPr')
        else:
            self._xml_empty_tag('sheetPr', attributes)

    def _write_page_set_up_pr(self):
        # Write the <pageSetUpPr> element.
        if not self.fit_page:
            return

        attributes = [('fitToPage', 1)]
        self._xml_empty_tag('pageSetUpPr', attributes)

    def _write_tab_color(self):
        # Write the <tabColor> element.
        color = self.tab_color

        if not color:
            return

        attributes = [('rgb', color)]

        self._xml_empty_tag('tabColor', attributes)

    def _write_outline_pr(self):
        # Write the <outlinePr> element.
        attributes = []

        if not self.outline_changed:
            return

        if self.outline_style:
            attributes.append(("applyStyles", 1))
        if not self.outline_below:
            attributes.append(("summaryBelow", 0))
        if not self.outline_right:
            attributes.append(("summaryRight", 0))
        if not self.outline_on:
            attributes.append(("showOutlineSymbols", 0))

        self._xml_empty_tag('outlinePr', attributes)

    def _write_row_breaks(self):
        # Write the <rowBreaks> element.
        page_breaks = self._sort_pagebreaks(self.hbreaks)

        if not page_breaks:
            return

        count = len(page_breaks)

        attributes = [
            ('count', count),
            ('manualBreakCount', count),
        ]

        self._xml_start_tag('rowBreaks', attributes)

        for row_num in (page_breaks):
            self._write_brk(row_num, 16383)

        self._xml_end_tag('rowBreaks')

    def _write_col_breaks(self):
        # Write the <colBreaks> element.
        page_breaks = self._sort_pagebreaks(self.vbreaks)

        if not page_breaks:
            return

        count = len(page_breaks)

        attributes = [
            ('count', count),
            ('manualBreakCount', count),
        ]

        self._xml_start_tag('colBreaks', attributes)

        for col_num in (page_breaks):
            self._write_brk(col_num, 1048575)

        self._xml_end_tag('colBreaks')

    def _write_brk(self, brk_id, brk_max):
        # Write the <brk> element.
        attributes = [
            ('id', brk_id),
            ('max', brk_max),
            ('man', 1)]

        self._xml_empty_tag('brk', attributes)

    def _write_merge_cells(self):
        # Write the <mergeCells> element.
        merged_cells = self.merge
        count = len(merged_cells)

        if not count:
            return

        attributes = [('count', count)]

        self._xml_start_tag('mergeCells', attributes)

        for merged_range in (merged_cells):

            # Write the mergeCell element.
            self._write_merge_cell(merged_range)

        self._xml_end_tag('mergeCells')

    def _write_merge_cell(self, merged_range):
        # Write the <mergeCell> element.
        (row_min, col_min, row_max, col_max) = merged_range

        # Convert the merge dimensions to a cell range.
        cell_1 = xl_rowcol_to_cell(row_min, col_min)
        cell_2 = xl_rowcol_to_cell(row_max, col_max)
        ref = cell_1 + ':' + cell_2

        attributes = [('ref', ref)]

        self._xml_empty_tag('mergeCell', attributes)

    def _write_hyperlinks(self):
        # Process any stored hyperlinks in row/col order and write the
        # <hyperlinks> element. The attributes are different for internal
        # and external links.
        hlink_refs = []
        display = None

        # Sort the hyperlinks into row order.
        row_nums = sorted(self.hyperlinks.keys())

        # Exit if there are no hyperlinks to process.
        if not row_nums:
            return

        # Iterate over the rows.
        for row_num in (row_nums):
            # Sort the hyperlinks into column order.
            col_nums = sorted(self.hyperlinks[row_num].keys())

            # Iterate over the columns.
            for col_num in (col_nums):
                # Get the link data for this cell.
                link = self.hyperlinks[row_num][col_num]
                link_type = link["link_type"]

                # If the cell isn't a string then we have to add the url as
                # the string to display.
                if (self.table
                        and self.table[row_num]
                        and self.table[row_num][col_num]):
                    cell = self.table[row_num][col_num]
                    if type(cell).__name__ != 'String':
                        display = link["url"]

                if link_type == 1:
                    # External link with rel file relationship.
                    self.rel_count += 1

                    hlink_refs.append([link_type,
                                       row_num,
                                       col_num,
                                       self.rel_count,
                                       link["str"],
                                       display,
                                       link["tip"]])

                    # Links for use by the packager.
                    self.external_hyper_links.append(['/hyperlink',
                                                      link["url"], 'External'])
                else:
                    # Internal link with rel file relationship.
                    hlink_refs.append([link_type,
                                       row_num,
                                       col_num,
                                       link["url"],
                                       link["str"],
                                       link["tip"]])

        # Write the hyperlink elements.
        self._xml_start_tag('hyperlinks')

        for args in (hlink_refs):
            link_type = args.pop(0)

            if link_type == 1:
                self._write_hyperlink_external(*args)
            elif link_type == 2:
                self._write_hyperlink_internal(*args)

        self._xml_end_tag('hyperlinks')

    def _write_hyperlink_external(self, row, col, id_num, location=None,
                                  display=None, tooltip=None):
        # Write the <hyperlink> element for external links.
        ref = xl_rowcol_to_cell(row, col)
        r_id = 'rId' + str(id_num)

        attributes = [
            ('ref', ref),
            ('r:id', r_id)]

        if location is not None:
            attributes.append(('location', location))
        if display is not None:
            attributes.append(('display', display))
        if tooltip is not None:
            attributes.append(('tooltip', tooltip))

        self._xml_empty_tag('hyperlink', attributes)

    def _write_hyperlink_internal(self, row, col, location=None, display=None,
                                  tooltip=None):
        # Write the <hyperlink> element for internal links.
        ref = xl_rowcol_to_cell(row, col)

        attributes = [
            ('ref', ref),
            ('location', location)]

        if tooltip is not None:
            attributes.append(('tooltip', tooltip))
        attributes.append(('display', display))

        self._xml_empty_tag('hyperlink', attributes)

    def _write_auto_filter(self):
        # Write the <autoFilter> element.
        if not self.autofilter_ref:
            return

        attributes = [('ref', self.autofilter_ref)]

        if self.filter_on:
            # Autofilter defined active filters.
            self._xml_start_tag('autoFilter', attributes)
            self._write_autofilters()
            self._xml_end_tag('autoFilter')

        else:
            # Autofilter defined without active filters.
            self._xml_empty_tag('autoFilter', attributes)

    def _write_autofilters(self):
        # Function to iterate through the columns that form part of an
        # autofilter range and write the appropriate filters.
        (col1, col2) = self.filter_range

        for col in range(col1, col2 + 1):
            # Skip if column doesn't have an active filter.
            if not col in self.filter_cols:
                continue

            # Retrieve the filter tokens and write the autofilter records.
            tokens = self.filter_cols[col]
            filter_type = self.filter_type[col]

            # Filters are relative to first column in the autofilter.
            self._write_filter_column(col - col1, filter_type, tokens)

    def _write_filter_column(self, col_id, filter_type, filters):
        # Write the <filterColumn> element.
        attributes = [('colId', col_id)]

        self._xml_start_tag('filterColumn', attributes)

        if filter_type == 1:
            # Type == 1 is the new XLSX style filter.
            self._write_filters(filters)
        else:
            # Type == 0 is the classic "custom" filter.
            self._write_custom_filters(filters)

        self._xml_end_tag('filterColumn')

    def _write_filters(self, filters):
        # Write the <filters> element.

        if len(filters) == 1 and filters[0] == 'blanks':
            # Special case for blank cells only.
            self._xml_empty_tag('filters', [('blank', 1)])
        else:
            # General case.
            self._xml_start_tag('filters')

            for autofilter in (filters):
                self._write_filter(autofilter)

            self._xml_end_tag('filters')

    def _write_filter(self, val):
        # Write the <filter> element.
        attributes = [('val', val)]

        self._xml_empty_tag('filter', attributes)

    def _write_custom_filters(self, tokens):
        # Write the <customFilters> element.
        if len(tokens) == 2:
            # One filter expression only.
            self._xml_start_tag('customFilters')
            self._write_custom_filter(*tokens)
            self._xml_end_tag('customFilters')
        else:
            # Two filter expressions.
            attributes = []

            # Check if the "join" operand is "and" or "or".
            if tokens[2] == 0:
                attributes = [('and', 1)]
            else:
                attributes = [('and', 0)]

            # Write the two custom filters.
            self._xml_start_tag('customFilters', attributes)
            self._write_custom_filter(tokens[0], tokens[1])
            self._write_custom_filter(tokens[3], tokens[4])
            self._xml_end_tag('customFilters')

    def _write_custom_filter(self, operator, val):
        # Write the <customFilter> element.
        attributes = []

        operators = {
            1: 'lessThan',
            2: 'equal',
            3: 'lessThanOrEqual',
            4: 'greaterThan',
            5: 'notEqual',
            6: 'greaterThanOrEqual',
            22: 'equal',
        }

        # Convert the operator from a number to a descriptive string.
        if operators[operator] is not None:
            operator = operators[operator]
        else:
            warn("Unknown operator = %s" % operator)

        # The 'equal' operator is the default attribute and isn't stored.
        if not operator == 'equal':
            attributes.append(('operator', operator))
        attributes.append(('val', val))

        self._xml_empty_tag('customFilter', attributes)

    def _write_sheet_protection(self):
        # Write the <sheetProtection> element.
        attributes = []

        if not self.protect_options:
            return

        options = self.protect_options

        if options['password']:
            attributes.append(('password', options['password']))
        if options['sheet']:
            attributes.append(('sheet', 1))
        if options['content']:
            attributes.append(('content', 1))
        if not options['objects']:
            attributes.append(('objects', 1))
        if not options['scenarios']:
            attributes.append(('scenarios', 1))
        if options['format_cells']:
            attributes.append(('formatCells', 0))
        if options['format_columns']:
            attributes.append(('formatColumns', 0))
        if options['format_rows']:
            attributes.append(('formatRows', 0))
        if options['insert_columns']:
            attributes.append(('insertColumns', 0))
        if options['insert_rows']:
            attributes.append(('insertRows', 0))
        if options['insert_hyperlinks']:
            attributes.append(('insertHyperlinks', 0))
        if options['delete_columns']:
            attributes.append(('deleteColumns', 0))
        if options['delete_rows']:
            attributes.append(('deleteRows', 0))
        if not options['select_locked_cells']:
            attributes.append(('selectLockedCells', 1))
        if options['sort']:
            attributes.append(('sort', 0))
        if options['autofilter']:
            attributes.append(('autoFilter', 0))
        if options['pivot_tables']:
            attributes.append(('pivotTables', 0))
        if not options['select_unlocked_cells']:
            attributes.append(('selectUnlockedCells', 1))

        self._xml_empty_tag('sheetProtection', attributes)
