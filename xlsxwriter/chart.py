###############################################################################
#
# Chart - A class for writing the Excel XLSX Worksheet file.
#
# Copyright 2013, John McNamara, jmcnamara@cpan.org
#
import re
from warnings import warn

from . import xmlwriter
from .utility import xl_color
from .utility import xl_rowcol_to_cell


class Chart(xmlwriter.XMLwriter):
    """
    A class for writing the Excel XLSX Chart file.


    """

    ###########################################################################
    #
    # Public API.
    #
    ###########################################################################

    def __init__(self, options=None):
        """
        Constructor.

        """

        super(Chart, self).__init__()

        self.subtype = None
        self.sheet_type = 0x0200
        self.orientation = 0x0
        self.series = []
        self.embedded = 0
        self.id = ''
        self.series_index = 0
        self.style_id = 2
        self.axis_ids = []
        self.axis2_ids = []
        self.cat_has_num_fmt = 0
        self.requires_category = 0
        self.legend_position = 'right'
        self.cat_axis_position = 'b'
        self.val_axis_position = 'l'
        self.formula_ids = {}
        self.formula_data = []
        self.horiz_cat_axis = 0
        self.horiz_val_axis = 1
        self.protection = 0
        self.chartarea = {}
        self.plotarea = {}
        self.x_axis = {}
        self.y_axis = {}
        self.y2_axis = {}
        self.x2_axis = {}
        self.chart_name = ''
        self.show_blanks = 'gap'
        self.show_hidden_data = 0
        self.show_crosses = 1
        self.width = 480
        self.height = 288
        self.x_scale = 1
        self.y_scale = 1
        self.x_offset = 0
        self.y_offset = 0
        self.table = None
        self.title_formula = None
        self.title_name = None
        self.cross_between = None
        self.default_marker = None
        self.series_gap = None
        self.series_overlap = None

        self._set_default_properties()

    def add_series(self, options):
        # Add a series and it's properties to a chart.

        # Check that the required input has been specified.
        if not 'values' in options:
            warn("Must specify 'values' in add_series()")
            return

        if self.requires_category and not 'categories' in options:
            warn("Must specify 'categories' in add_series() "
                 "for this chart type")

        # Convert list into a formula string.
        values = self._list_to_formula(options.get('values'))
        categories = self._list_to_formula(options.get('categories'))

        # Switch name and name_formula parameters if required.
        name, name_formula = self._process_names(options.get('name'),
                                                 options.get('name_formula'))

        # Get an id for the data equivalent to the range formula.
        cat_id = self._get_data_id(categories, options.get('categories_data'))
        val_id = self._get_data_id(values, options.get('values_data'))
        name_id = self._get_data_id(name_formula, options.get('name_data'))

        # Set the line properties for the series.
        line = self._get_line_properties(options.get('line'))

        # Allow 'border' as a synonym for 'line' in bar/column style charts.
        if options.get('border'):
            line = self._get_line_properties(options['border'])

        # Set the fill properties for the series.
        fill = self._get_fill_properties(options.get('fill'))

        # Set the marker properties for the series.
        marker = self._get_marker_properties(options.get('marker'))

        # Set the trendline properties for the series.
        trendline = self._get_trendline_properties(options.get('trendline'))

        # Set the error bars properties for the series.
        y_error_bars = self._get_error_bars_props(options.get('y_error_bars'))
        x_error_bars = self._get_error_bars_props(options.get('x_error_bars'))

        error_bars = {'x_error_bars': x_error_bars,
                      'y_error_bars': y_error_bars}

        # Set the point properties for the series.
        points = self._get_points_properties(options.get('points'))

        # Set the labels properties for the series.
        labels = self._get_labels_properties(options.get('data_labels'))

        # Set the "invert if negative" fill property.
        invert_if_neg = options.get('invert_if_negative', False)

        # Set the gap for Bar/Column charts.
        if options.get('gap'):
            self.series_gap = options['gap']

        # Set the overlap for Bar/Column charts.
        if options.get('overlap'):
            self.series_overlap = options['overlap']

        # Set the secondary axis properties.
        x2_axis = options.get('x2_axis')
        y2_axis = options.get('y2_axis')

        # Add the user supplied data to the internal structures.
        series = {
            'values': values,
            'categories': categories,
            'name': name,
            'name_formula': name_formula,
            'name_id': name_id,
            'val_data_id': val_id,
            'cat_data_id': cat_id,
            'line': line,
            'fill': fill,
            'marker': marker,
            'trendline': trendline,
            'labels': labels,
            'invert_if_neg': invert_if_neg,
            'x2_axis': x2_axis,
            'y2_axis': y2_axis,
            'points': points,
            'error_bars': error_bars,
        }

        self.series.append(series)

    def set_x_axis(self, options):
        # Set the properties of the X-axis.
        axis = self._convert_axis_args(self.x_axis, options)

        self.x_axis = axis

    def set_y_axis(self, options):
        # Set the properties of the Y-axis.
        axis = self._convert_axis_args(self.y_axis, options)

        self.y_axis = axis

    def set_x2_axis(self, options):
        # Set the properties of the secondary X-axis.
        axis = self._convert_axis_args(self.x2_axis, options)

        self.x2_axis = axis

    def set_y2_axis(self, options):
        # Set the properties of the secondary Y-axis.
        axis = self._convert_axis_args(self.y2_axis, options)

        self.y2_axis = axis

    def set_title(self, options):
        # Set the properties of the chart title.

        name, name_formula = self._process_names(options.get('name'),
                                                 options.get('name_formula'))

        data_id = self._get_data_id(name_formula, options.get('data'))

        self.title_name = name
        self.title_formula = name_formula
        self.title_data_id = data_id

        # Set the font properties if present.
        self.title_font = self._convert_font_args(options.get('name_font'))

    def set_legend(self, options):
        # Set the properties of the chart legend.

        self.legend_position = options.get('position', 'right')
        self.legend_delete_series = options.get('delete_series')

    def set_plotarea(self, options):
        # Set the properties of the chart plotarea.
        # Convert the user defined properties to internal properties.
        self.plotarea = self._get_area_properties(options)

    def set_chartarea(self, options):
        # Set the properties of the chart chartarea.
        # Convert the user defined properties to internal properties.
        self.chartarea = self._get_area_properties(options)

    def set_style(self, style_id):
        # Set one of the 42 built-in Excel chart styles. The default is 2.
        if style_id is None:
            style_id = 2

        if style_id < 0 or style_id > 42:
            style_id = 2

        self.style_id = style_id

    def show_blanks_as(self, option):
        # Set the option for displaying blank data in a chart.
        if not option:
            return

        valid_options = {
            'gap': 1,
            'zero': 1,
            'span': 1,
        }

        if not 'option' in valid_options:
            warn("Unknown show_blanks_as() option '%s'" % option)
            return

        self.show_blanks = option

    def show_hidden_data(self):
        # Display data in hidden rows or columns.
        self.show_hidden_data = 1

    def set_size(self, options):
        # Set dimensions or scale for the chart.
        self.width = options.get('width')
        self.height = options.get('height')
        self.x_scale = options.get('x_scale')
        self.y_scale = options.get('y_scale')
        self.x_offset = options.get('x_offset')
        self.x_offset = options.get('y_offset')

    def set_table(self, args):
        # Set properties for an axis data table.
        table = {
            'horizontal': 1,
            'vertical': 1,
            'outline': 1,
            'show_keys': 0,
        }

        if 'horizontal' in args:
            table['horizontal'] = args.get('horizontal')

        if 'vertical' in args:
            table['vertical'] = args.get('vertical')

        if 'outline' in args:
            table['outline'] = args.get('outline')

        if 'show_keys' in args:
            table['show_keys'] = args.get('show_keys')

        self.table = table

    def set_up_down_bars(self, options):
        # Set properties for the chart up-down bars.
        if options is None:
            return

        # Defaults.
        up_line = None
        up_fill = None
        down_line = None
        down_fill = None

        # Set properties for 'up' bar.
        if options.get('up'):
            # Map border to line.
            if 'border' in options['up']:
                options['up']['line'] = options['up']['border']

            if 'line' in options['up']:
                up_line = self._get_line_properties(options['up']['line'])

            if 'fill' in options['up']:
                up_line = self._get_line_properties(options['up']['fill'])

        # Set properties for 'down' bar.
        if options.get('down'):
            # Map border to line.
            if 'border' in options['down']:
                options['down']['line'] = options['down']['border']

            if 'line' in options['down']:
                down_line = self._get_line_properties(options['down']['line'])

            if 'fill' in options['down']:
                down_line = self._get_line_properties(options['down']['fill'])

        self.up_down_bars = {'up': {'line': up_line,
                                    'fill': up_fill,
                                    },
                             'down': {'line': down_line,
                                      'fill': down_fill,
                                      },
                             }

    def set_drop_lines(self, options):
        # Set properties for the chart drop lines.
        line = self._get_line_properties(options.get('line'))

        self.drop_lines = {'line': line}

    def set_high_low_lines(self, options):
        # Set properties for the chart high-low lines.
        line = self._get_line_properties(options.get('line'))

        self.hi_low_lines = {'line': line}

    ###########################################################################
    #
    # Private API.
    #
    ###########################################################################

    def _assemble_xml_file(self):
        # Assemble and write the XML file.

        # Write the XML declaration.
        self._xml_declaration()

        # Write the c:chartSpace element.
        self._write_chart_space()

        # Write the c:lang element.
        self._write_lang()

        # Write the c:style element.
        self._write_style()

        # Write the c:protection element.
        self._write_protection()

        # Write the c:chart element.
        self._write_chart()

        # Write the c:spPr element for the chartarea formatting.
        self._write_sp_pr(self.chartarea)

        # Write the c:printSettings element.
        if self.embedded:
            self._write_print_settings()

        # Close the worksheet tag.
        self._xml_end_tag('c:chartSpace')
        # Close the file.
        self._xml_close()

    def _convert_axis_args(self, axis, user_options):
        # Convert user defined axis values into private hash values.
        options = axis['defaults'].copy()
        options.update(user_options)

        name, name_formula = self._process_names(options.get('name'),
                                                 options.get('name_formula'))

        data_id = self._get_data_id(name_formula, options.get('data'))

        axis = {
            'defaults': axis['defaults'],
            'name': name,
            'formula': name_formula,
            'data_id': data_id,
            'reverse': options.get('reverse'),
            'min': options.get('min'),
            'max': options.get('max'),
            'minor_unit': options.get('minor_unit'),
            'major_unit': options.get('major_unit'),
            'minor_unit_type': options.get('minor_unit_type'),
            'major_unit_type': options.get('major_unit_type'),
            'log_base': options.get('log_base'),
            'crossing': options.get('crossing'),
            'position': options.get('position'),
            'label_position': options.get('label_position'),
            'num_format': options.get('num_format'),
            'num_format_linked': options.get('num_format_linked'),
        }

        if 'visible' in options:
            axis['visible'] = options.get('visible')
        else:
            axis['visible'] = 1

        # Map major_gridlines properties.
        if (options.get('major_gridlines')
                and options['major_gridlines']['visible']):
            axis['major_gridlines'] = \
                self._get_gridline_properties(options['major_gridlines'])

        # Map minor_gridlines properties.
        if (options.get('minor_gridlines')
                and options['minor_gridlines']['visible']):
            axis['minor_gridlines'] = \
                self._get_gridline_properties(options['minor_gridlines'])

        # Only use the first letter of bottom, top, left or right.
        if axis.get('position'):
            axis['position'] = axis['position'].lower()[0]

        # Set the font properties if present.
        axis['num_font'] = self._convert_font_args(options.get('num_font'))
        axis['name_font'] = self._convert_font_args(options.get('name_font'))

        return axis

    def _convert_font_args(self, options):
        # Convert user defined font values into private dict values.
        if not options:
            return

        font = {
            'name': options.get('name'),
            'color': options.get('color'),
            'size': options.get('size'),
            'bold': options.get('bold'),
            'italic': options.get('italic'),
            'underline': options.get('underline'),
            'pitch_family': options.get('pitch_family'),
            'charset': options.get('charset'),
            'baseline': options.get('baseline', 0),
        }

        # Convert font size units.
        if font['size']:
            font['size'] *= 100

        return font

    def _list_to_formula(self, data):
        # Convert and list of row col values to a range formula.

        # If it isn't an array ref it is probably a formula already.
        if type(data) is not list:
            return data

        sheet = data[0]
        range1 = xl_rowcol_to_cell(data[1], data[2], True, True)
        range2 = xl_rowcol_to_cell(data[3], data[4], True, True)

        return sheet + '!' + range1 + ':' + range2

    def _process_names(self, name, name_formula):
        # Switch name and name_formula parameters if required.

        # Name looks like a formula, use it to set name_formula.
        if name is not None and re.match(r'^=[^!]+!\$', name):
            name_formula = name
            name = ''

        return name, name_formula

    def _get_data_type(self, data):
        # Find the overall type of the data associated with a series.

        # Check for no data in the series.
        if data is None or len(data) == 0:
            return 'none'

        # Determine if data is numeric or strings.
        for token in (data):
            if token is None:
                continue

            try:
                float(token)
            except ValueError:
                # Not a number. Assume entire data series is string data.
                return 'str'

        # The series data was all numeric.
        return 'num'

    def _get_data_id(self, formula, data):
        # Assign an id to a each unique series formula or title/axis formula.
        # Repeated formulas such as for categories get the same id. If the
        # series or title has user specified data associated with it then
        # that is also stored. This data is used to populate cached Excel
        # data when creating a chart. If there is no user defined data then
        # it will be populated by the parent Workbook._add_chart_data().

        # Ignore series without a range formula.
        if not formula:
            return

        # Strip the leading '=' from the formula.
        if formula.startswith('='):
            formula = formula.lstrip('=')

        # Store the data id in a hash keyed by the formula and store the data
        # in a separate array with the same id.
        if not formula in self.formula_ids:
            # Haven't seen this formula before.
            formula_id = len(self.formula_data)

            self.formula_data.append(data)
            self.formula_ids[formula] = formula_id
        else:
            # Formula already seen. Return existing id.
            formula_id = self.formula_ids[formula]

            # Store user defined data if it isn't already there.
            if self.formula_data[formula_id] is None:
                self.formula_data[formula_id] = data

        return formula_id

    def _get_color(self, color):
        # Convert the user specified colour index or string to a rgb colour.
        return xl_color(color)

    def _get_line_properties(self, line):
        # Convert user line properties to the structure required internally.

        if not line:
            return {'defined': False}

        dash_types = {
            'solid': 'solid',
            'round_dot': 'sysDot',
            'square_dot': 'sysDash',
            'dash': 'dash',
            'dash_dot': 'dashDot',
            'long_dash': 'lgDash',
            'long_dash_dot': 'lgDashDot',
            'long_dash_dot_dot': 'lgDashDotDot',
            'dot': 'dot',
            'system_dash_dot': 'sysDashDot',
            'system_dash_dot_dot': 'sysDashDotDot',
        }

        # Check the dash type.
        dash_type = line.get('dash_type')

        if dash_type is not None:
            if dash_type in dash_types:
                line['dash_type'] = dash_types[dash_type]
            else:
                warn("Unknown dash type '%'" % dash_type)
                return

        line['defined'] = True

        return line

    def _get_fill_properties(self, fill):
        # Convert user fill properties to the structure required internally.

        if not fill:
            return {'defined': False}

        fill['defined'] = True

        return fill

    def _get_marker_properties(self, marker):
        # Convert user marker properties to the structure required internally.

        if not marker:
            return

        types = {
            'automatic': 'automatic',
            'none': 'none',
            'square': 'square',
            'diamond': 'diamond',
            'triangle': 'triangle',
            'x': 'x',
            'star': 'start',
            'dot': 'dot',
            'short_dash': 'dot',
            'dash': 'dash',
            'long_dash': 'dash',
            'circle': 'circle',
            'plus': 'plus',
            'picture': 'picture',
        }

        # Check for valid types.
        marker_type = marker.get('type')

        if marker_type is not None:
            if marker_type == 'automatic':
                marker['automatic'] = 1

            if marker_type in types:
                marker['type'] = types[marker_type]
            else:
                warn("Unknown marker type '%s" % marker_type)
                return

        # Set the line properties for the marker..
        line = self._get_line_properties(marker.get('line'))

        # Allow 'border' as a synonym for 'line'.
        if 'border' in marker:
            line = self._get_line_properties(marker['border'])

        # Set the fill properties for the marker.
        fill = self._get_fill_properties(marker.get('fill'))

        marker['line'] = line
        marker['fill'] = fill

        return marker

    def _get_trendline_properties(self, trendline):
        # Convert user trendline properties to structure required internally.

        if not trendline:
            return

        types = {
            'exponential': 'exp',
            'linear': 'linear',
            'log': 'log',
            'moving_average': 'movingAvg',
            'polynomial': 'poly',
            'power': 'power',
        }

        # Check the trendline type.
        trend_type = trendline.get('type')

        if trend_type in types:
            trendline['type'] = types[trend_type]
        else:
            warn("Unknown trendline type 'trend_type'" % trend_type)
            return

        # Set the line properties for the trendline..
        line = self._get_line_properties(trendline.get('line'))

        # Allow 'border' as a synonym for 'line'.
        if 'border' in trendline:
            line = self._get_line_properties(trendline['border'])

        # Set the fill properties for the trendline.
        fill = self._get_fill_properties(trendline.get('fill'))

        trendline['line'] = line
        trendline['fill'] = fill

        return trendline

    def _get_error_bars_props(self, options):
        # Convert user error bars properties to structure required internally.
        if not options:
            return

        # Default values.
        error_bars = {
            'type': 'fixedVal',
            'value': 1,
            'endcap': 1,
            'direction': 'both'
        }

        types = {
            'fixed': 'fixedVal',
            'percentage': 'percentage',
            'standard_deviation': 'stdDev',
            'standard_error': 'stdErr',
        }

        # Check the error bars type.
        error_type = options[type]

        if error_type in types:
            error_bars['type'] = types[error_type]
        else:
            warn("Unknown error bars type 'error_type" % error_type)
            return

        # Set the value for error types that require it.
        if 'value' in options:
            error_bars['value'] = options['value']

        # Set the end-cap style.
        if 'end_style' in options:
            error_bars['endcap'] = options['end_style']

        # Set the error bar direction.
        if 'direction' in options:
            if options['direction'] == 'minus':
                error_bars['direction'] = 'minus'
            elif options['direction'] == 'plus':
                error_bars['direction'] = 'plus'
            else:
                # Default to 'both'.
                pass

        # Set the line properties for the error bars.
        error_bars['line'] = self._get_line_properties(options.get('line'))

        return error_bars

    def _get_gridline_properties(self, options):
        # Convert user gridline properties to structure required internally.
        gridline = {}

        # Set the visible property for the gridline.
        gridline['visible'] = options.get('visible')

        # Set the line properties for the gridline.
        gridline['line'] = self._get_line_properties(options.get('line'))

        return gridline

    def _get_labels_properties(self, labels):
        # Convert user labels properties to the structure required internally.

        if not labels:
            return None

        # Map user defined label positions to Excel positions.
        position = labels.get('position')

        if position:
            positions = {
                'center': 'ctr',
                'right': 'r',
                'left': 'l',
                'top': 't',
                'above': 't',
                'bottom': 'b',
                'below': 'b',
                'inside_end': 'inEnd',
                'outside_end': 'outEnd',
                'best_fit': 'bestFit',
            }

            if position in positions:
                labels['position'] = positions[position]
            else:
                warn("Unknown label position '%s'" % position)
                return

        return labels

    def _get_area_properties(self, options):
        # Convert user area properties to the structure required internally.
        area = {}

        # Handle Excel::Writer::XLSX style properties.
        # Set the line properties for the chartarea.
        line = self._get_line_properties(options.get('line'))

        # Allow 'border' as a synonym for 'line'.
        if options.get('border'):
            line = self._get_line_properties(options['border'])

        # Set the fill properties for the chartarea.
        fill = self._get_fill_properties(options.get('fill'))

        area['line'] = line
        area['fill'] = fill

        return area

    def _get_points_properties(self, user_points):
        # Convert user points properties to structure required internally.
        points = []

        if not user_points:
            return

        for user_point in (user_points):
            point = {}

            if user_point is not None:

                # Set the line properties for the point.
                line = self._get_line_properties(user_point.get('line'))

                # Allow 'border' as a synonym for 'line'.
                if 'border' in user_point:
                    line = self._get_line_properties(user_point['border'])

                # Set the fill properties for the chartarea.
                fill = self._get_fill_properties(user_point.get('fill'))

                point['line'] = line
                point['fill'] = fill

            points.append(point)

        return points

    def _get_primary_axes_series(self):
        # Returns series which use the primary axes.
        primary_axes_series = []

        for series in (self.series):
            if not series['y2_axis']:
                primary_axes_series.append(series)

        return primary_axes_series

    def _get_secondary_axes_series(self):
        # Returns series which use the secondary axes.
        secondary_axes_series = []

        for series in (self.series):
            if series['y2_axis']:
                secondary_axes_series.append(series)

        return secondary_axes_series

    def _add_axis_ids(self, args):
        # Add unique ids for primary or secondary axes
        chart_id = 1 + int(self.id)
        axis_count = 1 + len(self.axis2_ids) + len(self.axis_ids)

        id1 = '5%03d%04d' % (chart_id, axis_count)
        id2 = '5%03d%04d' % (chart_id, axis_count + 1)

        if 'primary_axes' in args:
            self.axis_ids.append(id1)
            self.axis_ids.append(id2)

        if not 'primary_axes' in args:
            self.axis2_ids.append(id1)
            self.axis2_ids.append(id2)

    def _get_font_style_attributes(self, font):
        # _get_font_style_attributes.
        attributes = []

        if not font:
            return attributes

        if 'size' in font:
            attributes.append(('sz', font['size']))

        if 'bold' in font:
            attributes.append(('b', 1))

        if 'italic' in font:
            attributes.append(('i', 1))

        if 'underline' in font:
            attributes.append(('u', 'sng'))

        attributes.append(('baseline', 1))

        return attributes

    def _get_font_latin_attributes(self, font):
        # _get_font_latin_attributes.
        attributes = []

        if not font:
            return attributes

        if 'name' in font:
            attributes.append(('typeface', font['name']))

        if 'pitch_family' in font:
            attributes.append(('pitchFamily', font['pitch_family']))

        if 'charset' in font:
            attributes.append(('charset', font['charset']))

        return attributes

    def _set_default_properties(self):
        # Setup the default properties for a chart.

        self.x_axis['defaults'] = {
            'num_format': 'General',
            'major_gridlines': {'visible': 0}
        }

        self.y_axis['defaults'] = {
            'num_format': 'General',
            'major_gridlines': {'visible': 1}
        }

        self.x2_axis['defaults'] = {
            'num_format': 'General',
            'label_position': 'none',
            'crossing': 'max',
            'visible': 0
        }

        self.y2_axis['defaults'] = {
            'num_format': 'General',
            'major_gridlines': {'visible': 0},
            'position': 'right',
            'visible': 1
        }

        self.set_x_axis({})
        self.set_y_axis({})

        self.set_x2_axis({})
        self.set_y2_axis({})

    def _set_embedded_config_data(self):
        # Setup the default configuration data for an embedded chart.
        self.embedded = 1

    ###########################################################################
    #
    # XML methods.
    #
    ###########################################################################

    def _write_chart_space(self):
        # Write the <c:chartSpace> element.
        schema = 'http://schemas.openxmlformats.org/'
        xmlns_c = schema + 'drawingml/2006/chart'
        xmlns_a = schema + 'drawingml/2006/main'
        xmlns_r = schema + 'officeDocument/2006/relationships'

        attributes = [
            ('xmlns:c', xmlns_c),
            ('xmlns:a', xmlns_a),
            ('xmlns:r', xmlns_r),
        ]

        self._xml_start_tag('c:chartSpace', attributes)

    def _write_lang(self):
        # Write the <c:lang> element.
        val = 'en-US'

        attributes = [('val', val)]

        self._xml_empty_tag('c:lang', attributes)

    def _write_style(self):
        # Write the <c:style> element.
        style_id = self.style_id

        # Don't write an element for the default style, 2.
        if style_id == 2:
            return

        attributes = [('val', style_id)]

        self._xml_empty_tag('c:style', attributes)

    def _write_chart(self):
        # Write the <c:chart> element.
        self._xml_start_tag('c:chart')

        # Write the chart title elements.

        if self.title_formula:
            self._write_title_formula(self.title_formula, self.title_data_id,
                                      None, self.title_font)
        elif self.title_name:
            self._write_title_rich(self.title_name, None, self.title_font)

        # Write the c:plotArea element.
        self._write_plot_area()

        # Write the c:legend element.
        self._write_legend()

        # Write the c:plotVisOnly element.
        self._write_plot_vis_only()

        # Write the c:dispBlanksAs element.
        self._write_disp_blanks_as()

        self._xml_end_tag('c:chart')

    def _write_disp_blanks_as(self):
        # Write the <c:dispBlanksAs> element.
        val = self.show_blanks

        # Ignore the default value.
        if val == 'gap':
            return

        attributes = [('val', val)]

        self._xml_empty_tag('c:dispBlanksAs', attributes)

    def _write_plot_area(self):
        # Write the <c:plotArea> element.
        self._xml_start_tag('c:plotArea')

        # Write the c:layout element.
        self._write_layout()

        # Write  subclass chart type elements for primary and secondary axes.
        self._write_chart_type({'primary_axes': True})
        self._write_chart_type({'primary_axes': False})

        # Write c:catAx and c:valAx elements for series using primary axes.
        self._write_cat_axis({'x_axis': self.x_axis,
                              'y_axis': self.y_axis,
                              'axis_ids': self.axis_ids
                              })

        self._write_val_axis({'x_axis': self.x_axis,
                              'y_axis': self.y_axis,
                              'axis_ids': self.axis_ids
                              })

        # Write c:valAx and c:catAx elements for series using secondary axes.
        self._write_val_axis({'x_axis': self.x2_axis,
                              'y_axis': self.y2_axis,
                              'axis_ids': self.axis2_ids
                              })

        self._write_cat_axis({'x_axis': self.x2_axis,
                              'y_axis': self.y2_axis,
                              'axis_ids': self.axis2_ids
                              })

        # Write the c:dTable element.
        self._write_d_table()

        # Write the c:spPr element for the plotarea formatting.
        self._write_sp_pr(self.plotarea)

        self._xml_end_tag('c:plotArea')

    def _write_layout(self):
        # Write the <c:layout> element.
        self._xml_empty_tag('c:layout')

    def _write_chart_type(self, options):
        # Write the chart type element. This method should be overridden
        # by the subclasses.
        return

    def _write_grouping(self, val):
        # Write the <c:grouping> element.
        attributes = [('val', val)]

        self._xml_empty_tag('c:grouping', attributes)

    def _write_series(self, series):
        # Write the series elements.
        self._write_ser(series)

    def _write_ser(self, series):
        # Write the <c:ser> element.
        index = self.series_index
        self.series_index += 1

        self._xml_start_tag('c:ser')

        # Write the c:idx element.
        self._write_idx(index)

        # Write the c:order element.
        self._write_order(index)

        # Write the series name.
        self._write_series_name(series)

        # Write the c:spPr element.
        self._write_sp_pr(series)

        # Write the c:marker element.
        self._write_marker(series['marker'])

        # Write the c:invertIfNegative element.
        self._write_c_invert_if_negative(series['invert_if_neg'])

        # Write the c:dPt element.
        self._write_d_pt(series['points'])

        # Write the c:dLbls element.
        self._write_d_lbls(series['labels'])

        # Write the c:trendline element.
        self._write_trendline(series['trendline'])

        # Write the c:errBars element.
        self._write_error_bars(series['error_bars'])

        # Write the c:cat element.
        self._write_cat(series)

        # Write the c:val element.
        self._write_val(series)

        self._xml_end_tag('c:ser')

    def _write_idx(self, val):
        # Write the <c:idx> element.

        attributes = [('val', val)]

        self._xml_empty_tag('c:idx', attributes)

    def _write_order(self, val):
        # Write the <c:order> element.

        attributes = [('val', val)]

        self._xml_empty_tag('c:order', attributes)

    def _write_series_name(self, series):
        # Write the series name.
        if series.get('name_formula'):
            self._write_tx_formula(series['name_formula'], series['name_id'])
        elif series.get('name'):
            self._write_tx_value(series['name'])

    def _write_cat(self, series):
        # Write the <c:cat> element.
        formula = series['categories']
        data_id = series['cat_data_id']
        data = None

        if data_id is not None:
            data = self.formula_data[data_id]

        # Ignore <c:cat> elements for charts without category values.
        if not formula:
            return

        self._xml_start_tag('c:cat')

        # Check the type of cached data.
        cat_type = self._get_data_type(data)

        if cat_type == 'str':
            self.cat_has_num_fmt = 0
            # Write the c:numRef element.
            self._write_str_ref(formula, data, cat_type)
        else:
            self.cat_has_num_fmt = 1
            # Write the c:numRef element.
            self._write_num_ref(formula, data, cat_type)

        self._xml_end_tag('c:cat')

    def _write_val(self, series):
        # Write the <c:val> element.
        formula = series['values']
        data_id = series['val_data_id']
        data = self.formula_data[data_id]

        self._xml_start_tag('c:val')

        # Unlike Cat axes data should only be numeric.
        # Write the c:numRef element.
        self._write_num_ref(formula, data, 'num')

        self._xml_end_tag('c:val')

    def _write_num_ref(self, formula, data, ref_type):
        # Write the <c:numRef> element.
        self._xml_start_tag('c:numRef')

        # Write the c:f element.
        self._write_series_formula(formula)

        if ref_type == 'num':
            # Write the c:numCache element.
            self._write_num_cache(data)
        elif ref_type == 'str':
            # Write the c:strCache element.
            self._write_str_cache(data)

        self._xml_end_tag('c:numRef')

    def _write_str_ref(self, formula, data, ref_type):
        # Write the <c:strRef> element.

        self._xml_start_tag('c:strRef')

        # Write the c:f element.
        self._write_series_formula(formula)

        if ref_type == 'num':
            # Write the c:numCache element.
            self._write_num_cache(data)
        elif ref_type == 'str':
            # Write the c:strCache element.
            self._write_str_cache(data)

        self._xml_end_tag('c:strRef')

    def _write_series_formula(self, formula):
        # Write the <c:f> element.

        # Strip the leading '=' from the formula.
        if formula.startswith('='):
            formula = formula.lstrip('=')

        self._xml_data_element('c:f', formula)

    def _write_axis_ids(self, args):
        # Write the <c:axId> elements for the primary or secondary axes.

        # Generate the axis ids.
        self._add_axis_ids(args)

        if args['primary_axes']:

            # Write the axis ids for the primary axes.
            self._write_axis_id(self.axis_ids[0])
            self._write_axis_id(self.axis_ids[1])
        else:
            # Write the axis ids for the secondary axes.
            self._write_axis_id(self.axis2_ids[0])
            self._write_axis_id(self.axis2_ids[1])

    def _write_axis_id(self, val):
        # Write the <c:axId> element.

        attributes = [('val', val)]

        self._xml_empty_tag('c:axId', attributes)

    def _write_cat_axis(self, args):
        # Write the <c:catAx> element. Usually the X axis.
        x_axis = args['x_axis']
        y_axis = args['y_axis']
        axis_ids = args['axis_ids']

        # If there are no axis_ids then we don't need to write this element.
        if axis_ids is None or not len(axis_ids):
            return

        position = self.cat_axis_position
        horiz = self.horiz_cat_axis

        # Overwrite the default axis position with a user supplied value.
        if x_axis.get('position'):
            position = x_axis['position']

        self._xml_start_tag('c:catAx')

        self._write_axis_id(axis_ids[0])

        # Write the c:scaling element.
        self._write_scaling(x_axis.get('reverse'),
                            None,
                            None,
                            None)

        if not x_axis.get('visible'):
            self._write_delete(1)

        # Write the c:axPos element.
        self._write_axis_pos(position, y_axis.get('reverse'))

        # Write the c:majorGridlines element.
        self._write_major_gridlines(x_axis.get('major_gridlines'))

        # Write the c:minorGridlines element.
        self._write_minor_gridlines(x_axis.get('minor_gridlines'))

        # Write the axis title elements.
        if x_axis.get('formula'):
            self._write_title_formula(x_axis['formula'], x_axis['data_id'],
                                      horiz, x_axis['name_font'])
        elif x_axis.get('name'):
            self._write_title_rich(x_axis['name'], horiz, x_axis['name_font'])

        # Write the c:numFmt element.
        self._write_cat_number_format(x_axis)

        # Write the c:majorTickMark element.
        self._write_major_tick_mark(x_axis.get('major_tick_mark'))

        # Write the c:tickLblPos element.
        self._write_tick_label_pos(x_axis.get('label_position'))

        # Write the axis font elements.
        self._write_axis_font(x_axis.get('num_font'))

        # Write the c:crossAx element.
        self._write_cross_axis(axis_ids[1])

        if self.show_crosses or x_axis.get('visible'):

            # Note, the category crossing comes from the value axis.
            if (y_axis.get('crossing') is None
                    or y_axis.get('crossing') == 'max'):

                # Write the c:crosses element.
                self._write_crosses(y_axis.get('crossing'))
            else:

                # Write the c:crossesAt element.
                self._write_c_crosses_at(y_axis.get('crossing'))

        # Write the c:auto element.
        self._write_auto(1)

        # Write the c:labelAlign element.
        self._write_label_align('ctr')

        # Write the c:labelOffset element.
        self._write_label_offset(100)

        self._xml_end_tag('c:catAx')

    def _write_val_axis(self, args):
        # Write the <c:valAx> element. Usually the Y axis.
        x_axis = args['x_axis']
        y_axis = args['y_axis']
        axis_ids = args['axis_ids']
        position = args.get('position', self.val_axis_position)
        horiz = self.horiz_val_axis

        # If there are no axis_ids then we don't need to write this element.
        if axis_ids is None or not len(axis_ids):
            return

        # Overwrite the default axis position with a user supplied value.
        position = y_axis.get('position') or position

        self._xml_start_tag('c:valAx')

        self._write_axis_id(axis_ids[1])

        # Write the c:scaling element.
        self._write_scaling(y_axis.get('reverse'),
                            y_axis.get('min'),
                            y_axis.get('max'),
                            y_axis.get('log_base'))

        if not y_axis.get('visible'):
            self._write_delete(1)

        # Write the c:axPos element.
        self._write_axis_pos(position, x_axis.get('reverse'))

        # Write the c:majorGridlines element.
        self._write_major_gridlines(y_axis.get('major_gridlines'))

        # Write the c:minorGridlines element.
        self._write_minor_gridlines(y_axis.get('minor_gridlines'))

        # Write the axis title elements.
        if y_axis.get('formula'):
            self._write_title_formula(y_axis['formula'], y_axis['data_id'],
                                      horiz, y_axis['name_font'])
        elif y_axis.get('name'):
            self._write_title_rich(y_axis['name'],
                                   horiz,
                                   y_axis.get('name_font'))

        # Write the c:numberFormat element.
        self._write_number_format(y_axis)

        # Write the c:majorTickMark element.
        self._write_major_tick_mark(y_axis.get('major_tick_mark'))

        # Write the c:tickLblPos element.
        self._write_tick_label_pos(y_axis.get('label_position'))

        # Write the axis font elements.
        self._write_axis_font(y_axis.get('num_font'))

        # Write the c:crossAx element.
        self._write_cross_axis(axis_ids[0])

        # Note, the category crossing comes from the value axis.
        if x_axis.get('crossing') is None or x_axis['crossing'] == 'max':

            # Write the c:crosses element.
            self._write_crosses(x_axis.get('crossing'))
        else:

            # Write the c:crossesAt element.
            self._write_c_crosses_at(x_axis.get('crossing'))

        # Write the c:crossBetween element.
        self._write_cross_between()

        # Write the c:majorUnit element.
        self._write_c_major_unit(y_axis.get('major_unit'))

        # Write the c:minorUnit element.
        self._write_c_minor_unit(y_axis.get('minor_unit'))

        self._xml_end_tag('c:valAx')

    def _write_cat_val_axis(self, args):
        # Write the <c:valAx> element. This is for the second valAx
        # in scatter plots. Usually the X axis.
        x_axis = args['x_axis']
        y_axis = args['y_axis']
        axis_ids = args['axis_ids']
        position = args['position'] or self.val_axis_position
        horiz = self.horiz_val_axis

        # If there are no axis_ids then we don't need to write this element.
        if axis_ids is None or not len(axis_ids):
            return

        # Overwrite the default axis position with a user supplied value.
        position = x_axis.get('position') or position

        self._xml_start_tag('c:valAx')

        self._write_axis_id(axis_ids[0])

        # Write the c:scaling element.
        self._write_scaling(x_axis.get('reverse'),
                            x_axis.get('min'),
                            x_axis.get('max'),
                            x_axis.get('log_base'))

        if not x_axis.get('visible'):
            self._write_delete(1)

        # Write the c:axPos element.
        self._write_axis_pos(position, y_axis.get('reverse'))

        # Write the c:majorGridlines element.
        self._write_major_gridlines(x_axis.get('major_gridlines'))

        # Write the c:minorGridlines element.
        self._write_minor_gridlines(x_axis.get('minor_gridlines'))

        # Write the axis title elements.
        if x_axis.get('formula'):
            self._write_title_formula(x_axis['formula'], y_axis['data_id'],
                                      horiz, x_axis['name_font'])
        elif x_axis.get('name'):
            self._write_title_rich(x_axis['name_font'],
                                   horiz,
                                   x_axis['name_font'])

        # Write the c:numberFormat element.
        self._write_number_format(x_axis)

        # Write the c:majorTickMark element.
        self._write_major_tick_mark(x_axis.get('major_tick_mark'))

        # Write the c:tickLblPos element.
        self._write_tick_label_pos(x_axis.get('label_position'))

        # Write the axis font elements.
        self._write_axis_font(x_axis.get('num_font'))

        # Write the c:crossAx element.
        self._write_cross_axis(axis_ids[1])

        # Note, the category crossing comes from the value axis.
        if y_axis.get('crossing') is None or y_axis['crossing'] == 'max':

            # Write the c:crosses element.
            self._write_crosses(y_axis.get('crossing'))
        else:

            # Write the c:crossesAt element.
            self._write_c_crosses_at(y_axis.get('crossing'))

        # Write the c:crossBetween element.
        self._write_cross_between()

        # Write the c:majorUnit element.
        self._write_c_major_unit(x_axis.get('major_unit'))

        # Write the c:minorUnit element.
        self._write_c_minor_unit(x_axis.get('minor_unit'))

        self._xml_end_tag('c:valAx')

    def _write_date_axis(self, args):
        # Write the <c:dateAx> element. Usually the X axis.
        x_axis = args['x_axis']
        y_axis = args['y_axis']
        axis_ids = args['axis_ids']

        # If there are no axis_ids then we don't need to write this element.
        if axis_ids is None or not len(axis_ids):
            return

        position = self.cat_axis_position

        # Overwrite the default axis position with a user supplied value.
        position = x_axis.get('position') or position

        self._xml_start_tag('c:dateAx')

        self._write_axis_id(axis_ids[0])

        # Write the c:scaling element.
        self._write_scaling(x_axis.get('reverse'),
                            x_axis.get('min'),
                            x_axis.get('max'),
                            x_axis.get('log_base'))

        if not x_axis.get('visible'):
            self._write_delete(1)

        # Write the c:axPos element.
        self._write_axis_pos(position, y_axis.get('reverse'))

        # Write the c:majorGridlines element.
        self._write_major_gridlines(x_axis.get('major_gridlines'))

        # Write the c:minorGridlines element.
        self._write_minor_gridlines(x_axis.get('minor_gridlines'))

        # Write the axis title elements.
        if x_axis.get('formula'):
            self._write_title_formula(x_axis['formula'],
                                      x_axis['data_id'],
                                      None,
                                      x_axis['name_font'])
        elif x_axis.get('name'):
            self._write_title_rich(x_axis['name_font'],
                                   None,
                                   x_axis['name_font'])

        # Write the c:numFmt element.
        self._write_number_format(x_axis)

        # Write the c:majorTickMark element.
        self._write_major_tick_mark(x_axis.get('major_tick_mark'))

        # Write the c:tickLblPos element.
        self._write_tick_label_pos(x_axis.get('label_position'))

        # Write the axis font elements.
        self._write_axis_font(x_axis.get('num_font'))

        # Write the c:crossAx element.
        self._write_cross_axis(axis_ids[1])

        if self.show_crosses or x_axis.get('visible'):

            # Note, the category crossing comes from the value axis.
            if (y_axis.get('crossing') is None
                    or y_axis.get('crossing') == 'max'):

                # Write the c:crosses element.
                self._write_crosses(y_axis.get('crossing'))
            else:

                # Write the c:crossesAt element.
                self._write_c_crosses_at(y_axis.get('crossing'))

        # Write the c:auto element.
        self._write_auto(1)

        # Write the c:labelOffset element.
        self._write_label_offset(100)

        # Write the c:majorUnit element.
        self._write_c_major_unit(x_axis.get('major_unit'))

        # Write the c:majorTimeUnit element.
        if x_axis.get('major_unit'):
            self._write_c_major_time_unit(x_axis['major_unit_type'])

        # Write the c:minorUnit element.
        self._write_c_minor_unit(x_axis.get('minor_unit'))

        # Write the c:minorTimeUnit element.
        if x_axis.get('minor_unit'):
            self._write_c_minor_time_unit(x_axis['minor_unit_type'])

        self._xml_end_tag('c:dateAx')

    def _write_scaling(self, reverse, min_val, max_val, log_base):
        # Write the <c:scaling> element.

        self._xml_start_tag('c:scaling')

        # Write the c:logBase element.
        self._write_c_log_base(log_base)

        # Write the c:orientation element.
        self._write_orientation(reverse)

        # Write the c:max element.
        self._write_c_max(max_val)

        # Write the c:min element.
        self._write_c_min(min_val)

        self._xml_end_tag('c:scaling')

    def _write_c_log_base(self, val):
        # Write the <c:logBase> element.

        if not val:
            return

        attributes = [('val', val)]

        self._xml_empty_tag('c:logBase', attributes)

    def _write_orientation(self, reverse):
        # Write the <c:orientation> element.
        val = 'minMax'

        if reverse:
            val = 'maxMin'

        attributes = [('val', val)]

        self._xml_empty_tag('c:orientation', attributes)

    def _write_c_max(self, max_val):
        # Write the <c:max_val> element.

        if max_val is None:
            return

        attributes = [('val', max_val)]

        self._xml_empty_tag('c:max', attributes)

    def _write_c_min(self, min_val):
        # Write the <c:min_val> element.

        if min_val is None:
            return

        attributes = [('val', min_val)]

        self._xml_empty_tag('c:min', attributes)

    def _write_axis_pos(self, val, reverse):
        # Write the <c:axPos> element.

        if reverse:
            if val == 'l':
                val = 'r'
            if val == 'b':
                val = 't'

        attributes = [('val', val)]

        self._xml_empty_tag('c:axPos', attributes)

    def _write_number_format(self, axis):
        # Write the <c:numberFormat> element. Note: It is assumed that if
        # a user defined number format is supplied (i.e., non-default) then
        # the sourceLinked attribute is 0.
        # The user can override this if required.
        format_code = axis.get('num_format')
        source_linked = 1

        # Check if a user defined number format has been set.
        if (format_code is not None
                and format_code != axis['defaults']['num_format']):
            source_linked = 0

        # User override of sourceLinked.
        if axis.get('num_format_linked'):
            source_linked = 1

        attributes = [
            ('formatCode', format_code),
            ('sourceLinked', source_linked),
        ]

        self._xml_empty_tag('c:numFmt', attributes)

    def _write_cat_number_format(self, axis):
        # Write the <c:numFmt> element. Special case handler for category
        # axes which don't always have a number format.
        format_code = axis.get('num_format')
        source_linked = 1
        default_format = 1

        # Check if a user defined number format has been set.
        if (format_code is not None
                and format_code != axis['defaults']['num_format']):
            source_linked = 0
            default_format = 0

        # User override of linkedSource.
        if axis.get('num_format_linked'):
            source_linked = 1

        # Skip if cat doesn't have a num format (unless it is non-default).
        if not self.cat_has_num_fmt and default_format:
            return

        attributes = [
            ('formatCode', format_code),
            ('sourceLinked', source_linked),
        ]

        self._xml_empty_tag('c:numFmt', attributes)

    def _write_major_tick_mark(self, val):
        # Write the <c:majorTickMark> element.

        if not val:
            return

        attributes = [('val', val)]

        self._xml_empty_tag('c:majorTickMark', attributes)

    def _write_tick_label_pos(self, val=None):
        # Write the <c:tickLblPos> element.
        if val is None or val == 'next_to':
            val = 'nextTo'

        attributes = [('val', val)]

        self._xml_empty_tag('c:tickLblPos', attributes)

    def _write_cross_axis(self, val):
        # Write the <c:crossAx> element.

        attributes = [('val', val)]

        self._xml_empty_tag('c:crossAx', attributes)

    def _write_crosses(self, val=None):
        # Write the <c:crosses> element.
        if val is None:
            val = 'autoZero'

        attributes = [('val', val)]

        self._xml_empty_tag('c:crosses', attributes)

    def _write_c_crosses_at(self, val):
        # Write the <c:crossesAt> element.

        attributes = [('val', val)]

        self._xml_empty_tag('c:crossesAt', attributes)

    def _write_auto(self, val):
        # Write the <c:auto> element.

        attributes = [('val', val)]

        self._xml_empty_tag('c:auto', attributes)

    def _write_label_align(self, val):
        # Write the <c:labelAlign> element.

        attributes = [('val', val)]

        self._xml_empty_tag('c:lblAlgn', attributes)

    def _write_label_offset(self, val):
        # Write the <c:labelOffset> element.

        attributes = [('val', val)]

        self._xml_empty_tag('c:lblOffset', attributes)

    def _write_major_gridlines(self, gridlines):
        # Write the <c:majorGridlines> element.

        if not gridlines:
            return

        if not gridlines['visible']:
            return

        if gridlines['line']['defined']:
            self._xml_start_tag('c:majorGridlines')

            # Write the c:spPr element.
            self._write_sp_pr(gridlines)

            self._xml_end_tag('c:majorGridlines')
        else:
            self._xml_empty_tag('c:majorGridlines')

    def _write_minor_gridlines(self, gridlines):
        # Write the <c:minorGridlines> element.

        if not gridlines:
            return
        if not gridlines.visible:
            return

        if gridlines['line']['defined']:
            self._xml_start_tag('c:minorGridlines')

            # Write the c:spPr element.
            self._write_sp_pr(gridlines)

            self._xml_end_tag('c:minorGridlines')
        else:
            self._xml_empty_tag('c:minorGridlines')

    def _write_cross_between(self):
        # Write the <c:crossBetween> element.
        val = self.cross_between

        if val is None:
            val = 'between'

        attributes = [('val', val)]

        self._xml_empty_tag('c:crossBetween', attributes)

    def _write_c_major_unit(self, val):
        # Write the <c:majorUnit> element.

        if not val:
            return

        attributes = [('val', val)]

        self._xml_empty_tag('c:majorUnit', attributes)

    def _write_c_minor_unit(self, val):
        # Write the <c:minorUnit> element.

        if not val:
            return

        attributes = [('val', val)]

        self._xml_empty_tag('c:minorUnit', attributes)

    def _write_c_major_time_unit(self, val=None):
        # Write the <c:majorTimeUnit> element.
        if val is None:
            val = 'days'

        attributes = [('val', val)]

        self._xml_empty_tag('c:majorTimeUnit', attributes)

    def _write_c_minor_time_unit(self, val=None):
        # Write the <c:minorTimeUnit> element.
        if val is None:
            val = 'days'

        attributes = [('val', val)]

        self._xml_empty_tag('c:minorTimeUnit', attributes)

    def _write_legend(self):
        # Write the <c:legend> element.
        position = self.legend_position
        delete_series = []
        overlay = 0

        # if (self.legend_delete_series is not None
        #    and ref self.legend_delete_series == 'ARRAY'):
        #    delete_series =  self.legend_delete_series

        # if position =~ s/^overlay_//:
        #    overlay = 1
        allowed = {
            'right': 'r',
            'left': 'l',
            'top': 't',
            'bottom': 'b',
        }

        if position == 'none':
            return

        if not position in allowed:
            return

        position = allowed[position]

        self._xml_start_tag('c:legend')

        # Write the c:legendPos element.
        self._write_legend_pos(position)

        # Remove series labels from the legend.
        for index in (delete_series):

            # Write the c:legendEntry element.
            self._write_legend_entry(index)

        # Write the c:layout element.
        self._write_layout()

        # Write the c:overlay element.
        if overlay:
            self._write_overlay()

        self._xml_end_tag('c:legend')

    def _write_legend_pos(self, val):
        # Write the <c:legendPos> element.

        attributes = [('val', val)]

        self._xml_empty_tag('c:legendPos', attributes)

    def _write_legend_entry(self, index):
        # Write the <c:legendEntry> element.

        self._xml_start_tag('c:legendEntry')

        # Write the c:idx element.
        self._write_idx(index)

        # Write the c:delete element.
        self._write_delete(1)

        self._xml_end_tag('c:legendEntry')

    def _write_overlay(self):
        # Write the <c:overlay> element.
        val = 1

        attributes = [('val', val)]

        self._xml_empty_tag('c:overlay', attributes)

    def _write_plot_vis_only(self):
        # Write the <c:plotVisOnly> element.
        val = 1

        # Ignore this element if we are plotting hidden data.
        if self.show_hidden_data:
            return

        attributes = [('val', val)]

        self._xml_empty_tag('c:plotVisOnly', attributes)

    def _write_print_settings(self):
        # Write the <c:printSettings> element.
        self._xml_start_tag('c:printSettings')

        # Write the c:headerFooter element.
        self._write_header_footer()

        # Write the c:pageMargins element.
        self._write_page_margins()

        # Write the c:pageSetup element.
        self._write_page_setup()

        self._xml_end_tag('c:printSettings')

    def _write_header_footer(self):
        # Write the <c:headerFooter> element.
        self._xml_empty_tag('c:headerFooter')

    def _write_page_margins(self):
        # Write the <c:pageMargins> element.
        b = 0.75
        l = 0.7
        r = 0.7
        t = 0.75
        header = 0.3
        footer = 0.3

        attributes = [
            ('b', b),
            ('l', l),
            ('r', r),
            ('t', t),
            ('header', header),
            ('footer', footer),
        ]

        self._xml_empty_tag('c:pageMargins', attributes)

    def _write_page_setup(self):
        # Write the <c:pageSetup> element.
        self._xml_empty_tag('c:pageSetup')

    def _write_title_rich(self, title, horiz, font):
        # Write the <c:title> element for a rich string.

        self._xml_start_tag('c:title')

        # Write the c:tx element.
        self._write_tx_rich(title, horiz, font)

        # Write the c:layout element.
        self._write_layout()

        self._xml_end_tag('c:title')

    def _write_title_formula(self, title, data_id, horiz, font):
        # Write the <c:title> element for a rich string.

        self._xml_start_tag('c:title')

        # Write the c:tx element.
        self._write_tx_formula(title, data_id)

        # Write the c:layout element.
        self._write_layout()

        # Write the c:txPr element.
        self._write_tx_pr(horiz, font)

        self._xml_end_tag('c:title')

    def _write_tx_rich(self, title, horiz, font):
        # Write the <c:tx> element.

        self._xml_start_tag('c:tx')

        # Write the c:rich element.
        self._write_rich(title, horiz, font)

        self._xml_end_tag('c:tx')

    def _write_tx_value(self, title):
        # Write the <c:tx> element with a value such as for series names.

        self._xml_start_tag('c:tx')

        # Write the c:v element.
        self._write_v(title)

        self._xml_end_tag('c:tx')

    def _write_tx_formula(self, title, data_id):
        # Write the <c:tx> element.
        data = None

        if data_id is not None:
            data = self.formula_data[data_id]

        self._xml_start_tag('c:tx')

        # Write the c:strRef element.
        self._write_str_ref(title, data, 'str')

        self._xml_end_tag('c:tx')

    def _write_rich(self, title, horiz, font):
        # Write the <c:rich> element.

        self._xml_start_tag('c:rich')

        # Write the a:bodyPr element.
        self._write_a_body_pr(horiz)

        # Write the a:lstStyle element.
        self._write_a_lst_style()

        # Write the a:p element.
        self._write_a_p_rich(title, font)

        self._xml_end_tag('c:rich')

    def _write_a_body_pr(self, horiz):
        # Write the <a:bodyPr> element.
        rot = -5400000
        vert = 'horz'

        attributes = [
            ('rot', rot),
            ('vert', vert),
        ]

        if not horiz:
            attributes = []

        self._xml_empty_tag('a:bodyPr', attributes)

    def _write_a_lst_style(self):
        # Write the <a:lstStyle> element.
        self._xml_empty_tag('a:lstStyle')

    def _write_a_p_rich(self, title, font):
        # Write the <a:p> element for rich string titles.

        self._xml_start_tag('a:p')

        # Write the a:pPr element.
        self._write_a_p_pr_rich(font)

        # Write the a:r element.
        self._write_a_r(title, font)

        self._xml_end_tag('a:p')

    def _write_a_p_formula(self, font):
        # Write the <a:p> element for formula titles.

        self._xml_start_tag('a:p')

        # Write the a:pPr element.
        self._write_a_p_pr_formula(font)

        # Write the a:endParaRPr element.
        self._write_a_end_para_rpr()

        self._xml_end_tag('a:p')

    def _write_a_p_pr_rich(self, font):
        # Write the <a:pPr> element for rich string titles.

        self._xml_start_tag('a:pPr')

        # Write the a:defRPr element.
        self._write_a_def_rpr(font)

        self._xml_end_tag('a:pPr')

    def _write_a_p_pr_formula(self, font):
        # Write the <a:pPr> element for formula titles.

        self._xml_start_tag('a:pPr')

        # Write the a:defRPr element.
        self._write_a_def_rpr(font)

        self._xml_end_tag('a:pPr')

    def _write_a_def_rpr(self, font):
        # Write the <a:defRPr> element.
        has_color = 0

        style_attributes = self._get_font_style_attributes(font)
        latin_attributes = self._get_font_latin_attributes(font)

        if font and 'color' in font:
            has_color = 1

        if latin_attributes or has_color:
            self._xml_start_tag('a:defRPr', style_attributes)

            if has_color:
                self._write_a_solid_fill({'color': font['color']})

            if latin_attributes:
                self._write_a_latin(latin_attributes)

            self._xml_end_tag('a:defRPr')
        else:
            self._xml_empty_tag('a:defRPr', style_attributes)

    def _write_a_end_para_rpr(self):
        # Write the <a:endParaRPr> element.
        lang = 'en-US'

        attributes = [('lang', lang)]

        self._xml_empty_tag('a:endParaRPr', attributes)

    def _write_a_r(self, title, font):
        # Write the <a:r> element.

        self._xml_start_tag('a:r')

        # Write the a:rPr element.
        self._write_a_r_pr(font)

        # Write the a:t element.
        self._write_a_t(title)

        self._xml_end_tag('a:r')

    def _write_a_r_pr(self, font):
        # Write the <a:rPr> element.
        has_color = 0
        lang = 'en-US'

        style_attributes = self._get_font_style_attributes(font)
        latin_attributes = self._get_font_latin_attributes(font)

        if font and 'color' in font:
            has_color = 1

        # Add the lang type to the attributes.
        style_attributes.insert(0, ('lang', lang))

        if latin_attributes or has_color:
            self._xml_start_tag('a:rPr', style_attributes)

            if has_color:
                self._write_a_solid_fill({'color': font['color']})

            if latin_attributes:
                self._write_a_latin(latin_attributes)

            self._xml_end_tag('a:rPr')
        else:
            self._xml_empty_tag('a:rPr', style_attributes)

    def _write_a_t(self, title):
        # Write the <a:t> element.

        self._xml_data_element('a:t', title)

    def _write_tx_pr(self, horiz, font):
        # Write the <c:txPr> element.

        self._xml_start_tag('c:txPr')

        # Write the a:bodyPr element.
        self._write_a_body_pr(horiz)

        # Write the a:lstStyle element.
        self._write_a_lst_style()

        # Write the a:p element.
        self._write_a_p_formula(font)

        self._xml_end_tag('c:txPr')

    def _write_marker(self, marker):
        # Write the <c:marker> element.
        if marker is None:
            marker = self.default_marker

        if not marker:
            return
        if marker['automatic']:
            return

        self._xml_start_tag('c:marker')

        # Write the c:symbol element.
        self._write_symbol(marker['type'])

        # Write the c:size element.
        size = marker['size']
        if size:
            self._write_marker_size(size)

        # Write the c:spPr element.
        self._write_sp_pr(marker)

        self._xml_end_tag('c:marker')

    def _write_marker_value(self):
        # Write the <c:marker> element without a sub-element.
        style = self.default_marker

        if not style:
            return

        attributes = [('val', 1)]

        self._xml_empty_tag('c:marker', attributes)

    def _write_marker_size(self, val):
        # Write the <c:size> element.

        attributes = [('val', val)]

        self._xml_empty_tag('c:size', attributes)

    def _write_symbol(self, val):
        # Write the <c:symbol> element.

        attributes = [('val', val)]

        self._xml_empty_tag('c:symbol', attributes)

    def _write_sp_pr(self, series):
        # Write the <c:spPr> element.
        has_fill = False
        has_line = False

        if 'fill' in series and series['fill']['defined']:
            has_fill = True

        if 'line' in series and series['line']['defined']:
            has_line = True

        if not has_fill and not has_line:
            return

        self._xml_start_tag('c:spPr')

        # Write the fill elements for solid charts such as pie and bar.
        if 'fill' in series and series['fill']['defined']:
            if series.fill['none']:

                # Write the a:noFill element.
                self._write_a_no_fill()
            else:
                # Write the a:solidFill element.
                self._write_a_solid_fill(series['fill'])

        # Write the a:ln element.
        if 'line' in series and series['line']['defined']:
            self._write_a_ln(series['line'])

        self._xml_end_tag('c:spPr')

    def _write_a_ln(self, line):
        # Write the <a:ln> element.
        attributes = []

        # Add the line width as an attribute.
        width = line['width']

        if width:
            # Round width to nearest 0.25, like Excel.
            width = int((width + 0.125) * 4) / 4

            # Convert to internal units.
            width = int(0.5 + (12700 * width))

            attributes = [('w', width)]

        self._xml_start_tag('a:ln', attributes)

        # Write the line fill.
        if line['none']:

            # Write the a:noFill element.
            self._write_a_no_fill()
        elif line['color']:

            # Write the a:solidFill element.
            self._write_a_solid_fill(line)

        # Write the line/dash type.
        line_type = line['dash_type']
        if line_type:
            # Write the a:prstDash element.
            self._write_a_prst_dash(line_type)

        self._xml_end_tag('a:ln')

    def _write_a_no_fill(self):
        # Write the <a:noFill> element.
        self._xml_empty_tag('a:noFill')

    def _write_a_solid_fill(self, line):
        # Write the <a:solidFill> element.

        self._xml_start_tag('a:solidFill')

        if line['color']:

            color = self._get_color(line['color'])

            # Write the a:srgbClr element.
            self._write_a_srgb_clr(color)

        self._xml_end_tag('a:solidFill')

    def _write_a_srgb_clr(self, val):
        # Write the <a:srgbClr> element.

        attributes = [('val', val)]

        self._xml_empty_tag('a:srgbClr', attributes)

    def _write_a_prst_dash(self, val):
        # Write the <a:prstDash> element.

        attributes = [('val', val)]

        self._xml_empty_tag('a:prstDash', attributes)

    def _write_trendline(self, trendline):
        # Write the <c:trendline> element.

        if not trendline:
            return

        self._xml_start_tag('c:trendline')

        # Write the c:name element.
        self._write_name(trendline['name'])

        # Write the c:spPr element.
        self._write_sp_pr(trendline)

        # Write the c:trendlineType element.
        self._write_trendline_type(trendline['type'])

        # Write the c:order element for polynomial trendlines.
        if trendline['type'] == 'poly':
            self._write_trendline_order(trendline['order'])

        # Write the c:period element for moving average trendlines.
        if trendline['type'] == 'movingAvg':
            self._write_period(trendline['period'])

        # Write the c:forward element.
        self._write_forward(trendline['forward'])

        # Write the c:backward element.
        self._write_backward(trendline['backward'])

        self._xml_end_tag('c:trendline')

    def _write_trendline_type(self, val):
        # Write the <c:trendlineType> element.

        attributes = [('val', val)]

        self._xml_empty_tag('c:trendlineType', attributes)

    def _write_name(self, data):
        # Write the <c:name> element.

        if data is None:
            return

        self._xml_data_element('c:name', data)

    def _write_trendline_order(self, val):
        # Write the <c:order> element.
        # val = _[0] is not None ? _[0]: 2

        attributes = [('val', val)]

        self._xml_empty_tag('c:order', attributes)

    def _write_period(self, val):
        # Write the <c:period> element.
        # val = _[0] is not None ? _[0]: 2

        attributes = [('val', val)]

        self._xml_empty_tag('c:period', attributes)

    def _write_forward(self, val):
        # Write the <c:forward> element.

        if not val:
            return

        attributes = [('val', val)]

        self._xml_empty_tag('c:forward', attributes)

    def _write_backward(self, val):
        # Write the <c:backward> element.

        if not val:
            return

        attributes = [('val', val)]

        self._xml_empty_tag('c:backward', attributes)

    def _write_hi_low_lines(self):
        # Write the <c:hiLowLines> element.
        hi_low_lines = self.hi_low_lines

        if not hi_low_lines:
            return

        if hi_low_lines.line['defined']:

            self._xml_start_tag('c:hiLowLines')

            # Write the c:spPr element.
            self._write_sp_pr(hi_low_lines)

            self._xml_end_tag('c:hiLowLines')
        else:
            self._xml_empty_tag('c:hiLowLines')

    def _write_drop_lines(self):
        # Write the <c:dropLines> element.
        drop_lines = self.drop_lines

        if not drop_lines:
            return

        if drop_lines.line['defined']:

            self._xml_start_tag('c:dropLines')

            # Write the c:spPr element.
            self._write_sp_pr(drop_lines)

            self._xml_end_tag('c:dropLines')
        else:
            self._xml_empty_tag('c:dropLines')

    def _write_overlap(self, val):
        # Write the <c:overlap> element.

        if val is None:
            return

        attributes = [('val', val)]

        self._xml_empty_tag('c:overlap', attributes)

    def _write_num_cache(self, data):
        # Write the <c:numCache> element.
        if data:
            count = len(data)
        else:
            count = 0

        self._xml_start_tag('c:numCache')

        # Write the c:formatCode element.
        self._write_format_code('General')

        # Write the c:ptCount element.
        self._write_pt_count(count)

        for i in range(count):
            token = data[i]

            if token is None:
                continue

            try:
                float(token)
            except ValueError:
                # Write non-numeric data as 0.
                token = 0

            # Write the c:pt element.
            self._write_pt(i, token)

        self._xml_end_tag('c:numCache')

    def _write_str_cache(self, data):
        # Write the <c:strCache> element.
        count = len(data)

        self._xml_start_tag('c:strCache')

        # Write the c:ptCount element.
        self._write_pt_count(count)

        for i in range(count):
            # Write the c:pt element.
            self._write_pt(i, data[i])

        self._xml_end_tag('c:strCache')

    def _write_format_code(self, data):
        # Write the <c:formatCode> element.

        self._xml_data_element('c:formatCode', data)

    def _write_pt_count(self, val):
        # Write the <c:ptCount> element.

        attributes = [('val', val)]

        self._xml_empty_tag('c:ptCount', attributes)

    def _write_pt(self, idx, value):
        # Write the <c:pt> element.

        if value is None:
            return

        attributes = [('idx', idx)]

        self._xml_start_tag('c:pt', attributes)

        # Write the c:v element.
        self._write_v(value)

        self._xml_end_tag('c:pt')

    def _write_v(self, data):
        # Write the <c:v> element.

        self._xml_data_element('c:v', data)

    def _write_protection(self):
        # Write the <c:protection> element.
        if not self.protection:
            return

        self._xml_empty_tag('c:protection')

    def _write_d_pt(self, points):
        # Write the <c:dPt> elements.
        index = -1

        if not points:
            return

        for point in (points):
            index += 1
            if not point:
                continue

            self._write_d_pt_point(index, point)

    def _write_d_pt_point(self, index, point):
        # Write an individual <c:dPt> element.

            self._xml_start_tag('c:dPt')

            # Write the c:idx element.
            self._write_idx(index)

            # Write the c:spPr element.
            self._write_sp_pr(point)

            self._xml_end_tag('c:dPt')

    def _write_d_lbls(self, labels):
        # Write the <c:dLbls> element.

        if not labels:
            return

        self._xml_start_tag('c:dLbls')

        # Write the c:dLblPos element.
        if labels['position']:
            self._write_d_lbl_pos(labels['position'])

        # Write the c:showVal element.
        if labels['value']:
            self._write_show_val()

        # Write the c:showCatName element.
        if labels['category']:
            self._write_show_cat_name()

        # Write the c:showSerName element.
        if labels['series_name']:
            self._write_show_ser_name()

        # Write the c:showPercent element.
        if labels['percentage']:
            self._write_show_percent()

        # Write the c:showLeaderLines element.
        if labels['leader_lines']:
            self._write_show_leader_lines()

        self._xml_end_tag('c:dLbls')

    def _write_show_val(self):
        # Write the <c:showVal> element.
        val = 1

        attributes = [('val', val)]

        self._xml_empty_tag('c:showVal', attributes)

    def _write_show_cat_name(self):
        # Write the <c:showCatName> element.
        val = 1

        attributes = [('val', val)]

        self._xml_empty_tag('c:showCatName', attributes)

    def _write_show_ser_name(self):
        # Write the <c:showSerName> element.
        val = 1

        attributes = [('val', val)]

        self._xml_empty_tag('c:showSerName', attributes)

    def _write_show_percent(self):
        # Write the <c:showPercent> element.
        val = 1

        attributes = [('val', val)]

        self._xml_empty_tag('c:showPercent', attributes)

    def _write_show_leader_lines(self):
        # Write the <c:showLeaderLines> element.
        val = 1

        attributes = [('val', val)]

        self._xml_empty_tag('c:showLeaderLines', attributes)

    def _write_d_lbl_pos(self, val):
        # Write the <c:dLblPos> element.

        attributes = [('val', val)]

        self._xml_empty_tag('c:dLblPos', attributes)

    def _write_delete(self, val):
        # Write the <c:delete> element.

        attributes = [('val', val)]

        self._xml_empty_tag('c:delete', attributes)

    def _write_c_invert_if_negative(self, invert):
        # Write the <c:invertIfNegative> element.
        val = 1

        if not invert:
            return

        attributes = [('val', val)]

        self._xml_empty_tag('c:invertIfNegative', attributes)

    def _write_axis_font(self, font):
        # Write the axis font elements.

        if not font:
            return

        self._xml_start_tag('c:txPr')
        self._xml_empty_tag('a:bodyPr')
        self._write_a_lst_style()
        self._xml_start_tag('a:p')

        self._write_a_p_pr_rich(font)

        self._write_a_end_para_rpr()
        self._xml_end_tag('a:p')
        self._xml_end_tag('c:txPr')

    def _write_a_latin(self):
        # Write the <a:latin> element.
        attributes = _

        self._xml_empty_tag('a:latin', attributes)

    def _write_d_table(self):
        # Write the <c:dTable> element.
        table = self.table

        if not table:
            return

        self._xml_start_tag('c:dTable')

        if table.horizontal:

            # Write the c:showHorzBorder element.
            self._write_show_horz_border()

        if table.vertical:

            # Write the c:showVertBorder element.
            self._write_show_vert_border()

        if table.outline:

            # Write the c:showOutline element.
            self._write_show_outline()

        if table.show_keys:

            # Write the c:showKeys element.
            self._write_show_keys()

        self._xml_end_tag('c:dTable')

    def _write_show_horz_border(self):
        # Write the <c:showHorzBorder> element.
        attributes = [('val', 1)]

        self._xml_empty_tag('c:showHorzBorder', attributes)

    def _write_show_vert_border(self):
        # Write the <c:showVertBorder> element.
        attributes = [('val', 1)]

        self._xml_empty_tag('c:showVertBorder', attributes)

    def _write_show_outline(self):
        # Write the <c:showOutline> element.
        attributes = [('val', 1)]

        self._xml_empty_tag('c:showOutline', attributes)

    def _write_show_keys(self):
        # Write the <c:showKeys> element.
        attributes = [('val', 1)]

        self._xml_empty_tag('c:showKeys', attributes)

    def _write_error_bars(self, error_bars):
        # Write the X and Y error bars.

        if not error_bars:
            return

        if error_bars['x_error_bars']:
            self._write_err_bars(('x', error_bars['x_error_bars']))

        if error_bars['y_error_bars']:
            self._write_err_bars(('y', error_bars['y_error_bars']))

    def _write_err_bars(self, direction, error_bars):
        # Write the <c:errBars> element.

        if not error_bars:
            return

        self._xml_start_tag('c:errBars')

        # Write the c:errDir element.
        self._write_err_dir(direction)

        # Write the c:errBarType element.
        self._write_err_bar_type(error_bars['direction'])

        # Write the c:errValType element.
        self._write_err_val_type(error_bars['type'])

        if not error_bars['endcap']:

            # Write the c:noEndCap element.
            self._write_no_end_cap()

        if error_bars['type'] != 'stdErr':

            # Write the c:val element.
            self._write_error_val(error_bars['value'])

        # Write the c:spPr element.
        self._write_sp_pr(error_bars)

        self._xml_end_tag('c:errBars')

    def _write_err_dir(self, val):
        # Write the <c:errDir> element.

        attributes = [('val', val)]

        self._xml_empty_tag('c:errDir', attributes)

    def _write_err_bar_type(self, val):
        # Write the <c:errBarType> element.

        attributes = [('val', val)]

        self._xml_empty_tag('c:errBarType', attributes)

    def _write_err_val_type(self, val):
        # Write the <c:errValType> element.

        attributes = [('val', val)]

        self._xml_empty_tag('c:errValType', attributes)

    def _write_no_end_cap(self):
        # Write the <c:noEndCap> element.
        attributes = [('val', 1)]

        self._xml_empty_tag('c:noEndCap', attributes)

    def _write_error_val(self, val):
        # Write the <c:val> element for error bars.

        attributes = [('val', val)]

        self._xml_empty_tag('c:val', attributes)

    def _write_up_down_bars(self):
        # Write the <c:upDownBars> element.
        up_down_bars = self.up_down_bars

        if not up_down_bars:
            return

        self._xml_start_tag('c:upDownBars')

        # Write the c:gapWidth element.
        self._write_gap_width(150)

        # Write the c:upBars element.
        self._write_up_bars(up_down_bars['up'])

        # Write the c:downBars element.
        self._write_down_bars(up_down_bars['down'])

        self._xml_end_tag('c:upDownBars')

    def _write_gap_width(self, val):
        # Write the <c:gapWidth> element.

        if val is None:
            return

        attributes = [('val', val)]

        self._xml_empty_tag('c:gapWidth', attributes)

    def _write_up_bars(self, bar_format):
        # Write the <c:upBars> element.

        # if (format.line.or is not None format.fill.) is not None:
        if not 'TODO':

            self._xml_start_tag('c:upBars')

            # Write the c:spPr element.
            self._write_sp_pr(bar_format)

            self._xml_end_tag('c:upBars')
        else:
            self._xml_empty_tag('c:upBars')
