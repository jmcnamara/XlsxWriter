###############################################################################
#
# ChartLine - A class for writing the Excel XLSX Line charts.
#
# Copyright 2013, John McNamara, jmcnamara@cpan.org
#

from . import chart


class ChartLine(chart.Chart):
    """
    A class for writing the Excel XLSX Line charts.


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
        super(ChartLine, self).__init__()

        if options is None:
            options = {}

        self.subtype = options.get('subtype')

        if not self.subtype:
            self.subtype = 'straight'

        self.default_marker = {'type': 'none'}
        self.smooth_allowed = True

    ###########################################################################
    #
    # Private API.
    #
    ###########################################################################

    def _write_chart_type(self, args):
        # Override the virtual superclass method with a chart specific method.
        # Write the c:lineChart element.
        self._write_line_chart(args)

    ###########################################################################
    #
    # XML methods.
    #
    ###########################################################################

    def _write_line_chart(self, args):
        # Write the <c:lineChart> element.

        if args['primary_axes']:
            series = self._get_primary_axes_series()
        else:
            series = self._get_secondary_axes_series()

        if not len(series):
            return

        self._modify_series_formatting()

        self._xml_start_tag('c:lineChart')

        # Write the c:grouping element.
        self._write_grouping('standard')

        # Write the series elements.
        for data in series:
            self._write_ser(data)

        # Write the c:dropLines element.
        self._write_drop_lines()

        # Write the c:hiLowLines element.
        self._write_hi_low_lines()

        # Write the c:upDownBars element.
        self._write_up_down_bars()

        # Write the c:marker element.
        self._write_marker_value()

        # Write the c:axId elements
        self._write_axis_ids(args)

        self._xml_end_tag('c:lineChart')

    def _write_d_pt_point(self, index, point):
        # Write an individual <c:dPt> element. Override the parent method to
        # add markers.

        self._xml_start_tag('c:dPt')

        # Write the c:idx element.
        self._write_idx(index)

        self._xml_start_tag('c:marker')

        # Write the c:spPr element.
        self._write_sp_pr(point)

        self._xml_end_tag('c:marker')

        self._xml_end_tag('c:dPt')

    def _modify_series_formatting(self):
        # Add default formatting to the series data unless it has already been
        # specified by the user.
        subtype = self.subtype

        # The default scatter style "markers only" requires a line type.
        if subtype == 'marker_only':

            # Go through each series and define default values.
            for series in self.series:

                # Set a line type unless there is already a user defined type.
                if not series['line']['defined']:
                    series['line'] = {'width': 2.25,
                                      'none': 1,
                                      'defined': 1,
                                      }

        # Turn markers on for subtypes that have them.
        if 'marker' in subtype:

            # Go through each series and define default values.
            for series in self.series:
                # Set a marker type unless there is a user defined type.
                if series['marker'] is None or not series['marker']['defined']:
                    series['marker'] = {'type': 'automatic',
                                        'automatic': 1,
                                        'defined': 1,
                                        'line': self._get_line_properties(None),
                                        'fill': self._get_fill_properties(None)
                                        }

        # Turn on smoothing if required
        if 'smooth' in subtype:
            for series in self.series:
                if series['smooth'] is None:
                    series['smooth'] = True
