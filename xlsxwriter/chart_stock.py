###############################################################################
#
# ChartStock - A class for writing the Excel XLSX Stock charts.
#
# Copyright 2013, John McNamara, jmcnamara@cpan.org
#

from . import chart


class ChartStock(chart.Chart):
    """
    A class for writing the Excel XLSX Stock charts.

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
        super(ChartStock, self).__init__()

        if options is None:
            options = {}

        self.show_crosses = 0
        self.hi_low_lines = {}

        # Override and reset the default axis values.
        self.x_axis['defaults']['num_format'] = 'dd/mm/yyyy'
        self.x2_axis['defaults']['num_format'] = 'dd/mm/yyyy'

        self.set_x_axis({})
        self.set_x2_axis({})

    ###########################################################################
    #
    # Private API.
    #
    ###########################################################################

    def _write_chart_type(self, args):
        # Override the virtual superclass method with a chart specific method.
        # Write the c:stockChart element.
        self._write_stock_chart(args)

    ###########################################################################
    #
    # XML methods.
    #
    ###########################################################################

    def _write_stock_chart(self, args):
    # Write the <c:stockChart> element.
    # Overridden to add hi_low_lines().

        if args['primary_axes']:
            series = self._get_primary_axes_series()
        else:
            series = self._get_secondary_axes_series()

        if not len(series):
            return

        # Add default formatting to the series data.
        self._modify_series_formatting()

        self._xml_start_tag('c:stockChart')

        # Write the series elements.
        for data in series:
            self._write_ser(data)

        # Write the c:dropLines element.
        self._write_drop_lines()

        # Write the c:hiLowLines element.
        if args.get('primary_axes'):
            self._write_hi_low_lines()

        # Write the c:upDownBars element.
        self._write_up_down_bars()

        # Write the c:marker element.
        self._write_marker_value()

        # Write the c:axId elements
        self._write_axis_ids(args)

        self._xml_end_tag('c:stockChart')

    def _write_plot_area(self):
        # Overridden to use _write_date_axis() instead of _write_cat_axis().
        self._xml_start_tag('c:plotArea')

        # Write the c:layout element.
        self._write_layout(self.plotarea.get('layout'), 'plot')

        # Write the subclass chart elements for primary and secondary axes.
        self._write_chart_type({'primary_axes': 1})
        self._write_chart_type({'primary_axes': 0})

        # Write c:catAx and c:valAx elements for series using primary axes.
        self._write_date_axis({'x_axis': self.x_axis,
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

        self._write_date_axis({'x_axis': self.x2_axis,
                               'y_axis': self.y2_axis,
                               'axis_ids': self.axis2_ids
                               })

        # Write the c:dTable element.
        self._write_d_table()

        # Write the c:spPr element for the plotarea formatting.
        self._write_sp_pr(self.plotarea)

        self._xml_end_tag('c:plotArea')

    def _modify_series_formatting(self):
        # Add default formatting to the series data.

        index = 0

        for series in self.series:
            if index % 4 != 3:
                if not series['line']['defined']:
                    series['line'] = {'width': 2.25,
                                      'none': 1,
                                      'defined': 1}

                if series['marker'] is None:
                    if index % 4 == 2:
                        series['marker'] = {'type': 'dot', 'size': 3}
                    else:
                        series['marker'] = {'type': 'none'}

            index += 1
