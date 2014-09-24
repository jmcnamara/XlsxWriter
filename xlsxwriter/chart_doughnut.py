##############################################################################
#
# ChartDoughnut - A class for writing the Excel XLSX Donut charts.
#
# Copyright 2013-2014, John McNamara, jmcnamara@cpan.org
#

from . import chart_pie


class ChartDoughnut(chart_pie.ChartPie):
    """
    A class for writing the Excel XLSX Doughnut charts.
    """
    ###########################################################################
    #
    # XML methods.
    #
    ###########################################################################

    def _write_pie_chart(self, args):
        # Write the <c:pieChart> element.  Over-ridden method to remove
        # axis_id code since Pie charts don't require val and cat axes.
        self._xml_start_tag('c:doughnutChart')

        # Write the c:varyColors element.
        self._write_vary_colors()

        # Write the series elements.
        for data in self.series:
            self._write_ser(data)

        # Write the c:firstSliceAng element.
        self._write_first_slice_ang()

        self._xml_end_tag('c:doughnutChart')
