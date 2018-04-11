###############################################################################
#
# Tests for XlsxWriter.
#
# Copyright (c), 2013-2018, John McNamara, jmcnamara@cpan.org
#

from ..excel_comparsion_test import ExcelComparisonTest
from ...workbook import Workbook


class TestCompareXLSXFiles(ExcelComparisonTest):
    """
    Test file created by XlsxWriter against a file created by Excel.

    """

    def setUp(self):
        self.maxDiff = None

        filename = 'autofilter09.xlsx'

        test_dir = 'xlsxwriter/test/comparison/'
        self.got_filename = test_dir + '_test_' + filename
        self.exp_filename = test_dir + 'xlsx_files/' + filename
        self.txt_filename = test_dir + 'xlsx_files/' + 'autofilter_data.txt'

        self.ignore_files = []
        self.ignore_elements = {}

    def test_create_file(self):
        """
        Test the creation of a simple XlsxWriter file with an autofilter.
        This test checks a filter list.
        """

        workbook = Workbook(self.got_filename)
        worksheet = workbook.add_worksheet()

        # Set the autofilter.
        worksheet.autofilter('A1:D51')

        # Add filter criteria.
        worksheet.filter_column_list(0, ['East', 'South', 'North'])

        # Open a text file with autofilter example data.
        textfile = open(self.txt_filename)

        # Read the headers from the first line of the input file.
        headers = textfile.readline().strip("\n").split()

        # Write out the headers.
        worksheet.write_row('A1', headers)

        # Start writing data after the headers.
        row = 1

        # Read the rest of the text file and write it to the worksheet.
        for line in textfile:

            # Split the input data based on whitespace.
            data = line.strip("\n").split()

            # Convert the number data from the text file.
            for i, item in enumerate(data):
                try:
                    data[i] = float(item)
                except ValueError:
                    pass

            # Simulate a blank cell in the data.
            if row == 6:
                data[0] = ''

            # Get some of the field data.
            region = data[0]

            # Check for rows that match the filter.
            if region == 'North' or region == 'South' or region == 'East':
                # Row matches the filter, no further action required.
                pass
            else:
                # We need to hide rows that don't match the filter.
                worksheet.set_row(row, options={'hidden': True})

            # Write out the row data.
            worksheet.write_row(row, 0, data)

            # Move on to the next worksheet row.
            row += 1

        textfile.close()
        workbook.close()

        self.assertExcelEqual()
