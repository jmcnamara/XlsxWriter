##############################################################################
#
# An example of adding a Polars dataframe to a worksheet created by XlsxWriter.
#
# SPDX-License-Identifier: BSD-2-Clause
# Copyright 2013-2023, John McNamara, jmcnamara@cpan.org
#

import xlsxwriter
import polars as pl

with xlsxwriter.Workbook("polars_xlsxwriter.xlsx") as workbook:
    # Create a new worksheet.
    worksheet = workbook.add_worksheet()

    # Do something with the worksheet.
    worksheet.write("A1", "The data below is added by Polars")

    df = pl.DataFrame({"Data": [10, 20, 30, 20, 15, 30, 45]})

    # Write the Polars data to the worksheet created above, at an offset to
    # avoid overwriting the previous text.
    df.write_excel(workbook=workbook, worksheet="Sheet1", position="A2")
