##############################################################################
#
# An example of converting a Pandas dataframe to an xlsx file
# with column formats using Pandas and XlsxWriter.
#
# SPDX-License-Identifier: BSD-2-Clause
# Copyright 2013-2023, John McNamara, jmcnamara@cpan.org
#

import pandas as pd

# Create a Pandas dataframe from some data.
df = pd.DataFrame(
    {
        "Numbers": [1010, 2020, 3030, 2020, 1515, 3030, 4545],
        "Percentage": [0.1, 0.2, 0.33, 0.25, 0.5, 0.75, 0.45],
    }
)

# Create a Pandas Excel writer using XlsxWriter as the engine.
writer = pd.ExcelWriter("pandas_column_formats.xlsx", engine="xlsxwriter")

# Convert the dataframe to an XlsxWriter Excel object.
df.to_excel(writer, sheet_name="Sheet1")

# Get the xlsxwriter workbook and worksheet objects.
workbook = writer.book
worksheet = writer.sheets["Sheet1"]

# Add some cell formats.
format1 = workbook.add_format({"num_format": "#,##0.00"})
format2 = workbook.add_format({"num_format": "0%"})

# Note: It isn't possible to format any cells that already have a format such
# as the index or headers or any cells that contain dates or datetimes.

# Set the column width and format.
worksheet.set_column(1, 1, 18, format1)

# Set the format but not the column width.
worksheet.set_column(2, 2, None, format2)

# Close the Pandas Excel writer and output the Excel file.
writer.close()
