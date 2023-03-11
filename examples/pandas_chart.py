##############################################################################
#
# An example of converting a Pandas dataframe to an xlsx file with a chart
# using Pandas and XlsxWriter.
#
# SPDX-License-Identifier: BSD-2-Clause
# Copyright 2013-2023, John McNamara, jmcnamara@cpan.org
#

import pandas as pd


# Create a Pandas dataframe from some data.
df = pd.DataFrame({"Data": [10, 20, 30, 20, 15, 30, 45]})

# Create a Pandas Excel writer using XlsxWriter as the engine.
writer = pd.ExcelWriter("pandas_chart.xlsx", engine="xlsxwriter")

# Convert the dataframe to an XlsxWriter Excel object.
df.to_excel(writer, sheet_name="Sheet1")

# Get the xlsxwriter workbook and worksheet objects.
workbook = writer.book
worksheet = writer.sheets["Sheet1"]

# Create a chart object.
chart = workbook.add_chart({"type": "column"})

# Get the dimensions of the dataframe.
(max_row, max_col) = df.shape

# Configure the series of the chart from the dataframe data.
chart.add_series({"values": ["Sheet1", 1, 1, max_row, 1]})

# Insert the chart into the worksheet.
worksheet.insert_chart(1, 3, chart)

# Close the Pandas Excel writer and output the Excel file.
writer.close()
