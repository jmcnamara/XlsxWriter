##############################################################################
#
# An example of converting some string percentage data in a Pandas dataframe
# to percentage numbers in an xlsx file with using XlsxWriter.
#
# SPDX-License-Identifier: BSD-2-Clause
# Copyright 2013-2023, John McNamara, jmcnamara@cpan.org
#

import pandas as pd


# Create a Pandas dataframe from some data.
df = pd.DataFrame({"Names": ["Anna", "Arek", "Arun"], "Grade": ["100%", "70%", "85%"]})

# Convert the percentage strings to percentage numbers.
df["Grade"] = df["Grade"].str.replace("%", "")
df["Grade"] = df["Grade"].astype(float)
df["Grade"] = df["Grade"].div(100)

# Create a Pandas Excel writer using XlsxWriter as the engine.
writer = pd.ExcelWriter("pandas_percent.xlsx", engine="xlsxwriter")

# Convert the dataframe to an XlsxWriter Excel object.
df.to_excel(writer, sheet_name="Sheet1")

# Get the xlsxwriter workbook and worksheet objects.
workbook = writer.book
worksheet = writer.sheets["Sheet1"]

# Add a percent number format.
percent_format = workbook.add_format({"num_format": "0%"})

# Apply the number format to Grade column.
worksheet.set_column(2, 2, None, percent_format)

# Close the Pandas Excel writer and output the Excel file.
writer.close()
