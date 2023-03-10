##############################################################################
#
# A simple example of converting a Polars dataframe to an xlsx file using
# Polars and XlsxWriter.
#
# SPDX-License-Identifier: BSD-2-Clause
# Copyright 2013-2023, John McNamara, jmcnamara@cpan.org
#

import polars as pl

# Create a Pandas dataframe from some data.
df = pl.DataFrame({"Data": [10, 20, 30, 20, 15, 30, 45]})

# Write the dataframe to a new Excel file.
df.write_excel(workbook="polars_simple.xlsx")
