##############################################################################
#
# A simple example of converting a Polars dataframe to an xlsx file with
# custom formatting of the worksheet table.
#
# SPDX-License-Identifier: BSD-2-Clause
# Copyright 2013-2023, John McNamara, jmcnamara@cpan.org
#

from datetime import date
import polars as pl

# Create a Pandas dataframe with some sample data.
df = pl.DataFrame(
    {
        "Dates": [date(2023, 1, 1), date(2023, 1, 2), date(2023, 1, 3)],
        "Strings": ["Alice", "Bob", "Carol"],
        "Numbers": [0.12345, 100, -99.523],
    }
)

# Write the dataframe to a new Excel file with formatting options.
df.write_excel(
    workbook="polars_format_custom.xlsx",
    # Set an alternative table style.
    table_style="Table Style Medium 4",
    # See the floating point precision for reals.
    float_precision=6,
    # Set an alternative number/date format for Polar Date types.
    dtype_formats={pl.Date: "yyyy mm dd;@"},
    # Add totals to the numeric columns.
    column_totals=True,
    # Autofit the column widths.
    autofit=True,
)
