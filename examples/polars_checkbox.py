##############################################################################
#
# A example of displaying the boolean values in a Polars dataframe as checkboxes
# in an output xlsx file.
#
# SPDX-License-Identifier: BSD-2-Clause
#
# Copyright (c) 2013-2025, John McNamara, jmcnamara@cpan.org
#

import polars as pl

# Create a Pandas dataframe with some sample data.
df = pl.DataFrame(
    {
        "Region": ["North", "South", "East", "West"],
        "Target": [100, 70, 90, 120],
        "On-track": [False, True, True, False],
    }
)

# Write the dataframe to a new Excel file with formatting options.
df.write_excel(
    workbook="polars_checkbox.xlsx",
    # Set the checkbox format for the "On-track" boolean column.
    column_formats={"On-track": {"checkbox": True}},
    # Set an alternative table style.
    table_style="Table Style Light 9",
    # Autofit the column widths.
    autofit=True,
)
