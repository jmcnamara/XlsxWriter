##############################################################################
#
# An example of adding a Polars dataframe to a worksheet with a conditional
# format.
#
# SPDX-License-Identifier: BSD-2-Clause
# Copyright 2013-2023, John McNamara, jmcnamara@cpan.org
#

import polars as pl

df = pl.DataFrame({"Data": [10, 20, 30, 20, 15, 30, 45]})

df.write_excel(
    workbook="pandas_conditional.xlsx",
    conditional_formats={"Data": {"type": "3_color_scale"}},
)
