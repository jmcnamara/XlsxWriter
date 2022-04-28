.. SPDX-License-Identifier: BSD-2-Clause
   Copyright 2013-2022, John McNamara, jmcnamara@cpan.org

.. _ex_pandas_column_formats:

Example: Pandas Excel output with column formatting
===================================================

An example of converting a Pandas dataframe to an Excel file with column
formats using Pandas and XlsxWriter.

It isn't possible to format any cells that already have a format such as
the index or headers or any cells that contain dates or datetimes.

Note: This feature requires Pandas >= 0.16.

.. image:: _images/pandas_column_formats.png

.. literalinclude:: ../../../examples/pandas_column_formats.py
