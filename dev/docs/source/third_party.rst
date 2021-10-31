.. SPDX-License-Identifier: BSD-2-Clause
   Copyright 2013-2021, John McNamara, jmcnamara@cpan.org

.. _third_party:

Libraries that use or enhance XlsxWriter
========================================

The following are some libraries or applications that wrap or extend
XlsxWriter.


Pandas
------

Python `Pandas <https://pandas.pydata.org/>`_ is a Python data analysis
library. It can read, filter and re-arrange small and large data sets and
output them in a range of formats including Excel.

XlsxWriter is available as an Excel output engine in Pandas. See also See
:ref:`ewx_pandas`.


XlsxPandasFormatter
-------------------

`XlsxPandasFormatter
<https://github.com/webermarcolivier/xlsxpandasformatter>`_ is a helper class
that wraps the worksheet, workbook and dataframe objects written by Pandas
``to_excel()`` method using the ``xlsxwriter`` engine to allow consistent
formatting of cells.
