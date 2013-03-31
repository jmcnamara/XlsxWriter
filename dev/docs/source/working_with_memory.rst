.. _memory_perf: 

Working with Memory and Performance
===================================

The Python XlsxWriter module is based on the design of the Perl module
:ref:`Excel::Writer::XLSX <ewx>` which in turn is based on an older Perl
module called
`Spreadsheet::WriteExcel <http://search.cpan.org/~jmcnamara/Spreadsheet-WriteExcel/>`_.

Spreadsheet::WriteExcel was written to optimise speed and reduce memory
usage. However, these design goals meant that it wasn't easy to implement
features that many users requested such as writing formatting and data
separately.

As a result XlsxWriter (and Excel::Writer::XLSX) takes a different
design approach and holds a lot more data in memory so that it is functionally
more flexible.

The effect of this is that XlsxWriter can consume a lot of memory. In addition
the extended row and column ranges in Excel 2007+ mean that it is possible to
run out of memory creating large files.

Fortunately, this memory usage can be reduced almost completely by using the
:func:`Workbook` ``'reduce_memory'`` property::

    workbook = Workbook(filename, {'reduce_memory': True})

For larger file this also gives an increase in performance, see below.

The trade-off is that you won't be able to take advantage of any new features
that manipulate cell data after it is written. One such feature is Tables.

.. Note::
   One of the optimisations used to reduce memory usage is that cell
   strings aren't stored in an Excel structure call "shared strings" and
   instead are written "in-line". This is a documented Excel feature that is
   supported by most spreadsheet applications. One known exception is Apple
   Numbers for Mac where the string data isn't displayed.



Performance Figures
-------------------

The performance figures below show execution time and memory usage for
worksheets of size ``N`` rows x 50 columns with a 50/50 mixture of strings and
numbers. The figures are taken from an arbitrary, mid-range, machine. Specific
figures will vary from machine to machine but the trends should be the same.

XlsxWriter in normal operation mode, the execution time and memory usage
increase more of less linearly with the number of rows:

+-------+---------+----------+----------------+
| Rows  | Columns | Time (s) | Memory (bytes) |
+=======+=========+==========+================+
| 200   | 50      | 0.72     | 2050552        |
+-------+---------+----------+----------------+
| 400   | 50      | 1.45     | 4478272        |
+-------+---------+----------+----------------+
| 800   | 50      | 2.90     | 8083072        |
+-------+---------+----------+----------------+
| 1600  | 50      | 5.92     | 17799424       |
+-------+---------+----------+----------------+
| 3200  | 50      | 11.83    | 32218624       |
+-------+---------+----------+----------------+
| 6400  | 50      | 23.72    | 64792576       |
+-------+---------+----------+----------------+
| 12800 | 50      | 47.85    | 128760832      |
+-------+---------+----------+----------------+

XlsxWriter in ``reduce_memory`` mode, the execution time still increases
linearly with the number of rows but the memory usage is small and constant:

+-------+---------+----------+----------------+
| Rows  | Columns | Time (s) | Memory (bytes) |
+=======+=========+==========+================+
| 200   | 50      | 0.40     | 54248          |
+-------+---------+----------+----------------+
| 400   | 50      | 0.80     | 54248          |
+-------+---------+----------+----------------+
| 800   | 50      | 1.60     | 54248          |
+-------+---------+----------+----------------+
| 1600  | 50      | 3.19     | 54248          |
+-------+---------+----------+----------------+
| 3200  | 50      | 6.29     | 54248          |
+-------+---------+----------+----------------+
| 6400  | 50      | 12.74    | 54248          |
+-------+---------+----------+----------------+
| 12800 | 50      | 25.34    | 54248          |
+-------+---------+----------+----------------+

In the ``reduce_memory`` mode the performance is also increased. There will be
further optimisation in both modes in later releases.

These figures were generated using programs in the ``dev/performance``
directory of the XlsxWriter source code.




