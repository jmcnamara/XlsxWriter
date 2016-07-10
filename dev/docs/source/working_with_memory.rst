.. _memory_perf:

Working with Memory and Performance
===================================

By default XlsxWriter holds all cell data in memory. This is to allow future
features when formatting is applied separately from the data.

The effect of this is that XlsxWriter can consume a lot of memory and it is
possible to run out of memory when creating large files.

Fortunately, this memory usage can be reduced almost completely by using the
:func:`Workbook` ``'constant_memory'`` property::

    workbook = xlsxwriter.Workbook(filename, {'constant_memory': True})

The optimization works by flushing each row after a subsequent row is written.
In this way the largest amount of data held in memory for a worksheet is the
amount of data required to hold a single row of data.

Since each new row flushes the previous row, data must be written in sequential
row order when ``'constant_memory'`` mode is on::

    # With 'constant_memory' you must write data in row by column order.
    for row in range(0, row_max):
        for col in range(0, col_max):
            worksheet.write(row, col, some_data)

    # With 'constant_memory' this would only write the first column of data.
    for col in range(0, col_max):
        for row in range(0, row_max):
            worksheet.write(row, col, some_data)

Another optimization that is used to reduce memory usage is that cell strings
aren't stored in an Excel structure call "shared strings" and instead are
written "in-line". This is a documented Excel feature that is supported by
most spreadsheet applications. One known exception is Apple Numbers for Mac
where the string data isn't displayed.

The trade-off when using ``'constant_memory'`` mode is that you won't be able
to take advantage of any new features that manipulate cell data after it is
written. Currently the :func:`add_table()` method doesn't work in this mode
and :func:`merge_range()` and :func:`set_row()` only work for the current row.


For larger files ``'constant_memory'`` mode also gives an increase in execution
speed, see below.


Performance Figures
-------------------

The performance figures below show execution time and memory usage for
worksheets of size ``N`` rows x 50 columns with a 50/50 mixture of strings and
numbers. The figures are taken from an arbitrary, mid-range, machine. Specific
figures will vary from machine to machine but the trends should be the same.

XlsxWriter in normal operation mode: the execution time and memory usage
increase more of less linearly with the number of rows:

+-------+---------+----------+----------------+
| Rows  | Columns | Time (s) | Memory (bytes) |
+=======+=========+==========+================+
| 200   | 50      | 0.43     | 2346728        |
+-------+---------+----------+----------------+
| 400   | 50      | 0.84     | 4670904        |
+-------+---------+----------+----------------+
| 800   | 50      | 1.68     | 8325928        |
+-------+---------+----------+----------------+
| 1600  | 50      | 3.39     | 17855192       |
+-------+---------+----------+----------------+
| 3200  | 50      | 6.82     | 32279672       |
+-------+---------+----------+----------------+
| 6400  | 50      | 13.66    | 64862232       |
+-------+---------+----------+----------------+
| 12800 | 50      | 27.60    | 128851880      |
+-------+---------+----------+----------------+

XlsxWriter in ``constant_memory`` mode: the execution time still increases
linearly with the number of rows but the memory usage remains small and
constant:

+-------+---------+----------+----------------+
| Rows  | Columns | Time (s) | Memory (bytes) |
+=======+=========+==========+================+
| 200   | 50      | 0.37     | 62208          |
+-------+---------+----------+----------------+
| 400   | 50      | 0.74     | 62208          |
+-------+---------+----------+----------------+
| 800   | 50      | 1.46     | 62208          |
+-------+---------+----------+----------------+
| 1600  | 50      | 2.93     | 62208          |
+-------+---------+----------+----------------+
| 3200  | 50      | 5.90     | 62208          |
+-------+---------+----------+----------------+
| 6400  | 50      | 11.84    | 62208          |
+-------+---------+----------+----------------+
| 12800 | 50      | 23.63    | 62208          |
+-------+---------+----------+----------------+

In the ``constant_memory`` mode the performance is also increased slightly.

These figures were generated using programs in the ``dev/performance``
directory of the XlsxWriter repo.


Benchmark of Python Excel Writers
---------------------------------

If you wish to compare the performance of different Python Excel writing
modules there is a program called `bench_excel_writers.py
<https://raw.githubusercontent.com/jmcnamara/XlsxWriter/master/dev/performance/bench_excel_writers.py>`_
in the ``dev/performance`` directory of the XlsxWriter repo.

And here is the output for 10,000 rows x 50 columns using the latest version
of the modules at the time of writing::

    Versions:
        python      : 2.7.2
        openpyxl    : 2.2.1
        pyexcelerate: 0.6.6
        xlsxwriter  : 0.7.2
        xlwt        : 1.0.0

    Dimensions:
        Rows = 10000
        Cols = 50

    Times:
        pyexcelerate          :  10.63
        xlwt                  :  16.93
        xlsxwriter (optimized):  20.37
        xlsxwriter            :  24.24
        openpyxl   (optimized):  26.63
        openpyxl              :  35.75


As with any benchmark the results will depend on Python/module versions, CPU,
RAM and Disk I/O and on the benchmark itself. So make sure to verify these
results for your own setup.
