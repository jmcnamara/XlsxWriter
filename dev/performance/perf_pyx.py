##############################################################################
#
# Simple Python program to test the speed and memory usage of
# the XlsxWriter module.
#
# python perf_pyx.py [num_rows] [optimization_mode]
#
# Copyright 2013-2016, John McNamara, jmcnamara@cpan.org

import sys
import xlsxwriter
from time import clock
from pympler.asizeof import asizeof

# Default to 1000 rows and non-optimised.
if len(sys.argv) > 1:
    row_max = int(sys.argv[1]) // 2
else:
    row_max = 1000

if len(sys.argv) > 2 and int(sys.argv[2]) == 1:
    optimise = 1
else:
    optimise = 0

if len(sys.argv) > 3 and int(sys.argv[3]) == 1:
    get_memory_size = 1
else:
    get_memory_size = 0


col_max = 50

# Start timing after everything is loaded.
start_time = clock()

# Start of program being tested.
workbook = xlsxwriter.Workbook('py_ewx.xlsx',
                               {'constant_memory': optimise})
worksheet = workbook.add_worksheet()

worksheet.set_column(0, col_max, 18)

for row in range(0, row_max):
    for col in range(0, col_max):
        worksheet.write_string(row * 2, col, "Row: %d Col: %d" % (row, col))
    for col in range(0, col_max + 1):
        worksheet.write_number(row * 2 + 1, col, row + col)

# Get total memory size for workbook object before closing it.
if get_memory_size:
    total_size = asizeof(workbook)
else:
    total_size = 0

workbook.close()

# Get the elapsed time.
elapsed = clock() - start_time

# Print a simple CSV output for reporting.

print("%6d, %3d, %6.2f, %d" % (row_max * 2, col_max, elapsed, total_size))
