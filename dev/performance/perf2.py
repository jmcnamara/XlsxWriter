##############################################################################
#
# Benchmarking script for XlsxWriter
#

import os
import string
import sys
import random
from time import clock

from docopt import docopt
from pympler.asizeof import asizeof
import xlsxwriter

from xlsxwriter.utility import xl_rowcol_to_cell_fast

random.seed(42)
STR_LEN = 1
MAX_INT = 32676
MAX_FORMATS = 20
MAX_FORMAT_PROPS = 3
FORMAT_PROPERTIES = (('align', 'left'),
                     ('align', 'center'),
                     ('align', 'right'),
                     ('align', 'bottom'),
                     ('align', 'top'),
                     ('bold', True),
                     ('bold', False),
                     ('text_wrap', True),
                     ('text_wrap', False),
                     ('font_color', 'red'),
                     ('font_color', 'green'),
                     ('font_color', 'white'),
                     ('font_color', 'gray'),
                     ('font_color', 'blue'),
                     ('font_color', 'yellow'),
                     ('font_color', 'orange'),
                     ('font_color', 'black'),
                     ('font_color', 'purple'))


def benchmark_xlsx(rows, cols, optimise, memory_check):

    # Set up testing data (do not benchmark).
    # todo: Cache for later use (better benchmark perf).

    chars = (string.ascii_uppercase + string.digits + ' '
             + string.ascii_lowercase)

    strs = [''.join(random.choice(chars)
                    for _ in range(STR_LEN))
            for _ in range(rows * cols)]

    ints = [bool(random.randint(-MAX_INT, MAX_INT))
            for _ in xrange(rows*cols)]

    floats = [random.randrange(-float(MAX_INT), float(MAX_INT))
              for _ in xrange(rows*cols)]

    bools = [bool(random.randint(-1, 1))
             for x in xrange(rows*cols)]

    data_types = [strs, ints, floats, bools]

    # todo: Add more data types?

    len_data_types = len(data_types)
    locations = []
    for i in range(len_data_types):
        for x in range(rows):
            if len(locations) < x + 1:
                locations.append([])
            for y in range(cols):
                y_index = cols * i + y
                locations[x].append(xl_rowcol_to_cell_fast(x, y_index))

    # todo: Test urls.

    start_time = clock()

    # Start of program being tested.
    workbook = xlsxwriter.Workbook('xlsxw_perf_%s_%s.xlsx' % (rows, cols),
                                   {'constant_memory': optimise})
    worksheet = workbook.add_worksheet()

    formats = []
    for _ in xrange(MAX_FORMATS):
        properties = {}
        for i in xrange(MAX_FORMAT_PROPS):
            prop = random.choice(FORMAT_PROPERTIES)
            properties[prop[0]] = prop[1]
        formats.append(workbook.add_format(properties))

    # Create the actual spreadsheet.
    for i, data_type in enumerate(data_types):
        for row in range(rows):
            for col in range(cols):
                y_index = col + len_data_types * i
                # todo: Test comments.
                worksheet.write(locations[x][y_index],
                                data_type[row * cols + col],
                                random.choice(formats))

    # Get total memory size for workbook object before closing it.
    if memory_check:
        total_size = asizeof(workbook)
    else:
        total_size = 0

    workbook.close()

    # Get the elapsed time.
    elapsed = clock() - start_time

    # Print a simple CSV output for reporting.
    print("%10s %10s %10s %10s" % (rows, cols, elapsed, total_size))

    return elapsed, total_size


class fib:
    """
    Generator for the fibonacci sequence with offset start
    """
    def __init__(self, start, max):
        self.start = start
        self.max = max

    def __iter__(self):
        self.a = 0
        self.b = 1
        return self

    def next(self):
        while self.a < self.start:
            self.a, self.b = self.b, self.a + self.b
        fib = self.a
        if fib > self.max:
            raise StopIteration
        self.a, self.b = self.b, self.a + self.b
        return fib


def main():
    helpstr = """Open Shell to debug the current crunch environment

Usage:
  %(script)s (-h | --help)
  %(script)s [options]

Options:
  -h --help                     Show this screen
  -o --optimise                 optimise
  -m --memory-check             report on memory usage
  -r [rows]                     max rows to run
  -c [cols]                     max cols to run

    """

    arguments = docopt(helpstr % dict(script=os.path.basename(sys.argv[0])))

    ROW_MIN = 100
    COL_MIN = 100
    ROW_MAX = 400
    COL_MAX = 400

    optimise = arguments['--optimise']
    memory_check = arguments['--memory-check']
    rows = int(arguments.get('-r') or ROW_MAX)
    cols = int(arguments.get('-c') or COL_MAX)

    print("%10s %10s %10s %10s" % ('rows', 'cols', 'elapsed', 'size'))

    for r in fib(ROW_MIN, rows):
        for c in fib(COL_MIN, cols):
            benchmark_xlsx(r, c, optimise, memory_check)

if __name__ == '__main__':
    main()
