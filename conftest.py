# Config script for py.test.
import sys


collect_ignore = ['setup.py']

# Tests to ignore/skip in Python 2.5/Jython.
py25_ignore = [

    # 'f' not supported in'%Y-%m-%dT%H:%M:%S.%f'
    'xlsxwriter/test/worksheet/test_date_time_01.py',
    'xlsxwriter/test/worksheet/test_date_time_02.py',
    'xlsxwriter/test/worksheet/test_date_time_03.py',

    # No fractions or decimal.
    'xlsxwriter/test/comparison/test_types03.py',

    # No unicode_literals.
    'xlsxwriter/test/comparison/test_chart_axis25.py',
    'xlsxwriter/test/comparison/test_utf8_01.py',
    'xlsxwriter/test/comparison/test_utf8_03.py',
    'xlsxwriter/test/comparison/test_utf8_04.py',
    'xlsxwriter/test/comparison/test_utf8_05.py',
    'xlsxwriter/test/comparison/test_utf8_06.py',
    'xlsxwriter/test/comparison/test_utf8_07.py',
    'xlsxwriter/test/comparison/test_utf8_08.py',
    'xlsxwriter/test/comparison/test_utf8_09.py',
    'xlsxwriter/test/comparison/test_utf8_10.py',
    'xlsxwriter/test/comparison/test_utf8_11.py',
    'xlsxwriter/test/comparison/test_properties05.py',
    'xlsxwriter/test/comparison/test_defined_name04.py',
    'xlsxwriter/test/comparison/test_data_validation07.py',
    ]

if sys.version_info < (2, 6, 0):
    collect_ignore.extend(py25_ignore)
