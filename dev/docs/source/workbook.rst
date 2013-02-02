The Workbook Class
==================

The Workbook class is the main class exposed by the XlsxWriter module and it
is the only class that you will need to instantiate directly.

The Workbook class represents the entire spreadsheet as you see it in Excel and
internally it represents the Excel file as it is written on disk.

Constructor
-----------

A new Excel workbook is created using the ``Workbook()`` constructor and a
filename::

    from xlsxwriter import Workbook

    workbook  = Workbook('filename.xlsx')
    worksheet = workbook.add_worksheet()

    worksheet.write(0, 0, 'Hello Excel')

.. note:
   A later version of the module will support writing to filehandles like
   Excel::Writer::XLSX. However, that functionality isn't available yet.

It is recommended that the filename uses the extension ``.xlsx`` rather than
``.xls`` since the latter causes an Excel warning when used with the XLSX
format.

On Windows remember to escape an directory separators::

    workbook1 = Excel::Writer::XLSX.new("c:\\tmp\\filename.xlsx")
    workbook2 = Excel::Writer::XLSX.new(r'c:\tmp\filename.xlsx')


add_worksheet([sheetname])
--------------------------

The ``add_worksheet()`` method adds a new worksheet to a workbook.

At least one worksheet should be added to a new workbook. The Worksheet object
is used to write data and configure the worksheet.

The ``sheetname`` parameter is optional. If it is not specified the default
Excel convention will be followed, i.e. Sheet1, Sheet2, etc.::

    worksheet1 = workbook.add_worksheet()           # Sheet1
    worksheet2 = workbook.add_worksheet('Foglio2')  # Foglio2
    worksheet3 = workbook.add_worksheet('Data')     # Data
    worksheet4 = workbook.add_worksheet()           # Sheet4

The worksheet name must be a valid Excel worksheet name, i.e. it cannot
contain any of the following characters, ``[ ] : * ? / \`` and it must be less
than 32 characters. In addition, you cannot use the same, case insensitive,
``sheetname`` for more than one worksheet.

add_format(properties)
----------------------

The ``add_format()`` method can be used to create new Format objects which are
used to apply formatting to a cell. You can either define the properties at
creation time via a dictionary of property values or later via method calls::

    format1 = workbook.add_format(props); # Set properties at creation.
    format2 = workbook.add_format();      # Set properties later.

See the "CELL FORMATTING" section for more details about Format properties
and how to set them.


close()
-------

In general your Excel file will be closed automatically when your program ends
or when the Workbook object goes out of scope, however the ``close()`` method
can be used to explicitly close an Excel file::

    workbook.close()

An explicit ``close()`` is required if the file must be closed prior to
performing some external action on it such as copying it, reading its size or
attaching it to an email.

In addition, ``close()`` may be occasionally by required to prevent Python's
garbage collector from disposing of the Workbook, Worksheet and Format objects
in the wrong order.

In general, if an XlsxWriter file is created with a size of 0 bytes or fails
to be created for some unknown, silent, reason you should add ``close()``
to your program.