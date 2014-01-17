.. _workbook:

The Workbook Class
==================

The Workbook class is the main class exposed by the XlsxWriter module and it is
the only class that you will need to instantiate directly.

The Workbook class represents the entire spreadsheet as you see it in Excel and
internally it represents the Excel file as it is written on disk.

Constructor
-----------

.. py:function:: Workbook(filename [,options])

   Create a new XlsxWriter Workbook object.

   :param string filename: The name of the new Excel file to create.
   :param dict options:    Optional workbook parameters. See below.
   :rtype:                 A Workbook object.


The ``Workbook()`` constructor is used to create a new Excel workbook with a
given filename::

    import xlsxwriter

    workbook  = xlsxwriter.Workbook('filename.xlsx')
    worksheet = workbook.add_worksheet()

    worksheet.write(0, 0, 'Hello Excel')

.. image:: _images/workbook01.png

The constructor options are:

* **constant_memory**: Reduces the amount of data stored in memory so that
  large files can be written efficiently::

       workbook = xlsxwriter.Workbook(filename, {'constant_memory': True})

  Note, in this mode a row of data is written and then discarded when a cell
  in a new row is added via one of the worksheet ``write_()`` methods.
  Therefore, once this mode is active, data should be written in sequential
  row order.

  See :ref:`memory_perf` for more details.

* **tmpdir**: ``XlsxWriter`` stores worksheet data in a temporary files prior
  to assembling the final XLSX file. The temporary files are created in the
  system's temp directory. If the default temporary directory isn't accessible
  to your application, or doesn't contain enough space, you can specify an
  alternative location using the ``tempdir`` option::

       workbook = xlsxwriter.Workbook(filename, {'tmpdir': '/home/user/tmp'})

  The temporary directory must exist and will not be created.

* **in_memory**: To avoid the use of temporary files in the assembly of the
  final XLSX file, for example on servers that don't allow temp files such as
  the Google APP Engine, set the ``in_memory`` constructor option to ``True``::

       workbook = xlsxwriter.Workbook(filename, {'in_memory': True})

  This option overrides the ``constant_memory`` option.
* **strings_to_numbers**: Enable the
  :ref:`worksheet. <Worksheet>`:func:`write()` method to convert strings to
  numbers, where possible, using :func:`float()` in order to avoid an Excel
  warning about "Numbers Stored as Text". The default is ``False``::

      workbook = xlsxwriter.Workbook(filename, {'strings_to_numbers': True})

* **strings_to_formulas**: Enable the
  :ref:`worksheet. <Worksheet>`:func:`write()` method to convert strings to
  formulas. The default is ``True``::

      workbook = xlsxwriter.Workbook(filename, {'strings_to_formulas': False})

* **strings_to_urls**: Enable the
  :ref:`worksheet. <Worksheet>`:func:`write()` method to convert strings to
  urls. The default is ``True``::

      workbook = xlsxwriter.Workbook(filename, {'strings_to_urls': True})

* **default_date_format**: This option is used to specify a default date
  format string for use with the
  :ref:`worksheet. <Worksheet>`:func:`write_datetime()` method when an
  explicit format isn't given. See :ref:`working_with_dates_and_time` for more
  details::

      xlsxwriter.Workbook(filename, {'default_date_format': 'dd/mm/yy'})

* **date_1904**: Excel for Windows uses a default epoch of 1900 and Excel for
  Mac uses an epoch of 1904. However, Excel on either platform will convert
  automatically between one system and the other. XlsxWriter stores dates in
  the 1900 format by default. If you wish to change this you can use the
  ``date_1904`` workbook option. This option is mainly for enhanced
  compatibility with Excel and in general isn't required very often::

      workbook = xlsxwriter.Workbook(filename, {'date_1904': True})

When specifying a filename it is recommended that you use an ``.xlsx``
extension or Excel will generate a warning when opening the file.

It is possible to write files to in-memory strings using StringIO as follows::

    output = StringIO()
    workbook = xlsxwriter.Workbook(output)
    worksheet = workbook.add_worksheet()

    worksheet.write('A1', 'Hello')
    workbook.close()

    xlsx_data = output.getvalue()

To avoid the use of any temporary files and keep the entire file in-memory use
the ``in_memory`` constructor option shown above.

See also :ref:`ex_http_server`.


workbook.add_worksheet()
------------------------

.. function:: add_worksheet([sheetname])

   Add a new worksheet to a workbook.

   :param string sheetname: Optional worksheet name, defaults to Sheet1, etc.
   :rtype: A :ref:`worksheet <Worksheet>` object.

The ``add_worksheet()`` method adds a new worksheet to a workbook.

At least one worksheet should be added to a new workbook. The
:ref:`Worksheet <worksheet>` object is used to write data and configure a
worksheet in the workbook.

The ``sheetname`` parameter is optional. If it is not specified the default
Excel convention will be followed, i.e. Sheet1, Sheet2, etc.::

    worksheet1 = workbook.add_worksheet()           # Sheet1
    worksheet2 = workbook.add_worksheet('Foglio2')  # Foglio2
    worksheet3 = workbook.add_worksheet('Data')     # Data
    worksheet4 = workbook.add_worksheet()           # Sheet4

.. image:: _images/workbook02.png

The worksheet name must be a valid Excel worksheet name, i.e. it cannot contain
any of the characters ``' [ ] : * ? / \
'`` and it must be less than 32 characters.

In addition, you cannot use the same, case insensitive, ``sheetname`` for more
than one worksheet.

workbook.add_format()
---------------------

.. py:function:: add_format([properties])

   Create a new Format object to formats cells in worksheets.

   :param dictionary properties: An optional dictionary of format properties.
   :rtype: A :ref:`format <Format>` object.

The ``add_format()`` method can be used to create new :ref:`Format <Format>`
objects which are used to apply formatting to a cell. You can either define
the properties at creation time via a dictionary of property values or later
via method calls::

    format1 = workbook.add_format(props); # Set properties at creation.
    format2 = workbook.add_format();      # Set properties later.

See the :ref:`format` and :ref:`working_with_formats` sections for more details
about Format properties and how to set them.


workbook.add_chart()
--------------------

.. py:function:: add_chart(options)

   Create a chart object that can be added to a worksheet.

   :param dictionary options: An dictionary of chart type options.
   :rtype: A :ref:`Chart <chart_class>` object.

This method is use to create a new chart object that can be inserted into a
worksheet via the :func:`insert_chart()` Worksheet method::

    chart = workbook.add_chart({'type': 'column'})

The properties that can be set are::

    type    (required)
    subtype (optional)

* ``type``

  This is a required parameter. It defines the type of chart that will be
  created::

    chart = workbook.add_chart({'type': 'line'})

  The available types are::

    area
    bar
    column
    line
    pie
    radar
    scatter
    stock

* ``subtype``

  Used to define a chart subtype where available::

    workbook.add_chart({'type': 'bar', 'subtype': 'stacked'})

See the :ref:`chart_class` for a list of available chart subtypes.

See also :ref:`working_with_charts` and :ref:`chart_examples`.

workbook.add_chartsheet()
-------------------------

.. function:: add_chartsheet([sheetname])

   Add a new add_chartsheet to a workbook.

   :param string sheetname: Optional chartsheet name, defaults to Chart1, etc.
   :rtype: A :ref:`chartsheet <Chartsheet>` object.

The ``add_chartsheet()`` method adds a new chartsheet to a workbook.

.. image:: _images/chartsheet.png

See :ref:`chartsheet` for details.

The ``sheetname`` parameter is optional. If it is not specified the default
Excel convention will be followed, i.e. Chart1, Chart2, etc.

The chartsheet name must be a valid Excel worksheet name, i.e. it cannot contain
any of the characters ``' [ ] : * ? / \
'`` and it must be less than 32 characters.

In addition, you cannot use the same, case insensitive, ``sheetname`` for more
than one chartsheet.


workbook.close()
----------------

.. py:function:: close()

   Close the Workbook object and write the XLSX file.

In general your Excel file will be closed automatically when your program ends
or when the Workbook object goes out of scope, however the ``close()`` method
can be used to explicitly close an Excel file::

    workbook.close()

An explicit ``close()`` is required if the file must be closed prior to
performing some external action on it such as copying it, reading its size or
attaching it to an email.

In addition, ``close()`` may occasionally be required to prevent Python's
garbage collector from disposing of the Workbook, Worksheet and Format objects
in the wrong order.


workbook.set_properties()
-------------------------

.. py:function:: set_properties()

   Set the document properties such as Title, Author etc.

   :param dict properties: Dictionary of document properties.

The ``set_properties`` method can be used to set the document properties of the
Excel file created by ``XlsxWriter``. These properties are visible when you
use the ``Office Button -> Prepare -> Properties`` option in Excel and are
also available to external applications that read or index windows files.

The properties that can be set are:

* ``title``
* ``subject``
* ``author``
* ``manager``
* ``company``
* ``category``
* ``keywords``
* ``comments``
* ``status``

The properties are all optional and should be passed in dictionary format as
follows::

    workbook.set_properties({
        'title':    'This is an example spreadsheet',
        'subject':  'With document properties',
        'author':   'John McNamara',
        'manager':  'Dr. Heinz Doofenshmirtz',
        'company':  'of Wolves',
        'category': 'Example spreadsheets',
        'keywords': 'Sample, Example, Properties',
        'comments': 'Created with Python and XlsxWriter'})

.. image:: _images/doc_properties.png

See also :ref:`ex_doc_properties`.

workbook.define_name()
----------------------

.. py:function:: define_name()

   Create a defined name in the workbook to use as a variable.

   :param string name:    The defined name.
   :param string formula: The cell or range that the defined name refers to.

This method is used to defined a name that can be used to represent a value, a
single cell or a range of cells in a workbook. These defined names can then be
used in formulas::

    workbook.define_name('Exchange_rate', '=0.96')
    worksheet.write('B3', '=B2*Exchange_rate')

As in Excel a name defined like this is "global" to the workbook and can be
referred to from any worksheet::

    # Global workbook name.
    workbook.define_name('Sales',         '=Sheet1!$G$1:$H$10')

It is also possible to define a local/worksheet name by prefixing it with the
sheet name using the syntax ``'sheetname!definedname'``::

    # Local worksheet name.
    workbook.define_name('Sheet2!Sales', '=Sheet2!$G$1:$G$10')

If the sheet name contains spaces or special characters you must follow the
Excel convention and enclose it in single quotes::

    workbook.define_name("'New Data'!Sales", '=Sheet2!$G$1:$G$10')

See also the ``defined_name.py`` program in the examples directory.


workbook.worksheets()
---------------------

.. py:function:: worksheets()

   Return a list of the worksheet objects in the workbook.

   :rtype: A list of :ref:`worksheet <Worksheet>` objects.

The ``worksheets()`` method returns a list of the worksheets in a workbook.
This is useful if you want to repeat an operation on each worksheet in a
workbook::

    for worksheet in workbook.worksheets():
        worksheet.write('A1', 'Hello')

workbook.set_calc_mode()
------------------------

.. py:function:: set_calc_mode(mode)

   Set the Excel calculation mode for the workbook.

   :param string mode: The calculation mode string

Set the calculation mode for formulas in the workbook. This is mainly of use
for workbooks with slow formulas where you want to allow the user to
calculate them manually.

The ``mode`` parameter can be:

* ``auto``: The default. Excel will re-calculate formulas when a formula or 
  a value affecting the formula changes.

* ``auto_except_tables``: Excel will automatically re-calculate formulas
  except for tables.

* ``manual``: Only re-calculate formulas when the user requires it. Generally
  by pressing F9.
