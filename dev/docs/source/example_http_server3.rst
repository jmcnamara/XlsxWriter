.. _ex_http_server3:

Example: Simple HTTP Server (Python 3)
======================================

Example of using Python and XlsxWriter to create an Excel XLSX file in an in
memory string suitable for serving via SimpleHTTPRequestHandler or Django or with the
Google App Engine.

Even though the final file will be in memory, via the BytesIO object, the
XlsxWriter module uses temp files during assembly for efficiency. To avoid
this on servers that don't allow temp files set the ``in_memory`` constructor
option to ``True``.

The Python 3 Runtime Environment in Google App Engine now supports a
`filesystem with read/write access to /tmp <https://cloud.google.com/appengine/docs/standard/python3/runtime#filesystem>`_
which means that the ``in_memory`` option isn't required. The ``/tmp`` dir
isn't supported in the Python 2 Runtime Environment so that option is still
required there.

For a Python 2 example see :ref:`ex_http_server`.

.. literalinclude:: ../../../examples/http_server_py3.py
