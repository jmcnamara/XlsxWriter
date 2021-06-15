.. _ex_http_server:

Example: Simple HTTP Server
===========================

Example of using Python and XlsxWriter to create an Excel XLSX file in an in
memory string suitable for serving via SimpleHTTPRequestHandler or Django or with the
Google App Engine.

Even though the final file will be in memory, via the BytesIO object, the
XlsxWriter module uses temp files during assembly for efficiency. To avoid
this on servers that don't allow temp files set the ``in_memory`` constructor
option to ``True``.

The Python 3 Runtime Environment in Google App Engine supports a
`filesystem with read/write access to /tmp <https://cloud.google.com/appengine/docs/standard/python3/runtime#filesystem>`_
which means that the ``in_memory`` option isn't required.
required there.

.. literalinclude:: ../../../examples/http_server.py
