.. _ex_http_server:

Example: Simple HTTP Server
===========================

Example of using XlsxWriter to create an Excel XLSX file as an
in memory StringIO suitable for serving via SimpleHTTPServer or Django.

Note: This example doesn't currently work with the Google App Engine since it
needs to access a temporary directory, which isn't allowed on the GAE.

.. literalinclude:: ../../../examples/http_server.py

