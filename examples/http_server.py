##############################################################################
#
# Example of using Python and XlsxWriter to create an Excel XLSX file in an
# in memory string suitable for serving via SimpleHTTPServer or Django.
#
# Note, this doesn't currently work with the Google App Engine since it
# needs to access a temporary directory. See Github issue #28.
#
# Copyright 2013, John McNamara, jmcnamara@cpan.org
#

# Note: This is a Python 2 example. For Python 3 use http.server (or
#       equivalent) and io.StringIO.

import SimpleHTTPServer
import SocketServer
import StringIO

import xlsxwriter


class Handler(SimpleHTTPServer.SimpleHTTPRequestHandler):

    def do_GET(self):
        # Create a new workbook in memory and add a worksheet
        output = StringIO.StringIO()

        workbook = xlsxwriter.Workbook(output)
        worksheet = workbook.add_worksheet()

        # Write some test data.
        worksheet.write(0, 0, 'Hello, world!')

        # Close the workbook before streaming the data.
        workbook.close()

        # Rewind the buffer.
        output.seek(0)

        # Construct a server response.
        self.send_response(200)
        self.send_header('Content-Disposition', 'attachment; filename=test.xlsx')
        self.send_header('Content-type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
        self.end_headers()
        self.wfile.write(output.read())
        return


print("Server listening on port 8000...")
httpd = SocketServer.TCPServer(("", 8000), Handler)
httpd.serve_forever()
