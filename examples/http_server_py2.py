##############################################################################
#
# Example of using Python and XlsxWriter to create an Excel XLSX file in an
# in memory string suitable for serving via SimpleHTTPServer or Django
# or with the Google App Engine.
#
# Copyright 2013-2015, John McNamara, jmcnamara@cpan.org
#

# Note: This is a Python 2 example. For Python 3 see http_server_py3.py.

import SimpleHTTPServer
import SocketServer
import StringIO

import xlsxwriter


class Handler(SimpleHTTPServer.SimpleHTTPRequestHandler):

    def do_GET(self):
        # Create an in-memory output file for the new workbook.
        output = StringIO.StringIO()

        # Even though the final file will be in memory the module uses temp
        # files during assembly for efficiency. To avoid this on servers that
        # don't allow temp files, for example the Google APP Engine, set the
        # 'in_memory' constructor option to True:
        workbook = xlsxwriter.Workbook(output, {'in_memory': True})
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


print('Server listening on port 8000...')
httpd = SocketServer.TCPServer(('', 8000), Handler)
httpd.serve_forever()
