#######################################################################
#
# Example of how to create a workbook in memory with XlsxWriter and serve it via SimpleHTTPServer.
#
# Copyright 2013, John McNamara, jmcnamara@cpan.org
#

import SimpleHTTPServer
import SocketServer

import StringIO

from xlsxwriter.workbook import Workbook


class Handler(SimpleHTTPServer.SimpleHTTPRequestHandler):
    def do_GET(self):
        # Create a new workbook in memory and add a worksheet
        output = StringIO.StringIO()

        book = Workbook(output)
        sheet = book.add_worksheet('test')

        # Write some test data
        sheet.write(0, 0, 'Hello, world!')

        book.close()

        # set the buffer position to the beginning
        output.seek(0)

        # construct response
        self.send_response(200)
        self.send_header('Content-Disposition', 'attachment; filename=test.xlsx')
        self.send_header('Content-type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
        self.end_headers()
        self.wfile.write(output.read())
        return


# start server listening port 8000
httpd = SocketServer.TCPServer(("", 8000), Handler)
httpd.serve_forever()





