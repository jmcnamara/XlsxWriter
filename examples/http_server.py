##############################################################################
#
# Example of using Python and XlsxWriter to create an Excel XLSX file in an in
# memory string suitable for serving via SimpleHTTPRequestHandler or Django or
# with the Google App Engine.
#
# SPDX-License-Identifier: BSD-2-Clause
# Copyright 2013-2023, John McNamara, jmcnamara@cpan.org
#

import http.server
import socketserver
import io

import xlsxwriter


class Handler(http.server.SimpleHTTPRequestHandler):
    def do_GET(self):
        # Create an in-memory output file for the new workbook.
        output = io.BytesIO()

        # Even though the final file will be in memory the module uses temp
        # files during assembly for efficiency. To avoid this on servers that
        # don't allow temp files set the 'in_memory' constructor option to True.
        #
        # Note: The Python 3 Runtime Environment in Google App Engine supports
        # a filesystem with read/write access to /tmp which means that the
        # 'in_memory' option isn't required there and can be omitted. See:
        #
        # https://cloud.google.com/appengine/docs/standard/python3/runtime#filesystem
        #
        workbook = xlsxwriter.Workbook(output, {"in_memory": True})
        worksheet = workbook.add_worksheet()

        # Write some test data.
        worksheet.write(0, 0, "Hello, world!")

        # Close the workbook before streaming the data.
        workbook.close()

        # Rewind the buffer.
        output.seek(0)

        # Construct a server response.
        self.send_response(200)
        self.send_header("Content-Disposition", "attachment; filename=test.xlsx")
        self.send_header(
            "Content-type",
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
        self.end_headers()
        self.wfile.write(output.read())
        return


print("Server listening on port 8000...")
httpd = socketserver.TCPServer(("", 8000), Handler)
httpd.serve_forever()
