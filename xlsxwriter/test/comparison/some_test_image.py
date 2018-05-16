from warnings import warn
from struct import unpack
import os
import sys


class MyClass:

    def __init__(self):
        self.image_types = {}
        self.images = []

    def _get_image_properties(self, filename, image_data=False):
        # Extract dimension information from the image file.
        height = 0
        width = 0
        x_dpi = 96
        y_dpi = 96

        if not image_data:
            # Open the image file and read in the data.
            fh = open(filename, "rb")
            data = fh.read()
        else:
            # Read the image data from the user supplied byte stream.
            data = image_data.getvalue()

        # Get the image filename without the path.
        image_name = os.path.basename(filename)

        # Look for some common image file markers.
        marker1 = (unpack('3s', data[1:4]))[0]
        marker2 = (unpack('>H', data[:2]))[0]
        marker3 = (unpack('2s', data[:2]))[0]
        marker4 = (unpack('6s', data[:6]))[0]

        wmf_marker = b"\xd7\xcd\xc6\x9a\x00\x00"

        if sys.version_info < (2, 6, 0):
            # Python 2.5/Jython.
            png_marker = 'PNG'
            bmp_marker = 'BM'
        else:
            # Eval the binary literals for Python 2.5/Jython compatibility.
            png_marker = eval("b'PNG'")
            bmp_marker = eval("b'BM'")

        if marker1 == png_marker:
            self.image_types['png'] = 1
            (image_type, width, height, x_dpi, y_dpi) = self._process_png(data)

        elif marker2 == 0xFFD8:
            self.image_types['jpeg'] = 1
            (image_type, width, height, x_dpi, y_dpi) = self._process_jpg(data)

        elif marker3 == bmp_marker:
            self.image_types['bmp'] = 1
            (image_type, width, height) = self._process_bmp(data)

        elif marker4 == wmf_marker:
            # sanity check (standard metafile header)
            if (unpack('4s', data[22:26]))[0] != b"\x01\x00\t\x00":
                raise SyntaxError("Unsupported WMF file format")
            self.image_types['wmf'] = 1
            (image_type, width, height, x_dpi, y_dpi) = self._process_wmf(data)

        else:
            raise Exception("%s: Unknown or unsupported image file format."
                            % filename)

        # Check that we found the required data.
        if not height or not width:
            raise Exception("%s: no size data found in image file." % filename)

        # Store image data to copy it into file container.
        self.images.append([filename, image_type, image_data])

        if not image_data:
            fh.close()

        return image_type, width, height, image_name, x_dpi, y_dpi

    def _process_png(self, data):
        # Extract width and height information from a PNG file.
        offset = 8
        data_length = len(data)
        end_marker = False
        width = 0
        height = 0
        x_dpi = 96
        y_dpi = 96

        # Look for numbers rather than strings for Python 2.6/3 compatibility.
        marker_ihdr = 0x49484452  # IHDR
        marker_phys = 0x70485973  # pHYs
        marker_iend = 0X49454E44  # IEND

        # Search through the image data to read the height and width in the
        # IHDR element. Also read the DPI in the pHYs element.
        while not end_marker and offset < data_length:

            length = (unpack('>I', data[offset + 0:offset + 4]))[0]
            marker = (unpack('>I', data[offset + 4:offset + 8]))[0]

            # Read the image dimensions.
            if marker == marker_ihdr:
                width = (unpack('>I', data[offset + 8:offset + 12]))[0]
                height = (unpack('>I', data[offset + 12:offset + 16]))[0]

            # Read the image DPI.
            if marker == marker_phys:
                x_density = (unpack('>I', data[offset + 8:offset + 12]))[0]
                y_density = (unpack('>I', data[offset + 12:offset + 16]))[0]
                units = (unpack('b', data[offset + 16:offset + 17]))[0]

                if units == 1:
                    x_dpi = x_density * 0.0254
                    y_dpi = y_density * 0.0254

            if marker == marker_iend:
                end_marker = True
                continue

            offset = offset + length + 12

        return 'png', width, height, x_dpi, y_dpi

    def _process_jpg(self, data):
        # Extract width and height information from a JPEG file.
        offset = 2
        data_length = len(data)
        end_marker = False
        width = 0
        height = 0
        x_dpi = 96
        y_dpi = 96

        # Search through the image data to read the height and width in the
        # 0xFFC0/C2 element. Also read the DPI in the 0xFFE0 element.
        while not end_marker and offset < data_length:

            marker = (unpack('>H', data[offset + 0:offset + 2]))[0]
            length = (unpack('>H', data[offset + 2:offset + 4]))[0]

            # Read the image dimensions.
            if ((marker & 0xFFF0) == 0xFFC0
                and marker != 0xFFC4
                    and marker != 0xFFC8
                    and marker != 0xFFCC):
                height = (unpack('>H', data[offset + 5:offset + 7]))[0]
                width = (unpack('>H', data[offset + 7:offset + 9]))[0]
                print ("Marker: %04X, %d, %d" % (marker, height, width))

            # Read the image DPI.
            if marker == 0xFFE0:
                units = (unpack('b', data[offset + 11:offset + 12]))[0]
                x_density = (unpack('>H', data[offset + 12:offset + 14]))[0]
                y_density = (unpack('>H', data[offset + 14:offset + 16]))[0]

                if units == 1:
                    x_dpi = x_density
                    y_dpi = y_density

                if units == 2:
                    x_dpi = x_density * 2.54
                    y_dpi = y_density * 2.54

            if marker == 0xFFDA:
                end_marker = True
                continue

            offset = offset + length + 2

        return 'jpeg', width, height, x_dpi, y_dpi

    def _process_bmp(self, data):
        # Extract width and height information from a BMP file.
        width = (unpack('<L', data[18:22]))[0]
        height = (unpack('<L', data[22:26]))[0]
        return 'bmp', width, height

    def _process_wmf(self, data):
        # Extract width and height information from a WMF file.
        def short(c, o=0):
            v = word(c, o)
            if v >= 32768:
                v -= 65536
            return v

        def i16le(c, o=0):
            """
            Converts a 2-bytes (16 bits) string to an integer.

            c: string containing bytes to convert
            o: offset of bytes to convert in string
            """
            return unpack("<H", c[o:o+2])[0]

        word = i16le
        # get units per inch
        inch = i16le(data, 14)

        # get bounding box
        x0 = short(data, 6)
        y0 = short(data, 8)
        x1 = short(data, 10)
        y1 = short(data, 12)

        x_dpi = 96
        y_dpi = 96

        # normalize size to 72 dots per inch
        width = (x1 - x0) * x_dpi // inch
        height = (y1 - y0) * y_dpi // inch

        return 'wmf', width, height, x_dpi, y_dpi

if __name__ == '__main__':

    f_ = lambda x : os.path.join(os.path.dirname(__file__),'images', x )
    x = MyClass()
    print( "=================================")
    print (x._get_image_properties(f_('example_1.wmf')))
    print (x._get_image_properties(f_('example_2.wmf')))
    print (x._get_image_properties(f_('Python.wmf')))

    print ("=================================")