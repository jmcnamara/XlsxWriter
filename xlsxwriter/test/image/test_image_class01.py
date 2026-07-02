###############################################################################
#
# Tests for XlsxWriter.
#
# SPDX-License-Identifier: BSD-2-Clause
#
# Copyright (c), 2013-2025, John McNamara, jmcnamara@cpan.org
#
import unittest
from io import BytesIO
from struct import pack

from xlsxwriter.exceptions import UndefinedImageSize, UnsupportedImageFormat
from xlsxwriter.image import Image


class TestImageProperties(unittest.TestCase):
    """
    Test the properties of an Image object.
    """

    def test_image_properties01(self):
        """Test the Image class properties."""
        image = Image("xlsxwriter/test/comparison/images/red.png")

        self.assertEqual(image.image_type, "PNG")
        self.assertEqual(image.width, 32)
        self.assertEqual(image.height, 32)
        self.assertEqual(image.x_dpi, 96)
        self.assertEqual(image.y_dpi, 96)

    def test_image_properties02(self):
        """Test the Image class properties."""
        with open("xlsxwriter/test/comparison/images/red.png", "rb") as image_file:
            image_data = BytesIO(image_file.read())

        image = Image(image_data)

        self.assertEqual(image.image_type, "PNG")
        self.assertEqual(image.width, 32)
        self.assertEqual(image.height, 32)
        self.assertEqual(image.x_dpi, 96)
        self.assertEqual(image.y_dpi, 96)

    def test_image_properties03(self):
        """Test the Image class properties."""
        image = Image("xlsxwriter/test/comparison/images/red_64x20.png")

        self.assertEqual(image.image_type, "PNG")
        self.assertEqual(image.width, 64)
        self.assertEqual(image.height, 20)
        self.assertEqual(image.x_dpi, 96)
        self.assertEqual(image.y_dpi, 96)

    def test_image_properties04(self):
        """Test a top-down BMP (negative height)."""
        image = Image("xlsxwriter/test/comparison/images/red_topdown_32x32.bmp")

        self.assertEqual(image.image_type, "BMP")
        self.assertEqual(image.width, 32)
        self.assertEqual(image.height, 32)
        self.assertEqual(image.x_dpi, 96)
        self.assertEqual(image.y_dpi, 96)

    def test_image_properties05(self):
        """Test a top-down BMP (negative height)."""
        image = Image("xlsxwriter/test/comparison/images/red_topdown_20x12.bmp")

        self.assertEqual(image.image_type, "BMP")
        self.assertEqual(image.width, 20)
        self.assertEqual(image.height, 12)
        self.assertEqual(image.x_dpi, 96)
        self.assertEqual(image.y_dpi, 96)

    def test_image_properties06(self):
        """Test that a GIF with dimensions >= 32768 are read as unsigned."""
        image = Image("xlsxwriter/test/comparison/images/black_40000x45000.gif")

        self.assertEqual(image.image_type, "GIF")
        self.assertEqual(image.width, 40000)
        self.assertEqual(image.height, 45000)
        self.assertEqual(image.x_dpi, 96)
        self.assertEqual(image.y_dpi, 96)

    def test_image_too_small(self):
        """Test that a truncated/too-small file raises UnsupportedImageFormat."""
        # Create and test a 12-byte buffer with a valid BMP marker but is too
        # short to hold the image markers/dimensions read from the header.
        image_data = BytesIO(b"BM" + b"\x00" * 10)

        with self.assertRaises(UnsupportedImageFormat):
            Image(image_data)

    def test_image_wmf_zero_inch(self):
        """Test that a WMF with a zero inch value doesn't divide by zero."""
        # A placeable WMF header with a malformed zero "inch" scaling value.
        data = bytearray(44)
        data[0:4] = pack("<L", 0x9AC6CDD7)  # Placeable WMF marker.
        data[10:12] = pack("<h", 100)  # Bounding box x2.
        data[12:14] = pack("<h", 100)  # Bounding box y2.
        data[14:16] = pack("<H", 0)  # Logical units per inch (malformed).

        with self.assertRaises(UndefinedImageSize):
            Image(BytesIO(bytes(data)))

    def test_image_emf_zero_frame(self):
        """Test that an EMF with a zero frame extent doesn't divide by zero."""
        # An EMF header with a malformed zero rectangular frame.
        data = bytearray(44)
        data[0:4] = pack("<l", 1)  # EMF iType marker.
        data[16:20] = pack("<l", 99)  # Bounding box x2.
        data[20:24] = pack("<l", 99)  # Bounding box y2.
        # Rectangular frame is left as zero (malformed).
        data[40:44] = b" EMF"  # EMF signature marker.

        image = Image(BytesIO(bytes(data)))

        self.assertEqual(image.image_type, "EMF")
        self.assertEqual(image.width, 100)
        self.assertEqual(image.height, 100)
        self.assertEqual(image.x_dpi, 96)
        self.assertEqual(image.y_dpi, 96)
