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

from ...image import Image


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
