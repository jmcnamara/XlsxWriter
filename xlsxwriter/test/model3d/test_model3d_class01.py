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

from xlsxwriter.model3d import Model3D


class TestModel3DProperties(unittest.TestCase):
    """
    Test the properties of a Model3D object.
    """

    def test_model3d_properties01(self):
        """Test the Model3D class properties from file."""
        model = Model3D("xlsxwriter/test/comparison/models/Duck.glb")

        self.assertEqual(model.model_type, "GLB")
        self.assertIsNotNone(model._digest)
        self.assertGreater(model._width, 0)
        self.assertGreater(model._height, 0)

    def test_model3d_properties02(self):
        """Test the Model3D class properties from BytesIO."""
        with open("xlsxwriter/test/comparison/models/Duck.glb", "rb") as model_file:
            model_data = BytesIO(model_file.read())

        model = Model3D(model_data)

        self.assertEqual(model.model_type, "GLB")
        self.assertIsNotNone(model._digest)

    def test_model3d_properties03(self):
        """Test the Model3D class default dimensions."""
        model = Model3D("xlsxwriter/test/comparison/models/Duck.glb")

        # Default dimensions should be set
        self.assertIsInstance(model._width, float)
        self.assertIsInstance(model._height, float)

    def test_model3d_user_options(self):
        """Test the Model3D class with user options."""
        model = Model3D("xlsxwriter/test/comparison/models/Duck.glb")

        model._set_user_options({
            "width": 300,
            "height": 250,
            "x_offset": 10,
            "y_offset": 20,
            "description": "A 3D duck model",
            "decorative": True,
        })

        self.assertEqual(model._width, 300)
        self.assertEqual(model._height, 250)
        self.assertEqual(model._x_offset, 10)
        self.assertEqual(model._y_offset, 20)
        self.assertEqual(model._description, "A 3D duck model")
        self.assertTrue(model._decorative)

    def test_model3d_description_property(self):
        """Test the Model3D description property."""
        model = Model3D("xlsxwriter/test/comparison/models/Duck.glb")

        self.assertIsNone(model.description)

        model.description = "Test description"
        self.assertEqual(model.description, "Test description")

    def test_model3d_decorative_property(self):
        """Test the Model3D decorative property."""
        model = Model3D("xlsxwriter/test/comparison/models/Duck.glb")

        self.assertFalse(model.decorative)

        model.decorative = True
        self.assertTrue(model.decorative)

    def test_model3d_repr(self):
        """Test the Model3D __repr__ method."""
        model = Model3D("xlsxwriter/test/comparison/models/Duck.glb")

        repr_str = repr(model)
        self.assertIn("Model3D:", repr_str)
        self.assertIn("Duck.glb", repr_str)
        self.assertIn("GLB", repr_str)
