##############################################################################
#
# An example of inserting 3D models into a worksheet using the XlsxWriter
# Python module.
#
# SPDX-License-Identifier: BSD-2-Clause
#
# Copyright (c) 2013-2025, John McNamara, jmcnamara@cpan.org
#
import xlsxwriter

# Create a new Excel file and add a worksheet.
workbook = xlsxwriter.Workbook("models_3d.xlsx")
worksheet = workbook.add_worksheet()

# Widen the first column to make the text clearer.
worksheet.set_column("A:A", 30)

# Insert a 3D model. The model must be in GLB (binary glTF) format.
worksheet.write("A2", "Insert a 3D model in a cell:")
worksheet.insert_3d_model("B2", "duck.glb")

# Insert a 3D model with custom dimensions.
worksheet.write("A12", "Insert a 3D model with size:")
worksheet.insert_3d_model("B12", "duck.glb", {"width": 200, "height": 200})

# Insert a 3D model with an offset.
worksheet.write("A23", "Insert a 3D model with offset:")
worksheet.insert_3d_model("B23", "duck.glb", {"x_offset": 15, "y_offset": 10})

# Insert a 3D model with a description for accessibility.
worksheet.write("A34", "Insert a 3D model with alt text:")
worksheet.insert_3d_model(
    "B34", "duck.glb", {"description": "A 3D model of a rubber duck"}
)

workbook.close()
