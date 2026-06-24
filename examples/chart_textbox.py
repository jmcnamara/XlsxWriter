#######################################################################
#
# An example of creating Excel charts containing textboxes with Python
# and XlsxWriter.
#
# SPDX-License-Identifier: BSD-2-Clause
#
# Copyright (c) 2013-2025, John McNamara, jmcnamara@cpan.org
#
import xlsxwriter

workbook = xlsxwriter.Workbook("chart_textbox.xlsx")
worksheet = workbook.add_worksheet()
bold = workbook.add_format({"bold": 1})

# Add the worksheet data that the charts will refer to.
headings = ["Number", "Batch 1", "Batch 2"]
data = [
    [2, 3, 4, 5, 6, 7],
    [10, 40, 50, 20, 10, 50],
    [30, 60, 70, 50, 40, 30],
]

worksheet.write_row("A1", headings, bold)
worksheet.write_column("A2", data[0])
worksheet.write_column("B2", data[1])
worksheet.write_column("C2", data[2])


#######################################################################
#
# Create a new scatter chart with a simple textbox.
#
chart1 = workbook.add_chart({"type": "scatter"})

# Configure the first series.
chart1.add_series(
    {
        "name": "=Sheet1!$B$1",
        "categories": "=Sheet1!$A$2:$A$7",
        "values": "=Sheet1!$B$2:$B$7",
    }
)

# Configure second series. Note use of alternative syntax to define ranges.
chart1.add_series(
    {
        "name": ["Sheet1", 0, 2],
        "categories": ["Sheet1", 1, 0, 6, 0],
        "values": ["Sheet1", 1, 2, 6, 2],
    }
)

# Add a chart title and some axis labels.
chart1.set_title({"name": "Simple chart textbox"})
chart1.set_x_axis({"name": "Test number"})
chart1.set_y_axis({"name": "Sample length (mm)"})

# Set an Excel chart style.
chart1.set_style(11)

# Add a textbox to the chart (with specific relative anchor point).
chart1.add_textbox("Hello", {"x": 0.7, "y": 0.2})

# Insert the chart into the worksheet (with an offset).
worksheet.insert_chart("D2", chart1, {"x_offset": 25, "y_offset": 10})

#######################################################################
#
# Create a scatter chart sub-type with straight lines and markers and a multi-line rich text textbox.
#
chart2 = workbook.add_chart({"type": "scatter", "subtype": "straight_with_markers"})

# Configure the first series.
chart2.add_series(
    {
        "name": "=Sheet1!$B$1",
        "categories": "=Sheet1!$A$2:$A$7",
        "values": "=Sheet1!$B$2:$B$7",
    }
)

# Configure second series.
chart2.add_series(
    {
        "name": "=Sheet1!$C$1",
        "categories": "=Sheet1!$A$2:$A$7",
        "values": "=Sheet1!$C$2:$C$7",
    }
)

# Add a chart title and some axis labels.
chart2.set_title({"name": "Rich text in textbox"})
chart2.set_x_axis({"name": "Test number"})
chart2.set_y_axis({"name": "Sample length (mm)"})

# Set an Excel chart style.
chart2.set_style(12)

# Add a multi-line rich text textbox to the chart (with specific relative anchor point and size).
chart2.add_textbox(
    [
        {"align": "left", "runs": [{"text": "2023-2014", "font": {"underline": True}}]},
        {
            "align": "left",
            "runs": [
                {"text": "C", "font": {"italic": True, "name": "Times New Roman"}},
                {"text": "max", "font": {"italic": True, "baseline": -25000}},
                {"text": " = 161.3 (\xb14.3)"},
            ],
        },
        {
            "align": "left",
            "runs": [
                {"text": "k", "font": {"italic": True, "name": "Times New Roman"}},
                {"text": "bio", "font": {"italic": True, "baseline": -25000}},
                {"text": " = 0.624 (\xb10.070)"},
            ],
        },
        {
            "align": "left",
            "runs": [
                {"text": "k", "font": {"italic": True, "name": "Times New Roman"}},
                {"text": "exposure", "font": {"italic": True, "baseline": -25000}},
                {"text": " = 0"},
            ],
        },
    ],
    {"x": 0.6, "y": 0.18, "width": 0.4, "height": 0.6},
)

# Insert the chart into the worksheet (with an offset).
worksheet.insert_chart("D18", chart2, {"x_offset": 25, "y_offset": 10})

workbook.close()
