#######################################################################
#
# An example of creating Excel Radar charts with Python and XlsxWriter.
#
# SPDX-License-Identifier: BSD-2-Clause
#
# Copyright (c) 2013-2025, John McNamara, jmcnamara@cpan.org
#
import xlsxwriter

workbook = xlsxwriter.Workbook("chart_radar.xlsx")
worksheet = workbook.add_worksheet()
bold = workbook.add_format({"bold": 1})

# Add the worksheet data that the charts will refer to.
headings = ["Number", "Batch 1", "Batch 2"]
data = [
    [2, 3, 4, 5, 6, 7],
    [30, 60, 70, 50, 40, 30],
    [25, 40, 50, 30, 50, 40],
]

worksheet.write_row("A1", headings, bold)
worksheet.write_column("A2", data[0])
worksheet.write_column("B2", data[1])
worksheet.write_column("C2", data[2])

#######################################################################
#
# Create a new radar chart.
#
chart1 = workbook.add_chart({"type": "radar"})

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

# Add a chart title.
chart1.set_title({"name": "Results of sample analysis"})

# Set an Excel chart style.
chart1.set_style(11)

# Insert the chart into the worksheet (with an offset).
worksheet.insert_chart("D2", chart1, {"x_offset": 25, "y_offset": 10})

#######################################################################
#
# Create a radar chart with markers chart sub-type.
#
chart2 = workbook.add_chart({"type": "radar", "subtype": "with_markers"})

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

# Add a chart title.
chart2.set_title({"name": "Radar Chart With Markers"})

# Set an Excel chart style.
chart2.set_style(12)

# Insert the chart into the worksheet (with an offset).
worksheet.insert_chart("D18", chart2, {"x_offset": 25, "y_offset": 10})

#######################################################################
#
# Create a filled radar chart sub-type.
#
chart3 = workbook.add_chart({"type": "radar", "subtype": "filled"})

# Configure the first series.
chart3.add_series(
    {
        "name": "=Sheet1!$B$1",
        "categories": "=Sheet1!$A$2:$A$7",
        "values": "=Sheet1!$B$2:$B$7",
    }
)

# Configure second series.
chart3.add_series(
    {
        "name": "=Sheet1!$C$1",
        "categories": "=Sheet1!$A$2:$A$7",
        "values": "=Sheet1!$C$2:$C$7",
    }
)

# Add a chart title.
chart3.set_title({"name": "Filled Radar Chart"})

# Set an Excel chart style.
chart3.set_style(13)

# Insert the chart into the worksheet (with an offset).
worksheet.insert_chart("D34", chart3, {"x_offset": 25, "y_offset": 10})

workbook.close()
