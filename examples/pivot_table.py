###############################################################################
#
# An example of how to create pivot tables with the XlsxWriter module.
#
# A pivot table is a way of grouping and summarizing the data in a worksheet
# range.
#
# SPDX-License-Identifier: BSD-2-Clause
#
# Copyright (c) 2013-2026, John McNamara, jmcnamara@cpan.org
#
import xlsxwriter


def main():
    workbook = xlsxwriter.Workbook("pivot_table.xlsx")

    # Add a worksheet with some sample data for the pivot tables.
    worksheet = workbook.add_worksheet("Data")
    populate_pivot_data(workbook, worksheet)

    # -------------------------------------------------------------------------
    # 1. Add an empty pivot table.
    # -------------------------------------------------------------------------

    # Add a pivot table based on the worksheet data to a new worksheet.
    worksheet = workbook.add_worksheet("Pivot1")
    worksheet.add_pivot_table("A3", {"source": "Data!A1:D51"})

    # -------------------------------------------------------------------------
    # 2. Add a pivot table with a Row field and a Value field.
    # -------------------------------------------------------------------------

    # Add a pivot table that summarizes the "Volume" data by "Region". Note
    # that the data source can also be specified as a list.
    worksheet = workbook.add_worksheet("Pivot2")
    worksheet.add_pivot_table(
        2,
        0,
        {
            "source": ["Data", 0, 0, 50, 3],
            "rows": ["Region"],
            "values": [{"field": "Volume"}],
        },
    )

    # -------------------------------------------------------------------------
    # 3. Add a pivot table with a Row field, a Column field and a Value field.
    # -------------------------------------------------------------------------

    # Add a pivot table that summarizes the "Volume" data by "Region" and
    # "Item".
    worksheet = workbook.add_worksheet("Pivot3")
    worksheet.add_pivot_table(
        "A3",
        {
            "source": "Data!A1:D51",
            "rows": ["Region"],
            "columns": ["Item"],
            "values": [{"field": "Volume"}],
        },
    )

    # -------------------------------------------------------------------------
    # 4. Add a pivot table with a Filter field.
    # -------------------------------------------------------------------------

    # Add a pivot table that summarizes the "Volume" data by "Region", with a
    # filter on "Month". The filter is displayed in the first row and the
    # pivot table body starts in the third row.
    worksheet = workbook.add_worksheet("Pivot4")
    worksheet.add_pivot_table(
        "A1",
        {
            "source": "Data!A1:D51",
            "rows": ["Region"],
            "values": [{"field": "Volume"}],
            "filters": ["Month"],
        },
    )

    # -------------------------------------------------------------------------
    # 5. Add a pivot table with two Row fields.
    # -------------------------------------------------------------------------

    # Add a pivot table that summarizes the "Volume" data by "Region" and then
    # by "Item" within each region.
    worksheet = workbook.add_worksheet("Pivot5")
    worksheet.add_pivot_table(
        "A3",
        {
            "source": "Data!A1:D51",
            "rows": ["Region", "Item"],
            "values": [{"field": "Volume"}],
        },
    )

    # -------------------------------------------------------------------------
    # 6. Add a pivot table with several Value fields.
    # -------------------------------------------------------------------------

    # Add a pivot table that summarizes the "Volume" data by "Region" using
    # different functions.
    worksheet = workbook.add_worksheet("Pivot6")
    worksheet.add_pivot_table(
        "A3",
        {
            "source": "Data!A1:D51",
            "rows": ["Region"],
            "values": [
                {"field": "Volume"},
                {"field": "Volume", "function": "count"},
                {"field": "Volume", "function": "average"},
            ],
        },
    )

    # -------------------------------------------------------------------------
    # 7. Add a pivot table with a custom name and number format.
    # -------------------------------------------------------------------------

    # Add a pivot table that summarizes the "Volume" data by "Region", with a
    # custom name and number format.
    worksheet = workbook.add_worksheet("Pivot7")
    worksheet.add_pivot_table(
        "A3",
        {
            "source": "Data!A1:D51",
            "rows": ["Region"],
            "values": [
                {"field": "Volume", "name": "Total Volume", "num_format": "#,##0"},
            ],
        },
    )

    # -------------------------------------------------------------------------
    # 8. Add a pivot table with a different style and no grand totals.
    # -------------------------------------------------------------------------

    # Add a pivot table that summarizes the "Volume" data by "Region" and
    # "Item", without grand totals.
    worksheet = workbook.add_worksheet("Pivot8")
    worksheet.add_pivot_table(
        "A3",
        {
            "source": "Data!A1:D51",
            "rows": ["Region"],
            "columns": ["Item"],
            "values": [{"field": "Volume"}],
            "style": "Pivot Style Medium 9",
            "row_grand_totals": False,
            "column_grand_totals": False,
        },
    )

    # -------------------------------------------------------------------------
    # 9. Add a pivot table with a tabular layout.
    # -------------------------------------------------------------------------

    # Add a pivot table that summarizes the "Volume" data by "Region" and
    # "Item", with each row field in its own column.
    worksheet = workbook.add_worksheet("Pivot9")
    worksheet.add_pivot_table(
        "A3",
        {
            "source": "Data!A1:D51",
            "rows": ["Region", "Item"],
            "values": [{"field": "Volume"}],
            "layout": "tabular",
        },
    )

    # -------------------------------------------------------------------------
    # 10. Add two pivot tables to the same worksheet.
    # -------------------------------------------------------------------------

    # Add two pivot tables side by side that summarize the "Volume" data by
    # "Region" in one and by "Item" in the other.
    worksheet = workbook.add_worksheet("Pivot10")
    worksheet.add_pivot_table(
        "A3",
        {
            "source": "Data!A1:D51",
            "rows": ["Region"],
            "values": [{"field": "Volume"}],
        },
    )
    worksheet.add_pivot_table(
        "F3",
        {
            "source": "Data!A1:D51",
            "rows": ["Item"],
            "values": [{"field": "Volume"}],
        },
    )

    # -------------------------------------------------------------------------
    # 11. Add a pivot table to the same worksheet as the data.
    # -------------------------------------------------------------------------

    # Add a new worksheet with a copy of the sample data.
    worksheet = workbook.add_worksheet("Pivot11")
    populate_pivot_data(workbook, worksheet)

    # Add a pivot table to the right of the data that summarizes the "Volume"
    # data by "Region".
    worksheet.add_pivot_table(
        "G3",
        {
            "source": "Pivot11!A1:D51",
            "rows": ["Region"],
            "values": [{"field": "Volume"}],
        },
    )

    workbook.close()


def populate_pivot_data(workbook, worksheet):
    # Add the sample data for the pivot tables to the worksheet.
    data = [
        ["East", "Apple", 9000, "July"],
        ["East", "Apple", 5000, "April"],
        ["South", "Orange", 9000, "September"],
        ["North", "Apple", 2000, "November"],
        ["West", "Apple", 9000, "November"],
        ["South", "Pear", 7000, "October"],
        ["North", "Pear", 9000, "August"],
        ["West", "Orange", 1000, "December"],
        ["West", "Grape", 1000, "November"],
        ["South", "Pear", 10000, "April"],
        ["West", "Grape", 6000, "January"],
        ["South", "Orange", 3000, "May"],
        ["North", "Apple", 3000, "December"],
        ["South", "Apple", 7000, "February"],
        ["West", "Grape", 1000, "December"],
        ["East", "Grape", 8000, "February"],
        ["South", "Grape", 10000, "June"],
        ["West", "Pear", 7000, "December"],
        ["South", "Apple", 2000, "October"],
        ["East", "Grape", 7000, "December"],
        ["North", "Grape", 6000, "July"],
        ["East", "Pear", 8000, "February"],
        ["North", "Apple", 7000, "August"],
        ["North", "Orange", 7000, "July"],
        ["North", "Apple", 6000, "June"],
        ["South", "Grape", 8000, "September"],
        ["West", "Apple", 3000, "October"],
        ["South", "Orange", 10000, "November"],
        ["West", "Grape", 4000, "December"],
        ["North", "Orange", 5000, "August"],
        ["East", "Orange", 1000, "November"],
        ["East", "Orange", 4000, "October"],
        ["North", "Grape", 5000, "August"],
        ["East", "Apple", 1000, "July"],
        ["South", "Apple", 10000, "March"],
        ["East", "Grape", 7000, "October"],
        ["West", "Grape", 1000, "September"],
        ["East", "Grape", 10000, "October"],
        ["South", "Orange", 8000, "March"],
        ["North", "Apple", 4000, "July"],
        ["South", "Orange", 5000, "July"],
        ["West", "Apple", 4000, "June"],
        ["East", "Apple", 5000, "April"],
        ["North", "Pear", 3000, "August"],
        ["East", "Grape", 9000, "November"],
        ["North", "Orange", 8000, "October"],
        ["East", "Apple", 10000, "June"],
        ["South", "Pear", 1000, "December"],
        ["North", "Grape", 10000, "July"],
        ["East", "Grape", 6000, "February"],
    ]

    # Widen the columns for clarity.
    worksheet.set_column(0, 3, 12)

    # Write the header titles.
    header_format = workbook.add_format({"bold": True})
    worksheet.write_row(0, 0, ["Region", "Item", "Volume", "Month"], header_format)

    # Write the other worksheet data.
    for row, row_data in enumerate(data, 1):
        worksheet.write_row(row, 0, row_data)


if __name__ == "__main__":
    main()
