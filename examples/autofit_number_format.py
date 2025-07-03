from xlsxwriter.workbook import Workbook

workbook = Workbook("autofit.xlsx")
worksheet = workbook.add_worksheet()

# Write some worksheet data to demonstrate autofitting.
for i in range(12):
    worksheet.write(0, i, 1.123)
    worksheet.write(1, i, 12.123)
    worksheet.write(2, i, 123.123)
    worksheet.write(3, i, 123.123456789)

    format = workbook.add_format()
    if i == 0:
        format = None
    elif i == 1:
        format.set_num_format('0')
    elif i > 1:
        if i == 3:
            format.set_num_format(2)
        else:
            format.set_num_format('0.'+'0'*(i-1))

    worksheet.set_column(i,i, None, format)

# Autofit the worksheet.
worksheet.autofit()

workbook.close()
