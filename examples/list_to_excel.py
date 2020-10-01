# A simple module to generate excel from Dictionary using XlsxWriter Module


import xlsxwriter


def genExcel(headers, data, path):

	# Create a new workbook
	workbook = xlsxwriter.Workbook('%s.xlsx' % path)

	# Add work sheet to the workbook
	worksheet = workbook.add_worksheet()


	# Add headers / first row into the worksheet
	for h in range(len(headers)):
		worksheet.write(0, h, headers[h])

	# Loop over data and insert it into the worksheet
	counter_1 = 1
	for d in data:
		counter_2 = 0
		for a in d:
			worksheet.write(counter_1, counter_2, a)
			counter_2 += 1
		counter_1 += 1

	# Close the worksheet
	workbook.close()


# Variables for the excel
EXCEL_HEADERS = ['First Name','Last Name','Mobile','Email']
EXCEL_DATA = [
	['Ricky','Singh','8800XXXXXX','ricky@XXXXXX.XXX'],
	['Steve','Jobs','9900XXXXXX','steve@XXXX.XX.XX'],
	['Tony','Stark','XXXXXXX999','ironman@marvel.universe'],
]


# Generate the excel
genExcel(EXCEL_HEADERS, EXCEL_DATA,'test')
