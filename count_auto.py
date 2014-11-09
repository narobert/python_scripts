import xlrd
import xlsxwriter

# Counts the number of times faculty has guest lectured
file_location = "/Users/narobert/Lectures.xlsx"

workbook = xlrd.open_workbook(file_location)

sheet = workbook.sheet_by_index(0)

sheet2 = workbook.sheet_by_index(1)

count = 0

lectures = []

for rows in range(sheet2.nrows):

	if rows == 0:

		continue

	count = 0

	# Increments through list of names and appends to array
	for row in range(sheet.nrows):
		
		names = sheet.cell_value(row, 0)

		words = names.split()
	
		for word in words:

			if word == sheet2.cell_value(rows, 0):

              	 		count += 1

	print sheet2.cell_value(rows, 0), "has guest lectured", count, "times"

	lectures.append([sheet2.cell_value(rows, 0), count])

		
# Creates an excel spreadsheet of how many times faculty has guest lectured
workbook = xlsxwriter.Workbook('Count_Lectures.xlsx')
worksheet = workbook.add_worksheet()

row = 0
col = 0

for last, times in (lectures):
	worksheet.write(row, col, last)
	worksheet.write(row, col + 1, times)
	row += 1

workbook.close()	
