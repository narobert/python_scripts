import xlrd
import xlsxwriter

# Counts the number of times faculty has guest lectured
file_location = "/Users/narobert/Lectures.xlsx"

workbook = xlrd.open_workbook(file_location)

sheet = workbook.sheet_by_index(0)

sheet2 = workbook.sheet_by_index(1)

count1 = 0
count2 = 0
count3 = 0
count4 = 0

for row in range(sheet.nrows):
		
        # Replaces last_name manually using faculty list
	names = sheet.cell_value(row, 0)

	if names.count(sheet2.cell_value(1, 0)) == 1:
		count1 += 1

        if names.count(sheet2.cell_value(2, 0)) == 1:
		count2 += 1

        if names.count(sheet2.cell_value(3, 0)) == 1:
		count3 += 1
 
        if names.count(sheet2.cell_value(4, 0)) == 1:
		count4 += 1

print sheet2.cell_value(1, 0), "has guest lectured", count1, "times"
print sheet2.cell_value(2, 0), "has guest lectured", count2, "times"
print sheet2.cell_value(3, 0), "has guest lectured", count3, "times"
print sheet2.cell_value(4, 0), "has guest lectured", count4, "times"


# Creates an excel spreadsheet of how many times faculty has guest lectured
workbook = xlsxwriter.Workbook('Count_Lectures.xlsx')
worksheet = workbook.add_worksheet()

lectures = (
	[sheet2.cell_value(1, 0), count1],
	[sheet2.cell_value(2, 0), count2],
	[sheet2.cell_value(3, 0), count3],
	[sheet2.cell_value(4, 0), count4],
)

row = 0
col = 0

for last, times in (lectures):
	worksheet.write(row, col, last)
	worksheet.write(row, col + 1, times)
	row += 1

workbook.close()
