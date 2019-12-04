import xlsxwriter


def writeHeader(worksheet):
	headers = ['DATE', 'IN', '', 'OUT', 'IN', '', 'OUT', 'IN', '', 'OUT']

	row = 0
	col = 0
	
	for i in range(len(headers)):
		worksheet.write(row, col, headers[i])
		col += 1
		
def writeTimes(worksheet, days):
	times = ['8:00AM', '', '12:00PM', '1:00PM', '', '5:00PM']
	
	row = 1	
	
	for i in range(len(days)):
		col = 1
		for time in times:			
			worksheet.write(row, col, time)
			col += 1
		row += 1


def writeDates(worksheet, days):
	row = 0
	col = 0
	
	for day in days:
		row += 1
		worksheet.write(row, col, day)
		


