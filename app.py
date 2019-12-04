import datetime, xlsxwriter
from utils.dates import days, today	
from utils.excel import writeDates, writeHeader, writeTimes

workbook = xlsxwriter.Workbook('Timesheet ' + today + '.xlsx')
worksheet = workbook.add_worksheet()

worksheet.set_column(0, 0, 14)


writeHeader(worksheet)
writeDates(worksheet, days)
writeTimes(worksheet, days)
	
workbook.close()