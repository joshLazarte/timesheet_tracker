import datetime, xlsxwriter
from utils.dates import days, today	
from utils.excel import writeDates, writeHeader, writeTimes



workbook = xlsxwriter.Workbook('Timesheet ' + today + '.xlsx')
worksheet = workbook.add_worksheet()

writeHeader(worksheet)
writeDates(worksheet, days)
writeTimes(worksheet, days)
	
workbook.close()