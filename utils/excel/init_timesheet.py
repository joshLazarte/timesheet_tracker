import xlsxwriter
from ..dates import today

workbook = xlsxwriter.Workbook('Timesheet ' + today + '.xlsx')
timesheet = workbook.add_worksheet()


