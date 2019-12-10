from utils.dates import days, today
from utils.excel.vars import *
from utils.excel.init_timesheet import timesheet, workbook
import copy
from utils.excel.excel import *
import xlsxwriter

from utils.excel.format import *


workbook = xlsxwriter.Workbook('Timesheet ' + today + '.xlsx')
timesheet = workbook.add_worksheet()

# write PTO headers
#writeDataRow(timesheet, CONTENT_ROW_START, PTO_COL_START, PTO_HEADERS, HEADER_FORMAT)

# write timesheet headers
#writeDataRow(timesheet, CONTENT_ROW_START, TIMESHEET_COL_START, TIMESHEET_HEADERS, HEADER_FORMAT)

# write dates
#writeDataCol(timesheet, CONTENT_ROW_START + 1, TIMESHEET_COL_START, days, None)

# write default times
#writeMultiRows(timesheet, CONTENT_ROW_START + 1, CONTENT_ROW_END + 1, TIMESHEET_COL_START + 1, DEFAULT_TIMES, 1, None)

# write OT info
#writeMultiCols(timesheet, OT_ROW_START, PTO_COL_START, 4, OT_DETAILS, 2, BOLD)


# write PTO header and section

def writePtoSection():
	# merge comment cells	
	row = copy.copy(CONTENT_ROW_START)
	startCol = PTO_COL_START + 3
	endCol = startCol + 3
	
	for i in range(CONTENT_ROW_END):
		timesheet.merge_range(row, startCol, row, endCol, '')
		row += 1	
	
	# set column size
	timesheet.set_column(1,2,12)	
	
	#write headers and border
	writeDataRow(timesheet, CONTENT_ROW_START, PTO_COL_START, PTO_HEADERS, getPtoCellFormat)

	colList = [''] * (CONTENT_ROW_END - 1)
	rowList = [''] * len(PTO_HEADERS)
	writeDataCol(timesheet, CONTENT_ROW_START + 1, PTO_COL_START, colList, getPtoCellFormat)
	writeDataCol(timesheet, CONTENT_ROW_START + 1, TIMESHEET_COL_START, colList, getPtoCellFormat)
	writeDataRow(timesheet, CONTENT_ROW_END, PTO_COL_START, rowList, getPtoCellFormat)
	
	



def getPtoCellFormat(row, col):
	format = {}
	
	if row == CONTENT_ROW_START:
		format['bold'] = True
		format['top'] = 2
		format['align'] = 'center'
		
	if col == PTO_COL_START:
		format['left'] = 2
		
	if col == TIMESHEET_COL_START:
		format['right'] = 2
		
	if row == CONTENT_ROW_END:
		format['bottom'] = 2

	return workbook.add_format(format)

writePtoSection()

workbook.close()