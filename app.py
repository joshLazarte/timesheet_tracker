from utils.dates import days, today
from utils.excel.vars import *
from utils.excel.init_timesheet import timesheet, workbook
import copy
from utils.excel.excel import *
import xlsxwriter

from utils.excel.format import *


workbook = xlsxwriter.Workbook('Timesheet ' + today + '.xlsx')
timesheet = workbook.add_worksheet()

def writePtoSection():
	# set col width
	timesheet.set_column(PTO_COL_START, TIMESHEET_COL_START + len(PTO_HEADERS), 12)

	# merge comment cells	
	row = copy.copy(CONTENT_ROW_START)
	startCol = PTO_COL_START + 3
	endCol = startCol + 2
	
	for i in range(CONTENT_ROW_END):
		timesheet.merge_range(row, startCol, row, endCol, '')
		row += 1			
	
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
		format['bottom'] = 2		
		
	if col == PTO_COL_START:
		format['left'] = 2
		
	if col == TIMESHEET_COL_START:
		format['right'] = 2
		
	if row == CONTENT_ROW_END:
		format['bottom'] = 2

	return workbook.add_format(format)
	

def getTimeSheetFormat(row, col):
	format = {}
	format['align'] = 'center'
	
	if row == CONTENT_ROW_START:
		format['bold'] = True
		format['top'] = 2		
		format['bottom'] = 2		
		
	if col == TIMESHEET_COL_START:
		format['left'] = 2		
		
	if col == TIMESHEET_COL_START + len(TIMESHEET_HEADERS) - 1:
		format['right'] = 2
		
	if row == CONTENT_ROW_END:
		format['bottom'] = 2

	return workbook.add_format(format)
	
def getOTFormat(row, col):
	format = {}
	
	if row == CONTENT_ROW_START:
		format['bold'] = True
		format['top'] = 2		
		format['bottom'] = 2		
		
	if col == TIMESHEET_COL_START:
		format['left'] = 2		
		
	if col == TIMESHEET_COL_START + len(TIMESHEET_HEADERS) - 1:
		format['right'] = 2
		
	if row == CONTENT_ROW_END:
		format['bottom'] = 2
	
	return workbook.add_format(format)

def writeTimesheet():
	# write timesheet headers
	writeDataRow(timesheet, CONTENT_ROW_START, TIMESHEET_COL_START, TIMESHEET_HEADERS, getTimeSheetFormat)	
	timesheet.set_column(TIMESHEET_COL_START, TIMESHEET_COL_START + len(TIMESHEET_HEADERS), 12)
	
	# initial outline
	colList = [''] * (CONTENT_ROW_END - 1)
	rowList = [''] * len(TIMESHEET_HEADERS)
	writeDataCol(timesheet, CONTENT_ROW_START + 1, TIMESHEET_COL_START + len(TIMESHEET_HEADERS) - 1, colList, getTimeSheetFormat)
	writeDataRow(timesheet, CONTENT_ROW_END, TIMESHEET_COL_START, rowList, getTimeSheetFormat)
		
	# write dates
	writeDataCol(timesheet, CONTENT_ROW_START + 1, TIMESHEET_COL_START, days, getTimeSheetFormat)

	# write default times
	writeMultiRows(timesheet, CONTENT_ROW_START + 1, CONTENT_ROW_END + 1, TIMESHEET_COL_START + 1, DEFAULT_TIMES, 1, getTimeSheetFormat)

def getOTKeyFormat(row, col):
	format = {}
	format['bold'] = True
	format['left'] = 2	
	format['right'] = 2	
		
	if row == OT_ROW_START:
		format['top'] = 2			
		
	if row == OT_ROW_START + len(OT_DETAILS) - 1:
		format['bottom'] = 2
	
	return workbook.add_format(format)

def getOTValueFormat(row, startCol, col):
	format = {}		
	
	if col == startCol + 2:
		format['right'] = 2
	
	if row == OT_ROW_START:
		format['top'] = 2			
		
	if row == OT_ROW_START + len(OT_DETAILS) - 1:
		format['bottom'] = 2
	
	return workbook.add_format(format)
	
def writeOTDetails():	
	# write OT info		
	col = copy.copy(PTO_COL_START)	
	
	for i in range(4):
		row = copy.copy(OT_ROW_START)
		for data in OT_DETAILS:
			timesheet.write(row, col, data, getOTKeyFormat(row, col))			
			timesheet.write(row, col + 1, '', getOTValueFormat(row, col, col + 1))			
			timesheet.write(row, col + 2, '', getOTValueFormat(row, col, col + 2))			
			timesheet.merge_range(row, col+1, row, col+2, '')
			row += 1
		
		col += 4

writePtoSection()
writeTimesheet()
writeOTDetails()

workbook.close()