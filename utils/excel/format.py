from .vars import *
from .init_timesheet import timesheet, workbook
import copy

# set comments col width
comments_col = PTO_COL_START + (len(PTO_HEADERS) - 1)
timesheet.set_column(comments_col, comments_col, 20)

# outline PTO Section
def outlinePTOSection():
	col = copy.copy(PTO_COL_START)
	row = copy.copy(CONTENT_ROW_START)
	
	cell_format = workbook.add_format()
	cell_format.set_top()
	
	for i in range(5):
		timesheet.write(row, col, '', cell_format)
		col += 1
		
outlinePTOSection()

