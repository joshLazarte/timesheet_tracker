from .vars import *
from .init_timesheet import timesheet, workbook
import copy



# merge comments cells
def mergeCommentCells():
	startCol = PTO_COL_START + 3
	endCol = startCol + 2
	row = copy.copy(CONTENT_ROW_START) 
	
	for i in range(CONTENT_ROW_END):
		timesheet.merge_range(row, startCol, row, endCol, '')
		row += 1
		



	
	
	
	

mergeCommentCells()

