import copy

def writeDataRow(worksheet, rowStart, colStart, dataset, getFormat):
	row = copy.copy(rowStart)
	col = copy.copy(colStart)
	
	for data in dataset:
		worksheet.write(row, col, data, getFormat(row, col))
		col += 1
		
		
def writeDataCol(worksheet, rowStart, colStart, dataset, getFormat):
	row = copy.copy(rowStart)
	col = copy.copy(colStart)
	
	for data in dataset:
		worksheet.write(row, col, data, getFormat(row, col))
		row += 1
		
def writeMultiRows(worksheet, rowStart, rowEnd, colStart, dataset, rowSkip, getFormat):
	row = copy.copy(rowStart)	
	
	for i in range(rowEnd - rowStart):
		col = copy.copy(colStart)
		for data in dataset:
			worksheet.write(row, col, data, getFormat(row, col))
			col += 1
		row += rowSkip
		
def writeMultiCols(worksheet, rowStart, colStart, repeatNum, dataset, colSkip, getFormat):
	col = copy.copy(colStart)	
	
	for i in range(repeatNum):
		row = copy.copy(rowStart)
		for data in dataset:
			worksheet.write(row, col, data, getFormat(row, col))
			row += 1
		col += colSkip


	
	

		


