import copy

def writeDataRow(worksheet, rowStart, colStart, dataset, format):
	row = copy.copy(rowStart)
	col = copy.copy(colStart)
	
	for data in dataset:
		worksheet.write(row, col, data, format)
		col += 1
		
		
def writeDataCol(worksheet, rowStart, colStart, dataset, format):
	row = copy.copy(rowStart)
	col = copy.copy(colStart)
	
	for data in dataset:
		worksheet.write(row, col, data, format)
		row += 1
		
def writeMultiRows(worksheet, rowStart, rowEnd, colStart, dataset, rowSkip, format):
	row = copy.copy(rowStart)	
	
	for i in range(rowEnd - rowStart):
		col = copy.copy(colStart)
		for data in dataset:
			worksheet.write(row, col, data, format)
			col += 1
		row += rowSkip
		
def writeMultiCols(worksheet, rowStart, colStart, repeatNum, dataset, colSkip, format):
	col = copy.copy(colStart)	
	
	for i in range(repeatNum):
		row = copy.copy(rowStart)
		for data in dataset:
			worksheet.write(row, col, data, format)
			row += 1
		col += colSkip



		


