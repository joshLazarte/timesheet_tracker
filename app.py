from utils.dates import days
from utils.excel.vars import *
from utils.excel.init_timesheet import timesheet, workbook
import utils.excel.format
from utils.excel.excel import *



# write PTO headers
writeDataRow(timesheet, CONTENT_ROW_START, PTO_COL_START, PTO_HEADERS, HEADER_FORMAT)

# write timesheet headers
writeDataRow(timesheet, CONTENT_ROW_START, TIMESHEET_COL_START, TIMESHEET_HEADERS, HEADER_FORMAT)

# write dates
writeDataCol(timesheet, CONTENT_ROW_START + 1, TIMESHEET_COL_START, days, None)

# write default times
writeMultiRows(timesheet, CONTENT_ROW_START + 1, CONTENT_ROW_END + 1, TIMESHEET_COL_START + 1, DEFAULT_TIMES, 1, None)

# write OT info
writeMultiCols(timesheet, OT_ROW_START, PTO_COL_START, 4, OT_DETAILS, 2, BOLD)

workbook.close()