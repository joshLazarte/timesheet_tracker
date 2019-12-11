from .init_timesheet import timesheet, workbook

PTO_COL_START = 0
PTO_HEADERS = ['DATE', 'PAYCODE', 'HOURS', 'COMMENTS', '', '']

TIMESHEET_COL_START = PTO_COL_START + len(PTO_HEADERS)
TIMESHEET_HEADERS = ['DATE', 'IN', 'OUT', 'IN', 'OUT', 'IN', 'OUT', 'TOTAL HRS', 'ON CALL']
DEFAULT_TIMES = ['8:00AM', '12:00PM', '1:00PM', '5:00PM']

HEADER_ROW = 0
MAIN_HEADER = ['PTO', 'TIMESHEET']

CONTENT_ROW_START = HEADER_ROW + 1
CONTENT_ROW_END = CONTENT_ROW_START + 14

OT_ROW_START = CONTENT_ROW_END + 3
OT_DETAILS = ['REASON:', 'JOB #:', 'DATE:', 'TIMES:', 'TOTAL HR:']

# format vars
BOLD =  workbook.add_format({'bold': True})
CENTER = workbook.add_format({'align': 'center'})
HEADER_FORMAT = workbook.add_format({'bold': True, 'align': 'center'})

