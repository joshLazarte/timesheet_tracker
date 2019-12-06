from .init_timesheet import timesheet, workbook

PTO_COL_START = 0
TIMESHEET_COL_START = 4
HEADER_ROW = 0
CONTENT_ROW_START = 1
CONTENT_ROW_END = CONTENT_ROW_START + 14

OT_ROW_START = CONTENT_ROW_END + 2

MAIN_HEADER = ['PTO', 'TIMESHEET']
PTO_HEADERS = ['DATE', 'PAYCODE', 'HOURS', 'COMMENTS']
TIMESHEET_HEADERS = ['DATE', 'IN', '', 'OUT', 'IN', '', 'OUT', 'IN', '', 'OUT', 'TOTAL HRS', 'ON CALL']
DEFAULT_TIMES = ['8:00AM', '', '12:00PM', '1:00PM', '', '5:00PM']
OT_DETAILS = ['REASON: ', 'JOB NUMBER: ', 'DATE: ', 'HOURS WORKED: ', 'TOTAL HOURS: ' ]

# format vars
BOLD =  workbook.add_format({'bold': True})

