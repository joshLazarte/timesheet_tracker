import datetime

weekdays = [	
	'Mon',
	'Tue',
	'Wed',
	'Thu',
	'Fri',
	'Sat',
	'Sun',
]
	
	
def buildDayStrings(startDate, numDays):
	dayStrings = []
	delta = datetime.timedelta(days=1)
	current = startDate
	
	for i in range(numDays):
		dateString = weekdays[current.weekday()] + ' ' + current.strftime('%m/%d');
		dayStrings.append(dateString)
		current += delta
	
	return dayStrings
	
	
days = buildDayStrings(datetime.datetime(2019, 12, 6), numDays=14)

today = datetime.datetime.now().strftime('%m-%d')