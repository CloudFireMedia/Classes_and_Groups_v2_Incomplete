function Test_Export() {
	ExportEvents({
		'regular_events_calendar': 'TEST',
		'new_events_calendar': 'TEST (NEW)',
		'exclude_dates_ss_url': 'https://docs.google.com/spreadsheets/d/1d0-hBf96ilIpAO67LR86leEq09jYP2866uWC48bJloc/edit',
		'populate_days': [
			'sunday',
			'monday',
			'tuesday',
			'wednesday',
			'thursday',
			'friday',
			'saturday'
		]
	});
}

function Test_ExcludeEvents() {
	var exclusion_dates = GetExclusionDates('https://docs.google.com/spreadsheets/d/1d0-hBf96ilIpAO67LR86leEq09jYP2866uWC48bJloc/edit'),
		regularEventsCalendars = CalendarApp.getCalendarsByName('TEST'),
		newEventsCalendars = CalendarApp.getCalendarsByName('TEST (NEW)');

	ExcludeEvents([regularEventsCalendars[0], newEventsCalendars[0]], exclusion_dates);
}

function Test_DeleteEvents() {
	DeleteEvents(['TEST', 'TEST (NEW)']);
}