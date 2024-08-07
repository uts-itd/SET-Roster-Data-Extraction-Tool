const SRDET = (() => {
	function extractName(cellValue) {
		const parenthesisIndex = cellValue.indexOf('(');

		if (parenthesisIndex > 0)
			return cellValue.substring(0, parenthesisIndex - 1).trim();

		return cellValue.trim();
	}

	function extractTime(cellValue) {
		const pattern = /(from?|from ?)?(2[0-3]|[01]?[0-9])[\.\:]([0-5][0-9])([ -]?|( - )?|( -)?|(- )?)(til?|til ?)?(2[0-3]|[01]?[0-9])[\.\:]([0-5][0-9])|((from?|from ?)|(til?|til ?))(2[0-3]|[01]?[0-9])[\.\:]([0-5][0-9])/g;

		let matches = cellValue.match(pattern);
		let timeString;

		matches === null ? timeString = matches : timeString = matches[0];

		return timeString;
	}

	function extractOT(cellValue) {
		const pattern = /\bOT\b/;

		return cellValue.match(pattern) !== null ? true : false;
	}

	function isTimeRange(timeString) {
		let pattern = /(from?|from ?)?(2[0-3]|[01]?[0-9])[\.\:]([0-5][0-9])([ -]?|( - )?|( -)?|(- )?)(til?|til ?)?(2[0-3]|[01]?[0-9])[\.\:]([0-5][0-9])/g;

		let matches = timeString.match(pattern);

		if (matches === null)
			return false;

		return true;
	}

	function convertTime(timeString) {
		let timeStringArr = timeString.split('.');

		let hour = timeStringArr[0] / 1;
		let minutes = timeStringArr[1] / 60;

		return hour + minutes;
	}

	function getTimeString(columnIndex) {
		const timeStrings = new Map([
			[3, "8.00-9.00"],
			[4, "9.00-10.00"],
			[5, "10.00-11.00"],
			[6, "11.00-12.00"],
			[7, "12.00-1.00"],
			[8, "1.00-2.00"],
			[9, "2.00-3.00"],
			[10, "3.00-4.00"],
			[11, "4.00-5.00"],
			[12, "5.00-6.00"],
			[13, "6.00-7.00"]
		]);

		return timeStrings.get(columnIndex);
	}

	// not used
	function extractTimeRanges(row) {
		const timePattern = /[0-9]{1,2}-[0-9]{1,2}/;
		const timeMap = new Map();
		let keyIndex = 2;

		row.forEach(item => {
			if(item.match(timePattern)) {
				const timeStringArr = item.split('-').map(time => time + '.00');
				const timeString = timeStringArr[0] + '-' + timeStringArr[1];

				timeMap.set(keyIndex++, timeString);
			}
		});

		return timeMap;
	}

	function excelDateToJSDate(serial) {
		const utc_days = Math.floor(serial - 25569);
		const utc_value = utc_days * 86400;
		const date_info = new Date(utc_value * 1000);

		return date_info;
	}

	// Gets the start and end time from the timeString or cellValue if time is present
	function getTime(timeString, cellValue) {
		const STARTSEMIPHORE = 'from';
		const ENDSEMIPHORE = 'til';
		const timeStringArr = timeString.split('-');

		// Set start and end time based of column header (timeString)
		let startTime = timeStringArr[0].replace(STARTSEMIPHORE, '').trim();
		let endTime = timeStringArr[1].replace(ENDSEMIPHORE, '').trim();

		// Checks if there are any time override values in the cells (e.g. John Doe (from 9.30)
		const timeStringOverride = SRDET.extractTime(cellValue);
		
		if (timeStringOverride !== null) {
			const timeStringOverrideArr = timeStringOverride.split('-');

			// [from 9.30, til 3.30], [from 9.30], [til 3.30]
			timeStringOverrideArr.forEach(time => {
				if (time.includes(STARTSEMIPHORE))
					startTime = SRDET.convertTime(time.replace(STARTSEMIPHORE, '').trim());

				if (time.includes(ENDSEMIPHORE))
					endTime = SRDET.convertTime(time.replace(ENDSEMIPHORE, '').trim());
			});
		}

		return [+startTime, +endTime];	
	}


	return {
		extractName,
		extractTime,
		extractOT,
		extractTimeRanges,
		isTimeRange,
		convertTime,
		getTimeString,
		getTime,
		excelDateToJSDate,
	};
})();

module.exports = SRDET;
