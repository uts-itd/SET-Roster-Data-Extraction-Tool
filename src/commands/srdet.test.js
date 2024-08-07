const SRDET = require('./srdet');

describe('extractName() takes a string value of from a roster cell and returns the name', () => {
	const nameString = "John Doe ()";

	const extractedName = SRDET.extractName(nameString);

	test('Name extracted is "John Doe"', () => {
		expect(extractedName).toBe('John Doe');
	});
});

describe('extractTime() takes a string value from a roster cell and returns the time string found in parenthesis', () => {
	const nameStrings = [
		'John Doe (from 9.30 OT)', 
		'John Doe (til 9.30)',
		'John Doe (from 9.15 - til 9.45)',
		'John Doe (9.15-9.45)',
		'John Doe'
	];

	const extractedTimes = nameStrings.map(str => SRDET.extractTime(str));

	test('Time extracted from "John Doe (from 9.30 OT)" is "from 9.30"', () => {
		expect(extractedTimes[0]).toBe('from 9.30');
	});

	test('Time extracted from "John Doe (til 9.30)" is "til 9.30"', () => {
		expect(extractedTimes[1]).toBe('til 9.30');
	});

	test('Time extracted from "John Doe (from 9.15 - til 9.45)" is "from 9.15 - til 9.45"2', () => {
		expect(extractedTimes[2]).toBe('from 9.15 - til 9.45');
	});

	test('Time extracted from "John Doe (9.15-9.45)" is "9.15-9.45"', () => {
		expect(extractedTimes[3]).toBe('9.15-9.45');
	});

	test('Time extracted from "John Doe" is Null', () => {
		expect(extractedTimes[4]).toBeNull();
	});
});

describe('extractOT() returns true if the string "OT" is found, otherwise false', () => {
	const nameStrings = [
		'John Doe (OT)',
		'John Doe'
	];

	const extractedOTs = nameStrings.map(str => SRDET.extractOT(str));

	test('String "John Doe (OT)" returns true', () => {
		expect(extractedOTs[0]).toBeTruthy();
	});

	test('String "John Doe" returns false', () => {
		expect(extractedOTs[1]).toBeFalsy();
	});
});

describe('isTimeRange() checks if the time string is a range', () => {
	const timeStrings = [
		'9.30',
		'9.30-10.30'
	];

	const isTimeRanges = timeStrings.map(str => SRDET.isTimeRange(str));

	test('String "9.30" is falsy', () => {
		expect(isTimeRanges[0]).toBeFalsy();
	});

	test('String "9.30-10.30" is truthy', () => {
		expect(isTimeRanges[1]).toBeTruthy();
	});
});

describe('convertTime() converts a time string to double', () => {
	const timeString = '9.30';

	const timeDbl = SRDET.convertTime(timeString);

	test('Converts "9.30" to 9.5', () => {
		expect(timeDbl).toBe(9.5);
	});
});

describe('extractTimeRanges() converts the timeRow to a Map of timeStrings', () => {
	const row = [
		'Service Points',
		'Details',
		'8-9',
		'9-10',
		'10-11',
		'11-12',
		'12-1',
		'1-2',
		'2-3',
		'3-4',
		'4-5',
		'5-6',
		'6-7'
	];

	const expectedMap = new Map([
		[2, '8.00-9.00'],
		[3, '9.00-10.00'],
		[4, '10.00-11.00'],
		[5, '11.00-12.00'],
		[6, '12.00-1.00'],
		[7, '1.00-2.00'],
		[8, '2.00-3.00'],
		[9, '3.00-4.00'],
		[10, '4.00-5.00'],
		[11, '5.00-6.00'],
		[12, '6.00-7.00']
	]);

	const timeRangeMap = SRDET.extractTimeRanges(row);

	test('timeRangeMap is length of 11', () => {
		expect(timeRangeMap.size).toBe(11);
	});

	test.skip('timeRangeMap keys are all integers', () => {
		expect(timeRangeMap.keys().every(key => typeof(key) === 'number')).toBeTruthy();
	});

	test('timeRangeMap is same as expectedMap', () => {
		expect(timeRangeMap).toEqual(expectedMap);
	});
});

describe('excelDateToJSDate() convert excel date serial to JS Date format', () => {
	const dateSerial = 45108;
	const expectedDate = new Date('2023-07-01T00:00:00.000Z');

	const convertedDate = SRDET.excelDateToJSDate(dateSerial);

	test('45108 is converted to "2023-07-01T00:00:00.000Z"', () => {
		expect(convertedDate).toEqual(expectedDate);
	});
});

// timeString = h.mm
// timeRange = h.mm-h.mm
