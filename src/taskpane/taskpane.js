/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console, document, Excel, Office */

Office.onReady((info) => {
	if (info.host === Office.HostType.Excel) {
		document.querySelector("#create-extraction-sheet").addEventListener(
			'click', () => tryCatch(extractData)
		);

		document.getElementById("app-body").style.display = "flex";
	}
});

async function createRosterDataSheet() {
	await Excel.run(async (context) => {
		const wSheetName = 'Roster Data';
		const worksheet = context.workbook.worksheets.add(wSheetName);

		worksheet.activate();

		await context.sync();
	});
}

async function createRosterDataTable() {
	await Excel.run(async (context) => {
		const rosterDataSheet = context.workbook.worksheets.getItem("Roster Data");
		
		const dataTable = rosterDataSheet.tables.add("A1:I1", true);
		dataTable.name = "rosterData";

		dataTable.getHeaderRowRange().values =	
			[["Name", "Service Point", "Date", "Start", "End", "Time", "OT", "Value", "Address"]];

		await context.sync();
	});	
}

async function extractData() {
	await Excel.run(async (context) => {
		if (await rosterDataSheetExists(context) === false)
			await createRosterDataSheet();

		if (await rosterDataTableExists(context) === false)
			await createRosterDataTable();

		const rosterTables = context.workbook.worksheets.getItem('Roster').tables;

		rosterTables.load('items/name');
		await context.sync();

		const tableNames = rosterTables.items.map(table => table.name);

		// Table format = Name, Service Point, Date, Start, End, Time, OT, Value, Address
		let rosterData = [];

		// Iterate through the tables
		for (const tableName of tableNames) {
			const table = rosterTables.getItem(tableName);

			rosterData = rosterData.concat(await extractRosterData(context, table));
		}

		const rosterDataTable = context.workbook.tables.getItem('rosterData');

		rosterDataTable.rows.add(null, rosterData);


		// format table
		rosterDataTable.getRange().format.autofitColumns();
		rosterDataTable.columns.getItem('Date').getDataBodyRange().numberFormat = 'dd/mm/yyyy';
	})

	/*
	 * HELPER FUNCTIONS
	 * ----------------
	 */

	// Extracts roster data from a single table (i.e. day)
	async function extractRosterData(context, table) {
		table.rows.load('items/values');
		await context.sync();

		const rows = table.rows.items;
		const rosterData = [];
		const date = await getDate(context, table);
		let address = '';

		const STARTTIMESEMIPHORE = 'from';
		const ENDTIMESEMIPHORE = 'til';
		
		rows.forEach(row => {
			const servicePoint = row.values[0][0];

			for (let colIndex = 2; colIndex < 14; colIndex++) {
				const cellValue = row.values[0][colIndex];

				if (cellValue !== '') {
					// Get name
					const name = extractName(cellValue);

					// Get start and end times
					const timeArr = getTime(getTimeString(colIndex), cellValue);
					const startTime = timeArr[0];
					const endTime = timeArr[1];

					// get Time
					const time = endTime > startTime ? endTime - startTime : endTime + 12 - startTime;
					
					// get OT
					const dateObj = excelDateToJSDate(date);
					let OT = '';

					if (extractOT(cellValue) === 'OT' || dateObj.getDay() == 6 || dateObj.getDay() == 7)
						OT = 'OT';

					const value = cellValue;

					rosterData.push([name, servicePoint, date, startTime, endTime, time, OT, value, address]);
				}
			}
		});

		return rosterData;
	}

	// Converts Excel date serial number to Date object
	function excelDateToJSDate(serial) {
		const utc_days = Math.floor(serial - 25569);
		const utc_value = utc_days * 86400;
		const date_info = new Date(utc_value * 1000);

		return date_info;
	}

	// Gets the date value of the rostered day from the date row above the header row of the table
	async function getDate(context, table) {
		const dateRange = table.getHeaderRowRange().getOffsetRange(-1, 0);

		dateRange.load('values');
		await context.sync();
	
		return dateRange.values[0][0];
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
		const timeStringOverride = extractTimeString(cellValue);
		
		if (timeStringOverride !== null) {
			const timeStringOverrideArr = timeStringOverride.split('-');

			// [from 9.30, til 3.30], [from 9.30], [til 3.30]
			timeStringOverrideArr.forEach(time => {
				if (time.includes(STARTSEMIPHORE))
					startTime = convertTime(time.replace(STARTSEMIPHORE, '').trim());

				if (time.includes(ENDSEMIPHORE))
					endTime = convertTime(time.replace(ENDSEMIPHORE, '').trim());
			});
		}

		return [+startTime, +endTime];	
	}

	// Extracts the staff name from the range
	function extractName(rangeValue) {
		let parenthesisIndex = rangeValue.indexOf("(");

		if (parenthesisIndex > 0)
			return rangeValue.substring(0, parenthesisIndex - 1).trim();

		return rangeValue.trim();
	}

	// Gets time string based of column index
	function getTimeString(columnIndex) {
		let timeString;

		switch (columnIndex) {
			case 2: 
				timeString = "7.00-8.00";
				break;
			case 3: 
				timeString = "8.00-9.00";
				break;
			case 4: 
				timeString = "9.00-10.00";
				break;
			case 5: 
				timeString = "10.00-11.00";
				break;
			case 6: 
				timeString = "11.00-12.00";
				break;
			case 7: 
				timeString = "12.00-1.00";
				break;
			case 8:
				timeString = "1.00-2.00";
				break;
			case 9:
				timeString = "2.00-3.00";
				break;
			case 10:
				timeString = "3.00-4.00";
				break;
			case 11:
				timeString = "4.00-5.00";
				break;
			case 12:
				timeString = "5.00-6.00";
				break;
			case 13:
				timeString = "6.00-7.00";
				break;
			default:
				throw new error("Invalid column index to determine time string");	
		}

		return timeString;
	}

	// Extracts the time string from the cell value (i.e. if it's not a time that starts/ends on the hour)
	function extractTimeString(rangeValue) {
		const pattern = /(from?|from ?)?(2[0-3]|[01]?[0-9])[\.\:]([0-5][0-9])([ -]?|( - )?|( -)?|(- )?)(til?|til ?)?(2[0-3]|[01]?[0-9])[\.\:]([0-5][0-9])|((from?|from ?)|(til?|til ?))(2[0-3]|[01]?[0-9])[\.\:]([0-5][0-9])/g;

		let matches = rangeValue.match(pattern);
		let timeString;

		matches === null ? timeString = matches : timeString = matches[0];

		return timeString;
	}

	// Returns true if the time string is a range (e.g. 9.00-10.00)
	function isTimeRange(timeString) {
		let pattern = /(from?|from ?)?(2[0-3]|[01]?[0-9])[\.\:]([0-5][0-9])([ -]?|( - )?|( -)?|(- )?)(til?|til ?)?(2[0-3]|[01]?[0-9])[\.\:]([0-5][0-9])/g;

		let matches = timeString.match(pattern);

		if (matches === null)
			return false;

		return true;
	}

	// Extracts "OT" from the range value
	function extractOT(cellValue) {
		const pattern = /\bOT\b/;

		return cellValue.match !== null ? true : false;
	}

	// Converts time to a double (e.g 9.30 -> 9.5)
	function convertTime(timeString) {
		let timeStringArr = timeString.split('.');

		let hour = timeStringArr[0] / 1;
		let minutes = timeStringArr[1] / 60;

		return hour + minutes;
	}

	async function rosterDataSheetExists(context) {
		const sheets = context.workbook.worksheets;
		sheets.load('items/name');

		await context.sync();

		if (sheets.items.find(sheet => sheet.name === "Roster Data"))
			return true;

		return false;
	}

	async function rosterDataTableExists(context) {
		const tables = context.workbook.worksheets.getItem("Roster Data").tables;
		tables.load("items/name");
		
		await context.sync();
		
		if (tables.items.find(table => table.name === "roster data"))
			return true;
		
		return false;
	}

	// Gets a list of service points. (NOT USED)
	async function extractServicePoints(context) {
		const rosterSheet = context.workbook.worksheets.getItem("Roster");
		const rosterTables = rosterSheet.tables;
		
		// Get list of table names on Roster sheet, these should correspond to the day's of the week
		rosterTables.load('items/name');
		await context.sync();

		let tableNames = [];
		rosterTables.items.forEach(table => tableNames.push(table.name));
		
		// Iterate over the list of table names to retrieve the service points
		let servicePoints = [];

		for (const tableName of tableNames) {
			const table = rosterTables.getItem(tableName);
			let servicePointRange = table.columns.getItem('Service Point').getDataBodyRange().load('values');

			await context.sync();

			let bodyValues = servicePointRange.values;

			// Reduce to 1D array
			bodyValues = bodyValues.reduce((acc, item) => acc.concat(item));

			servicePoints = [].concat(...bodyValues);
		}

		// Remove duplicates
		servicePoints = [...new Set(servicePoints)];

		// Remove header items
		const headerItems = ['UTS', 'HOME'];

		headerItems.forEach(item => {
			let indexOfItem = servicePoints.indexOf(item);

			if (indexOfItem >= 0)
				servicePoints.splice(indexOfItem, 1);
		});

		return servicePoints;
	}
}

async function tryCatch(callback) {
	try {
		await callback();
	} catch (error) {
		// TODO: display error in UI.
		console.error(error);
	}
}
