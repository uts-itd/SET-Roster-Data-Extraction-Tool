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

		// Extract list of service points from roster sheet
		//let servicePoints = await extractServicePoints(context);

		// Extract roster data
		const rosterSheet = context.workbook.worksheets.getItem('Roster');
		const rosterTables = rosterSheet.tables;

		rosterTables.load('items/name');
		await context.sync();

		const tableNames = [];
		rosterTables.items.forEach(table => tableNames.push(table.name));

		// Table format = Name, Service Point, Date, Start, End, Time, OT, Value, Address
		let rosterData = [];

		// Iterate through the tables
		for (const tableName of tableNames) {
			const table = rosterTables.getItem(tableName);

			rosterData = rosterData.concat(await extractRosterData(context, table));
		}

		let rosterDataTable = context.workbook.tables.getItem('rosterData');
		await context.sync();

		rosterDataTable.rows.add(null, rosterData);

		await context.sync();
	})

	/*
	 * HELPER FUNCTIONS
	 * ----------------
	 */

	// Converts Excel Date serial number to Date object
	function excelDateToJSDate(serial) {
		let utc_days = Math.floor(serial - 25569);
		let utc_value = utc_days * 86400;
		let date_info = new Date(utc_value * 1000);

		return date_info;
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

	// Extracts roster data from a single table (i.e. day)
	async function extractRosterData(context, table) {
		let rows = table.rows.load('items/values');
		let dateRange = table.getHeaderRowRange().getOffsetRange(-1,0);

		dateRange.load('values');
		await context.sync();
		
		let numberOfRows = rows.items.length;
		let rosterData = [];

		let name = '';
		let servicePoint = '';
		let date = excelDateToJSDate(dateRange.values[0][0]);
		let startTime = '';
		let endTime = '';
		let time = '';
		let OT = '';
		let value = '';
		let address = '';
		
		const STARTTIMESEMIPHORE = 'from';
		const ENDTIMESEMIPHORE = 'til';
		
		// Iterate through the rows
		for (let rowIndex = 0; rowIndex < numberOfRows; rowIndex++) {
			// Get Service Point
			servicePoint = rows.items[rowIndex].values[0][0];

			for (let colIndex = 2; colIndex < 14; colIndex++) {
				let cellValue = rows.items[rowIndex].values[0][colIndex];

				if (cellValue !== '') {
					// Get name
					name = extractName(cellValue);

					// Get start and end times
					let timeString = getTimeString(colIndex);
					let timeStringArr = timeString.split('-');
					
					startTime = timeStringArr[0].replace(STARTTIMESEMIPHORE, '').trim();
					endTime = timeStringArr[1].replace(ENDTIMESEMIPHORE, '').trim();

					timeStringOverride = extractTimeString(cellValue);

					if (timeStringOverride !== null)
						timeString = timeStringOverride;

					if (isTimeRange(timeString)) {
						let timeStringArr = timeString.split("-");

						startTime = convertTime(timeStringArr[0].replace(STARTTIMESEMIPHORE,'').trim());
						endTime = convertTime(timeStringArr[1].replace(ENDTIMESEMIPHORE, '').trim());

					} else {
						if (timeString.includes(STARTTIMESEMIPHORE))
							startTime = convertTime(timeString.replace(STARTTIMESEMIPHORE, '').trim());

						if (timeString.includes(ENDTIMESEMIPHORE))
							endTime = convertTime(timeString.replace(ENDTIMESEMIPHORE, '').trim());
					}

					// get Time
					if (endTime > startTime)
						time = endTime - startTime;
					else
						time = endTime + 12 - startTime;
					
					// get OT
					if (extractOT(cellValue)) {
						OT = 'OT';
					}

					if (date.getDay() == 6 || 
						date.getDay() == 7)
						OT = 'OT';

					value = cellValue;

					rosterData.push([name, servicePoint, date, startTime, endTime, time, OT, value, address]);
				}
			}
		};

		return rosterData;
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
