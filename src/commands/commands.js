/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global Office */

Office.onReady(() => {
  // If needed, Office.js is ready to be called.
});

/**
 * Shows a notification when the add-in command is executed.
 * @param event {Office.AddinCommands.Event}
 */
function action(event) {
  const message = {
    type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
    message: "Performed action.",
    icon: "Icon.80x80",
    persistent: true,
  };

  // Show a notification message.
  Office.context.mailbox.item.notificationMessages.replaceAsync("action", message);

  // Be sure to indicate when the add-in command function is complete.
  event.completed();
}

// My SRDET stuff here:
async function extractData(args) {
	await Excel.run(async (context) => {
		const names = context.workbook.names;
		const tables = context.workbook.tables;
		const sheets = context.workbook.worksheets;

		context.workbook.load(
			'worksheets/items/name' + 
			', tables/items/rows/items/values' +
			', tables/items/name' +
			', names/items/arrayValues/values' +
			', names/items/name'
		);

		await context.sync();

		// create Roster Data Sheet if it doesn't exist
		createRosterDataSheet(sheets).activate();

		// create rosterData table if it doesn't exist
		const rosterDataTable = createRosterDataTable(tables);
		const tableNames = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];
		let rosterData = [];

		tableNames.forEach(tableName => {
			const rosterTable = tables.items.find(item => item.name === tableName);
			const date = names.items.find(item => item.name === `${tableName}_Date`).arrayValues.values[0][0];
			
			rosterData = rosterData.concat(extractRosterData(rosterTable, date));
		});

		rosterDataTable.rows.add(null, rosterData);

		// format table
		rosterDataTable.getRange().format.autofitColumns();
		rosterDataTable.columns.getItem('Date').getDataBodyRange().numberFormat = 'dd/mm/yyyy';

		args.completed();
	});
}

// Create the table "rosterTable"
function createRosterDataTable(tables) {
	const tableName = 'rosterData';
	let dataTable = tables.items.find(item => item.name === tableName);

	if (dataTable === undefined) {
		dataTable = tables.add(`'Roster Data'!A1:I1`, true);

		dataTable.name = tableName;
		dataTable.getHeaderRowRange().values = 
			[["Name", "Service Point", "Date", "Start", "End", "Time", "OT", "Value", "Address"]];
	} else {
		dataTable.rows.deleteRows(dataTable.rows.items);
	}

	return dataTable;
}

// Create the worksheet "Roster Data"
function createRosterDataSheet(sheets) {
	const wSheetName = 'Roster Data';

	let rosterDataSheet = sheets.items.find(sheet => sheet.name === wSheetName);

	if (rosterDataSheet === undefined)
		rosterDataSheet = sheets.add(wSheetName);

	return rosterDataSheet;
}

// Extract data from a single roster table, i.e Monday
function extractRosterData(table, date) {
	const rows = table.rows.items;
	const rosterData = [];
	const address = '';
	
	const STARTTIMESEMIPHORE = 'from';
	const ENDTIMESEMIPHORE = 'til';

	rows.forEach(row => {
		const servicePoint = row.values[0][0];

		for (let colIndex = 2; colIndex < 14; colIndex++) {
			const cellValue = row.values[0][colIndex];

			if (cellValue !== '') {
				const name = extractName(cellValue);

				// Get start and end times
				const timeArray = getTime(getTimeString(colIndex), cellValue);
				const startTime = timeArray[0];
				const endTime = timeArray[1];

				// Calculate the time (hours)
				const time = endTime > startTime ?
					endTime - startTime :
					endTime + 12 - startTime;

				// Get OT
				const dateObj = excelDateToJSDate(date);
				let ot = '';

				if (extractOT(cellValue) === 'OT' || dateObj.getDay() == 6 || dateObj.getDay() == 7)
					ot = 'OT';

				const value = cellValue;

				rosterData.push([
					name, servicePoint, date, startTime, endTime, time, ot, value, address
				]);
			}
		}
	});

	return rosterData;
}

// Extracts the staff name from the range
function extractName(rangeValue) {
	const parenthesisIndex = rangeValue.indexOf('(');

	if (parenthesisIndex > 0)
		return rangeValue.substring(0, parenthesisIndex - 1).trim();

	return rangeValue.trim();
}

// Gets the start and end time from the timeString or cellValue if time is present
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

// Converts Excel date serial number to Date object
function excelDateToJSDate(serial) {
	const utc_days = Math.floor(serial - 25569);
	const utc_value = utc_days * 86400;
	const date_info = new Date(utc_value * 1000);

	return date_info;
}

// Extracts "OT" from the range value
function extractOT(cellValue) {
	const pattern = /\bOT\b/;

	return cellValue.match !== null ? true : false;
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

// Converts time to a double (e.g 9.30 -> 9.5)
function convertTime(timeString) {
	let timeStringArr = timeString.split('.');

	let hour = timeStringArr[0] / 1;
	let minutes = timeStringArr[1] / 60;

	return hour + minutes;
}

async function extractData2(args) {
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

Office.actions.associate("extractData", extractData);


// Register the function with Office.
Office.actions.associate("action", action);
