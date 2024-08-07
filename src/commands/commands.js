const SRDET = require('./srdet');

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
async function extractData(event) {
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

		event.completed();
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

		for (let colIndex = 3; colIndex < 14; colIndex++) {
			const cellValue = row.values[0][colIndex];

			if (cellValue !== '') {
				const name = SRDET.extractName(cellValue);

				// Get start and end times
				const timeArray = SRDET.getTime(SRDET.getTimeString(colIndex), cellValue);
				const startTime = timeArray[0];
				const endTime = timeArray[1];

				// Calculate the time (hours)
				const time = endTime > startTime ?
					endTime - startTime :
					endTime + 12 - startTime;

				// Get OT
				const dateObj = SRDET.excelDateToJSDate(date);
				let ot = '';

				if (SRDET.extractOT(cellValue) === 'OT' || dateObj.getDay() == 6 || dateObj.getDay() == 7)
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
