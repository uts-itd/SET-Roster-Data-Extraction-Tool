/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console, document, Excel, Office */

Office.onReady((info) => {
	if (info.host === Office.HostType.Excel) {
		document.querySelector("#create-extraction-sheet").addEventListener(
			'click', () => tryCatch(createDataTable)
		);

		document.getElementById("app-body").style.display = "flex";
	}
});

async function createDataTable() {
	await Excel.run(async (context) => {
		const srdetSheet = createWorksheet(context);

		createRosterExtractionTable(srdetSheet);
		srdetSheet.activate();

		await context.sync();
	});

	function createWorksheet(context) {
		const wSheetName = 'SRDET';
		const worksheet = context.workbook.worksheets.add(wSheetName);

		return worksheet;
	}

	function createRosterExtractionTable(worksheet) {
		const dataTable = worksheet.tables.add("A1:H2", true);
		dataTable.name = "srdet";

		dataTable.getHeaderRowRange().values =		
			[["Name","Date","Start","End","Time","OT","Value","Address"]];
			
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
