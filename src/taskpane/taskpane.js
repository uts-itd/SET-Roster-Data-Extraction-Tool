/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console, document, Excel, Office */

Office.onReady((info) => {
	if (info.host === Office.HostType.Excel) {
		document.querySelector('.author-text>a').addEventListener('click', openLink);
		document.getElementById("app-body").style.display = "flex";
	}
});

function openLink() {
	const url = 'https://github.com/kenyachan/';
	
	window.open(url);
}

async function tryCatch(callback) {
	try {
		await callback();
	} catch (error) {
		// TODO: display error in UI.
		console.error(error);
	}
}
