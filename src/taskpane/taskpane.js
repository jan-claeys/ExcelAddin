/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console, document, Excel, Office */
import axios from "axios";

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    if (!Office.context.requirements.isSetSupported("ExcelApi", "1.7")) {
      console.log("Sorry. The tutorial add-in uses Excel.js APIs that are not available in your version of Office.");
    }

    // Assign event handlers and other initialization logic.
    document.getElementById("download").onclick = download;
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
  }
});

async function download() {
  await Excel.run(async (context) => {
    const res = await axios.get('http://localhost:4000');
    const data = res.data;
    const keys = Object.keys(data[0]);

    const currentWorksheet = context.workbook.worksheets.getActiveWorksheet();
    const expensesTable = currentWorksheet.tables.add(`A1:${numberToLetters(keys.length - 1)}1`, true /*hasHeaders*/);
    expensesTable.name = "ExpensesTable";

    expensesTable.getHeaderRowRange().values = [keys];
    expensesTable.rows.add(null /*add at the end*/, data.map(Object.values));

    expensesTable.columns.getItemAt(3).getRange().numberFormat = [["\u20AC#,##0.00"]];
    expensesTable.getRange().format.autofitColumns();
    expensesTable.getRange().format.autofitRows();

    await context.sync();
  }).catch(function (error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
      console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
  });
}

//0=A 25=Z
function numberToLetters(num) {
  let letters = ''
  while (num >= 0) {
      letters = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'[num % 26] + letters
      num = Math.floor(num / 26) - 1
  }
  return letters
}
