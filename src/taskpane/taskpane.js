/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console, document, Excel, Office */
import axios from "axios";

const baseUrl = API_URL;
let newValues = [];

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    if (!Office.context.requirements.isSetSupported("ExcelApi", "1.7")) {
      console.log("Sorry. The tutorial add-in uses Excel.js APIs that are not available in your version of Office.");
    }

    if (localStorage.getItem("newValues")) {
      newValues = JSON.parse(localStorage.getItem("newValues"));
    }
    // Assign event handlers and other initialization logic.
    document.getElementById("download").onclick = download;

    loadEntities();
    document.getElementById("entities").onchange = entitiesOnChange;
  }
});

async function download() {
  await Excel.run(async (context) => {


    const sheet = context.workbook.worksheets.getActiveWorksheet();

    clearSheet(sheet, context);

    const tableId = document.getElementById("tables").value;

    localStorage.setItem("tableId", tableId);

    try {
      const res = await axios.get(baseUrl + `/tables/${tableId}`);
      const data = res.data;
      const keys = Object.keys(data[0]);

      const table = sheet.tables.add(`A1:${indexToLetters(keys.length - 1)}1`, true /*hasHeaders*/ );
      table.name = "table";

      table.getHeaderRowRange().values = [keys];
      table.rows.add(null /*add at the end*/ , data.map(Object.values));

      table.columns.getItemAt(3).getRange().numberFormat = [
        ["\u20AC#,##0.00"]
      ];
      table.getRange().format.autofitColumns();
      table.getRange().format.autofitRows();

      table.onChanged.add(onTableChanged);

      await context.sync();
    } catch (error) {
      throw error;
    }
  }).catch(function (error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
      console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
  });
}

async function onTableChanged(eventArgs) {
  await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const table = sheet.tables.getItem("table");

    const range = eventArgs.getRange(context);

    range.format.fill.color = "#ffff00";

    const row = table.rows.getItemAt(parseInt(eventArgs.address.charAt(1)) - 2).load("values");
    let headers = table.getHeaderRowRange().load("values");

    const columnIndex = letterToIndex(eventArgs.address.charAt(0));

    await context.sync();
    const rowValues = row.values;
    const headersValue = headers.values;

    newValues.push({
      code: rowValues[0][0],
      table: localStorage.getItem("tableId"),
      value: eventArgs.details.valueAfter,
      column: headersValue[0][columnIndex - 1].replace(" (u)", "")
    });

    localStorage.setItem("newValues", JSON.stringify(newValues));

    await context.sync();
  });
}

function clearSheet(sheet, context) {
  const range = sheet.getRange("A1:KK20000");
  range.delete(Excel.DeleteShiftDirection.up);

  localStorage.removeItem("newValues");
  newValues = [];
  context.sync();
}

//0=A 25=Z
function indexToLetters(num) {
  let letters = "";
  while (num >= 0) {
    letters = "ABCDEFGHIJKLMNOPQRSTUVWXYZ" [num % 26] + letters;
    num = Math.floor(num / 26) - 1;
  }
  return letters;
}

// A=1 Z=26
function letterToIndex(letter) {
  let index = 0;
  for (let i = 0; i < letter.length; i++) {
    let value = letter.charCodeAt(i) - 64;

    index = index * 26 + value;
  }

  return index;
}

function loadEntities() {
  const select = document.getElementById("entities");
  axios
    .get(baseUrl + "/entities")
    .then((res) => {
      const entities = res.data;
      entities.forEach((entity) => select.add(new Option(entity.Entity, entity.Entity)));
      select.dispatchEvent(new Event("change"));
    })
    .catch(console.log);
}

function entitiesOnChange(e) {
  const select = document.getElementById("tables");
  select.innerHTML = "";
  axios
    .get(baseUrl + `/entities/${e.target.value}/tables`)
    .then((res) => {
      const tables = res.data;
      tables.forEach((table) => select.add(new Option(table.Table, table.ID)));
    })
    .catch(console.log);
}