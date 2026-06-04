const SPREADSHEET_ID = "";
const SHEET_NAME = "\u8a18\u5e33\u660e\u7d30";
const HEADERS = [
  "id",
  "type",
  "amount",
  "category",
  "categoryLabel",
  "subcategory",
  "payment",
  "paymentLabel",
  "date",
  "note",
  "createdAt",
  "updatedAt"
];

function doPost(event) {
  const payload = JSON.parse(event.postData.contents || "{}");
  const sheet = getOrCreateSheet_();

  if (payload.action === "delete") {
    deleteById_(sheet, payload.id);
  } else {
    upsertRow_(sheet, payload);
  }

  return json_({ ok: true, rows: Math.max(sheet.getLastRow() - 1, 0) });
}

function doGet(event) {
  const sheet = getOrCreateSheet_();
  const payload = {
    ok: true,
    spreadsheetId: getSpreadsheet_().getId(),
    sheetName: sheet.getName(),
    rows: readRows_(sheet)
  };
  const callback = event && event.parameter && event.parameter.callback;

  if (callback) {
    return ContentService
      .createTextOutput(callback + "(" + JSON.stringify(payload) + ")")
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }

  return json_(payload);
}

function setup() {
  const sheet = getOrCreateSheet_();
  return {
    spreadsheetId: getSpreadsheet_().getId(),
    sheetName: sheet.getName(),
    rows: sheet.getLastRow()
  };
}

function getOrCreateSheet_() {
  const spreadsheet = getSpreadsheet_();
  let sheet = spreadsheet.getSheetByName(SHEET_NAME);

  if (!sheet) {
    sheet = spreadsheet.insertSheet(SHEET_NAME);
  }

  const firstRow = sheet.getRange(1, 1, 1, HEADERS.length).getValues()[0];
  const needsHeader = firstRow.every(function(cell) {
    return cell === "";
  });

  if (needsHeader) {
    sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
    sheet.setFrozenRows(1);
  }

  return sheet;
}

function getSpreadsheet_() {
  if (SPREADSHEET_ID) {
    return SpreadsheetApp.openById(SPREADSHEET_ID);
  }

  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  if (!spreadsheet) {
    throw new Error("Set SPREADSHEET_ID or create this script from the target Google Sheet.");
  }

  return spreadsheet;
}

function upsertRow_(sheet, payload) {
  const row = HEADERS.map(function(key) {
    return payload[key] == null ? "" : payload[key];
  });
  const rowIndex = findRowById_(sheet, payload.id);

  if (rowIndex > 0) {
    sheet.getRange(rowIndex, 1, 1, HEADERS.length).setValues([row]);
  } else {
    sheet.appendRow(row);
  }
}

function readRows_(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  const values = sheet.getRange(2, 1, lastRow - 1, HEADERS.length).getValues();
  return values
    .filter(function(row) {
      return row[0];
    })
    .map(function(row) {
      const item = {};
      HEADERS.forEach(function(key, index) {
        const value = row[index];
        item[key] = normalizeCellValue_(key, value);
      });
      return item;
    });
}

function normalizeCellValue_(key, value) {
  if (key === "date" && value instanceof Date) {
    return Utilities.formatDate(value, "Asia/Taipei", "yyyy-MM-dd");
  }

  if (value instanceof Date) {
    return value.getTime();
  }

  return value;
}

function deleteById_(sheet, id) {
  const rowIndex = findRowById_(sheet, id);
  if (rowIndex > 1) {
    sheet.deleteRow(rowIndex);
  }
}

function findRowById_(sheet, id) {
  if (!id || sheet.getLastRow() < 2) return -1;

  const values = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues();
  const index = values.findIndex(function(row) {
    return row[0] === id;
  });
  return index >= 0 ? index + 2 : -1;
}

function json_(payload) {
  return ContentService
    .createTextOutput(JSON.stringify(payload))
    .setMimeType(ContentService.MimeType.JSON);
}
