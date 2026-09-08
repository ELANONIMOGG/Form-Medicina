const SHEET_NAME = "Respuestas";

function getSheet_() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = spreadsheet.getSheetByName(SHEET_NAME);
  if (!sheet) sheet = spreadsheet.insertSheet(SHEET_NAME);
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(["ID", "Fecha", "Respuestas"]);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function jsonResponse_(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

function doGet() {
  const sheet = getSheet_();
  const values = sheet.getDataRange().getValues();
  const rows = values.slice(1).filter(row => row[0]).map(row => ({
    id: row[0],
    fecha: row[1] instanceof Date ? row[1].toISOString() : row[1],
    respuestas: JSON.parse(row[2] || "{}")
  }));
  return jsonResponse_(rows);
}

function doPost(event) {
  const payload = JSON.parse(event.postData.contents || "{}");
  const sheet = getSheet_();

  if (payload.action === "delete") {
    const values = sheet.getRange(2, 1, Math.max(sheet.getLastRow() - 1, 1), 1).getValues();
    const rowIndex = values.findIndex(row => String(row[0]) === String(payload.id));
    if (rowIndex >= 0) sheet.deleteRow(rowIndex + 2);
    return jsonResponse_({ok: true});
  }

  const id = Utilities.getUuid();
  sheet.appendRow([id, new Date(), JSON.stringify(payload.respuestas || {})]);
  return jsonResponse_({ok: true, id: id});
}
