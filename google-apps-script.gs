const SHEET_NAME = "Respuestas";

function getSheet_() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = spreadsheet.getSheetByName(SHEET_NAME) || spreadsheet.getSheets()[0];
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(["ID", "Fecha"]);
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
  const headers = values[0] || [];
  const rows = values.slice(1).filter(row => row[0]).map(row => {
    if (headers[2] === "Respuestas" && row[2]) {
      return {
        id: row[0],
        fecha: row[1] instanceof Date ? row[1].toISOString() : row[1],
        respuestas: JSON.parse(row[2] || "{}")
      };
    }
    const respuestas = {};
    headers.slice(2).forEach((header, index) => {
      if (header) respuestas[header] = row[index + 2] || "";
    });
    return {
      id: row[0],
      fecha: row[1] instanceof Date ? row[1].toISOString() : row[1],
      respuestas
    };
  });
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

  const respuestas = payload.respuestas || {};
  const currentHeaders = sheet.getRange(1, 1, 1, Math.max(sheet.getLastColumn(), 2)).getValues()[0];
  const newFields = Object.keys(respuestas).filter(field => !currentHeaders.includes(field));
  if (newFields.length) {
    sheet.getRange(1, currentHeaders.length + 1, 1, newFields.length).setValues([newFields]);
  }
  const headers = currentHeaders.concat(newFields);
  const id = Utilities.getUuid();
  const row = [id, new Date(), ...headers.slice(2).map(field => {
    const value = respuestas[field];
    return Array.isArray(value) ? value.join(", ") : value || "";
  })];
  sheet.getRange(sheet.getLastRow() + 1, 1, 1, row.length).setValues([row]);
  return jsonResponse_({ok: true, id: id});
}
