const SHEET_NAME = "Students";

function doGet(e) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const students = [];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (row[0] === "") continue;
    const student = {};
    headers.forEach((h, j) => student[h] = row[j]);
    students.push(student);
  }
  return ContentService
    .createTextOutput(JSON.stringify(students))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  const body = JSON.parse(e.postData.contents);
  const action = body.action;

  // SAVE - புது student
  if (action === "save") {
    const id = Date.now().toString();
    sheet.appendRow([
      id,
      body.name || "",
      body.room || "",
      body.contact || "",
      body.parent || "",
      body.aadhaar || "",
      body.aadhaarStatus || "Pending",
      body.totalFee || 0,
      body.paid || 0,
      body.balance || 0,
      body.status || "PENDING"
    ]);
    return ok("Student saved!");
  }

  // UPDATE FEE - existing student
  if (action === "updateFee") {
    const id = body.id.toString();
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0].toString() === id) {
        const total   = Number(body.totalFee) || 0;
        const paid    = Number(body.paid) || 0;
        const balance = total - paid;
        sheet.getRange(i + 1, 8).setValue(total);
        sheet.getRange(i + 1, 9).setValue(paid);
        sheet.getRange(i + 1, 10).setValue(balance);
        sheet.getRange(i + 1, 11).setValue(balance <= 0 ? "PAID" : "PENDING");
        return ok("Fee updated!");
      }
    }
    return ok("Student not found!");
  }

  // DELETE
  if (action === "delete") {
    const id = body.id.toString();
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0].toString() === id) {
        sheet.deleteRow(i + 1);
        return ok("Deleted!");
      }
    }
    return ok("Not found!");
  }

  return ok("Unknown action");
}

function ok(msg) {
  return ContentService
    .createTextOutput(JSON.stringify({ message: msg }))
    .setMimeType(ContentService.MimeType.JSON);
}