/**
 * Google Apps Script Web App endpoint for dashboard data.
 * Deploy as Web App (Execute as: Me, Access: Anyone with the link).
 */
function doGet() {
  const sheetName = "Construction Financial Data";
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);

  if (!sheet) {
    return ContentService.createTextOutput(
      JSON.stringify({ error: `Sheet "${sheetName}" not found`, rows: [] })
    ).setMimeType(ContentService.MimeType.JSON);
  }

  const values = sheet.getDataRange().getDisplayValues();
  if (values.length < 2) {
    return ContentService.createTextOutput(JSON.stringify({ rows: [] })).setMimeType(
      ContentService.MimeType.JSON
    );
  }

  const headers = values[0].map((header) => String(header).trim());
  const rows = values.slice(1).reduce((acc, row) => {
    const rowObj = {};
    headers.forEach((header, index) => {
      rowObj[header] = row[index];
    });

    const hasValue = Object.values(rowObj).some((value) => String(value).trim() !== "");
    if (hasValue) acc.push(rowObj);
    return acc;
  }, []);

  return ContentService.createTextOutput(
    JSON.stringify({
      generatedAt: new Date().toISOString(),
      sheetName,
      rows,
    })
  ).setMimeType(ContentService.MimeType.JSON);
}
