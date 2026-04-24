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

  const headerRowIndex = findHeaderRowIndex(values);
  const headers = values[headerRowIndex].map((header) => String(header).trim());
  const rows = values.slice(headerRowIndex + 1).reduce((acc, row) => {
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
      headerRowIndex: headerRowIndex + 1,
      rows,
    })
  ).setMimeType(ContentService.MimeType.JSON);
}

function findHeaderRowIndex(values) {
  const expectedHeaders = [
    "activity id",
    "activity",
    "planned value",
    "actual cost",
    "earned value",
    "complete",
    "cost variance",
    "budget variance",
    "budget status",
  ];

  let bestIndex = 0;
  let bestScore = -1;

  for (let i = 0; i < values.length; i += 1) {
    const normalized = values[i]
      .map((cell) => String(cell).toLowerCase().replace(/[^a-z0-9]+/g, " ").trim())
      .filter(Boolean);

    if (!normalized.length) continue;

    const score = expectedHeaders.reduce((count, alias) => {
      const found = normalized.some((cell) => cell.includes(alias));
      return count + (found ? 1 : 0);
    }, 0);

    if (score > bestScore) {
      bestScore = score;
      bestIndex = i;
    }
  }

  return bestIndex;
}
