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
  if (!values.length) {
    return ContentService.createTextOutput(JSON.stringify({ rows: [] })).setMimeType(
      ContentService.MimeType.JSON
    );
  }

  const expectedHeaderAliases = [
    "activity id",
    "activity",
    "planned value",
    "actual cost",
    "earned value",
    "complete",
    "cost variance",
    "budget",
  ];

  const normalizeHeader = (value) =>
    String(value)
      .toLowerCase()
      .replace(/[^a-z0-9]+/g, " ")
      .replace(/\s+/g, " ")
      .trim();

  const findHeaderRowIndex = (rows) => {
    let bestIndex = 0;
    let bestScore = 0;

    rows.forEach((row, index) => {
      const normalizedCells = row.map((cell) => normalizeHeader(cell)).filter(Boolean);
      if (!normalizedCells.length) return;

      const score = expectedHeaderAliases.reduce((count, alias) => {
        const hasAlias = normalizedCells.some((cell) => cell.includes(alias));
        return count + (hasAlias ? 1 : 0);
      }, 0);

      if (score > bestScore) {
        bestScore = score;
        bestIndex = index;
      }
    });

    return bestIndex;
  };

  const headerRowIndex = findHeaderRowIndex(values);
  const headers = values[headerRowIndex].map((header) => String(header).trim());

  if (!headers.some((header) => header)) {
    return ContentService.createTextOutput(
      JSON.stringify({ error: "No header row detected.", rows: [] })
    ).setMimeType(ContentService.MimeType.JSON);
  }

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
      headerRowIndex,
      rows,
    })
  ).setMimeType(ContentService.MimeType.JSON);
}
