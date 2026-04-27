(function attachDataBridge(global) {
  const DEFAULT_DATA_SOURCE_URL =
    "https://script.google.com/macros/s/AKfycbw7PFHT-JOj2gbEjXqg9MlmZU7YvER8P1BrZEpH9uPDFivnwOBnUw6VVDO16NsUL6ks/exec";

  const EXPECTED_HEADER_ALIASES = [
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

  const extractRowsFromPayload = (payload) => {
    if (Array.isArray(payload)) return payload;
    if (!payload || typeof payload !== "object") return [];
    if (Array.isArray(payload.rows)) return payload.rows;
    if (Array.isArray(payload.data)) return payload.data;
    if (Array.isArray(payload.items)) return payload.items;
    return [];
  };

  const findHeaderRowIndex = (sheet) => {
    const rows = XLSX.utils.sheet_to_json(sheet, {
      header: 1,
      defval: "",
      blankrows: false,
      raw: false,
    });

    let bestIndex = 0;
    let bestScore = 0;

    rows.forEach((row, index) => {
      const normalizedCells = row.map((cell) => normalizeHeader(cell)).filter(Boolean);
      const score = EXPECTED_HEADER_ALIASES.reduce((count, alias) => {
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

  const toGoogleSheetCsvUrl = (inputUrl) => {
    if (!inputUrl) return "";
    const trimmed = inputUrl.trim();
    if (!trimmed) return "";
    if (/output=csv/i.test(trimmed)) return trimmed;

    try {
      const parsed = new URL(trimmed);
      const match = parsed.pathname.match(/\/spreadsheets\/d\/([a-zA-Z0-9-_]+)/);
      if (!match) return "";
      const sheetId = match[1];
      const gid = parsed.searchParams.get("gid") || "0";
      return `https://docs.google.com/spreadsheets/d/${sheetId}/gviz/tq?tqx=out:csv&gid=${gid}`;
    } catch {
      return "";
    }
  };

  const isAppsScriptWebAppUrl = (inputUrl) => {
    if (!inputUrl) return false;
    try {
      const parsed = new URL(inputUrl.trim());
      return parsed.hostname === "script.google.com" && /\/macros\/s\/.+\/exec$/.test(parsed.pathname);
    } catch {
      return false;
    }
  };

  const fetchRowsFromSource = async (providedUrl = "") => {
    const rawUrl = providedUrl || DEFAULT_DATA_SOURCE_URL;
    const trimmedUrl = rawUrl.trim();
    const isWebAppSource = isAppsScriptWebAppUrl(trimmedUrl);
    const csvUrl = isWebAppSource ? "" : toGoogleSheetCsvUrl(trimmedUrl);

    if (!isWebAppSource && !csvUrl) {
      throw new Error("Invalid URL. Use a valid Google Sheet or Apps Script Web App URL.");
    }

    if (isWebAppSource) {
      const response = await fetch(trimmedUrl, { cache: "no-store" });
      if (!response.ok) throw new Error(`Unable to fetch Apps Script Web App (HTTP ${response.status})`);

      const contentType = response.headers.get("content-type") || "";
      if (contentType.toLowerCase().includes("application/json")) {
        const payload = await response.json();
        if (payload?.error) throw new Error(payload.error);
        return {
          rows: extractRowsFromPayload(payload),
          sourceName: `Apps Script Web App (${payload?.sheetName || "sheet"})`,
        };
      }

      const rawText = await response.text();
      const workbook = XLSX.read(rawText, { type: "string" });
      const firstSheetName = workbook.SheetNames[0];
      const sheet = workbook.Sheets[firstSheetName];
      const headerRowIndex = findHeaderRowIndex(sheet);
      const rows = XLSX.utils.sheet_to_json(sheet, { raw: false, defval: "", range: headerRowIndex });
      return {
        rows,
        sourceName: `Apps Script Web App CSV "${firstSheetName}"`,
      };
    }

    const response = await fetch(csvUrl, { cache: "no-store" });
    if (!response.ok) throw new Error(`Unable to fetch Google Sheet (HTTP ${response.status})`);

    const csvText = await response.text();
    const workbook = XLSX.read(csvText, { type: "string" });
    const firstSheetName = workbook.SheetNames[0];
    const sheet = workbook.Sheets[firstSheetName];
    const headerRowIndex = findHeaderRowIndex(sheet);
    const rows = XLSX.utils.sheet_to_json(sheet, { raw: false, defval: "", range: headerRowIndex });

    return {
      rows,
      sourceName: `Google Sheet "${firstSheetName}"`,
    };
  };

  global.DataBridge = {
    DEFAULT_DATA_SOURCE_URL,
    fetchRowsFromSource,
  };
})(window);
