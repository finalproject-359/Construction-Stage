const totalPlannedEl = document.getElementById("totalPlanned");
const totalActualEl = document.getElementById("totalActual");
const totalCvEl = document.getElementById("totalCv");
const projectStatusEl = document.getElementById("projectStatus");
const statusCardEl = document.getElementById("statusCard");
const messageEl = document.getElementById("message");
const tableBodyEl = document.getElementById("activityTableBody");

const DATA_SOURCE_URL =
  "https://script.google.com/macros/s/AKfycbxaaigY2kno4qhfMVbt2nYSG2bO4T7475KAwxIJeZHAi_nyJ7_pqHq7UzzVgb8kXm79SA/exec";

const chartDependencyWarning =
  typeof window.Chart === "undefined"
    ? "Chart.js is not available. Charts will be skipped until the dependency loads."
    : "";

const formatCurrency = (value) =>
  new Intl.NumberFormat("en-PH", {
    style: "currency",
    currency: "PHP",
    minimumFractionDigits: 2,
    maximumFractionDigits: 2,
  }).format(Number.isFinite(value) ? value : 0);

const formatCompactCurrency = (value) => {
  const numericValue = Number.isFinite(value) ? value : 0;
  const absoluteValue = Math.abs(numericValue);

  if (absoluteValue < 1000) {
    return formatCurrency(numericValue);
  }

  return new Intl.NumberFormat("en-PH", {
    style: "currency",
    currency: "PHP",
    notation: "compact",
    maximumFractionDigits: 1,
  }).format(numericValue);
};

const parseNumber = (value) => {
  if (value === null || value === undefined || value === "") return 0;
  if (typeof value === "number") return Number.isFinite(value) ? value : 0;
  const stringValue = String(value).trim();
  const isAccountingNegative = /^\(.*\)$/.test(stringValue);
  const cleaned = stringValue.replace(/[^0-9.-]/g, "");
  const parsed = Number(cleaned);
  if (!Number.isFinite(parsed)) return 0;
  return isAccountingNegative ? -Math.abs(parsed) : parsed;
};

const normalize = (value, fallback = "N/A") => {
  if (value === null || value === undefined || value === "") return fallback;
  return String(value).trim() || fallback;
};

const formatPercent = (value) => `${parseNumber(value).toFixed(2)}%`;

const normalizeHeader = (value) =>
  String(value)
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, " ")
    .replace(/\s+/g, " ")
    .trim();

const compactHeader = (value) => normalizeHeader(value).replace(/\s+/g, "");

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

const getCell = (row, aliases) => {
  const normalizedAliases = aliases.map((alias) => normalizeHeader(alias));
  const compactAliases = normalizedAliases.map((alias) => alias.replace(/\s+/g, ""));
  for (const key of Object.keys(row)) {
    const normalizedKey = normalizeHeader(key);
    const normalizedKeyCompact = compactHeader(key);
    const matched = normalizedAliases.some((alias, index) => {
      const aliasCompact = compactAliases[index];
      const isShortAlias = aliasCompact.length <= 2;

      if (isShortAlias) {
        return normalizedKeyCompact === aliasCompact;
      }

      return (
        normalizedKey === alias ||
        normalizedKey.includes(alias) ||
        alias.includes(normalizedKey) ||
        normalizedKeyCompact === aliasCompact ||
        normalizedKeyCompact.includes(aliasCompact) ||
        aliasCompact.includes(normalizedKeyCompact)
      );
    });
    if (matched) {
      return row[key];
    }
  }
  return null;
};

const isSummaryLabel = (value) => {
  const text = normalize(value, "").toLowerCase();
  return (
    text.includes("total") ||
    text.includes("summary") ||
    text.includes("grand total") ||
    text.includes("overall")
  );
};

const isValidActivityId = (value) => {
  if (value === null || value === undefined) return false;
  const normalizedValue = String(value).trim();
  if (!normalizedValue) return false;
  return !isSummaryLabel(normalizedValue);
};

const findHeaderRowIndex = (sheet) => {
  const rows = XLSX.utils.sheet_to_json(sheet, {
    header: 1,
    defval: "",
    blankrows: false,
    raw: false,
  });

  if (!rows.length) return 0;

  let bestIndex = 0;
  let bestScore = 0;

  rows.forEach((row, index) => {
    const normalizedCells = row.map((cell) => normalizeHeader(cell)).filter(Boolean);
    if (!normalizedCells.length) return;

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

const extractDashboardRows = (rawRows) =>
  rawRows
    .map((row) => {
      const activityId = normalize(
        getCell(row, ["activity id", "id", "activityid", "activity code", "wbs", "task id"]),
        ""
      );
      if (!isValidActivityId(activityId)) return null;

      const plannedCost = parseNumber(
        getCell(row, [
          "planned cost",
          "planned value",
          "planned value (pv)",
          "planned value pv",
          "planned",
          "pv",
          "plannedcost",
          "budget",
          "budget value",
          "budgeted cost",
        ])
      );
      const actualCost = parseNumber(
        getCell(row, [
          "actual cost",
          "actual cost (ac)",
          "actual cost ac",
          "actual",
          "ac",
          "actualcost",
        ])
      );
      const providedCv = parseNumber(
        getCell(row, ["cost variance", "cv", "cost variance (cv)"])
      );
      const computedCv = plannedCost - actualCost;
      const cv = providedCv !== 0 ? providedCv : computedCv;
      const providedCostUsed = parseNumber(
        getCell(row, ["% cost used", "cost used %", "cost used percent"])
      );
      const providedBudgetVariance = parseNumber(
        getCell(row, ["budget variance", "budget variance %", "budget variance percent"])
      );
      const providedBudgetStatus = normalize(
        getCell(row, ["budget status", "status"]),
        ""
      );

      return {
        projectId: normalize(getCell(row, ["project id", "projectid", "project"]), "Unspecified"),
        activityId,
        activity: normalize(getCell(row, ["activity", "activity name"]), "Unspecified"),
        plannedCost,
        actualCost,
        ev: parseNumber(
          getCell(row, ["earned value (ev)", "earned value ev", "earned value", "ev"])
        ),
        percentComplete: parseNumber(
          getCell(row, [
            "% complete",
            "percent complete",
            "complete %",
            "completion %",
            "progress %",
          ])
        ),
        cv,
        costUsedPercent: providedCostUsed || (plannedCost ? (actualCost / plannedCost) * 100 : 0),
        budgetVariancePercent:
          providedBudgetVariance || (plannedCost ? (cv / plannedCost) * 100 : 0),
        budgetStatus: providedBudgetStatus || (cv >= 0 ? "On Budget" : "Over Budget"),
      };
    })
    .filter(
      (row) =>
        row &&
        (row.plannedCost !== 0 ||
          row.actualCost !== 0 ||
          row.ev !== 0 ||
          row.percentComplete !== 0)
    );

const calculateTotalsFromRows = (rows) =>
  rows.reduce(
    (acc, row) => {
      acc.planned += row.plannedCost;
      acc.actual += row.actualCost;
      acc.cv += row.cv;
      return acc;
    },
    { planned: 0, actual: 0, cv: 0 }
  );

const extractMetrics = (rawRows) => {
  const rows = extractDashboardRows(rawRows);
  return {
    rows,
    totals: calculateTotalsFromRows(rows),
    totalSource: "activity-id-rows",
  };
};

const renderKpis = (totals) => {
  const safeTotals = {
    planned: parseNumber(totals?.planned),
    actual: parseNumber(totals?.actual),
    cv: parseNumber(totals?.cv),
  };

  totalPlannedEl.textContent = formatCurrency(safeTotals.planned);
  totalActualEl.textContent = formatCurrency(safeTotals.actual);
  totalCvEl.textContent = formatCurrency(safeTotals.cv);

  statusCardEl.classList.remove("status-under", "status-over");

  if (safeTotals.actual < safeTotals.planned) {
    projectStatusEl.textContent = "Under Budget";
    statusCardEl.classList.add("status-under");
  } else if (safeTotals.actual > safeTotals.planned) {
    projectStatusEl.textContent = "Over Budget";
    statusCardEl.classList.add("status-over");
  } else {
    projectStatusEl.textContent = "On Budget";
  }
};

const renderTable = (rows) => {
  if (!rows.length) {
    tableBodyEl.innerHTML =
      '<tr><td colspan="10" class="placeholder">No valid rows found in Construction Financial Data sheet.</td></tr>';
    return;
  }

  tableBodyEl.innerHTML = rows
    .map(
      (row) => `
      <tr>
        <td>${row.activityId}</td>
        <td>${row.activity}</td>
        <td>${formatCurrency(row.plannedCost)}</td>
        <td>${formatCurrency(row.actualCost)}</td>
        <td>${formatCurrency(row.ev)}</td>
        <td>${formatPercent(row.percentComplete)}</td>
        <td>${formatCurrency(row.cv)}</td>
        <td>${formatPercent(row.costUsedPercent)}</td>
        <td>${formatPercent(row.budgetVariancePercent)}</td>
        <td>${row.budgetStatus}</td>
      </tr>
    `
    )
    .join("");
};

const showMessage = (text, isError = false) => {
  messageEl.textContent = text;
  messageEl.style.color = isError ? "#dc2626" : "#6b7280";
};

const processWorksheet = (sheet, sourceName = "workbook") => {
  if (!sheet) {
    showMessage('Sheet data not found. Please load a valid worksheet.', true);
    return;
  }

  const headerRowIndex = findHeaderRowIndex(sheet);
  const rawRows = XLSX.utils.sheet_to_json(sheet, {
    raw: false,
    defval: "",
    range: headerRowIndex,
  });

  if (!rawRows.length) {
    renderKpis({ planned: 0, actual: 0, cv: 0 });
    renderTable([]);
    showMessage(`No rows detected from ${sourceName}. Check if the sheet has headers and values.`, true);
    return;
  }

  const { rows, totals, totalSource } = extractMetrics(rawRows);

  renderKpis(totals);
  renderTable(rows);

  const sourceLabel =
    totalSource === "activity-id-rows"
      ? "activity rows with valid Activity ID"
      : "sum of activity rows";

  const dependencySuffix = chartDependencyWarning ? ` ${chartDependencyWarning}` : "";
  const headerMessage =
    headerRowIndex > 0 ? ` Header row auto-detected at spreadsheet row ${headerRowIndex + 1}.` : "";

  showMessage(
    `Loaded ${rows.length} activity row(s) from ${sourceName}. KPI totals source: ${sourceLabel}.${headerMessage}${dependencySuffix}`
  );
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
    return (
      parsed.hostname === "script.google.com" &&
      /\/macros\/s\/.+\/exec$/.test(parsed.pathname)
    );
  } catch {
    return false;
  }
};

const processRows = (rawRows, sourceName = "web app") => {
  if (!Array.isArray(rawRows) || !rawRows.length) {
    renderKpis({ planned: 0, actual: 0, cv: 0 });
    renderTable([]);
    showMessage(`No rows detected from ${sourceName}. Check if the sheet has headers and values.`, true);
    return;
  }

  const { rows, totals, totalSource } = extractMetrics(rawRows);

  renderKpis(totals);
  renderTable(rows);

  const sourceLabel =
    totalSource === "activity-id-rows"
      ? "activity rows with valid Activity ID"
      : "sum of activity rows";

  const dependencySuffix = chartDependencyWarning ? ` ${chartDependencyWarning}` : "";

  showMessage(
    `Loaded ${rows.length} activity row(s) from ${sourceName}. KPI totals source: ${sourceLabel}.${dependencySuffix}`
  );
};

const loadGoogleSheet = async (providedUrl = "") => {
  const rawUrl = providedUrl || DATA_SOURCE_URL;
  const trimmedUrl = rawUrl.trim();
  const isWebAppSource = isAppsScriptWebAppUrl(trimmedUrl);
  const csvUrl = isWebAppSource ? "" : toGoogleSheetCsvUrl(trimmedUrl);

  if (!isWebAppSource && !csvUrl) {
    showMessage(
      "Invalid URL. Paste a valid Google Sheet link or Apps Script Web App URL and try again.",
      true
    );
    return;
  }

  try {
    showMessage("Loading data source...");

    if (isWebAppSource) {
      const response = await fetch(trimmedUrl);
      if (!response.ok) {
        throw new Error(`Unable to fetch Apps Script Web App (HTTP ${response.status})`);
      }

      const payload = await response.json();
      if (payload?.error) {
        throw new Error(payload.error);
      }

      processRows(payload?.rows || [], `Apps Script Web App (${payload?.sheetName || "unknown sheet"})`);
      return;
    }

    const response = await fetch(csvUrl);
    if (!response.ok) {
      throw new Error(`Unable to fetch Google Sheet (HTTP ${response.status})`);
    }

    const csvText = await response.text();
    const workbook = XLSX.read(csvText, { type: "string" });
    const firstSheetName = workbook.SheetNames[0];
    const sheet = workbook.Sheets[firstSheetName];

    processWorksheet(sheet, `Google Sheet "${firstSheetName}"`);
  } catch (error) {
    showMessage(
      `Error loading data source: ${error.message}. Ensure the sheet is shared or the Web App is deployed for access.`,
      true
    );
  }
};

const initializeDataSource = () => {
  if (DATA_SOURCE_URL) {
    loadGoogleSheet(DATA_SOURCE_URL);
  }
};

initializeDataSource();
