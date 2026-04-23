const fileInput = document.getElementById("fileInput");
const totalPlannedEl = document.getElementById("totalPlanned");
const totalActualEl = document.getElementById("totalActual");
const totalCvEl = document.getElementById("totalCv");
const projectStatusEl = document.getElementById("projectStatus");
const statusCardEl = document.getElementById("statusCard");
const messageEl = document.getElementById("message");
const tableBodyEl = document.getElementById("activityTableBody");

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

const formatPercent = (value) => `${parseNumber(value).toFixed(1)}%`;

const normalizeHeader = (value) =>
  String(value)
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, " ")
    .replace(/\s+/g, " ")
    .trim();

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
  for (const key of Object.keys(row)) {
    const normalizedKey = normalizeHeader(key);
    if (normalizedAliases.includes(normalizedKey)) {
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
      const activityId = normalize(getCell(row, ["activity id", "id", "activityid"]), "");
      if (!isValidActivityId(activityId)) return null;

      const plannedCost = parseNumber(
        getCell(row, [
          "planned cost",
          "planned value",
          "planned value (pv)",
          "planned",
          "pv",
          "plannedcost",
          "budgeted cost",
        ])
      );
      const actualCost = parseNumber(
        getCell(row, ["actual cost", "actual cost (ac)", "actual", "ac", "actualcost"])
      );
      const computedCv = plannedCost - actualCost;

      return {
        projectId: normalize(getCell(row, ["project id", "projectid", "project"]), "Unspecified"),
        activityId,
        activity: normalize(getCell(row, ["activity", "activity name"]), "Unspecified"),
        plannedCost,
        actualCost,
        ev: parseNumber(getCell(row, ["earned value (ev)", "earned value", "ev"])),
        percentComplete: parseNumber(getCell(row, ["% complete", "percent complete", "complete %"])),
        cv: computedCv,
        costUsedPercent: plannedCost ? (actualCost / plannedCost) * 100 : 0,
        budgetVariancePercent: plannedCost ? (computedCv / plannedCost) * 100 : 0,
        budgetStatus: computedCv >= 0 ? "On Budget" : "Over Budget",
      };
    })
    .filter(Boolean);

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
      '<tr><td colspan="10" class="placeholder">No valid rows found in Dashboard sheet.</td></tr>';
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

const processWorkbook = (arrayBuffer) => {
  const workbook = XLSX.read(arrayBuffer, { type: "array" });
  const preferredSheetName = workbook.SheetNames.find(
    (name) => name.trim().toLowerCase() === "dashboard"
  );
  const sheet = workbook.Sheets[preferredSheetName || "Dashboard"];

  if (!sheet) {
    showMessage('Sheet "Dashboard" not found. Please upload the correct file.', true);
    return;
  }

  const headerRowIndex = findHeaderRowIndex(sheet);
  const rawRows = XLSX.utils.sheet_to_json(sheet, {
    defval: "",
    range: headerRowIndex,
  });
  const { rows, totals, totalSource } = extractMetrics(rawRows);

  renderKpis(totals);
  renderTable(rows);

  const sourceLabel =
    totalSource === "activity-id-rows"
      ? "activity rows with valid Activity ID"
      : "sum of activity rows";

  const dependencySuffix = chartDependencyWarning ? ` ${chartDependencyWarning}` : "";
  const headerMessage =
    headerRowIndex > 0 ? ` Detected headers on row ${headerRowIndex + 1}.` : "";
  showMessage(
    `Loaded ${rows.length} activity row(s). KPI totals source: ${sourceLabel}.${headerMessage}${dependencySuffix}`
  );
};

fileInput.addEventListener("change", async (event) => {
  const file = event.target.files?.[0];
  if (!file) return;

  try {
    const buffer = await file.arrayBuffer();
    processWorkbook(buffer);
  } catch (error) {
    const isChartReferenceError =
      error instanceof ReferenceError && /chart is not defined/i.test(error.message);
    const errorText = isChartReferenceError
      ? "Error reading file: Chart.js failed to load (Chart is not defined). Please refresh and try again."
      : `Error reading file: ${error.message}`;
    showMessage(errorText, true);
  }
});
