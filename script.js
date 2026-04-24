const totalPlannedEl = document.getElementById("totalPlanned");
const totalActualEl = document.getElementById("totalActual");
const totalCvEl = document.getElementById("totalCv");
const projectStatusEl = document.getElementById("projectStatus");
const statusCardEl = document.getElementById("statusCard");
const physicalProgressEl = document.getElementById("physicalProgress");
const costSpentEl = document.getElementById("costSpent");
const efficiencyGapEl = document.getElementById("efficiencyGap");
const efficiencyCardEl = document.getElementById("efficiencyCard");
const messageEl = document.getElementById("message");
const tableBodyEl = document.getElementById("activityTableBody");
const overrunTableBodyEl = document.getElementById("overrunTableBody");

const DATA_SOURCE_URL =
  "https://script.google.com/macros/s/AKfycbxaaigY2kno4qhfMVbt2nYSG2bO4T7475KAwxIJeZHAi_nyJ7_pqHq7UzzVgb8kXm79SA/exec";

const chartDependencyWarning =
  typeof window.Chart === "undefined"
    ? "Chart.js is not available. Graphs are disabled."
    : "";

let dashboardRows = [];
let varianceChart = null;
let costChart = null;
let evmTrendChart = null;
let efficiencyScatterChart = null;

const getNiceStep = (rawStep) => {
  if (!Number.isFinite(rawStep) || rawStep <= 0) return 1;
  const exponent = Math.floor(Math.log10(rawStep));
  const magnitude = 10 ** exponent;
  const normalized = rawStep / magnitude;

  if (normalized <= 1) return 1 * magnitude;
  if (normalized <= 2) return 2 * magnitude;
  if (normalized <= 5) return 5 * magnitude;
  return 10 * magnitude;
};

const buildLinearAxisRange = (values, { includeZero = true, targetTickCount = 6 } = {}) => {
  const numericValues = values.filter(Number.isFinite);
  if (!numericValues.length) {
    return { min: 0, max: 1, stepSize: 1 };
  }

  let minValue = Math.min(...numericValues);
  let maxValue = Math.max(...numericValues);

  if (includeZero) {
    minValue = Math.min(0, minValue);
    maxValue = Math.max(0, maxValue);
  }

  if (minValue === maxValue) {
    const baseline = minValue === 0 ? 1 : Math.abs(minValue);
    const padding = Math.max(baseline * 0.2, 1);
    minValue -= padding;
    maxValue += padding;
  }

  const rawStep = (maxValue - minValue) / Math.max(targetTickCount - 1, 1);
  const stepSize = getNiceStep(rawStep);
  const min = Math.floor(minValue / stepSize) * stepSize;
  const max = Math.ceil(maxValue / stepSize) * stepSize;

  return { min, max, stepSize };
};

const formatCurrency = (value) =>
  new Intl.NumberFormat("en-PH", {
    style: "currency",
    currency: "PHP",
    minimumFractionDigits: 2,
    maximumFractionDigits: 2,
  }).format(Number.isFinite(value) ? value : 0);

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

const formatSignedPercent = (value) => {
  const numericValue = parseNumber(value);
  const sign = numericValue > 0 ? "+" : "";
  return `${sign}${numericValue.toFixed(2)}%`;
};

const getVarianceBand = (cv, plannedCost) => {
  if (!plannedCost) return "neutral";
  const variancePercent = (parseNumber(cv) / parseNumber(plannedCost)) * 100;
  if (variancePercent < -5) return "severe-over";
  if (variancePercent < 0) return "mild-over";
  return "under";
};

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
  const keyEntries = Object.keys(row).map((key) => ({
    key,
    normalized: normalizeHeader(key),
    compact: compactHeader(key),
  }));

  for (const alias of aliases) {
    const normalizedAlias = normalizeHeader(alias);
    const compactAlias = normalizedAlias.replace(/\s+/g, "");

    const exactMatch = keyEntries.find(
      (entry) => entry.normalized === normalizedAlias || entry.compact === compactAlias
    );
    if (exactMatch) return row[exactMatch.key];

    const prefixedMatch = keyEntries.find((entry) =>
      entry.normalized.startsWith(`${normalizedAlias} `)
    );
    if (prefixedMatch && compactAlias.length > 2 && normalizedAlias.includes(" ")) {
      return row[prefixedMatch.key];
    }
  }

  return null;
};

const isSummaryLabel = (value) => {
  const text = normalize(value, "").toLowerCase();
  return text.includes("total") || text.includes("summary") || text.includes("grand total");
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

const extractDashboardRows = (rawRows) =>
  rawRows
    .map((row, index) => {
      const detectedActivityId = normalize(
        getCell(row, ["activity id", "id", "activity code", "wbs", "task id"]),
        ""
      );
      const activity = normalize(getCell(row, ["activity", "activity name"]), "");
      const hasValidActivityId = isValidActivityId(detectedActivityId);
      const hasValidActivityName = activity !== "" && !isSummaryLabel(activity);

      if (!hasValidActivityId && !hasValidActivityName) return null;

      const plannedCost = parseNumber(getCell(row, ["planned value", "planned cost", "pv", "budget"]));
      const actualCost = parseNumber(getCell(row, ["actual cost", "ac", "actual"]));
      const rawCv = getCell(row, ["cost variance", "cv"]);
      const hasProvidedCv =
        rawCv !== null && rawCv !== undefined && String(rawCv).trim() !== "";
      const providedCv = parseNumber(rawCv);
      const cv = hasProvidedCv ? providedCv : plannedCost - actualCost;

      return {
        activityId: hasValidActivityId ? detectedActivityId : `ROW-${index + 1}`,
        activity: activity || "Unspecified",
        plannedCost,
        actualCost,
        ev: parseNumber(getCell(row, ["earned value", "ev"])),
        percentComplete: parseNumber(getCell(row, ["% complete", "percent complete", "progress %"])),
        cv,
        costUsedPercent: plannedCost ? (actualCost / plannedCost) * 100 : 0,
        budgetVariancePercent: plannedCost ? (cv / plannedCost) * 100 : 0,
        budgetStatus: cv >= 0 ? "Under Budget" : "Over Budget",
      };
    })
    .filter(
      (row) =>
        row &&
        (row.plannedCost !== 0 || row.actualCost !== 0 || row.ev !== 0 || row.percentComplete !== 0)
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

const calculateProgressMetrics = (rows, totals) => {
  const weightedProgressValue = rows.reduce(
    (acc, row) => acc + row.plannedCost * (row.percentComplete / 100),
    0
  );
  const physicalProgressPercent = totals.planned ? (weightedProgressValue / totals.planned) * 100 : 0;
  const costSpentPercent = totals.planned ? (totals.actual / totals.planned) * 100 : 0;
  const efficiencyGapPercent = physicalProgressPercent - costSpentPercent;

  return { physicalProgressPercent, costSpentPercent, efficiencyGapPercent };
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
  efficiencyCardEl.classList.remove("status-under", "status-over");

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

const renderProgressKpis = (metrics) => {
  physicalProgressEl.textContent = formatPercent(metrics.physicalProgressPercent);
  costSpentEl.textContent = formatPercent(metrics.costSpentPercent);
  efficiencyGapEl.textContent = formatSignedPercent(metrics.efficiencyGapPercent);

  efficiencyCardEl.classList.remove("status-under", "status-over");
  if (metrics.efficiencyGapPercent < 0) {
    efficiencyCardEl.classList.add("status-over");
  } else if (metrics.efficiencyGapPercent > 0) {
    efficiencyCardEl.classList.add("status-under");
  }
};

const renderTable = (rows) => {
  if (!rows.length) {
    tableBodyEl.innerHTML =
      '<tr><td colspan="10" class="placeholder">No valid rows found in data source.</td></tr>';
    return;
  }

  tableBodyEl.innerHTML = rows
    .map(
      (row) => `
      <tr class="variance-row variance-${getVarianceBand(row.cv, row.plannedCost)}">
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

const renderOverrunTable = (rows) => {
  const overrunRows = rows
    .filter((row) => row.cv < 0)
    .sort((a, b) => a.cv - b.cv)
    .slice(0, 5);

  if (!overrunRows.length) {
    overrunTableBodyEl.innerHTML =
      '<tr><td colspan="5" class="placeholder">No overrun activities found.</td></tr>';
    return;
  }

  overrunTableBodyEl.innerHTML = overrunRows
    .map(
      (row) => `
      <tr>
        <td>${row.activity}</td>
        <td>${formatCurrency(row.cv)}</td>
        <td>${formatPercent(row.budgetVariancePercent)}</td>
        <td>${formatPercent(row.percentComplete)}</td>
        <td>${formatPercent(row.costUsedPercent)}</td>
      </tr>
    `
    )
    .join("");
};

const showMessage = (text, isError = false) => {
  messageEl.textContent = text;
  messageEl.style.color = isError ? "#dc2626" : "#667085";
};

const destroyCharts = () => {
  if (varianceChart) {
    varianceChart.destroy();
    varianceChart = null;
  }

  if (costChart) {
    costChart.destroy();
    costChart = null;
  }

  if (evmTrendChart) {
    evmTrendChart.destroy();
    evmTrendChart = null;
  }

  if (efficiencyScatterChart) {
    efficiencyScatterChart.destroy();
    efficiencyScatterChart = null;
  }
};

const generateCharts = (rows) => {
  if (typeof window.Chart === "undefined") {
    showMessage(chartDependencyWarning, true);
    return;
  }

  if (!rows.length) {
    destroyCharts();
    showMessage("No data available to generate charts.", true);
    return;
  }

  const labels = rows.map((row) => row.activity);
  const varianceValues = rows.map((row) => row.cv);
  const varianceAxis = buildLinearAxisRange(varianceValues, { includeZero: true, targetTickCount: 7 });

  const plannedSeries = rows.map((row) => row.plannedCost);
  const actualSeries = rows.map((row) => row.actualCost);

  const costAxis = buildLinearAxisRange([...plannedSeries, ...actualSeries], {
    includeZero: true,
    targetTickCount: 7,
  });

  destroyCharts();

  varianceChart = new Chart(document.getElementById("varianceChart"), {
    type: "bar",
    data: {
      labels,
      datasets: [
        {
          label: "Cost Variance",
          data: varianceValues,
          backgroundColor: rows.map((row) => (row.cv >= 0 ? "#22c55e" : "#ef4444")),
          borderRadius: 6,
        },
      ],
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: { legend: { display: false } },
      scales: {
        y: {
          min: varianceAxis.min,
          max: varianceAxis.max,
          ticks: {
            stepSize: varianceAxis.stepSize,
            callback: (value) => formatCurrency(value),
          },
        },
      },
    },
  });

  costChart = new Chart(document.getElementById("costChart"), {
    type: "line",
    data: {
      labels,
      datasets: [
        {
          label: "Planned Cost (PV)",
          data: plannedSeries,
          borderColor: "#2f55ff",
          backgroundColor: "#2f55ff",
          tension: 0.3,
        },
        {
          label: "Actual Cost (AC)",
          data: actualSeries,
          borderColor: "#10b981",
          backgroundColor: "#10b981",
          tension: 0.3,
        },
      ],
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      scales: {
        y: {
          min: costAxis.min,
          max: costAxis.max,
          ticks: {
            stepSize: costAxis.stepSize,
            callback: (value) => formatCurrency(value),
          },
        },
      },
    },
  });

  const cumulativeSeries = rows.reduce(
    (acc, row, index) => {
      const previousPv = index ? acc.pv[index - 1] : 0;
      const previousEv = index ? acc.ev[index - 1] : 0;
      const previousAc = index ? acc.ac[index - 1] : 0;

      acc.labels.push(`${index + 1}. ${row.activity}`);
      acc.pv.push(previousPv + row.plannedCost);
      acc.ev.push(previousEv + row.ev);
      acc.ac.push(previousAc + row.actualCost);
      return acc;
    },
    { labels: [], pv: [], ev: [], ac: [] }
  );

  const evmAxis = buildLinearAxisRange(
    [...cumulativeSeries.pv, ...cumulativeSeries.ev, ...cumulativeSeries.ac],
    { includeZero: true, targetTickCount: 7 }
  );

  evmTrendChart = new Chart(document.getElementById("evmTrendChart"), {
    type: "line",
    data: {
      labels: cumulativeSeries.labels,
      datasets: [
        {
          label: "Cumulative PV",
          data: cumulativeSeries.pv,
          borderColor: "#2f55ff",
          backgroundColor: "#2f55ff",
          tension: 0.25,
        },
        {
          label: "Cumulative EV",
          data: cumulativeSeries.ev,
          borderColor: "#7c3aed",
          backgroundColor: "#7c3aed",
          tension: 0.25,
        },
        {
          label: "Cumulative AC",
          data: cumulativeSeries.ac,
          borderColor: "#10b981",
          backgroundColor: "#10b981",
          tension: 0.25,
        },
      ],
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      scales: {
        y: {
          min: evmAxis.min,
          max: evmAxis.max,
          ticks: {
            stepSize: evmAxis.stepSize,
            callback: (value) => formatCurrency(value),
          },
        },
      },
    },
  });

  const scatterPoints = rows.map((row) => ({
    x: row.percentComplete,
    y: row.costUsedPercent,
    activity: row.activity,
  }));

  efficiencyScatterChart = new Chart(document.getElementById("efficiencyScatterChart"), {
    type: "scatter",
    data: {
      datasets: [
        {
          label: "Activities",
          data: scatterPoints,
          backgroundColor: rows.map((row) => (row.costUsedPercent > row.percentComplete ? "#ef4444" : "#22c55e")),
          pointRadius: 5,
          pointHoverRadius: 7,
        },
        {
          label: "Balanced Spend (y = x)",
          type: "line",
          data: [
            { x: 0, y: 0 },
            { x: 100, y: 100 },
          ],
          borderColor: "#64748b",
          borderDash: [6, 6],
          pointRadius: 0,
          fill: false,
        },
      ],
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        tooltip: {
          callbacks: {
            label: (context) => {
              if (context.dataset.label !== "Activities") return context.dataset.label;
              const point = context.raw;
              return `${point.activity}: ${point.x.toFixed(2)}% complete, ${point.y.toFixed(2)}% cost used`;
            },
          },
        },
      },
      scales: {
        x: {
          min: 0,
          max: 100,
          title: { display: true, text: "% Complete" },
          ticks: { callback: (value) => `${value}%` },
        },
        y: {
          min: 0,
          max: 100,
          title: { display: true, text: "% Cost Used" },
          ticks: { callback: (value) => `${value}%` },
        },
      },
    },
  });
};

const processRows = (rawRows, sourceName = "web app") => {
  if (!Array.isArray(rawRows) || !rawRows.length) {
    dashboardRows = [];
    renderKpis({ planned: 0, actual: 0, cv: 0 });
    renderProgressKpis({ physicalProgressPercent: 0, costSpentPercent: 0, efficiencyGapPercent: 0 });
    renderTable([]);
    renderOverrunTable([]);
    destroyCharts();
    showMessage(`No rows detected from ${sourceName}.`, true);
    return;
  }

  const rows = extractDashboardRows(rawRows);
  dashboardRows = rows;
  const totals = calculateTotalsFromRows(rows);
  const progressMetrics = calculateProgressMetrics(rows, totals);

  renderKpis(totals);
  renderProgressKpis(progressMetrics);
  renderTable(rows);
  renderOverrunTable(rows);
  generateCharts(rows);

  showMessage(`Loaded ${rows.length} activity row(s) from ${sourceName}. Charts refreshed.`);
};

const setupServiceWorkerUpdates = async () => {
  if (!("serviceWorker" in navigator)) {
    showMessage("This browser does not support background app updates.", true);
    return;
  }

  try {
    const registration = await navigator.serviceWorker.register("./service-worker.js");

    navigator.serviceWorker.addEventListener("controllerchange", () => {
      window.location.reload();
    });

    setInterval(() => {
      registration.update();
    }, 5 * 60 * 1000);
  } catch (error) {
    showMessage(`Service worker setup failed: ${error.message}`, true);
  }
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

const loadGoogleSheet = async (providedUrl = "") => {
  const rawUrl = providedUrl || DATA_SOURCE_URL;
  const trimmedUrl = rawUrl.trim();
  const isWebAppSource = isAppsScriptWebAppUrl(trimmedUrl);
  const csvUrl = isWebAppSource ? "" : toGoogleSheetCsvUrl(trimmedUrl);

  if (!isWebAppSource && !csvUrl) {
    showMessage("Invalid URL. Use a valid Google Sheet or Apps Script Web App URL.", true);
    return;
  }

  try {
    showMessage("Loading data source...");

    if (isWebAppSource) {
      const response = await fetch(trimmedUrl);
      if (!response.ok) throw new Error(`Unable to fetch Apps Script Web App (HTTP ${response.status})`);
      const payload = await response.json();
      if (payload?.error) throw new Error(payload.error);
      processRows(payload?.rows || [], `Apps Script Web App (${payload?.sheetName || "sheet"})`);
      return;
    }

    const response = await fetch(csvUrl);
    if (!response.ok) throw new Error(`Unable to fetch Google Sheet (HTTP ${response.status})`);

    const csvText = await response.text();
    const workbook = XLSX.read(csvText, { type: "string" });
    const firstSheetName = workbook.SheetNames[0];
    const sheet = workbook.Sheets[firstSheetName];
    const headerRowIndex = findHeaderRowIndex(sheet);
    const rawRows = XLSX.utils.sheet_to_json(sheet, { raw: false, defval: "", range: headerRowIndex });

    processRows(rawRows, `Google Sheet "${firstSheetName}"`);
  } catch (error) {
    showMessage(`Error loading data source: ${error.message}`, true);
  }
};

setupServiceWorkerUpdates();
loadGoogleSheet(DATA_SOURCE_URL);
