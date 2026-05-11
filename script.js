const totalPlannedEl = document.getElementById("totalPlanned");
const totalActualEl = document.getElementById("totalActual");
const totalCvEl = document.getElementById("totalCv");
const projectStatusEl = document.getElementById("projectStatus");
const statusCardEl = document.getElementById("statusCard");
const physicalProgressEl = document.getElementById("physicalProgress");
const costSpentEl = document.getElementById("costSpent");
const efficiencyGapEl = document.getElementById("efficiencyGap");
const miniCompleteEl = document.getElementById("miniComplete");
const miniCostEl = document.getElementById("miniCost");
const miniCompleteBarEl = document.querySelector(".mini-fill.blue");
const miniCostBarEl = document.querySelector(".mini-fill.green");
const efficiencyCardEl = document.getElementById("efficiencyCard");
const messageEl = document.getElementById("message");
const tableBodyEl = document.getElementById("activityTableBody");
const overrunTableBodyEl = document.getElementById("overrunTableBody");
const gapTableBodyEl = document.getElementById("gapTableBody");
const varianceDisplayEl = document.getElementById("varianceDisplay");
const varianceStatusEl = document.getElementById("varianceStatus");
const projectFilterEl = document.getElementById("projectFilter");
const dateStartFilterEl = document.getElementById("dateStartFilter");
const dateEndFilterEl = document.getElementById("dateEndFilter");

const DATA_SOURCE_URL = window.DataBridge?.DEFAULT_DATA_SOURCE_URL || "";
const USE_COST_MANAGEMENT_ONLY = false;

const chartDependencyWarning =
  typeof window.Chart === "undefined"
    ? "Chart.js is not available. Graphs are disabled."
    : "";

let activitySummaryRows = [];
let varianceChart = null;
let costChart = null;
let dashboardRefreshTimer = null;
let isDashboardFetchInFlight = false;
let latestDashboardSignature = "";

const DASHBOARD_CACHE_KEY = "constructionStageDashboardRows";
const COST_ACTIVITIES_LOCAL_STORAGE_KEY = "constructionStageActivities";
const LEGACY_COST_ACTIVITIES_LOCAL_STORAGE_KEY = "constructionStageCostActivities";
const DAILY_COSTS_LOCAL_STORAGE_KEY = "constructionStageDailyCosts";
const DASHBOARD_CACHE_TTL_MS = 5 * 1000;
const DASHBOARD_REFRESH_INTERVAL_MS = 3 * 1000;

const getProjectFilterPrefill = () => {
  try {
    const url = new URL(window.location.href);
    const projectId = String(url.searchParams.get("projectId") || "").trim();
    const projectName = String(url.searchParams.get("project") || "").trim();
    if (projectId && projectName) return `${projectId} - ${projectName}`;
    return projectId || projectName || "";
  } catch {
    return "";
  }
};

const projectFilterPrefill = String(getProjectFilterPrefill() || "").trim().toLowerCase();
let hasAppliedProjectFilterPrefill = false;
const EXTENSION_BRIDGE_DISCONNECT_MESSAGE =
  "Could not establish connection. Receiving end does not exist.";

window.addEventListener("unhandledrejection", (event) => {
  const reasonMessage =
    event?.reason && typeof event.reason === "object"
      ? event.reason.message || ""
      : String(event?.reason || "");

  if (reasonMessage.includes(EXTENSION_BRIDGE_DISCONNECT_MESSAGE)) {
    console.warn(
      "Ignored a browser-extension bridge rejection because no receiving end is available."
    );
    event.preventDefault();
  }
});

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

const escapeHtml = (value) =>
  String(value ?? "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");

const normalizeProgressPercent = (value) => {
  const numericValue = parseNumber(value);
  if (!Number.isFinite(numericValue) || numericValue <= 0) return 0;
  return numericValue <= 1 ? numericValue * 100 : numericValue;
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

const normalizeIdentityKey = (value) => String(value || "").trim().toLowerCase();
const makeDashboardCompositeKey = ({ projectId = "", activityId = "", costId = "" } = {}) =>
  [projectId, activityId, costId].map(normalizeIdentityKey).join("::");

const formatActivityCostIdentity = (row = {}) => {
  const activityId = String(row.activityId || "").trim();
  const costId = String(row.costId || "").trim();
  if (activityId && costId && normalizeIdentityKey(activityId) !== normalizeIdentityKey(costId)) {
    return `${activityId} / ${costId}`;
  }
  return activityId || costId || "-";
};


const normalizeHeader = (value) =>
  String(value)
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, " ")
    .replace(/\s+/g, " ")
    .trim();

const compactHeader = (value) => normalizeHeader(value).replace(/\s+/g, "");

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

const extractDashboardRows = (rawRows) =>
  rawRows
    .map((row, index) => {
      const detectedActivityId = normalize(
        getCell(row, ["activity id", "activity id/cost id", "activity code", "wbs", "task id"]),
        ""
      );
      const detectedCostId = normalize(
        getCell(row, ["cost id", "cost code", "cost_id", "costid"]),
        ""
      );
      const activity = normalize(
        getCell(row, [
          "activity",
          "activity name",
          "activity title",
          "name",
          "task",
          "task name",
          "description",
        ]),
        ""
      );
      const hasValidActivityId = isValidActivityId(detectedActivityId);
      const hasValidCostId = isValidActivityId(detectedCostId);
      const hasValidActivityName = activity !== "" && !isSummaryLabel(activity);

      if (!hasValidActivityId && !hasValidCostId && !hasValidActivityName) return null;

      const projectId = normalize(getCell(row, ["project id", "project code", "projectid", "code"]), "");
      const projectName = normalize(getCell(row, ["project name", "project", "project title"]), "");
      const project = projectId && projectName
        ? `${projectId} - ${projectName}`
        : projectId || projectName || "No Project ID";
      const startDate = normalizeDateOnly(getCell(row, ["planned start", "start date", "start"]));
      const finishDate = normalizeDateOnly(getCell(row, ["planned finish", "finish date", "end date", "finish"]));
      const plannedCost = parseNumber(
        getCell(row, ["planned value", "planned cost", "total budget", "pv", "budget"])
      );
      const actualCost = parseNumber(
        getCell(row, ["actual cost", "total spent", "ac", "actual"])
      );
      const percentComplete = normalizeProgressPercent(
        getCell(row, ["% complete", "percent complete", "progress %", "progress", "completion"] )
      );
      const rawEarnedValue = getCell(row, ["earned value", "earned value/day", "ev"]);
      const hasProvidedEarnedValue =
        rawEarnedValue !== null && rawEarnedValue !== undefined && String(rawEarnedValue).trim() !== "";
      const ev = hasProvidedEarnedValue ? parseNumber(rawEarnedValue) : plannedCost * (percentComplete / 100);
      const rawCv = getCell(row, ["cost variance", "cv"]);
      const hasProvidedCv =
        rawCv !== null && rawCv !== undefined && String(rawCv).trim() !== "";
      const providedCv = parseNumber(rawCv);
      const cv = hasProvidedCv ? providedCv : ev - actualCost;

      return {
        activityId: hasValidActivityId ? detectedActivityId : "",
        costId: hasValidCostId ? detectedCostId : "",
        activity: activity || (hasValidActivityId ? `Activity ${detectedActivityId}` : hasValidCostId ? `Cost ${detectedCostId}` : "Unnamed Activity"),
        projectId,
        projectName,
        project,
        startDate,
        finishDate,
        plannedCost,
        actualCost,
        ev,
        percentComplete,
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

const normalizeDateOnly = (value) => {
  const raw = normalize(value, "");
  if (!raw) return "";
  const date = new Date(raw);
  if (Number.isNaN(date.getTime())) return "";
  return date.toISOString().split("T")[0];
};

const rowMatchesDateFilter = (row, startDate, endDate) => {
  if (!startDate && !endDate) return true;
  const rowStart = normalizeDateOnly(row.startDate) || normalizeDateOnly(row.date);
  const rowEnd = normalizeDateOnly(row.finishDate) || rowStart;
  if (!rowStart && !rowEnd) return false;
  if (startDate && rowEnd && rowEnd < startDate) return false;
  if (endDate && rowStart && rowStart > endDate) return false;
  return true;
};

const getFilteredRows = (rows) => {
  const selectedProject = String(projectFilterEl?.value || "all").trim().toLowerCase();
  const startDate = String(dateStartFilterEl?.value || "").trim();
  const endDate = String(dateEndFilterEl?.value || "").trim();

  return rows.filter((row) => {
    const projectMatches = selectedProject === "all"
      || normalize(row.project, "").trim().toLowerCase() === selectedProject;
    return projectMatches && rowMatchesDateFilter(row, startDate, endDate);
  });
};

const syncFilterOptionsFromRows = (rows) => {
  if (!projectFilterEl) return;
  const projects = Array.from(
    new Set(rows.map((row) => normalize(row.project, "Unspecified")).filter(Boolean))
  ).sort((a, b) => a.localeCompare(b));

  const currentSelection = normalize(projectFilterEl.value, "").trim().toLowerCase();
  const prefilledSelection = !hasAppliedProjectFilterPrefill ? projectFilterPrefill : "";

  projectFilterEl.innerHTML = `<option value="all">All Projects</option>${projects
    .map((project) => `<option value="${escapeHtml(project.toLowerCase())}">${escapeHtml(project)}</option>`)
    .join("")}`;

  const selected = prefilledSelection || (currentSelection === "all" ? "" : currentSelection);
  const selectedExists = selected && projects.some((project) => project.toLowerCase() === selected);

  if (selectedExists) {
    projectFilterEl.value = selected;
  } else {
    projectFilterEl.value = "all";
  }

  hasAppliedProjectFilterPrefill = true;
};

const renderDashboardFromRows = (rows, summaryRows = rows) => {
  const totals = calculateTotalsFromRows(rows);
  const progressMetrics = calculateProgressMetrics(rows, totals);
  renderKpis(totals);
  renderProgressKpis(progressMetrics, totals);
  renderTable(summaryRows);
  renderOverrunTable(rows);
  renderGapTable(rows);
  generateCharts(rows);
};

const applyFiltersAndRender = () => {
  const filteredRows = getFilteredRows(activitySummaryRows);
  renderDashboardFromRows(filteredRows, filteredRows);
};

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

  if (totalPlannedEl) totalPlannedEl.textContent = formatCurrency(safeTotals.planned);
  if (totalActualEl) totalActualEl.textContent = formatCurrency(safeTotals.actual);
  if (totalCvEl) totalCvEl.textContent = formatCurrency(safeTotals.cv);

  if (statusCardEl) statusCardEl.classList.remove("status-under", "status-over");
  if (efficiencyCardEl) efficiencyCardEl.classList.remove("status-under", "status-over");

  if (safeTotals.actual < safeTotals.planned) {
    if (projectStatusEl) projectStatusEl.textContent = "Under Budget";
    if (statusCardEl) statusCardEl.classList.add("status-under");
  } else if (safeTotals.actual > safeTotals.planned) {
    if (projectStatusEl) projectStatusEl.textContent = "Over Budget";
    if (statusCardEl) statusCardEl.classList.add("status-over");
  } else {
    if (projectStatusEl) projectStatusEl.textContent = "On Budget";
  }
};

const renderProgressKpis = (metrics, totals) => {
  const earnedValue = (parseNumber(totals?.planned) * parseNumber(metrics.physicalProgressPercent)) / 100;
  const cpi = parseNumber(totals?.actual) ? earnedValue / parseNumber(totals?.actual) : 0;

  if (physicalProgressEl) physicalProgressEl.textContent = formatCurrency(earnedValue);
  if (costSpentEl) costSpentEl.textContent = cpi.toFixed(2);
  if (efficiencyGapEl) efficiencyGapEl.textContent = `${formatPercent(metrics.physicalProgressPercent)} / ${formatPercent(metrics.costSpentPercent)}`;
  if (miniCompleteEl) miniCompleteEl.textContent = formatPercent(metrics.physicalProgressPercent);
  if (miniCostEl) miniCostEl.textContent = formatPercent(metrics.costSpentPercent);
  if (miniCompleteBarEl) miniCompleteBarEl.style.width = `${Math.max(0, Math.min(100, metrics.physicalProgressPercent))}%`;
  if (miniCostBarEl) miniCostBarEl.style.width = `${Math.max(0, Math.min(100, metrics.costSpentPercent))}%`;

  if (efficiencyCardEl) efficiencyCardEl.classList.remove("status-under", "status-over");
  if (metrics.efficiencyGapPercent < 0) {
    if (efficiencyCardEl) efficiencyCardEl.classList.add("status-over");
  } else if (metrics.efficiencyGapPercent > 0) {
    if (efficiencyCardEl) efficiencyCardEl.classList.add("status-under");
  }


  if (varianceDisplayEl) varianceDisplayEl.textContent = formatCurrency(parseNumber(totals?.cv));
  if (varianceStatusEl) varianceStatusEl.textContent = parseNumber(totals?.cv) < 0 ? "Over Budget" : "Under Budget";
};

const renderGapTable = (rows) => {
  if (!gapTableBodyEl) return;
  if (!rows.length) {
    gapTableBodyEl.innerHTML = `<tr><td colspan="6" class="placeholder">No data loaded yet.</td></tr>`;
    return;
  }

  gapTableBodyEl.innerHTML = rows
    .slice()
    .sort((a,b) => (b.percentComplete - b.costUsedPercent) - (a.percentComplete - a.costUsedPercent))
    .map((row) => {
      const gap = row.percentComplete - row.costUsedPercent;
      const status = gap >= 0 ? "On Track" : gap > -5 ? "Slightly Over Budget" : "Over Budget";
      const interpretation = gap >= 0 ? "Progress is ahead of cost." : gap > -5 ? "Cost is slightly ahead of progress." : "Cost is ahead of progress.";
      const gapClass = gap >= 0 ? "positive" : "negative";
      const statusClass = gap >= 0 ? "ok" : gap > -5 ? "warn" : "bad";
      return `<tr>
        <td>${escapeHtml(row.activity)}</td>
        <td>${formatPercent(row.percentComplete)}</td>
        <td>${formatPercent(row.costUsedPercent)}</td>
        <td><div class="gap-cell"><strong>${formatSignedPercent(gap)}</strong><span class="gap-track"><span class="gap-fill ${gapClass}" style="width:${Math.min(100, Math.abs(gap) * 4)}%"></span></span></div></td>
        <td><span class="status-pill ${statusClass}">${status}</span></td>
        <td>${interpretation}</td>
      </tr>`;
    }).join("");
};

const renderTable = (rows) => {
  if (!tableBodyEl) return;

  if (!rows.length) {
    tableBodyEl.innerHTML =
      '<tr><td colspan="10" class="placeholder">No valid rows found in data source.</td></tr>';
    return;
  }

  tableBodyEl.innerHTML = rows
    .map(
      (row) => `
      <tr class="variance-row variance-${getVarianceBand(row.cv, row.plannedCost)}">
        <td>${escapeHtml(formatActivityCostIdentity(row))}</td>
        <td>${escapeHtml(row.activity)}</td>
        <td>${formatCurrency(row.plannedCost)}</td>
        <td>${formatCurrency(row.actualCost)}</td>
        <td>${formatCurrency(row.ev)}</td>
        <td>${formatPercent(row.percentComplete)}</td>
        <td>${formatPercent(row.costUsedPercent)}</td>
        <td>${formatCurrency(row.cv)}</td>
        <td>${(row.actualCost ? row.ev / row.actualCost : 0).toFixed(2)}</td>
        <td><span class="status-pill ${row.cv >= 0 ? "ok" : "bad"}">${row.cv >= 0 ? "On Track" : "Over Budget"}</span></td>
      </tr>
    `
    )
    .join("") + `
      ${(() => {
        const totals = rows.reduce(
          (acc, row) => {
            acc.planned += row.plannedCost;
            acc.actual += row.actualCost;
            acc.ev += row.ev;
            acc.cv += row.cv;
            return acc;
          },
          { planned: 0, actual: 0, ev: 0, cv: 0 }
        );
        const aggregateCompletePercent = totals.planned ? (totals.ev / totals.planned) * 100 : 0;
        const aggregateCostUsedPercent = totals.planned ? (totals.actual / totals.planned) * 100 : 0;
        const aggregateCpi = totals.actual ? totals.ev / totals.actual : 0;

        return `<tr>
          <td></td>
          <td><strong>TOTAL</strong></td>
          <td><strong>${formatCurrency(totals.planned)}</strong></td>
          <td><strong>${formatCurrency(totals.actual)}</strong></td>
          <td><strong>${formatCurrency(totals.ev)}</strong></td>
          <td><strong>${formatPercent(aggregateCompletePercent)}</strong></td>
          <td><strong>${formatPercent(aggregateCostUsedPercent)}</strong></td>
          <td><strong>${formatCurrency(totals.cv)}</strong></td>
          <td><strong>${aggregateCpi.toFixed(2)}</strong></td>
          <td><span class="status-pill ${totals.cv >= 0 ? "ok" : "bad"}">${totals.cv >= 0 ? "On Track" : "Over Budget"}</span></td>
        </tr>`;
      })()}`;
};

const renderOverrunTable = (rows) => {
  if (!overrunTableBodyEl) return;
  const overrunRows = rows
    .filter((row) => row.cv < 0)
    .sort((a, b) => a.cv - b.cv)
    .slice(0, 5);

  if (!overrunRows.length) {
    overrunTableBodyEl.innerHTML =
      '<tr><td colspan="4" class="placeholder">No overrun activities found.</td></tr>';
    return;
  }

  overrunTableBodyEl.innerHTML = overrunRows
    .map(
      (row) => `
      <tr>
        <td>${escapeHtml(row.activity)}</td>
        <td>${formatCurrency(row.cv)}</td>
        <td>${formatSignedPercent(row.percentComplete - row.costUsedPercent)}</td>
        <td><span class="status-pill bad">Over Budget</span></td>
      </tr>
    `
    )
    .join("");
};

const showMessage = (text, isError = false) => {
  if (!messageEl) return;
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
  const completeSeries = rows.map((row) => row.percentComplete);
  const costUsedSeries = rows.map((row) => row.costUsedPercent);
  const varianceValues = rows.map((row) => row.actualCost - row.plannedCost);
  const varianceAxis = buildLinearAxisRange(varianceValues, { includeZero: true, targetTickCount: 7 });
  const costAxis = buildLinearAxisRange(varianceValues, { includeZero: true, targetTickCount: 7 });

  destroyCharts();

  const varianceCanvas = document.getElementById("varianceChart");
  const costCanvas = document.getElementById("costChart");
  if (!varianceCanvas || !costCanvas) {
    showMessage("Dashboard chart containers are missing. Refresh or restore dashboard markup.", true);
    return;
  }

  varianceChart = new Chart(varianceCanvas, {
    type: "line",
    data: {
      labels,
      datasets: [
        {
          label: "Progress",
          data: completeSeries,
          borderColor: "#2f55ff",
          backgroundColor: "#2f55ff",
          tension: 0.25,
        },
        {
          label: "% Cost Used",
          data: costUsedSeries,
          borderColor: "#16a34a",
          backgroundColor: "#16a34a",
          tension: 0.25,
        },
      ],
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: { legend: { display: false } },
      scales: {
        y: {
          min: 0,
          max: 100,
          ticks: {
            callback: (value) => `${value}%`,
          },
        },
      },
    },
  });

  costChart = new Chart(costCanvas, {
    type: "bar",
    data: {
      labels,
      datasets: [
        {
          label: "Cost Variance (AC - PC)",
          data: varianceValues,
          backgroundColor: rows.map((row) => (row.actualCost - row.plannedCost <= 0 ? "#22c55e" : "#ef4444")),
          borderRadius: 6,
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

};

const processRows = (rawRows, sourceName = "web app") => {
  if (!Array.isArray(rawRows) || !rawRows.length) {
    if (activitySummaryRows.length) {
      showMessage(
        `Live source temporarily returned no rows from ${sourceName}. Retaining the last ${activitySummaryRows.length} activity row(s).`,
        true
      );
      return;
    }

    activitySummaryRows = [];
    renderKpis({ planned: 0, actual: 0, cv: 0 });
    renderProgressKpis({ physicalProgressPercent: 0, costSpentPercent: 0, efficiencyGapPercent: 0 });
    renderTable([]);
    renderOverrunTable([]);
    renderGapTable([]);
    if (varianceDisplayEl) varianceDisplayEl.textContent = formatCurrency(0);
    if (varianceStatusEl) varianceStatusEl.textContent = "Balanced";
    destroyCharts();
    showMessage(`No rows detected from ${sourceName}.`, true);
    return;
  }

  const rows = extractDashboardRows(rawRows);
  const nextSignature = JSON.stringify(rows);
  if (nextSignature === latestDashboardSignature) {
    showMessage(`Live sync active. No new updates from ${sourceName}.`);
    return;
  }

  latestDashboardSignature = nextSignature;
  localStorage.setItem(DASHBOARD_CACHE_KEY, JSON.stringify({ savedAt: Date.now(), rows }));
  activitySummaryRows = rows;
  syncFilterOptionsFromRows(rows);
  applyFiltersAndRender();

  const asOfEl = document.querySelector(".as-of-text");
  if (asOfEl) {
    asOfEl.textContent = `Data as of: ${new Date().toLocaleString()}`;
  }

  showMessage(`Loaded ${rows.length} activity row(s) from ${sourceName}. Charts refreshed.`);
};

const processActivitySummaryRows = (rawRows) => {
  if (!Array.isArray(rawRows) || !rawRows.length) {
    activitySummaryRows = [];
    applyFiltersAndRender();
    return;
  }

  activitySummaryRows = extractDashboardRows(rawRows);
  applyFiltersAndRender();
};

const loadRowsFromCostManagementLocalData = () => {
  const safeParseArray = (raw) => {
    try {
      const parsed = JSON.parse(raw || "[]");
      return Array.isArray(parsed) ? parsed : [];
    } catch {
      return [];
    }
  };

  const activities = safeParseArray(localStorage.getItem(COST_ACTIVITIES_LOCAL_STORAGE_KEY));
  const legacyActivities = safeParseArray(localStorage.getItem(LEGACY_COST_ACTIVITIES_LOCAL_STORAGE_KEY));
  const mergedActivities = activities.length ? activities : legacyActivities;
  if (!mergedActivities.length) return [];
  const dailyCosts = safeParseArray(localStorage.getItem(DAILY_COSTS_LOCAL_STORAGE_KEY));

  const actualCostByCompositeKey = dailyCosts.reduce((map, item) => {
    const projectId = String(item?.projectId || item?.project_id || "").trim();
    const activityId = String(item?.activityId || item?.activity_id || "").trim();
    const costId = String(item?.costId || item?.cost_id || "").trim();
    const compositeKey = makeDashboardCompositeKey({ projectId, activityId, costId });
    const activityFallbackKey = makeDashboardCompositeKey({ projectId, activityId, costId: "" });
    if (!projectId || (!activityId && !costId)) return map;
    map.set(compositeKey, (map.get(compositeKey) || 0) + parseNumber(item?.actualCost));
    if (activityFallbackKey !== compositeKey) {
      map.set(activityFallbackKey, (map.get(activityFallbackKey) || 0) + parseNumber(item?.actualCost));
    }
    return map;
  }, new Map());

  return mergedActivities.map((activity) => {
    const activityId = String(activity?.activityRefId || activity?.activityId || activity?.id || "").trim();
    const costId = String(activity?.costId || activity?.cost_id || "").trim();
    const projectId = String(activity?.projectId || "").trim();
    const projectName = String(activity?.projectName || activity?.project || "").trim();
    const compositeKey = makeDashboardCompositeKey({ projectId, activityId, costId });
    const activityFallbackKey = makeDashboardCompositeKey({ projectId, activityId, costId: "" });
    return {
      "Project ID": projectId,
      "Project Name": projectName,
      "Activity ID": activityId,
      "Cost ID": costId,
      Activity: String(activity?.name || activity?.activity || (activityId ? `Activity ${activityId}` : costId ? `Cost ${costId}` : "Unnamed Activity")).trim(),
      "Planned Cost": parseNumber(activity?.plannedCost),
      "Actual Cost": actualCostByCompositeKey.get(compositeKey) || actualCostByCompositeKey.get(activityFallbackKey) || 0,
      "Progress": parseNumber(activity?.progressPercent),
    };
  });
};

const mergeRemoteRowsWithCostManagementActuals = (remoteRows, localRows) => {
  if (!Array.isArray(remoteRows) || !remoteRows.length) return [];
  if (!Array.isArray(localRows) || !localRows.length) return remoteRows;

  const localByCompositeKey = new Map();
  const localByProjectActivityKey = new Map();
  localRows.forEach((row) => {
    const projectId = String(row?.["Project ID"] || "").trim();
    const activityId = String(row?.["Activity ID"] || "").trim();
    const costId = String(row?.["Cost ID"] || "").trim();
    const compositeKey = makeDashboardCompositeKey({ projectId, activityId, costId });
    const projectActivityKey = makeDashboardCompositeKey({ projectId, activityId, costId: "" });
    if (projectId && (activityId || costId)) localByCompositeKey.set(compositeKey, row);
    if (projectId && activityId && !localByProjectActivityKey.has(projectActivityKey)) {
      localByProjectActivityKey.set(projectActivityKey, row);
    }
  });

  return remoteRows.map((row) => {
    const projectId = String(getCell(row, ["project id", "project code", "projectid", "code"]) || "").trim();
    const activityId = String(
      getCell(row, ["activity id", "id", "activity code", "wbs", "task id"]) || ""
    ).trim();
    const costId = String(getCell(row, ["cost id", "cost code", "cost_id", "costid"]) || "").trim();
    const compositeKey = makeDashboardCompositeKey({ projectId, activityId, costId });
    const projectActivityKey = makeDashboardCompositeKey({ projectId, activityId, costId: "" });
    const local = localByCompositeKey.get(compositeKey) || localByProjectActivityKey.get(projectActivityKey);
    if (!local) return row;

    return {
      ...row,
      "Actual Cost":
        parseNumber(local?.["Actual Cost"]) || parseNumber(getCell(row, ["actual cost", "total spent", "ac", "actual"])),
      "Project ID": local?.["Project ID"] || projectId,
      "Project Name": local?.["Project Name"] || getCell(row, ["project name", "project", "project title"]) || "",
      "Activity ID": local?.["Activity ID"] || activityId,
      "Cost ID": local?.["Cost ID"] || costId,
    };
  });
};

const buildRowsFromActivitiesAndCosts = (activities, costs) => {
  const activitiesList = Array.isArray(activities) ? activities : [];
  const costsList = Array.isArray(costs) ? costs : [];
  const activityByProjectAndActivityId = new Map();

  activitiesList.forEach((activity) => {
    const projectId = String(activity?.projectId || activity?.project_id || "").trim();
    const activityId = String(
      activity?.activityRefId || activity?.activityId || activity?.activity_id || activity?.id || ""
    ).trim();
    if (!projectId || !activityId) return;
    activityByProjectAndActivityId.set(makeDashboardCompositeKey({ projectId, activityId, costId: "" }), activity);
  });

  const rows = [];
  const costBackedActivityKeys = new Set();

  costsList.forEach((cost) => {
    const projectId = String(cost?.projectId || cost?.project_id || "").trim();
    const activityId = String(
      cost?.activityId || cost?.activity_id || cost?.activityRefId || cost?.activity_ref_id || ""
    ).trim();
    const costId = String(cost?.costId || cost?.cost_id || cost?.id || "").trim();
    if (!projectId || (!activityId && !costId)) return;

    const activityKey = makeDashboardCompositeKey({ projectId, activityId, costId: "" });
    const matchedActivity = activityByProjectAndActivityId.get(activityKey);
    if (activityId) costBackedActivityKeys.add(activityKey);

    const projectName = String(
      cost?.project || cost?.projectName || matchedActivity?.project || matchedActivity?.projectName || ""
    ).trim();
    const activityName = String(
      cost?.activity || cost?.activityName || matchedActivity?.activity || matchedActivity?.name || (activityId ? `Activity ${activityId}` : `Cost ${costId}`)
    ).trim();

    rows.push({
      "Project ID": projectId,
      "Project Name": projectName,
      "Activity ID": activityId,
      "Cost ID": costId,
      Activity: activityName,
      "Planned Cost": parseNumber(cost?.plannedCost ?? cost?.planned_cost ?? matchedActivity?.plannedValue),
      "Actual Cost": parseNumber(cost?.actualCost ?? cost?.actual_cost),
      "Progress": parseNumber(matchedActivity?.percentComplete ?? matchedActivity?.progress ?? cost?.progress ?? cost?.percentComplete ?? 0),
      "Earned Value": parseNumber(cost?.earnedValue ?? cost?.earned_value),
      "Planned Start": matchedActivity?.plannedStart || matchedActivity?.startDate || "",
      "Planned Finish": matchedActivity?.plannedFinish || matchedActivity?.finishDate || "",
    });
  });

  activitiesList.forEach((activity) => {
    const projectId = String(activity?.projectId || activity?.project_id || "").trim();
    const activityId = String(
      activity?.activityRefId || activity?.activityId || activity?.activity_id || activity?.id || ""
    ).trim();
    const activityKey = makeDashboardCompositeKey({ projectId, activityId, costId: "" });
    if (!projectId || !activityId || costBackedActivityKeys.has(activityKey)) return;

    rows.push({
      "Project ID": projectId,
      "Project Name": String(activity?.project || activity?.projectName || "").trim(),
      "Activity ID": activityId,
      "Cost ID": "",
      Activity: String(activity?.activity || activity?.name || `Activity ${activityId}`).trim(),
      "Planned Cost": parseNumber(activity?.plannedValue ?? activity?.plannedCost),
      "Actual Cost": parseNumber(activity?.actualCost),
      "Progress": parseNumber(activity?.percentComplete ?? activity?.progress ?? activity?.completion ?? 0),
      "Earned Value": parseNumber(activity?.earnedValue),
      "Planned Start": activity?.plannedStart || activity?.startDate || "",
      "Planned Finish": activity?.plannedFinish || activity?.finishDate || "",
    });
  });

  return rows;
};

const hydrateDashboardFromCache = () => {
  const cached = localStorage.getItem(DASHBOARD_CACHE_KEY);
  if (!cached) return;

  try {
    const parsedCache = JSON.parse(cached);
    const rows = Array.isArray(parsedCache) ? parsedCache : parsedCache?.rows;
    const savedAt = Array.isArray(parsedCache) ? 0 : Number(parsedCache?.savedAt || 0);
    if (!Array.isArray(rows) || !rows.length) return;
    if (savedAt && Date.now() - savedAt > DASHBOARD_CACHE_TTL_MS) return;
    latestDashboardSignature = JSON.stringify(rows);
    activitySummaryRows = rows;
    syncFilterOptionsFromRows(rows);
    applyFiltersAndRender();
    showMessage(
      `Loaded ${rows.length} cached activity row(s). Verifying against live source now...`
    );
  } catch {
    localStorage.removeItem(DASHBOARD_CACHE_KEY);
  }
};

const refreshDashboardData = async ({ force = false } = {}) => {
  if (isDashboardFetchInFlight) return;

  isDashboardFetchInFlight = true;
  try {
    const localRows = loadRowsFromCostManagementLocalData();

    if (USE_COST_MANAGEMENT_ONLY) {
      if (localRows.length) {
        processRows(localRows, "Cost Management local storage");
        showMessage("Connected to Cost Management local data only.");
      } else {
        showMessage("No Cost Management local data found yet.", true);
      }
      return;
    }

    if (!DATA_SOURCE_URL.trim()) {
      if (localRows.length) {
        processRows(localRows, "Cost Management local storage");
        showMessage("Using Cost Management local data because no live data source URL is configured.");
      } else {
        processRows([], "local storage");
        showMessage("No dashboard data source URL configured and no local Cost Management data found.", true);
      }
      return;
    }

    if (force) {
      showMessage("Loading data source...");
    }

    if (typeof window.DataBridge.fetchDashboardBundleFromSource === "function") {
      try {
        const { activities, costs, sourceName } = await window.DataBridge.fetchDashboardBundleFromSource(
          DATA_SOURCE_URL
        );
        const bundleRows = buildRowsFromActivitiesAndCosts(activities, costs);
        if (bundleRows.length) {
          const enrichedBundleRows = mergeRemoteRowsWithCostManagementActuals(bundleRows, localRows);
          processRows(enrichedBundleRows, sourceName);
          return;
        }
      } catch (bundleError) {
        console.warn("Bundle fetch failed. Falling back to row-based fetch.", bundleError);
      }
    }

    const { rows, sourceName } = await window.DataBridge.fetchRowsFromSource(DATA_SOURCE_URL);

    if (Array.isArray(rows) && rows.length) {
      const enrichedRows = mergeRemoteRowsWithCostManagementActuals(rows, localRows);
      processRows(enrichedRows, sourceName);
      return;
    }

    if (localRows.length) {
      processRows(localRows, "Cost Management local storage");
      showMessage("Live source returned no rows. Showing Cost Management local data.");
    } else {
      processRows([], sourceName);
    }
  } catch (error) {
    const localRows = loadRowsFromCostManagementLocalData();
    if (localRows.length) {
      processRows(localRows, "Cost Management local storage");
      showMessage("Connected to Cost Management local data. Live source is temporarily unavailable.");
    } else {
      showMessage(`Error loading data source: ${error.message}`, true);
    }
  } finally {
    isDashboardFetchInFlight = false;
  }
};


const handleRealtimeStorageSync = (event) => {
  const trackedKeys = new Set([
    COST_ACTIVITIES_LOCAL_STORAGE_KEY,
    DAILY_COSTS_LOCAL_STORAGE_KEY,
    DASHBOARD_CACHE_KEY,
  ]);

  if (event?.key && !trackedKeys.has(event.key)) return;
  refreshDashboardData({ force: true });
};

const setupRealtimeDashboardSync = () => {
  if (dashboardRefreshTimer) {
    clearInterval(dashboardRefreshTimer);
  }

  dashboardRefreshTimer = setInterval(() => {
    refreshDashboardData();
  }, DASHBOARD_REFRESH_INTERVAL_MS);

  window.addEventListener("focus", () => refreshDashboardData({ force: true }));
  window.addEventListener("online", () => refreshDashboardData({ force: true }));
  document.addEventListener("visibilitychange", () => {
    if (document.visibilityState === "visible") {
      refreshDashboardData({ force: true });
    }
  });
  window.addEventListener("storage", handleRealtimeStorageSync);
};

const setupServiceWorkerUpdates = async () => {
  if (!("serviceWorker" in navigator)) {
    return;
  }

  try {
    const registration = await navigator.serviceWorker.register("./service-worker.js", { updateViaCache: "none" });

    const requestImmediateActivation = () => {
      if (registration.waiting) {
        registration.waiting.postMessage({ type: "SKIP_WAITING" });
      }
    };

    registration.addEventListener("updatefound", () => {
      const newWorker = registration.installing;
      if (!newWorker) return;

      newWorker.addEventListener("statechange", () => {
        if (newWorker.state === "installed" && navigator.serviceWorker.controller) {
          requestImmediateActivation();
        }
      });
    });

    navigator.serviceWorker.addEventListener("controllerchange", () => {
      window.location.reload();
    });

    requestImmediateActivation();
    registration.update();

    setInterval(() => {
      registration.update();
    }, 60 * 1000);

    window.addEventListener("focus", () => {
      registration.update();
    });

    document.addEventListener("visibilitychange", () => {
      if (document.visibilityState === "visible") {
        registration.update();
      }
    });
  } catch (error) {
    console.warn("Service worker setup failed:", error);
  }
};

if (projectFilterEl) projectFilterEl.addEventListener("change", applyFiltersAndRender);
if (dateStartFilterEl) dateStartFilterEl.addEventListener("change", applyFiltersAndRender);
if (dateEndFilterEl) dateEndFilterEl.addEventListener("change", applyFiltersAndRender);

setupServiceWorkerUpdates();
hydrateDashboardFromCache();
refreshDashboardData({ force: true });
setupRealtimeDashboardSync();
