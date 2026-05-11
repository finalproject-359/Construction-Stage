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
const varianceBreakdownEl = document.querySelector(".breakdown-donut");
const projectFilterEl = document.getElementById("projectFilter");
const dateStartFilterEl = document.getElementById("dateStartFilter");
const dateEndFilterEl = document.getElementById("dateEndFilter");
const activeProjectCountEl = document.getElementById("activeProjectCount");
const dashboardRiskLevelEl = document.getElementById("dashboardRiskLevel");

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
const DASHBOARD_CACHE_TTL_MS = 30 * 60 * 1000;
const DASHBOARD_REFRESH_INTERVAL_MS = 30 * 1000;
const DASHBOARD_MIN_STABLE_ROW_RATIO = 0.5;

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

const formatPercent = (value, maximumFractionDigits = 1) =>
  `${parseNumber(value).toFixed(maximumFractionDigits)}%`;

const truncateChartLabel = (value, maxLength = 22) => {
  const label = normalize(value, "Untitled");
  return label.length > maxLength ? `${label.slice(0, maxLength - 1)}…` : label;
};

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
  if (!row || typeof row !== "object") return null;

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

const firstNonEmptyCell = (row, aliases, fallback = "") => {
  for (const alias of aliases) {
    const value = getCell(row, [alias]);
    if (value !== null && value !== undefined && String(value).trim() !== "") {
      return value;
    }
  }
  return fallback;
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

const normalizeDateOnly = (value) => {
  const raw = normalize(value, "");
  if (!raw) return "";
  const date = new Date(raw);
  if (Number.isNaN(date.getTime())) return "";
  return date.toISOString().split("T")[0];
};

const getDateSortTime = (value) => {
  const normalizedDate = normalizeDateOnly(value);
  if (!normalizedDate) return Number.POSITIVE_INFINITY;
  const timestamp = new Date(`${normalizedDate}T00:00:00`).getTime();
  return Number.isFinite(timestamp) ? timestamp : Number.POSITIVE_INFINITY;
};

const compareDashboardRowsByStartPriority = (a, b) => {
  const startDiff = getDateSortTime(a?.startDate) - getDateSortTime(b?.startDate);
  if (startDiff) return startDiff;

  const finishDiff = getDateSortTime(a?.finishDate) - getDateSortTime(b?.finishDate);
  if (finishDiff) return finishDiff;

  const projectDiff = normalize(a?.project, "").localeCompare(normalize(b?.project, ""));
  if (projectDiff) return projectDiff;

  const activityIdDiff = formatActivityCostIdentity(a).localeCompare(formatActivityCostIdentity(b));
  if (activityIdDiff) return activityIdDiff;

  const activityDiff = normalize(a?.activity, "").localeCompare(normalize(b?.activity, ""));
  if (activityDiff) return activityDiff;

  return parseNumber(a?.sourceIndex) - parseNumber(b?.sourceIndex);
};

const sortDashboardRowsByStartPriority = (rows) =>
  Array.isArray(rows) ? [...rows].sort(compareDashboardRowsByStartPriority) : [];

const extractDashboardRows = (rawRows) =>
  rawRows
    .map((row, index) => {
      const detectedActivityId = normalize(
        firstNonEmptyCell(row, [
          "activity id",
          "activity id/cost id",
          "activity code",
          "activity ref id",
          "activity reference id",
          "source activity id",
          "wbs",
          "task id",
          "id",
        ]),
        ""
      );
      const detectedCostId = normalize(
        firstNonEmptyCell(row, ["cost id", "cost code", "cost_id", "costid"]),
        ""
      );
      const activity = normalize(
        firstNonEmptyCell(row, [
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

      const projectId = normalize(firstNonEmptyCell(row, ["project id", "project code", "projectid", "code"]), "");
      const projectName = normalize(firstNonEmptyCell(row, ["project name", "project", "project title"]), "");
      const project = projectId && projectName
        ? `${projectId} - ${projectName}`
        : projectId || projectName || "No Project ID";
      const startDate = normalizeDateOnly(getCell(row, ["planned start", "start date", "start"]));
      const finishDate = normalizeDateOnly(getCell(row, ["planned finish", "finish date", "end date", "finish"]));
      const plannedCost = parseNumber(
        firstNonEmptyCell(row, ["planned value", "planned cost", "planned cost/day", "planned cost per day", "total budget", "pv", "budget"])
      );
      const actualCost = parseNumber(
        firstNonEmptyCell(row, ["actual cost", "actual cost/day", "actual cost per day", "total spent", "ac", "actual", "cost", "amount"])
      );
      const percentComplete = normalizeProgressPercent(
        firstNonEmptyCell(row, ["% complete", "percent complete", "progress %", "progress", "progress/day", "completion"] )
      );
      const rawEarnedValue = firstNonEmptyCell(row, ["earned value", "earned value/day", "ev"]);
      const hasProvidedEarnedValue =
        rawEarnedValue !== null && rawEarnedValue !== undefined && String(rawEarnedValue).trim() !== "";
      const ev = hasProvidedEarnedValue ? parseNumber(rawEarnedValue) : plannedCost * (percentComplete / 100);
      const rawCv = firstNonEmptyCell(row, ["cost variance", "cv"]);
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
        sourceIndex: index,
      };
    })
    .filter(
      (row) =>
        row &&
        (row.plannedCost !== 0 || row.actualCost !== 0 || row.ev !== 0 || row.percentComplete !== 0)
    )
    .sort(compareDashboardRowsByStartPriority);

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

  return sortDashboardRowsByStartPriority(
    rows.filter((row) => {
      const projectMatches = selectedProject === "all"
        || normalize(row.project, "").trim().toLowerCase() === selectedProject;
      return projectMatches && rowMatchesDateFilter(row, startDate, endDate);
    })
  );
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


const updateDashboardSummary = (rows, totals, progressMetrics) => {
  const activityCount = rows.length;
  const uniqueProjects = new Set(rows.map((row) => normalize(row.project, "Unspecified"))).size;
  const overBudgetCount = rows.filter((row) => parseNumber(row.cv) < 0).length;
  const actualCost = parseNumber(totals?.actual);
  const earnedValue = (parseNumber(totals?.planned) * parseNumber(progressMetrics?.physicalProgressPercent)) / 100;
  const cpi = actualCost ? earnedValue / actualCost : 0;

  if (activeProjectCountEl) {
    const activityLabel = activityCount === 1 ? "activity" : "activities";
    const projectLabel = uniqueProjects === 1 ? "project" : "projects";
    activeProjectCountEl.textContent = activityCount
      ? `${activityCount} ${activityLabel} · ${uniqueProjects} ${projectLabel}`
      : "No active data";
  }

  if (dashboardRiskLevelEl) {
    if (!activityCount) {
      dashboardRiskLevelEl.textContent = "Awaiting data";
    } else if (!actualCost) {
      dashboardRiskLevelEl.textContent = "Awaiting actual costs";
    } else if (parseNumber(totals?.cv) < 0 || cpi < 1) {
      const alertLabel = overBudgetCount === 1 ? "cost alert" : "cost alerts";
      dashboardRiskLevelEl.textContent = `${overBudgetCount} ${alertLabel} · CPI ${cpi.toFixed(2)}`;
    } else {
      dashboardRiskLevelEl.textContent = `On track · CPI ${cpi.toFixed(2)}`;
    }
  }
};

const renderDashboardFromRows = (rows, summaryRows = rows) => {
  const orderedRows = sortDashboardRowsByStartPriority(rows);
  const orderedSummaryRows = summaryRows === rows ? orderedRows : sortDashboardRowsByStartPriority(summaryRows);
  const totals = calculateTotalsFromRows(orderedRows);
  const progressMetrics = calculateProgressMetrics(orderedRows, totals);
  renderKpis(totals);
  renderProgressKpis(progressMetrics, totals);
  updateDashboardSummary(orderedRows, totals, progressMetrics);
  renderTable(orderedSummaryRows);
  renderOverrunTable(orderedRows);
  renderGapTable(orderedRows);
  generateCharts(orderedRows);
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


  const totalCv = parseNumber(totals?.cv);
  const totalPlanned = parseNumber(totals?.planned);
  const totalActual = parseNumber(totals?.actual);
  const unfavorableShare = totalPlanned ? Math.max(0, Math.min(100, (totalActual / totalPlanned) * 100)) : 0;

  if (varianceDisplayEl) varianceDisplayEl.textContent = formatCurrency(totalCv);
  if (varianceStatusEl) varianceStatusEl.textContent = totalCv < 0 ? "Over Budget" : "Under Budget";
  if (varianceBreakdownEl) {
    varianceBreakdownEl.style.setProperty("--cost-used-share", `${unfavorableShare}%`);
    varianceBreakdownEl.setAttribute(
      "aria-label",
      `Cost used ${formatPercent(unfavorableShare)} of approved budget`
    );
  }
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

const buildDashboardChartOptions = ({ yAxis, valueFormatter, legendPosition = "bottom" } = {}) => ({
  responsive: true,
  maintainAspectRatio: false,
  animation: { duration: 260 },
  interaction: { mode: "index", intersect: false },
  plugins: {
    legend: {
      display: true,
      position: legendPosition,
      align: "start",
      labels: {
        boxWidth: 10,
        boxHeight: 10,
        color: "#475467",
        font: { size: 12, weight: "700" },
        usePointStyle: true,
        padding: 16,
      },
    },
    tooltip: {
      backgroundColor: "#0f172a",
      borderColor: "rgba(255,255,255,0.12)",
      borderWidth: 1,
      padding: 12,
      titleFont: { weight: "800" },
      bodyFont: { weight: "600" },
      callbacks: {
        title: (items) => labelsFromChartItems(items).join(""),
        label: (context) => {
          const label = context.dataset.label || "Value";
          const value = context.parsed?.y ?? context.parsed;
          return `${label}: ${valueFormatter ? valueFormatter(value) : value}`;
        },
      },
    },
  },
  scales: {
    x: {
      grid: { display: false },
      ticks: {
        color: "#667085",
        font: { size: 11, weight: "700" },
        maxRotation: 0,
        callback: function tickLabel(value) {
          return truncateChartLabel(this.getLabelForValue(value));
        },
      },
    },
    y: {
      ...yAxis,
      border: { display: false },
      grid: { color: "rgba(148, 163, 184, 0.2)" },
      ticks: {
        color: "#667085",
        font: { size: 11, weight: "700" },
        ...(yAxis?.ticks || {}),
      },
    },
  },
});

const labelsFromChartItems = (items = []) => {
  const firstItem = items[0];
  const sourceLabels = firstItem?.chart?.data?.labels || [];
  const label = sourceLabels[firstItem?.dataIndex] || firstItem?.label || "";
  return [label];
};

const updateOrCreateChart = (existingChart, canvas, config) => {
  if (!existingChart) return new Chart(canvas, config);
  existingChart.data = config.data;
  existingChart.options = config.options;
  existingChart.update("none");
  return existingChart;
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
  const varianceValues = rows.map((row) => row.cv);
  const costAxis = buildLinearAxisRange(varianceValues, { includeZero: true, targetTickCount: 6 });

  const varianceCanvas = document.getElementById("varianceChart");
  const costCanvas = document.getElementById("costChart");
  if (!varianceCanvas || !costCanvas) {
    showMessage("Dashboard chart containers are missing. Refresh or restore dashboard markup.", true);
    return;
  }

  const sharedLineStyles = {
    pointRadius: 3,
    pointHoverRadius: 5,
    borderWidth: 3,
    tension: 0.35,
    fill: false,
  };

  varianceChart = updateOrCreateChart(varianceChart, varianceCanvas, {
    type: "line",
    data: {
      labels,
      datasets: [
        {
          ...sharedLineStyles,
          label: "Progress",
          data: completeSeries,
          borderColor: "#2f55ff",
          backgroundColor: "#2f55ff",
          pointBackgroundColor: "#ffffff",
          pointBorderColor: "#2f55ff",
        },
        {
          ...sharedLineStyles,
          label: "% Cost Used",
          data: costUsedSeries,
          borderColor: "#16a34a",
          backgroundColor: "#16a34a",
          pointBackgroundColor: "#ffffff",
          pointBorderColor: "#16a34a",
        },
      ],
    },
    options: buildDashboardChartOptions({
      valueFormatter: (value) => formatPercent(value),
      yAxis: {
        min: 0,
        max: 100,
        ticks: { callback: (value) => `${value}%` },
      },
    }),
  });

  costChart = updateOrCreateChart(costChart, costCanvas, {
    type: "bar",
    data: {
      labels,
      datasets: [
        {
          label: "Cost Variance (EV - AC)",
          data: varianceValues,
          backgroundColor: rows.map((row) => (row.cv >= 0 ? "rgba(34, 197, 94, 0.84)" : "rgba(239, 68, 68, 0.84)")),
          borderColor: rows.map((row) => (row.cv >= 0 ? "#16a34a" : "#dc2626")),
          borderWidth: 1,
          borderRadius: 10,
          borderSkipped: false,
          maxBarThickness: 42,
        },
      ],
    },
    options: buildDashboardChartOptions({
      valueFormatter: formatCurrency,
      yAxis: {
        min: costAxis.min,
        max: costAxis.max,
        ticks: {
          stepSize: costAxis.stepSize,
          callback: (value) => formatCurrency(value),
        },
      },
    }),
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
  if (!rows.length && activitySummaryRows.length) {
    showMessage(
      `Live source returned an empty parsed result from ${sourceName}. Keeping the last stable ${activitySummaryRows.length} activity row(s).`,
      true
    );
    return;
  }

  const previousRowCount = activitySummaryRows.length;
  const droppedTooMuch =
    previousRowCount > 2 && rows.length > 0 && rows.length < previousRowCount * DASHBOARD_MIN_STABLE_ROW_RATIO;
  if (droppedTooMuch) {
    showMessage(
      `Live source returned only ${rows.length} parsed row(s), down from ${previousRowCount}. Keeping the last stable dashboard data until the next sync.`,
      true
    );
    return;
  }

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
    asOfEl.textContent = `Synced ${new Date().toLocaleString()}`;
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
      "Planned Start": activity?.plannedStart || activity?.plannedStartDate || activity?.startDate || "",
      "Planned Finish": activity?.plannedFinish || activity?.plannedFinishDate || activity?.finishDate || "",
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
        firstNonEmptyCell(local, ["Actual Cost"], null) !== null
          ? parseNumber(local?.["Actual Cost"])
          : parseNumber(firstNonEmptyCell(row, ["actual cost", "actual cost/day", "total spent", "ac", "actual"])),
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
  const readBundleCell = (row, aliases, fallback = "") => firstNonEmptyCell(row, aliases, fallback);

  const getActivityProjectId = (activity) =>
    String(readBundleCell(activity, ["project id", "projectId", "project_id", "project code", "projectCode"])).trim();
  const getActivityId = (activity) =>
    String(
      readBundleCell(activity, [
        "activity id",
        "activityId",
        "activity_id",
        "activity ref id",
        "activityRefId",
        "id",
      ])
    ).trim();
  const getCostProjectId = (cost) =>
    String(readBundleCell(cost, ["project id", "projectId", "project_id", "project code", "projectCode"])).trim();
  const getCostActivityId = (cost) =>
    String(
      readBundleCell(cost, [
        "activity id",
        "activityId",
        "activity_id",
        "activity ref id",
        "activityRefId",
        "activity_ref_id",
        "source activity id",
      ])
    ).trim();
  const getCostId = (cost) =>
    String(readBundleCell(cost, ["cost id", "costId", "cost_id", "id"])).trim();

  activitiesList.forEach((activity) => {
    const projectId = getActivityProjectId(activity);
    const activityId = getActivityId(activity);
    if (!projectId || !activityId) return;
    activityByProjectAndActivityId.set(makeDashboardCompositeKey({ projectId, activityId, costId: "" }), activity);
  });

  const rows = [];
  const costBackedActivityKeys = new Set();

  costsList.forEach((cost) => {
    const projectId = getCostProjectId(cost);
    const activityId = getCostActivityId(cost);
    const costId = getCostId(cost);
    if (!projectId || (!activityId && !costId)) return;

    const activityKey = makeDashboardCompositeKey({ projectId, activityId, costId: "" });
    const matchedActivity = activityByProjectAndActivityId.get(activityKey);
    if (activityId) costBackedActivityKeys.add(activityKey);

    const projectName = String(
      readBundleCell(cost, ["project", "project name", "projectName"], "") ||
        readBundleCell(matchedActivity, ["project", "project name", "projectName"], "")
    ).trim();
    const activityName = String(
      readBundleCell(cost, ["activity", "activity name", "activityName", "name"], "") ||
        readBundleCell(matchedActivity, ["activity", "activity name", "activityName", "name"], "") ||
        (activityId ? `Activity ${activityId}` : `Cost ${costId}`)
    ).trim();
    const plannedCost = parseNumber(
      readBundleCell(cost, ["planned cost", "plannedCost", "planned_cost", "planned value", "plannedValue"], "") ||
        readBundleCell(matchedActivity, ["planned value", "plannedValue", "planned cost", "plannedCost"], "")
    );
    const progress = parseNumber(
      readBundleCell(matchedActivity, ["percent complete", "percentComplete", "progress", "completion"], "") ||
        readBundleCell(cost, ["progress", "percent complete", "percentComplete", "progress/day"], 0)
    );
    const providedEarnedValue = readBundleCell(cost, ["earned value", "earnedValue", "earned_value", "earned value/day"], "");

    rows.push({
      "Project ID": projectId,
      "Project Name": projectName,
      "Activity ID": activityId,
      "Cost ID": costId,
      Activity: activityName,
      "Planned Cost": plannedCost,
      "Actual Cost": parseNumber(readBundleCell(cost, ["actual cost", "actualCost", "actual_cost", "actual cost/day", "cost", "amount"], 0)),
      "Progress": progress,
      "Earned Value": providedEarnedValue !== "" ? parseNumber(providedEarnedValue) : plannedCost * (progress / 100),
      "Planned Start": readBundleCell(matchedActivity, ["planned start", "plannedStart", "start date", "startDate"], ""),
      "Planned Finish": readBundleCell(matchedActivity, ["planned finish", "plannedFinish", "finish date", "finishDate"], ""),
    });
  });

  activitiesList.forEach((activity) => {
    const projectId = getActivityProjectId(activity);
    const activityId = getActivityId(activity);
    const activityKey = makeDashboardCompositeKey({ projectId, activityId, costId: "" });
    if (!projectId || !activityId || costBackedActivityKeys.has(activityKey)) return;
    const plannedCost = parseNumber(readBundleCell(activity, ["planned value", "plannedValue", "planned cost", "plannedCost"], 0));
    const progress = parseNumber(readBundleCell(activity, ["percent complete", "percentComplete", "progress", "completion"], 0));
    const providedEarnedValue = readBundleCell(activity, ["earned value", "earnedValue", "earned_value"], "");

    rows.push({
      "Project ID": projectId,
      "Project Name": String(readBundleCell(activity, ["project", "project name", "projectName"], "")).trim(),
      "Activity ID": activityId,
      "Cost ID": "",
      Activity: String(readBundleCell(activity, ["activity", "activity name", "activityName", "name"], `Activity ${activityId}`)).trim(),
      "Planned Cost": plannedCost,
      "Actual Cost": parseNumber(readBundleCell(activity, ["actual cost", "actualCost", "actual_cost"], 0)),
      "Progress": progress,
      "Earned Value": providedEarnedValue !== "" ? parseNumber(providedEarnedValue) : plannedCost * (progress / 100),
      "Planned Start": readBundleCell(activity, ["planned start", "plannedStart", "start date", "startDate"], ""),
      "Planned Finish": readBundleCell(activity, ["planned finish", "plannedFinish", "finish date", "finishDate"], ""),
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
    const cacheIsStale = savedAt && Date.now() - savedAt > DASHBOARD_CACHE_TTL_MS;
    latestDashboardSignature = JSON.stringify(rows);
    activitySummaryRows = rows;
    syncFilterOptionsFromRows(rows);
    applyFiltersAndRender();
    showMessage(
      cacheIsStale
        ? `Showing the last saved ${rows.length} dashboard row(s) while refreshing live data...`
        : `Loaded ${rows.length} cached activity row(s). Verifying against live source now...`
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

    if (localRows.length && !activitySummaryRows.length) {
      processRows(localRows, "Cost Management local storage");
      showMessage("Live source returned no rows. Showing Cost Management local data.");
    } else {
      processRows([], sourceName);
    }
  } catch (error) {
    const localRows = loadRowsFromCostManagementLocalData();
    if (activitySummaryRows.length) {
      showMessage(`Live source is temporarily unavailable (${error.message}). Keeping the last stable dashboard data.`, true);
    } else if (localRows.length) {
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
      showMessage("Dashboard update installed. Continue working; refresh manually when convenient.");
    });

    requestImmediateActivation();
    registration.update();

    setInterval(() => {
      registration.update();
    }, 10 * 60 * 1000);

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
