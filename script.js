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
const dateRangeFilterEl = document.getElementById("dateRangeFilter");
const customDateRangeFieldsEl = document.getElementById("customDateRangeFields");
const dateRangeSelectWrapEl = customDateRangeFieldsEl?.closest(".date-range-select-wrap");
const activitySummarySortEl = document.getElementById("activitySummarySort");
const activeProjectCountEl = document.getElementById("activeProjectCount");
const dashboardRiskLevelEl = document.getElementById("dashboardRiskLevel");
const exportReportButtonEl = document.getElementById("exportReportButton");
const exportReportOptionsEl = document.getElementById("exportReportOptions");

const DATA_SOURCE_URL = window.DataBridge?.DEFAULT_DATA_SOURCE_URL || "";
const chartDependencyWarning =
  typeof window.Chart === "undefined"
    ? "Chart.js is not available. Graphs are disabled."
    : "";

let activitySummaryRows = [];
let varianceChart = null;
let costChart = null;
let isDashboardFetchInFlight = false;
let latestDashboardSignature = "";
let lastStableDashboardRows = [];
let lastStableDashboardSavedAt = 0;

const DASHBOARD_REFRESH_INTERVAL_MS = 30 * 1000;
let dashboardRefreshTimer = null;
let pendingDashboardRefreshRequested = false;

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
      const loggedDate = normalizeDateOnly(
        getCell(row, ["activity date", "report date", "actual date", "date logged", "date"])
      );
      const startDate = normalizeDateOnly(
        getCell(row, ["actual start", "actual start date", "planned start", "start date", "start", "activity start"])
      ) || loggedDate;
      const finishDate = normalizeDateOnly(
        getCell(row, ["actual finish", "actual finish date", "planned finish", "finish date", "end date", "finish", "activity finish"])
      ) || loggedDate;
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
      const cv = ev - actualCost;

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

const buildLocalDate = (year, month, day) => {
  const date = new Date(year, month - 1, day);
  if (
    date.getFullYear() !== year ||
    date.getMonth() !== month - 1 ||
    date.getDate() !== day
  ) {
    return null;
  }
  return date;
};

const parseDateParts = (value) => {
  if (value instanceof Date) {
    return Number.isNaN(value.getTime()) ? null : value;
  }

  if (typeof value === "number" && Number.isFinite(value)) {
    const excelEpoch = Date.UTC(1899, 11, 30);
    const serialDate = new Date(excelEpoch + value * 24 * 60 * 60 * 1000);
    if (Number.isNaN(serialDate.getTime())) return null;
    return buildLocalDate(
      serialDate.getUTCFullYear(),
      serialDate.getUTCMonth() + 1,
      serialDate.getUTCDate()
    );
  }

  const raw = normalize(value, "");
  if (!raw) return null;

  const isoMatch = raw.match(/^(\d{4})[-/](\d{1,2})[-/](\d{1,2})/);
  if (isoMatch) {
    return buildLocalDate(Number(isoMatch[1]), Number(isoMatch[2]), Number(isoMatch[3]));
  }

  const delimitedMatch = raw.match(/^(\d{1,2})[-/.](\d{1,2})[-/.](\d{2,4})$/);
  if (delimitedMatch) {
    const first = Number(delimitedMatch[1]);
    const second = Number(delimitedMatch[2]);
    const yearValue = Number(delimitedMatch[3]);
    const year = yearValue < 100 ? 2000 + yearValue : yearValue;
    const dayFirst = first > 12 && second <= 12;
    const month = dayFirst ? second : first;
    const day = dayFirst ? first : second;
    return buildLocalDate(year, month, day);
  }

  const parsedDate = new Date(raw);
  return Number.isNaN(parsedDate.getTime()) ? null : parsedDate;
};

const normalizeDateOnly = (value) => {
  const date = parseDateParts(value);
  return date ? formatDateInputValue(date) : "";
};

const formatDateInputValue = (date) => {
  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, "0");
  const day = String(date.getDate()).padStart(2, "0");
  return `${year}-${month}-${day}`;
};

const setCustomDateRangeFieldsVisibility = () => {
  if (!customDateRangeFieldsEl || !dateRangeFilterEl) return;

  const shouldShowCustomFields = dateRangeFilterEl.value === "custom";
  customDateRangeFieldsEl.hidden = !shouldShowCustomFields;
  dateRangeSelectWrapEl?.classList.toggle("is-custom-range", shouldShowCustomFields);
  dateStartFilterEl?.toggleAttribute("disabled", !shouldShowCustomFields);
  dateEndFilterEl?.toggleAttribute("disabled", !shouldShowCustomFields);
};

const syncDateRangeInputConstraints = () => {
  if (!dateStartFilterEl || !dateEndFilterEl) return;

  const startDate = String(dateStartFilterEl.value || "").trim();
  const endDate = String(dateEndFilterEl.value || "").trim();
  dateEndFilterEl.min = startDate;
  dateStartFilterEl.max = endDate;

  if (startDate && endDate && endDate < startDate) {
    dateEndFilterEl.value = startDate;
    dateStartFilterEl.max = startDate;
  }
};

const resolvePresetDateRange = (selectedRange) => {
  if (selectedRange === "all" || selectedRange === "custom") return null;

  const today = new Date();
  const startDate = new Date(today);
  const endDate = new Date(today);

  if (selectedRange === "today") {
    return { startDate, endDate };
  }

  if (selectedRange === "week") {
    const dayOfWeek = today.getDay();
    const daysSinceMonday = dayOfWeek === 0 ? 6 : dayOfWeek - 1;
    startDate.setDate(today.getDate() - daysSinceMonday);
    endDate.setDate(startDate.getDate() + 6);
    return { startDate, endDate };
  }

  if (selectedRange === "month") {
    startDate.setDate(1);
    endDate.setMonth(today.getMonth() + 1, 0);
    return { startDate, endDate };
  }

  if (selectedRange === "year") {
    startDate.setMonth(0, 1);
    endDate.setMonth(11, 31);
    return { startDate, endDate };
  }

  return { startDate, endDate };
};

const updateDateRangeFilterValues = () => {
  if (!dateStartFilterEl || !dateEndFilterEl || !dateRangeFilterEl) return;

  const selectedRange = dateRangeFilterEl.value;
  setCustomDateRangeFieldsVisibility();

  if (selectedRange === "all") {
    dateStartFilterEl.value = "";
    dateEndFilterEl.value = "";
    syncDateRangeInputConstraints();
    return;
  }

  if (selectedRange === "custom") {
    syncDateRangeInputConstraints();
    return;
  }

  const presetRange = resolvePresetDateRange(selectedRange);
  if (!presetRange) {
    syncDateRangeInputConstraints();
    return;
  }

  dateStartFilterEl.value = formatDateInputValue(presetRange.startDate);
  dateEndFilterEl.value = formatDateInputValue(presetRange.endDate);
  syncDateRangeInputConstraints();
};

const handleDateRangeInputChange = () => {
  if (dateRangeFilterEl) dateRangeFilterEl.value = "custom";
  setCustomDateRangeFieldsVisibility();
  syncDateRangeInputConstraints();
  applyFiltersAndRender();
};

const rowMatchesDateFilter = (row, startDate, endDate, options = {}) => {
  if (!startDate && !endDate) return true;
  const normalizedRowStart = normalizeDateOnly(row.startDate) || normalizeDateOnly(row.date);
  const normalizedRowEnd = normalizeDateOnly(row.finishDate) || normalizedRowStart;
  const rowStart = normalizedRowStart && normalizedRowEnd && normalizedRowStart > normalizedRowEnd
    ? normalizedRowEnd
    : normalizedRowStart;
  const rowEnd = normalizedRowStart && normalizedRowEnd && normalizedRowStart > normalizedRowEnd
    ? normalizedRowStart
    : normalizedRowEnd;

  if (!rowStart && !rowEnd) return false;

  if (startDate && rowEnd && rowEnd < startDate) return false;
  if (endDate && rowStart && rowStart > endDate) return false;
  return true;
};

const getFilteredRows = (rows) => {
  const selectedProject = String(projectFilterEl?.value || "all").trim().toLowerCase();
  const startDate = String(dateStartFilterEl?.value || "").trim();
  const endDate = String(dateEndFilterEl?.value || "").trim();

  return rows
    .filter((row) => {
      const projectMatches = selectedProject === "all"
      || normalize(row.project, "").trim().toLowerCase() === selectedProject;
      if (!projectMatches) return false;
      if (!startDate && !endDate) return true;

      if (Array.isArray(row.dailyEntries) && row.dailyEntries.length) {
        return row.dailyEntries.some((entry) =>
          rowMatchesDateFilter({ startDate: entry.date, finishDate: entry.date }, startDate, endDate)
        );
      }

      const normalizedRowDate = normalizeDateOnly(row.date);
      if (normalizedRowDate) {
        return rowMatchesDateFilter({ startDate: normalizedRowDate, finishDate: normalizedRowDate }, startDate, endDate);
      }

      const normalizedRowStart = normalizeDateOnly(row.startDate);
      const normalizedRowEnd = normalizeDateOnly(row.finishDate) || normalizedRowStart;
      if (!normalizedRowStart && !normalizedRowEnd) return false;
      return rowMatchesDateFilter(
        { startDate: normalizedRowStart, finishDate: normalizedRowEnd },
        startDate,
        endDate
      );
    })
    .map((row) => {
      if (!Array.isArray(row.dailyEntries) || !row.dailyEntries.length) return row;
      if (!startDate && !endDate) return row;

      const matchingEntries = row.dailyEntries.filter((entry) =>
        rowMatchesDateFilter({ startDate: entry.date, finishDate: entry.date }, startDate, endDate)
      );
      if (!matchingEntries.length) return null;

      const actualCost = matchingEntries.reduce((sum, entry) => sum + parseNumber(entry.actualCost), 0);
      const evFromEntries = matchingEntries.reduce((sum, entry) => sum + parseNumber(entry.earnedValue), 0);
      const progressFromEntries = matchingEntries.reduce((sum, entry) => sum + parseNumber(entry.progress), 0);
      const firstDate = matchingEntries
        .map((entry) => normalizeDateOnly(entry.date))
        .filter(Boolean)
        .sort()[0] || row.startDate;
      const lastDate = matchingEntries
        .map((entry) => normalizeDateOnly(entry.date))
        .filter(Boolean)
        .sort()
        .slice(-1)[0] || row.finishDate;

      const percentComplete = Number.isFinite(progressFromEntries)
        ? progressFromEntries
        : parseNumber(row.percentComplete);
      const ev = Number.isFinite(evFromEntries)
        ? evFromEntries
        : (parseNumber(row.plannedCost) * (percentComplete / 100));

      return {
        ...row,
        actualCost,
        ev,
        cv: ev - actualCost,
        percentComplete,
        costUsedPercent: parseNumber(row.plannedCost) ? (actualCost / parseNumber(row.plannedCost)) * 100 : 0,
        startDate: firstDate,
        finishDate: lastDate,
        date: lastDate || firstDate || row.date || "",
      };
    })
    .filter(Boolean);
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
  const earnedValue = parseNumber(totals?.ev);
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
  const totals = calculateTotalsFromRows(rows);
  const progressMetrics = calculateProgressMetrics(rows, totals);
  renderKpis(totals);
  renderProgressKpis(progressMetrics, totals);
  updateDashboardSummary(rows, totals, progressMetrics);
  renderTable(summaryRows);
  renderOverrunTable(rows);
  renderGapTable(rows);
  generateCharts(rows);
};

const getActivitySummarySortValue = () => String(activitySummarySortEl?.value || "default");

const getRowSortDateValue = (row) => {
  const normalizedDate =
    normalizeDateOnly(row.startDate) || normalizeDateOnly(row.finishDate) || normalizeDateOnly(row.date);
  if (!normalizedDate) return null;
  const timestamp = new Date(`${normalizedDate}T00:00:00`).getTime();
  return Number.isFinite(timestamp) ? timestamp : null;
};

const sortActivitySummaryRows = (rows) => {
  const sortValue = getActivitySummarySortValue();
  if (sortValue === "default") return rows.slice();

  const sortedRows = rows.map((row, index) => ({ row, index }));
  const compareNumeric = (getValue, direction = "desc") => (a, b) => {
    const left = getValue(a.row);
    const right = getValue(b.row);
    const comparison = direction === "asc" ? left - right : right - left;
    return comparison || a.index - b.index;
  };
  const compareDates = (direction = "desc") => (a, b) => {
    const left = getRowSortDateValue(a.row);
    const right = getRowSortDateValue(b.row);
    if (left === null && right === null) return a.index - b.index;
    if (left === null) return 1;
    if (right === null) return -1;
    const comparison = direction === "asc" ? left - right : right - left;
    return comparison || a.index - b.index;
  };

  if (sortValue === "actual-cost-desc") {
    sortedRows.sort(compareNumeric((row) => parseNumber(row.actualCost), "desc"));
  } else if (sortValue === "actual-cost-asc") {
    sortedRows.sort(compareNumeric((row) => parseNumber(row.actualCost), "asc"));
  } else if (sortValue === "date-desc") {
    sortedRows.sort(compareDates("desc"));
  } else if (sortValue === "date-asc") {
    sortedRows.sort(compareDates("asc"));
  }

  return sortedRows.map((item) => item.row);
};

const applyFiltersAndRender = () => {
  const filteredRows = getFilteredRows(activitySummaryRows);
  const sortedSummaryRows = sortActivitySummaryRows(filteredRows);
  renderDashboardFromRows(filteredRows, sortedSummaryRows);
};

const buildExportRows = (rows) =>
  rows.map((row) => ({
    Project: normalize(row.project),
    Activity: normalize(row.activity),
    "Activity ID": normalize(row.activityId || row.costId || "-"),
    "Start Date": normalizeDateOnly(row.startDate) || normalize(row.startDate, "-"),
    "Finish Date": normalizeDateOnly(row.finishDate) || normalize(row.finishDate, "-"),
    "Planned Cost": parseNumber(row.plannedCost),
    "Actual Cost": parseNumber(row.actualCost),
    "Earned Value": parseNumber(row.ev),
    "Cost Variance": parseNumber(row.cv),
    "Percent Complete": parseNumber(row.percentComplete),
    "Cost Used %": parseNumber(row.costUsedPercent),
  }));

const exportDashboardReport = (format = "xlsx") => {
  const filteredRows = getFilteredRows(activitySummaryRows);
  if (!filteredRows.length) {
    showMessage("No filtered dashboard rows available for export.");
    return;
  }
  const exportRows = buildExportRows(filteredRows);
  const dateStamp = new Date().toISOString().slice(0, 10);
  const filenameBase = `costrack-dashboard-report-${dateStamp}`;

  if (format === "csv") {
    const worksheet = XLSX.utils.json_to_sheet(exportRows);
    const csv = XLSX.utils.sheet_to_csv(worksheet);
    const blob = new Blob([csv], { type: "text/csv;charset=utf-8;" });
    const link = document.createElement("a");
    link.href = URL.createObjectURL(blob);
    link.download = `${filenameBase}.csv`;
    link.click();
    URL.revokeObjectURL(link.href);
  } else {
    const workbook = XLSX.utils.book_new();
    const worksheet = XLSX.utils.json_to_sheet(exportRows);
    XLSX.utils.book_append_sheet(workbook, worksheet, "Dashboard Report");
    XLSX.writeFile(workbook, `${filenameBase}.xlsx`);
  }

  showMessage(`Exported ${filteredRows.length} row(s) as ${format.toUpperCase()}.`);
};

const calculateTotalsFromRows = (rows) => {
  const totals = rows.reduce(
    (acc, row) => {
      acc.planned += parseNumber(row.plannedCost);
      acc.actual += parseNumber(row.actualCost);
      acc.ev += parseNumber(row.ev);
      return acc;
    },
    { planned: 0, actual: 0, ev: 0, cv: 0 }
  );
  totals.cv = totals.ev - totals.actual;
  return totals;
};

const calculateProgressMetrics = (rows, totals) => {
  // The progress card is intended to compare field progress against budget burn.
  // Use the explicit % complete/progress values instead of EV/planned so manually
  // supplied earned values do not make the displayed progress look lower or higher
  // than the activity progress data shown elsewhere on the dashboard.
  const physicalProgressPercent = calculateWeightedPercentTotal(rows, "percentComplete");
  const costSpentPercent = totals.planned
    ? (totals.actual / totals.planned) * 100
    : calculateWeightedPercentTotal(rows, "costUsedPercent");
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
  const earnedValue = parseNumber(totals?.ev);
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

// Total progress is a planned-cost weighted average of the row progress values,
// so larger activities contribute proportionally to the summary percentage.
const calculateWeightedPercentTotal = (rows, percentKey) => {
  const weightedTotals = rows.reduce(
    (acc, row) => {
      const plannedCost = parseNumber(row.plannedCost);
      const percentValue = parseNumber(row[percentKey]);

      if (plannedCost > 0) {
        acc.weightedPercent += plannedCost * percentValue;
        acc.weight += plannedCost;
      } else {
        acc.unweightedPercent += percentValue;
        acc.unweightedCount += 1;
      }

      return acc;
    },
    { weightedPercent: 0, weight: 0, unweightedPercent: 0, unweightedCount: 0 }
  );

  if (weightedTotals.weight > 0) return weightedTotals.weightedPercent / weightedTotals.weight;
  return weightedTotals.unweightedCount > 0
    ? weightedTotals.unweightedPercent / weightedTotals.unweightedCount
    : 0;
};

// Total % Cost Used is the overall budget burn: total actual cost / total planned cost.
const calculateAggregateCostUsedPercent = (rows, totals) => {
  if (totals.planned) return (totals.actual / totals.planned) * 100;
  return calculateWeightedPercentTotal(rows, "costUsedPercent");
};

const calculateActivitiesPerformanceTotals = (rows) => {
  const totals = calculateTotalsFromRows(rows);
  const aggregateCompletePercent = calculateWeightedPercentTotal(rows, "percentComplete");
  const aggregateCostUsedPercent = calculateAggregateCostUsedPercent(rows, totals);
  const aggregateCpi = totals.actual ? totals.ev / totals.actual : 0;

  return {
    ...totals,
    aggregateCompletePercent,
    aggregateCostUsedPercent,
    aggregateCpi,
  };
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
        const totals = calculateActivitiesPerformanceTotals(rows);

        return `<tr>
          <td></td>
          <td><strong>TOTAL</strong></td>
          <td><strong>${formatCurrency(totals.planned)}</strong></td>
          <td><strong>${formatCurrency(totals.actual)}</strong></td>
          <td><strong>${formatCurrency(totals.ev)}</strong></td>
          <td><strong>${formatPercent(totals.aggregateCompletePercent)}</strong></td>
          <td><strong>${formatPercent(totals.aggregateCostUsedPercent)}</strong></td>
          <td><strong>${formatCurrency(totals.cv)}</strong></td>
          <td><strong>${totals.aggregateCpi.toFixed(2)}</strong></td>
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

const getDashboardErrorMessage = (error) => {
  const rawMessage = String(error?.message || error || "Unknown error").trim();
  if (/aborted|abort/i.test(rawMessage)) {
    return "Live data source timed out before it responded.";
  }
  return rawMessage || "Unknown error";
};

const updateDashboardSyncedAt = (sourceTimestamp = "") => {
  const asOfEl = document.querySelector(".as-of-text");
  if (!asOfEl) return;

  const sourceDate = sourceTimestamp ? new Date(sourceTimestamp) : null;
  const sourceLabel = sourceDate && Number.isFinite(sourceDate.getTime())
    ? ` · Source ${sourceDate.toLocaleTimeString()}`
    : "";
  asOfEl.textContent = `Live sync ${new Date().toLocaleTimeString()}${sourceLabel}`;
};

const rememberStableDashboardRows = (rows, sourceName = "dashboard data") => {
  lastStableDashboardRows = rows.slice();
  lastStableDashboardSavedAt = Date.now();
};

const hasStableDashboardRows = () => lastStableDashboardRows.length || activitySummaryRows.length;

const getStableDashboardRows = () =>
  lastStableDashboardRows.length ? lastStableDashboardRows : activitySummaryRows;


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

const resetDashboardToEmptySource = (sourceName = "live source", generatedAt = "") => {
  activitySummaryRows = [];
  lastStableDashboardRows = [];
  latestDashboardSignature = "";
  renderKpis({ planned: 0, actual: 0, cv: 0 });
  renderProgressKpis({ physicalProgressPercent: 0, costSpentPercent: 0, efficiencyGapPercent: 0 }, { planned: 0, actual: 0, ev: 0, cv: 0 });
  updateDashboardSummary([], { planned: 0, actual: 0, ev: 0, cv: 0 }, { physicalProgressPercent: 0, costSpentPercent: 0, efficiencyGapPercent: 0 });
  renderTable([]);
  renderOverrunTable([]);
  renderGapTable([]);
  if (varianceDisplayEl) varianceDisplayEl.textContent = formatCurrency(0);
  if (varianceStatusEl) varianceStatusEl.textContent = "Balanced";
  destroyCharts();
  updateDashboardSyncedAt(generatedAt);
  showMessage(`Live source ${sourceName} currently has no dashboard rows. Dashboard cleared to match Google Sheets.`);
};

const processRows = (rawRows, sourceName = "web app", { acceptEmpty = false, generatedAt = "" } = {}) => {
  if (!Array.isArray(rawRows) || !rawRows.length) {
    if (acceptEmpty) {
      resetDashboardToEmptySource(sourceName, generatedAt);
      return;
    }

    if (hasStableDashboardRows()) {
      const stableRows = getStableDashboardRows();
      activitySummaryRows = stableRows.slice();
      syncFilterOptionsFromRows(activitySummaryRows);
      applyFiltersAndRender();
      showMessage(
        `Live source temporarily returned no rows from ${sourceName}. Keeping the last stable ${activitySummaryRows.length} activity row(s) visible.`,
        true
      );
      return;
    }

    activitySummaryRows = [];
    lastStableDashboardRows = [];
    latestDashboardSignature = "";
    renderKpis({ planned: 0, actual: 0, cv: 0 });
    renderProgressKpis({ physicalProgressPercent: 0, costSpentPercent: 0, efficiencyGapPercent: 0 });
    updateDashboardSummary([], { planned: 0, actual: 0, ev: 0, cv: 0 }, { physicalProgressPercent: 0, costSpentPercent: 0, efficiencyGapPercent: 0 });
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
  if (!rows.length) {
    if (acceptEmpty) {
      resetDashboardToEmptySource(sourceName, generatedAt);
      return;
    }

    if (hasStableDashboardRows()) {
      const stableRows = getStableDashboardRows();
      activitySummaryRows = stableRows.slice();
      syncFilterOptionsFromRows(activitySummaryRows);
      applyFiltersAndRender();
      showMessage(
        `Live source returned an empty parsed result from ${sourceName}. Keeping the last stable ${activitySummaryRows.length} activity row(s) visible.`,
        true
      );
      return;
    }

    processRows([], sourceName);
    return;
  }

  // Accept the latest successful Google Sheets payload even when it contains
  // fewer rows than the previous dashboard state. Projects, activities, or costs
  // can be legitimately archived/deleted in the source sheet, and blocking
  // smaller payloads would keep stale rows visible instead of current data.

  const nextSignature = JSON.stringify(rows);
  if (nextSignature === latestDashboardSignature) {
    rememberStableDashboardRows(rows, sourceName);
    updateDashboardSyncedAt(generatedAt);
    showMessage(`Dashboard data is already up to date from ${sourceName}.`);
    return;
  }

  latestDashboardSignature = nextSignature;
  rememberStableDashboardRows(rows, sourceName);
  activitySummaryRows = rows;
  syncFilterOptionsFromRows(rows);
  applyFiltersAndRender();

  updateDashboardSyncedAt(generatedAt);
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

const buildRowsFromActivitiesAndCosts = (activities, costs, dailyCosts = []) => {
  const activitiesList = Array.isArray(activities) ? activities : [];
  const costsList = Array.isArray(costs) ? costs : [];
  const dailyCostsList = Array.isArray(dailyCosts) ? dailyCosts : [];
  const activityByProjectAndActivityId = new Map();
  const dailyCostsByCompositeKey = new Map();
  const dailyCostsByProjectActivityKey = new Map();
  const dailyEntriesByCompositeKey = new Map();
  const dailyEntriesByProjectActivityKey = new Map();
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
  const getDailyActualCost = (dailyCost) =>
    parseNumber(readBundleCell(dailyCost, ["actual cost/day", "actualCostPerDay", "actual_cost_per_day", "actual cost", "actualCost", "actual_cost", "cost", "amount"], 0));
  const getDailyEarnedValue = (dailyCost) =>
    parseNumber(readBundleCell(dailyCost, ["earned value/day", "earnedValuePerDay", "earned_value_per_day", "earned value", "earnedValue", "earned_value", "ev"], 0));
  const getDailyProgress = (dailyCost) =>
    parseNumber(readBundleCell(dailyCost, ["progress/day", "progressPerDay", "progress_per_day", "progress", "percent complete", "percentComplete", "% complete"], 0));
  const getDailyDate = (dailyCost) =>
    normalizeDateOnly(readBundleCell(dailyCost, ["date", "entry date", "actual date", "report date", "activity date"], ""));

  const addDailyCostTotal = (map, key, dailyCost) => {
    if (!key || key === "::") return;
    const existing = map.get(key) || {
      actualCost: 0,
      earnedValue: 0,
      progress: 0,
      count: 0,
      firstDate: "",
      lastDate: "",
    };
    const dailyDate = getDailyDate(dailyCost);
    existing.actualCost += getDailyActualCost(dailyCost);
    existing.earnedValue += getDailyEarnedValue(dailyCost);
    existing.progress += getDailyProgress(dailyCost);
    existing.count += 1;
    if (dailyDate) {
      if (!existing.firstDate || dailyDate < existing.firstDate) existing.firstDate = dailyDate;
      if (!existing.lastDate || dailyDate > existing.lastDate) existing.lastDate = dailyDate;
    }
    map.set(key, existing);
  };
  const addDailyEntry = (map, key, dailyCost) => {
    if (!key || key === "::") return;
    const existing = map.get(key) || [];
    existing.push({
      date: getDailyDate(dailyCost),
      actualCost: getDailyActualCost(dailyCost),
      earnedValue: getDailyEarnedValue(dailyCost),
      progress: getDailyProgress(dailyCost),
    });
    map.set(key, existing);
  };

  dailyCostsList.forEach((dailyCost) => {
    const projectId = getCostProjectId(dailyCost);
    const activityId = getCostActivityId(dailyCost);
    const costId = getCostId(dailyCost);
    if (!projectId || (!activityId && !costId)) return;

    addDailyCostTotal(
      dailyCostsByCompositeKey,
      makeDashboardCompositeKey({ projectId, activityId, costId }),
      dailyCost
    );
    addDailyEntry(
      dailyEntriesByCompositeKey,
      makeDashboardCompositeKey({ projectId, activityId, costId }),
      dailyCost
    );
    if (activityId) {
      addDailyCostTotal(
        dailyCostsByProjectActivityKey,
        makeDashboardCompositeKey({ projectId, activityId, costId: "" }),
        dailyCost
      );
      addDailyEntry(
        dailyEntriesByProjectActivityKey,
        makeDashboardCompositeKey({ projectId, activityId, costId: "" }),
        dailyCost
      );
    }
  });

  const getDailyTotalsForRow = ({ projectId, activityId, costId }) =>
    dailyCostsByCompositeKey.get(makeDashboardCompositeKey({ projectId, activityId, costId })) ||
    (activityId
      ? dailyCostsByProjectActivityKey.get(makeDashboardCompositeKey({ projectId, activityId, costId: "" }))
      : null) ||
    null;
  const getDailyEntriesForRow = ({ projectId, activityId, costId }) =>
    dailyEntriesByCompositeKey.get(makeDashboardCompositeKey({ projectId, activityId, costId })) ||
    (activityId
      ? dailyEntriesByProjectActivityKey.get(makeDashboardCompositeKey({ projectId, activityId, costId: "" }))
      : null) ||
    [];

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
    const dailyTotals = getDailyTotalsForRow({ projectId, activityId, costId });
    const dailyEntries = getDailyEntriesForRow({ projectId, activityId, costId });

    rows.push({
      "Project ID": projectId,
      "Project Name": projectName,
      "Activity ID": activityId,
      "Cost ID": costId,
      Activity: activityName,
      "Planned Cost": plannedCost,
      "Actual Cost": dailyTotals?.count
        ? dailyTotals.actualCost
        : parseNumber(readBundleCell(cost, ["actual cost", "actualCost", "actual_cost", "actual cost/day", "cost", "amount"], 0)),
      "Progress": dailyTotals?.count ? dailyTotals.progress : progress,
      "Earned Value": dailyTotals?.count
        ? dailyTotals.earnedValue
        : providedEarnedValue !== ""
          ? parseNumber(providedEarnedValue)
          : plannedCost * (progress / 100),
      "Planned Start": readBundleCell(matchedActivity, ["planned start", "plannedStart", "start date", "startDate"], ""),
      "Planned Finish": readBundleCell(matchedActivity, ["planned finish", "plannedFinish", "finish date", "finishDate"], ""),
      "Actual Start": dailyTotals?.firstDate || "",
      "Actual Finish": dailyTotals?.lastDate || "",
      Date: dailyTotals?.lastDate || dailyTotals?.firstDate || "",
      dailyEntries,
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
    const dailyTotals = getDailyTotalsForRow({ projectId, activityId, costId: "" });
    const dailyEntries = getDailyEntriesForRow({ projectId, activityId, costId: "" });

    rows.push({
      "Project ID": projectId,
      "Project Name": String(readBundleCell(activity, ["project", "project name", "projectName"], "")).trim(),
      "Activity ID": activityId,
      "Cost ID": "",
      Activity: String(readBundleCell(activity, ["activity", "activity name", "activityName", "name"], `Activity ${activityId}`)).trim(),
      "Planned Cost": plannedCost,
      "Actual Cost": dailyTotals?.count
        ? dailyTotals.actualCost
        : parseNumber(readBundleCell(activity, ["actual cost", "actualCost", "actual_cost"], 0)),
      "Progress": dailyTotals?.count ? dailyTotals.progress : progress,
      "Earned Value": dailyTotals?.count
        ? dailyTotals.earnedValue
        : providedEarnedValue !== ""
          ? parseNumber(providedEarnedValue)
          : plannedCost * (progress / 100),
      "Planned Start": readBundleCell(activity, ["planned start", "plannedStart", "start date", "startDate"], ""),
      "Planned Finish": readBundleCell(activity, ["planned finish", "plannedFinish", "finish date", "finishDate"], ""),
      "Actual Start": dailyTotals?.firstDate || "",
      "Actual Finish": dailyTotals?.lastDate || "",
      Date: dailyTotals?.lastDate || dailyTotals?.firstDate || "",
      dailyEntries,
    });
  });

  return rows;
};

const refreshDashboardData = async ({ force = false } = {}) => {
  if (isDashboardFetchInFlight) {
    pendingDashboardRefreshRequested = force || pendingDashboardRefreshRequested;
    return;
  }

  isDashboardFetchInFlight = true;
  try {
    if (!DATA_SOURCE_URL.trim()) {
      processRows([], "Google Sheets");
      showMessage("No Google Sheets data source URL configured. Dashboard is not connected to local data.", true);
      return;
    }

    if (force) {
      const stableCount = getStableDashboardRows().length;
      showMessage(stableCount ? "Refreshing data source while keeping current dashboard visible..." : "Loading data source...");
    }

    if (typeof window.DataBridge.fetchDashboardBundleFromSource === "function") {
      try {
        const { activities, costs, dailyCosts, dashboardRows, sourceName, generatedAt } = await window.DataBridge.fetchDashboardBundleFromSource(
          DATA_SOURCE_URL,
          { bypassCache: force && hasStableDashboardRows() }
        );
        const bundleRows = buildRowsFromActivitiesAndCosts(activities, costs, dailyCosts);
        const sourceRows = bundleRows.length ? bundleRows : dashboardRows;
        if (Array.isArray(sourceRows)) {
          processRows(sourceRows, sourceName, { acceptEmpty: true, generatedAt });
          return;
        }
      } catch (bundleError) {
        console.warn("Bundle fetch failed. Falling back to row-based fetch.", bundleError);
      }
    }

    const { rows, sourceName, generatedAt } = await window.DataBridge.fetchRowsFromSource(DATA_SOURCE_URL, {
      bypassCache: force && hasStableDashboardRows(),
    });

    if (Array.isArray(rows)) {
      processRows(rows, sourceName, { acceptEmpty: true, generatedAt });
      return;
    }

    processRows([], sourceName || "Google Sheets", { acceptEmpty: true, generatedAt });
  } catch (error) {
    if (hasStableDashboardRows()) {
      const stableRows = getStableDashboardRows();
      activitySummaryRows = stableRows.slice();
      syncFilterOptionsFromRows(activitySummaryRows);
      applyFiltersAndRender();
      const dashboardErrorMessage = getDashboardErrorMessage(error);
      showMessage(
        `Google Sheets is still catching up (${dashboardErrorMessage}). Keeping the last stable live dashboard data visible.`
      );
    } else {
      showMessage(`Error loading Google Sheets data source: ${getDashboardErrorMessage(error)}`, true);
    }
  } finally {
    isDashboardFetchInFlight = false;
    if (pendingDashboardRefreshRequested) {
      pendingDashboardRefreshRequested = false;
      setTimeout(() => refreshDashboardData({ force: true }), 0);
    }
  }
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
if (dateRangeFilterEl) {
  dateRangeFilterEl.addEventListener("change", () => {
    updateDateRangeFilterValues();
    applyFiltersAndRender();
  });
  updateDateRangeFilterValues();
}
if (dateStartFilterEl) dateStartFilterEl.addEventListener("change", handleDateRangeInputChange);
if (dateEndFilterEl) dateEndFilterEl.addEventListener("change", handleDateRangeInputChange);
if (activitySummarySortEl) activitySummarySortEl.addEventListener("change", applyFiltersAndRender);
if (exportReportButtonEl && exportReportOptionsEl) {
  exportReportButtonEl.addEventListener("click", () => {
    const expanded = exportReportButtonEl.getAttribute("aria-expanded") === "true";
    exportReportButtonEl.setAttribute("aria-expanded", String(!expanded));
    exportReportOptionsEl.hidden = expanded;
  });
  document.addEventListener("click", (event) => {
    if (!exportReportOptionsEl.hidden && !event.target.closest(".dashboard-export-menu")) {
      exportReportOptionsEl.hidden = true;
      exportReportButtonEl.setAttribute("aria-expanded", "false");
    }
  });
  exportReportOptionsEl.addEventListener("click", (event) => {
    const button = event.target.closest("button[data-export-format]");
    if (!button) return;
    exportDashboardReport(button.dataset.exportFormat || "xlsx");
    exportReportOptionsEl.hidden = true;
    exportReportButtonEl.setAttribute("aria-expanded", "false");
  });
}

const refreshDashboardIfVisible = async ({ force = false } = {}) => {
  if (!force && document.visibilityState === "hidden") return;
  if (document.activeElement instanceof HTMLElement) {
    const isTyping =
      document.activeElement.matches("input, textarea")
      || document.activeElement.isContentEditable;
    if (isTyping) return;
  }

  await refreshDashboardData({ force });
};

const setupDashboardRealtimeSync = () => {
  if (dashboardRefreshTimer) {
    clearInterval(dashboardRefreshTimer);
  }

  dashboardRefreshTimer = setInterval(() => {
    refreshDashboardIfVisible();
  }, DASHBOARD_REFRESH_INTERVAL_MS);

  window.addEventListener("focus", () => refreshDashboardIfVisible({ force: true }));
  window.addEventListener("online", () => refreshDashboardIfVisible({ force: true }));
  window.addEventListener("pageshow", () => refreshDashboardIfVisible({ force: true }));
  window.addEventListener("google-sheet:changed", () => refreshDashboardIfVisible({ force: true }));
  document.addEventListener("visibilitychange", () => {
    if (document.visibilityState === "visible") {
      refreshDashboardIfVisible({ force: true });
    }
  });
};

const runDashboardIntro = () => {
  const introEl = document.getElementById("dashboardIntro");
  const primaryLogoWrapEl = document.getElementById("introLogoWrapPrimary");
  const secondaryLogoWrapEl = document.getElementById("introLogoWrapSecondary");
  const particlesHostEl = document.getElementById("introParticles");
  const sidebarBrandLogoEl = document.querySelector(".brand .brand-logo");
  const sidebarBrandCopyEl = document.querySelector(".brand .brand-copy h2");
  const sidebarFooterLogoEl = document.querySelector(".sidebar-footer-logo img");
  const pageEl = document.body;
  if (
    !introEl
    || !primaryLogoWrapEl
    || !secondaryLogoWrapEl
    || !particlesHostEl
    || !sidebarBrandLogoEl
    || !pageEl
  ) return;

  pageEl.classList.add("intro-running");

  const emitParticleBurst = (count = 16, spread = 190) => {
    for (let i = 0; i < count; i += 1) {
      const particle = document.createElement("span");
      particle.className = "intro-particle";
      const size = 2 + Math.random() * 4;
      particle.style.width = `${size}px`;
      particle.style.height = `${size}px`;
      particle.style.left = `${primaryLogoWrapEl.offsetLeft + (Math.random() * 120 - 18)}px`;
      particle.style.top = `${primaryLogoWrapEl.offsetTop + (Math.random() * 120 - 18)}px`;
      particle.style.setProperty("--tx", `${(Math.random() - 0.5) * spread}px`);
      particle.style.setProperty("--ty", `${(Math.random() - 0.55) * spread}px`);
      particlesHostEl.appendChild(particle);
      particle.addEventListener("animationend", () => particle.remove(), { once: true });
    }
  };

  requestAnimationFrame(() => introEl.classList.add("intro-logo-build"));
  setTimeout(() => introEl.classList.add("intro-logo-enter"), 260);
  setTimeout(() => {
    introEl.classList.add("intro-logo-glow");
    emitParticleBurst(22, 210);
  }, 900);
  setTimeout(() => introEl.classList.add("reveal-dashboard"), 2000);

  setTimeout(() => {
    const logoTargetRect = sidebarBrandLogoEl.getBoundingClientRect();
    const brandTextRect = sidebarBrandCopyEl?.getBoundingClientRect();
    const primaryTargetX = Math.round(logoTargetRect.left);
    const primaryTargetY = Math.round(
      brandTextRect
        ? brandTextRect.top + ((brandTextRect.height - logoTargetRect.height) / 2)
        : logoTargetRect.top
    );
    introEl.style.setProperty("--intro-target-x-primary", `${primaryTargetX}px`);
    introEl.style.setProperty("--intro-target-y-primary", `${primaryTargetY}px`);

    if (sidebarFooterLogoEl) {
      const footerRect = sidebarFooterLogoEl.getBoundingClientRect();
      introEl.style.setProperty("--intro-target-x-secondary", `${Math.round(footerRect.left)}px`);
      introEl.style.setProperty("--intro-target-y-secondary", `${Math.round(footerRect.top)}px`);
    } else {
      introEl.style.setProperty("--intro-target-x-secondary", `${primaryTargetX}px`);
      introEl.style.setProperty("--intro-target-y-secondary", `${primaryTargetY}px`);
    }

    introEl.classList.add("move-logo", "split-logo");
    emitParticleBurst(24, 260);
    emitParticleBurst(18, 180);
    pageEl.classList.add("intro-main-reveal");
  }, 3000);

  setTimeout(() => {
    introEl.classList.add("is-complete");
    pageEl.classList.remove("intro-running");
  }, 5000);
};

if (document.body?.classList.contains("page-dashboard")) {
  const shouldPlayDashboardIntro = sessionStorage.getItem("costrackPlayDashboardIntro") === "true";
  const introEl = document.getElementById("dashboardIntro");

  if (shouldPlayDashboardIntro) {
    sessionStorage.removeItem("costrackPlayDashboardIntro");
    runDashboardIntro();
  } else if (introEl) {
    introEl.classList.add("is-complete");
  }
}

setupServiceWorkerUpdates();
refreshDashboardIfVisible({ force: true });
setupDashboardRealtimeSync();
