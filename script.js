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
const loadingStateEl = document.getElementById("loadingState");
const tableBodyEl = document.getElementById("activityTableBody");
const overrunTableBodyEl = document.getElementById("overrunTableBody");
const gapTableBodyEl = document.getElementById("gapTableBody");
const varianceDisplayEl = document.getElementById("varianceDisplay");
const varianceStatusEl = document.getElementById("varianceStatus");

const DATA_SOURCE_URL = window.DataBridge?.DEFAULT_DATA_SOURCE_URL || "";

const chartDependencyWarning =
  typeof window.Chart === "undefined"
    ? "Chart.js is not available. Graphs are disabled."
    : "";

let dashboardRows = [];
let varianceChart = null;
let costChart = null;
let dashboardRefreshTimer = null;
let isDashboardFetchInFlight = false;
let latestDashboardSignature = "";

const DASHBOARD_CACHE_KEY = "constructionStageDashboardRows";
const DASHBOARD_CACHE_TTL_MS = 5 * 1000;
const DASHBOARD_REFRESH_INTERVAL_MS = 15 * 1000;
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
        getCell(row, ["activity id", "id", "activity code", "wbs", "task id"]),
        ""
      );
      const activity = normalize(getCell(row, ["activity", "activity name"]), "");
      const hasValidActivityId = isValidActivityId(detectedActivityId);
      const hasValidActivityName = activity !== "" && !isSummaryLabel(activity);

      if (!hasValidActivityId && !hasValidActivityName) return null;

      const plannedCost = parseNumber(getCell(row, ["planned value", "planned cost", "pv", "budget"]));
      const actualCost = parseNumber(getCell(row, ["actual cost", "ac", "actual"]));
      const percentComplete = normalizeProgressPercent(
        getCell(row, ["% complete", "percent complete", "progress %", "progress", "completion"] )
      );
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
        ev: plannedCost * (percentComplete / 100),
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

const renderProgressKpis = (metrics, totals) => {
  const earnedValue = (parseNumber(totals?.planned) * parseNumber(metrics.physicalProgressPercent)) / 100;
  const cpi = parseNumber(totals?.actual) ? earnedValue / parseNumber(totals?.actual) : 0;

  physicalProgressEl.textContent = formatCurrency(earnedValue);
  costSpentEl.textContent = cpi.toFixed(2);
  efficiencyGapEl.textContent = `${formatPercent(metrics.physicalProgressPercent)} / ${formatPercent(metrics.costSpentPercent)}`;
  if (miniCompleteEl) miniCompleteEl.textContent = formatPercent(metrics.physicalProgressPercent);
  if (miniCostEl) miniCostEl.textContent = formatPercent(metrics.costSpentPercent);
  if (miniCompleteBarEl) miniCompleteBarEl.style.width = `${Math.max(0, Math.min(100, metrics.physicalProgressPercent))}%`;
  if (miniCostBarEl) miniCostBarEl.style.width = `${Math.max(0, Math.min(100, metrics.costSpentPercent))}%`;

  efficiencyCardEl.classList.remove("status-under", "status-over");
  if (metrics.efficiencyGapPercent < 0) {
    efficiencyCardEl.classList.add("status-over");
  } else if (metrics.efficiencyGapPercent > 0) {
    efficiencyCardEl.classList.add("status-under");
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
        <td>${row.activity}</td>
        <td>${formatPercent(row.percentComplete)}</td>
        <td>${formatPercent(row.costUsedPercent)}</td>
        <td><div class="gap-cell"><strong>${formatSignedPercent(gap)}</strong><span class="gap-track"><span class="gap-fill ${gapClass}" style="width:${Math.min(100, Math.abs(gap) * 4)}%"></span></span></div></td>
        <td><span class="status-pill ${statusClass}">${status}</span></td>
        <td>${interpretation}</td>
      </tr>`;
    }).join("");
};

const renderTable = (rows) => {
  if (!rows.length) {
    tableBodyEl.innerHTML =
      '<tr><td colspan="9" class="placeholder">No valid rows found in data source.</td></tr>';
    return;
  }

  tableBodyEl.innerHTML = rows
    .map(
      (row) => `
      <tr class="variance-row variance-${getVarianceBand(row.cv, row.plannedCost)}">
        <td>${row.activity}</td>
        <td>${formatCurrency(row.plannedCost)}</td>
        <td>${formatCurrency(row.actualCost)}</td>
        <td>${formatCurrency(row.ev)}</td>
        <td>${formatPercent(row.percentComplete)}</td>
        <td>${formatPercent(row.costUsedPercent)}</td>
        <td>${formatCurrency(row.actualCost - row.ev)}</td>
        <td>${(row.actualCost ? row.ev / row.actualCost : 0).toFixed(2)}</td>
        <td><span class="status-pill ${row.cv >= 0 ? "ok" : "bad"}">${row.cv >= 0 ? "On Track" : "Over Budget"}</span></td>
      </tr>
    `
    )
    .join("") + `
      <tr>
        <td><strong>TOTAL</strong></td>
        <td><strong>${formatCurrency(rows.reduce((a, r) => a + r.plannedCost, 0))}</strong></td>
        <td><strong>${formatCurrency(rows.reduce((a, r) => a + r.actualCost, 0))}</strong></td>
        <td><strong>${formatCurrency(rows.reduce((a, r) => a + r.ev, 0))}</strong></td>
        <td><strong>${formatPercent(rows.reduce((a, r) => a + r.percentComplete, 0) / rows.length)}</strong></td>
        <td><strong>${formatPercent(rows.reduce((a, r) => a + r.costUsedPercent, 0) / rows.length)}</strong></td>
        <td><strong>${formatCurrency(rows.reduce((a, r) => a + (r.actualCost - r.ev), 0))}</strong></td>
        <td><strong>${(rows.reduce((a, r) => a + r.ev, 0) / Math.max(rows.reduce((a, r) => a + r.actualCost, 0), 1)).toFixed(2)}</strong></td>
        <td><span class="status-pill bad">Over Budget</span></td>
      </tr>`;
};

const renderOverrunTable = (rows) => {
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
        <td>${row.activity}</td>
        <td>${formatCurrency(row.actualCost - row.ev)}</td>
        <td>${formatSignedPercent(row.percentComplete - row.costUsedPercent)}</td>
        <td><span class="status-pill bad">Over Budget</span></td>
      </tr>
    `
    )
    .join("");
};

const showMessage = (text, isError = false) => {
  messageEl.textContent = text;
  messageEl.style.color = isError ? "#dc2626" : "#667085";
};

const setLoadingState = (isLoading) => {
  if (!loadingStateEl) return;
  loadingStateEl.classList.toggle("hidden", !isLoading);
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

  varianceChart = new Chart(document.getElementById("varianceChart"), {
    type: "line",
    data: {
      labels,
      datasets: [
        {
          label: "% Complete",
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

  costChart = new Chart(document.getElementById("costChart"), {
    type: "bar",
    data: {
      labels,
      datasets: [
        {
          label: "Cost Variance (AC - PV)",
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
  const nextSignature = JSON.stringify(rows);
  if (nextSignature === latestDashboardSignature) {
    showMessage(`Live sync active. No new updates from ${sourceName}.`);
    return;
  }

  latestDashboardSignature = nextSignature;
  localStorage.setItem(DASHBOARD_CACHE_KEY, JSON.stringify({ savedAt: Date.now(), rows }));
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
    dashboardRows = rows;
    const totals = calculateTotalsFromRows(rows);
    const progressMetrics = calculateProgressMetrics(rows, totals);
    renderKpis(totals);
    renderProgressKpis(progressMetrics);
    renderTable(rows);
    renderOverrunTable(rows);
    generateCharts(rows);
    showMessage(
      `Loaded ${rows.length} cached activity row(s). Verifying against live source now...`
    );
  } catch {
    localStorage.removeItem(DASHBOARD_CACHE_KEY);
  }
};

const refreshDashboardData = async ({ force = false } = {}) => {
  if (isDashboardFetchInFlight) return;
  if (!force && document.visibilityState === "hidden") return;
  if (!DATA_SOURCE_URL.trim()) return;

  isDashboardFetchInFlight = true;
  setLoadingState(true);
  try {
    if (force) {
      showMessage("Loading data source...");
    }
    const { rows, sourceName } = await window.DataBridge.fetchRowsFromSource(DATA_SOURCE_URL);
    processRows(rows, sourceName);
  } catch (error) {
    showMessage(`Error loading data source: ${error.message}`, true);
  } finally {
    isDashboardFetchInFlight = false;
    setLoadingState(false);
  }
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

setupServiceWorkerUpdates();
if (DATA_SOURCE_URL.trim()) {
  localStorage.removeItem(DASHBOARD_CACHE_KEY);
  hydrateDashboardFromCache();
  refreshDashboardData({ force: true });
  setupRealtimeDashboardSync();
} else {
  showMessage("No data source configured. Add your new Google Apps Script or Google Sheet URL to DATA_SOURCE_URL.");
}
