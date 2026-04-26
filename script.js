const totalActivitiesEl = document.getElementById("totalActivities");
const completedActivitiesEl = document.getElementById("completedActivities");
const inProgressActivitiesEl = document.getElementById("inProgressActivities");
const notStartedActivitiesEl = document.getElementById("notStartedActivities");
const delayedActivitiesEl = document.getElementById("delayedActivities");
const messageEl = document.getElementById("message");
const tableBodyEl = document.getElementById("activityTableBody");

const DATA_SOURCE_URL =
  "https://script.google.com/macros/s/AKfycbxaaigY2kno4qhfMVbt2nYSG2bO4T7475KAwxIJeZHAi_nyJ7_pqHq7UzzVgb8kXm79SA/exec";

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
  }

  return null;
};

const activityStatus = (row) => {
  if (row.percentComplete >= 100) return "Completed";
  if (row.cv < 0 && row.percentComplete < 100) return "Delayed";
  if (row.percentComplete <= 0) return "Not Started";
  return "In Progress";
};

const renderStatusBadge = (status) => `<span class="badge badge-${status.toLowerCase().replace(/\s+/g, "-")}">${status}</span>`;

const formatDate = (date) =>
  date.toLocaleDateString("en-US", { month: "short", day: "numeric", year: "numeric" });

const buildPlannedDates = (index) => {
  const start = new Date(Date.UTC(2024, 3, 15 + index * 7));
  const finish = new Date(start);
  finish.setUTCDate(start.getUTCDate() + 14);
  return { start: formatDate(start), finish: formatDate(finish) };
};

const extractRows = (rawRows) =>
  rawRows
    .map((row, index) => {
      const activityId = normalize(
        getCell(row, ["activity id", "id", "activity code", "wbs", "task id"]),
        `ACT-${index + 1}`
      );
      const activity = normalize(getCell(row, ["activity", "activity name"]), "");
      if (!activity) return null;

      const plannedCost = parseNumber(getCell(row, ["planned value", "planned cost", "pv", "budget"]));
      const actualCost = parseNumber(getCell(row, ["actual cost", "ac", "actual"]));
      const rawCv = getCell(row, ["cost variance", "cv"]);
      const hasProvidedCv = rawCv !== null && rawCv !== undefined && String(rawCv).trim() !== "";
      const cv = hasProvidedCv ? parseNumber(rawCv) : plannedCost - actualCost;
      const percentComplete = parseNumber(
        getCell(row, ["% complete", "percent complete", "progress", "progress %"])
      );
      const status = activityStatus({ cv, percentComplete });
      const dates = buildPlannedDates(index);

      return {
        activityId,
        activity,
        project: "Pati Piso, Limas",
        activityType: "Construction",
        status,
        plannedStart: dates.start,
        plannedFinish: dates.finish,
        progress: Math.max(0, Math.min(percentComplete, 100)),
        costStatus: cv > 0 ? "Under Budget" : cv < 0 ? "Over Budget" : "On Budget",
      };
    })
    .filter(Boolean);

const renderSummaryCards = (rows) => {
  const counts = rows.reduce(
    (acc, row) => {
      acc.total += 1;
      if (row.status === "Completed") acc.completed += 1;
      if (row.status === "In Progress") acc.inProgress += 1;
      if (row.status === "Not Started") acc.notStarted += 1;
      if (row.status === "Delayed") acc.delayed += 1;
      return acc;
    },
    { total: 0, completed: 0, inProgress: 0, notStarted: 0, delayed: 0 }
  );

  totalActivitiesEl.textContent = counts.total;
  completedActivitiesEl.textContent = counts.completed;
  inProgressActivitiesEl.textContent = counts.inProgress;
  notStartedActivitiesEl.textContent = counts.notStarted;
  delayedActivitiesEl.textContent = counts.delayed;
};

const renderTable = (rows) => {
  if (!rows.length) {
    tableBodyEl.innerHTML = '<tr><td colspan="10" class="placeholder">No valid rows found in data source.</td></tr>';
    return;
  }

  tableBodyEl.innerHTML = rows
    .map(
      (row) => `
      <tr>
        <td>${row.activityId}</td>
        <td>${row.activity}</td>
        <td>${row.project}</td>
        <td>${row.activityType}</td>
        <td>${renderStatusBadge(row.status)}</td>
        <td>${row.plannedStart}</td>
        <td>${row.plannedFinish}</td>
        <td>
          <div class="progress-wrap">
            <div class="progress-track"><div class="progress-fill" style="width:${row.progress}%"></div></div>
            <span>${row.progress.toFixed(0)}%</span>
          </div>
        </td>
        <td>${renderStatusBadge(row.costStatus)}</td>
        <td class="actions">⋮</td>
      </tr>
    `
    )
    .join("");
};

const showMessage = (text, isError = false) => {
  messageEl.textContent = text;
  messageEl.style.color = isError ? "#dc2626" : "#667085";
};

const processRows = (rawRows, sourceName = "web app") => {
  if (!Array.isArray(rawRows) || !rawRows.length) {
    renderSummaryCards([]);
    renderTable([]);
    showMessage(`No rows detected from ${sourceName}.`, true);
    return;
  }

  const rows = extractRows(rawRows);
  renderSummaryCards(rows);
  renderTable(rows);
  showMessage(`Loaded ${rows.length} activity row(s) from ${sourceName}.`);
};

const setupServiceWorkerUpdates = async () => {
  if (!("serviceWorker" in navigator)) return;

  try {
    const registration = await navigator.serviceWorker.register("./service-worker.js");
    navigator.serviceWorker.addEventListener("controllerchange", () => window.location.reload());
    setInterval(() => registration.update(), 5 * 60 * 1000);
  } catch (error) {
    showMessage(`Service worker setup failed: ${error.message}`, true);
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

  if (!isAppsScriptWebAppUrl(trimmedUrl)) {
    showMessage("Invalid URL. Use a valid Apps Script Web App URL.", true);
    return;
  }

  try {
    showMessage("Loading data source...");
    const response = await fetch(trimmedUrl, { cache: "no-store" });
    if (!response.ok) throw new Error(`Unable to fetch Apps Script Web App (HTTP ${response.status})`);
    const payload = await response.json();
    if (payload?.error) throw new Error(payload.error);
    processRows(payload?.rows || [], `Apps Script Web App (${payload?.sheetName || "sheet"})`);
  } catch (error) {
    showMessage(`Error loading data source: ${error.message}`, true);
  }
};

setupServiceWorkerUpdates();
loadGoogleSheet(DATA_SOURCE_URL);
