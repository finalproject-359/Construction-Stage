
const topSearch = document.getElementById("costTopSearch");
const listSearch = document.getElementById("projectListSearch");
const projectsList = document.getElementById("costProjectsList");
const projectsEmpty = document.getElementById("costProjectsEmpty");
const selectionView = document.getElementById("costSelectionView");
const detailsView = document.getElementById("costDetailsView");
const selectedProjectBannerHost = document.getElementById("selectedProjectBannerHost");

const hasProjectSelectionInUrl = (() => {
  const query = new URLSearchParams(window.location.search);
  return Boolean(String(query.get("projectId") || "").trim() || String(query.get("project") || "").trim());
})();

if (hasProjectSelectionInUrl) {
  selectionView?.classList.add("hidden");
  detailsView?.classList.remove("hidden");
}

const safeJsonParse = (raw, fallback = []) => {
  try {
    const parsed = JSON.parse(raw || "[]");
    return Array.isArray(parsed) ? parsed : fallback;
  } catch {
    return fallback;
  }
};

const DAILY_COSTS_LOCAL_STORAGE_KEY = "constructionStageDailyCosts";
const COST_ACTIVITIES_LOCAL_STORAGE_KEY = "constructionStageActivities";
const LEGACY_COST_ACTIVITIES_LOCAL_STORAGE_KEY = "constructionStageCostActivities";

const loadFromLocalStorageArray = (key) => {
  try {
    return safeJsonParse(window.localStorage?.getItem(key) || "[]", []);
  } catch {
    return [];
  }
};

const persistToLocalStorage = (key, value) => {
  try {
    window.localStorage?.setItem(key, JSON.stringify(Array.isArray(value) ? value : []));
  } catch (error) {
    console.warn(`Unable to persist ${key} to local storage:`, error);
  }
};


let projectsState = [];
let costActivitiesState = [];
let dailyCostsState = [];

const loadProjects = () => projectsState.slice();
const loadDailyCosts = () => dailyCostsState.slice();
const saveDailyCosts = (items) => {
  dailyCostsState = Array.isArray(items) ? items.slice() : [];
};
const postToDataSource = async (resource, action, payload) => {
  const endpoint = window.DataBridge?.DEFAULT_DATA_SOURCE_URL;
  if (!endpoint) throw new Error("Google Sheets endpoint is not configured.");

  const requestPayload = { resource, action, ...payload };

  const parseResponsePayload = async (response) => {
    try {
      return await response.json();
    } catch {
      return null;
    }
  };

  const postWithFormat = async (format) =>
    fetch(endpoint, {
      method: "POST",
      headers: format === "json" ? { "Content-Type": "application/json" } : undefined,
      body:
        format === "json"
          ? JSON.stringify(requestPayload)
          : new URLSearchParams({ payload: JSON.stringify(requestPayload) }),
    });

  let response;
  let body;

  const sendViaGet = async () => {
    const url = new URL(endpoint);
    url.searchParams.set("payload", JSON.stringify(requestPayload));
    response = await fetch(url.toString(), { cache: "no-store" });
    body = await parseResponsePayload(response);
  };

  try {
    response = await postWithFormat("form");
    body = await parseResponsePayload(response);

    const needsJsonFallback =
      !response.ok ||
      (body?.ok === false &&
        /invalid payload|invalid payload parameter json|invalid json payload/i.test(String(body.error)));

    if (needsJsonFallback) {
      response = await postWithFormat("json");
      body = await parseResponsePayload(response);
    }
  } catch (error) {
    const maybeCorsIssue =
      endpoint.includes("script.google.com/macros/s/") &&
      /failed to fetch|networkerror|cors/i.test(String(error?.message || ""));

    if (maybeCorsIssue) {
      try {
        await sendViaGet();
      } catch {
        // Continue to shared error handling.
      }
    }

    if (!response || body?.ok === false) {
      const guidance = maybeCorsIssue
        ? "CORS check failed for POST. Verify your Google Apps Script Web App is deployed to Anyone and use the latest /exec deployment URL."
        : "Unable to reach the Google Sheet endpoint.";
      throw new Error(`${guidance} If this endpoint was recently changed, update DATA_SOURCE_URL in data-service.js.`);
    }
  }

  if (!response.ok) throw new Error(`Google Sheets sync failed (HTTP ${response.status}).`);
  if (body?.ok === false || body?.error) throw new Error(String(body.error || "Google Sheets sync failed."));
  return body;
};

const getValueByAliases = (source, aliases = []) => {
  if (!source || typeof source !== "object") return undefined;
  for (const alias of aliases) if (Object.prototype.hasOwnProperty.call(source, alias)) return source[alias];
  const normalizedEntries = Object.keys(source).map((key) => ({ key, normalized: String(key).toLowerCase().replace(/[^a-z0-9]/g, "") }));
  for (const alias of aliases) {
    const normalizedAlias = String(alias).toLowerCase().replace(/[^a-z0-9]/g, "");
    const matched = normalizedEntries.find((entry) => entry.normalized === normalizedAlias);
    if (matched) return source[matched.key];
  }
  // Fallback for headers with suffix/prefix qualifiers like "Planned Cost (PHP)".
  // Avoid over-broad matches (e.g., alias "costId" matching key "id").
  for (const alias of aliases) {
    const normalizedAlias = String(alias).toLowerCase().replace(/[^a-z0-9]/g, "");
    if (normalizedAlias.length < 4) continue;
    const matched = normalizedEntries.find((entry) => {
      if (entry.normalized.length < 4) return false;
      return entry.normalized.includes(normalizedAlias) || normalizedAlias.includes(entry.normalized);
    });
    if (matched) return source[matched.key];
  }
  return undefined;
};
const parseBudgetValue = (value) => Number(String(value ?? "0").replace(/[^\d.-]/g, "")) || 0;
const normalizeProject = (project = {}) => ({
  id: String(getValueByAliases(project, ["id", "projectId", "project_id"]) || "").trim(),
  name: String(getValueByAliases(project, ["name", "project", "projectName", "project_name"]) || "Untitled Project").trim(),
  code: String(getValueByAliases(project, ["code", "projectCode", "project_code"]) || "-").trim(),
  status: String(getValueByAliases(project, ["status", "projectStatus", "project_status"]) || "Not Started").trim(),
  budget: parseBudgetValue(getValueByAliases(project, ["budget", "plannedCost", "planned_value"])),
});
const escapeHtml = (value) => String(value).replaceAll("&", "&amp;").replaceAll("<", "&lt;").replaceAll(">", "&gt;").replaceAll('"', "&quot;").replaceAll("'", "&#39;");
const formatBudget = (value) => new Intl.NumberFormat("en-PH", { style: "currency", currency: "PHP", minimumFractionDigits: 2 }).format(value || 0);
const formatHumanDate = (value) => {
  const date = new Date(value);
  if (Number.isNaN(date.getTime())) return String(value || "-");
  return date.toLocaleDateString("en-US", { month: "short", day: "numeric", year: "numeric" });
};
const formatLongHumanDate = (value) => {
  const date = new Date(value);
  if (Number.isNaN(date.getTime())) return String(value || "-");
  return date.toLocaleDateString("en-US", { month: "long", day: "numeric", year: "numeric" });
};

const normalizeDateKey = (value) => {
  const date = new Date(value);
  if (Number.isNaN(date.getTime())) return String(value || "").trim();
  return date.toISOString().slice(0, 10);
};

const PH_FIXED_HOLIDAYS = new Set([
  "01-01",
  "04-09",
  "05-01",
  "06-12",
  "08-21",
  "11-01",
  "11-30",
  "12-08",
  "12-25",
  "12-30",
  "12-31",
]);
const PH_YEAR_SPECIFIC_HOLIDAYS = {
  2024: ["02-10", "04-10", "06-17"],
  2025: ["01-29", "03-31", "06-06"],
  2026: ["02-17", "03-20", "05-27"],
  2027: ["02-07", "03-10", "05-17"],
  2028: ["01-27", "02-27", "05-05"],
  2029: ["02-14", "02-15", "04-24"],
  2030: ["02-03", "02-05", "04-13"],
};
const formatMonthDayKey = (date) => `${String(date.getMonth() + 1).padStart(2, "0")}-${String(date.getDate()).padStart(2, "0")}`;
const getEasterSunday = (year) => {
  const a = year % 19;
  const b = Math.floor(year / 100);
  const c = year % 100;
  const d = Math.floor(b / 4);
  const e = b % 4;
  const f = Math.floor((b + 8) / 25);
  const g = Math.floor((b - f + 1) / 3);
  const h = (19 * a + b - d - g + 15) % 30;
  const i = Math.floor(c / 4);
  const k = c % 4;
  const l = (32 + 2 * e + 2 * i - h - k) % 7;
  const m = Math.floor((a + 11 * h + 22 * l) / 451);
  const month = Math.floor((h + l - 7 * m + 114) / 31) - 1;
  const day = ((h + l - 7 * m + 114) % 31) + 1;
  return new Date(year, month, day);
};
const getPhilippineHolidaySetForYear = (year) => {
  const holidays = new Set(PH_FIXED_HOLIDAYS);
  (PH_YEAR_SPECIFIC_HOLIDAYS[year] || []).forEach((holiday) => holidays.add(holiday));
  const lastDayOfAugust = new Date(year, 8, 0);
  const nationalHeroesDay = new Date(lastDayOfAugust);
  while (nationalHeroesDay.getDay() !== 1) nationalHeroesDay.setDate(nationalHeroesDay.getDate() - 1);
  holidays.add(formatMonthDayKey(nationalHeroesDay));
  const easterSunday = getEasterSunday(year);
  [-3, -2, -1].forEach((offset) => {
    const holidayDate = new Date(easterSunday);
    holidayDate.setDate(easterSunday.getDate() + offset);
    holidays.add(formatMonthDayKey(holidayDate));
  });
  return holidays;
};
const isWorkingDate = (date) => {
  const day = date.getDay();
  if (day === 0 || day === 6) return false;
  const holidays = getPhilippineHolidaySetForYear(date.getFullYear());
  return !holidays.has(formatMonthDayKey(date));
};

const buildAllDateRangeOptions = (startDate, finishDate) => {
  const start = new Date(startDate);
  const end = new Date(finishDate);
  if (Number.isNaN(start.getTime()) || Number.isNaN(end.getTime()) || start > end) return [];
  const options = [];
  const cursor = new Date(start);
  while (cursor <= end) {
    options.push(cursor.toISOString().slice(0, 10));
    cursor.setDate(cursor.getDate() + 1);
  }
  return options;
};

const buildDateRangeOptions = (startDate, finishDate) => {
  const options = buildAllDateRangeOptions(startDate, finishDate);
  return options.filter((value) => isWorkingDate(new Date(value)));
};
const formatProjectIdentityLabel = (project) => {
  const projectId = String(project?.id || "").trim();
  const projectName = String(project?.name || "").trim();
  if (!projectId) return projectName || "Untitled Project";
  if (!projectName) return projectId;
  return `${projectId} - ${projectName}`;
};
const toDateInputValue = (value) => {
  const d = new Date(value);
  if (Number.isNaN(d.getTime())) return "";
  return d.toISOString().slice(0, 10);
};
const computeDurationDays = (startDate, finishDate, fallback = 0) => {
  const fallbackValue = Number(fallback) || 0;
  if (fallbackValue > 0) return fallbackValue;

  const start = new Date(startDate);
  const finish = new Date(finishDate);
  if (!Number.isNaN(start.getTime()) && !Number.isNaN(finish.getTime()) && finish >= start) {
    return Math.max(1, Math.round((finish.getTime() - start.getTime()) / 86400000) + 1);
  }
  return 0;
};
const normalizeCostActivity = (activity = {}) => {
  const startDate = toDateInputValue(getValueByAliases(activity, ["startDate", "plannedStart", "planned_start"]));
  const finishDate = toDateInputValue(getValueByAliases(activity, ["finishDate", "plannedFinish", "planned_finish"]));
  const explicitDuration = Number(String(getValueByAliases(activity, ["durationDays", "duration_days", "duration"]) || "0").replace(/[^\d.-]/g, "")) || 0;

  return {
    id: String(getValueByAliases(activity, ["activityId", "activity_id", "activity id", "sourceActivityId", "source_activity_id", "source activity id", "code", "id"]) || "").trim(),
    costId: String(getValueByAliases(activity, ["costId", "cost_id", "cost id", "costCode", "cost_code", "cost code"]) || "").trim(),
    activityRefId: String(getValueByAliases(activity, ["activityRefId", "activity_ref_id", "activity ref id", "sourceActivityId", "source_activity_id", "source activity id", "activityId", "activity_id", "activity id", "id", "code"]) || "").trim(),
    projectId: String(getValueByAliases(activity, ["projectId", "project_id", "project id", "project", "projectName", "project_name", "project name"]) || "").trim(),
    projectName: String(getValueByAliases(activity, ["project", "projectName", "project_name", "project name"]) || "").trim(),
    name: String(getValueByAliases(activity, ["name", "activity", "activityName", "activity_name"]) || "Untitled Activity").trim(),
    startDate,
    finishDate,
    durationDays: computeDurationDays(startDate, finishDate, explicitDuration),
    plannedCost: parseBudgetValue(getValueByAliases(activity, ["plannedCost", "planned_cost", "planned cost", "plannedValue", "planned_value", "planned value", "budget"])),
  };
};

const getActivityRefId = (activity = {}) => String(activity.activityRefId || activity.id || "").trim();
const getCostActivityProjectKey = (activity = {}) => String(activity.projectId || activity.projectName || "").trim();
const getCostActivityKey = (activity = {}) => `${getCostActivityProjectKey(activity)}::${getActivityRefId(activity)}`;

const loadCostActivities = () => costActivitiesState.slice();

const normalizeRemoteActivity = (row = {}) => {
  const normalizedId = String(getValueByAliases(row, ["id", "activityId", "activity_id", "activity id", "code"]) || "").trim();
  const projectId = String(getValueByAliases(row, ["projectId", "project_id", "project id", "project", "projectName", "project_name", "project name"]) || "").trim();
  const startDate = toDateInputValue(getValueByAliases(row, ["startDate", "plannedStart", "planned_start", "planned start"]));
  const finishDate = toDateInputValue(getValueByAliases(row, ["finishDate", "plannedFinish", "planned_finish", "planned finish"]));
  const explicitDuration = Number(String(getValueByAliases(row, ["durationDays", "duration_days", "duration", "duration day"]) || "0").replace(/[^\d.-]/g, "")) || 0;

  return normalizeCostActivity({
    id: normalizedId,
    activityRefId: normalizedId,
    projectId,
    projectName: getValueByAliases(row, ["project", "projectName", "project_name", "project name"]),
    name: getValueByAliases(row, ["name", "activity", "activityName", "activity_name"]),
    costId: getValueByAliases(row, ["costId", "cost_id", "cost id", "costCode", "cost_code", "cost code"]),
    startDate,
    finishDate,
    durationDays: computeDurationDays(startDate, finishDate, explicitDuration),
    plannedCost: getValueByAliases(row, ["plannedCost", "planned_cost", "planned cost", "plannedValue", "planned_value", "planned value", "budget"]),
  });
};
const loadRemoteCostActivities = async (projectFilter = {}) => {
  const resourceRows = await loadActivitiesFromResourceEndpoint(projectFilter);
  if (resourceRows.length) return resourceRows;

  if (!window.DataBridge?.fetchRowsFromSource) return [];
  try {
    const { rows } = await window.DataBridge.fetchRowsFromSource(window.DataBridge.DEFAULT_DATA_SOURCE_URL);
    return (rows || [])
      .map(normalizeRemoteActivity)
      .filter((item) => getCostActivityProjectKey(item) && item.id);
  } catch (error) {
    console.warn("Unable to load cost activities from Google Sheets:", error);
    return [];
  }
};

const normalizeRemoteProject = (row = {}) => normalizeProject({
  id: getValueByAliases(row, ["id", "projectId", "project_id", "project id"]),
  name: getValueByAliases(row, ["name", "project", "projectName", "project_name", "project name"]),
  code: getValueByAliases(row, ["code", "projectCode", "project_code", "project code"]),
  status: getValueByAliases(row, ["status", "projectStatus", "project_status", "project status"]),
  budget: getValueByAliases(row, ["budget", "plannedCost", "planned_cost", "plannedValue", "planned value"]),
});

const loadRemoteProjects = async () => {
  if (!window.DataBridge?.DEFAULT_DATA_SOURCE_URL) return [];
  try {
    const url = new URL(window.DataBridge.DEFAULT_DATA_SOURCE_URL);
    url.searchParams.set("resource", "projects");
    url.searchParams.set("_ts", String(Date.now()));
    const response = await fetch(url.toString(), { cache: "no-store" });
    if (!response.ok) return [];
    const payload = await response.json();
    const rows = Array.isArray(payload?.projects) ? payload.projects : [];
    return rows.map(normalizeRemoteProject).filter((item) => item.id);
  } catch (error) {
    console.warn("Unable to load projects from resource endpoint:", error);
    return [];
  }
};

const loadActivitiesFromResourceEndpoint = async (projectFilter = {}) => {
  const dataSourceUrl = window.DataBridge?.DEFAULT_DATA_SOURCE_URL;
  if (!dataSourceUrl) return [];
  try {
    const url = new URL(dataSourceUrl);
    url.searchParams.set("resource", "activities");
    const projectIdFilter = String(projectFilter?.projectId || "").trim();
    const projectNameFilter = String(projectFilter?.projectName || "").trim();
    // Send only one project filter at a time to avoid overly strict server-side AND matching.
    // Prefer projectId because it is the canonical key; fall back to project name when ID is missing.
    if (projectIdFilter) {
      url.searchParams.set("projectId", projectIdFilter);
    } else if (projectNameFilter) {
      url.searchParams.set("project", projectNameFilter);
    }
    url.searchParams.set("_ts", String(Date.now()));
    const response = await fetch(url.toString(), { cache: "no-store" });
    if (!response.ok) return [];
    const payload = await response.json();
    const rows = Array.isArray(payload?.activities) ? payload.activities : [];
    return rows
      .map(normalizeRemoteActivity)
      .filter((item) => getCostActivityProjectKey(item) && item.id);
  } catch (error) {
    console.warn("Unable to load cost activities from activities resource endpoint:", error);
    return [];
  }
};


const normalizeProjectIdentityValue = (value = "") => {
  const raw = String(value || "").trim();
  if (!raw) return { raw: "", normalizedRaw: "", parsed: { id: "", name: "" } };
  return {
    raw,
    normalizedRaw: normalizeLookup(raw),
    parsed: splitProjectIdentityLabel(raw),
  };
};

const resolveProjectIdFromDailyCost = (dailyCost = {}, lookups = buildProjectIdentityLookups(loadProjects())) => {
  const candidateValues = [
    getValueByAliases(dailyCost, ["projectId", "project_id", "project id", "project"]),
    getValueByAliases(dailyCost, ["projectName", "project_name", "project name"]),
  ].map(normalizeProjectIdentityValue);

  for (const candidate of candidateValues) {
    if (!candidate.raw) continue;
    if (lookups.byId.has(candidate.normalizedRaw)) return lookups.byId.get(candidate.normalizedRaw) || candidate.raw;
    if (lookups.byName.has(candidate.normalizedRaw)) return lookups.byName.get(candidate.normalizedRaw) || candidate.raw;
    if (candidate.parsed.id && lookups.byId.has(candidate.parsed.id)) return lookups.byId.get(candidate.parsed.id) || candidate.raw;
    if (candidate.parsed.name && lookups.byName.has(candidate.parsed.name)) return lookups.byName.get(candidate.parsed.name) || candidate.raw;
  }

  return candidateValues.find((entry) => entry.raw)?.raw || "";
};

const loadRemoteDailyCosts = async (projectFilter = {}) => {
  const dataSourceUrl = window.DataBridge?.DEFAULT_DATA_SOURCE_URL;
  if (!dataSourceUrl) return [];
  const lookups = buildProjectIdentityLookups(loadProjects());
  try {
    const url = new URL(dataSourceUrl);
    url.searchParams.set("resource", "daily_costs");
    if (projectFilter?.projectId) url.searchParams.set("projectId", String(projectFilter.projectId));
    url.searchParams.set("_ts", String(Date.now()));
    const response = await fetch(url.toString(), { cache: "no-store" });
    if (!response.ok) return [];
    const payload = await response.json();
    const rows = Array.isArray(payload?.dailyCosts) ? payload.dailyCosts : [];
    return rows.map((row) => ({
      projectId: resolveProjectIdFromDailyCost(row, lookups),
      activityId: String(getValueByAliases(row, ["activityId", "activity_id", "activity id"]) || "").trim(),
      costId: String(getValueByAliases(row, ["costId", "cost_id", "cost id"]) || "").trim(),
      date: normalizeDateKey(getValueByAliases(row, ["date"])),
      actualCost: parseBudgetValue(getValueByAliases(row, ["actualCost", "actual_cost", "amount"])),
    })).filter((r) => r.projectId && r.activityId && r.date);
  } catch (error) {
    console.warn("Unable to load daily costs from resource endpoint:", error);
    return [];
  }
};

const extractActivityRefIdFromCostRow = (row = {}) => {
  const directActivityId = String(getValueByAliases(row, ["activityId", "activity_id", "activity id", "sourceActivityId", "source_activity_id"]) || "").trim();
  if (directActivityId) return directActivityId;

  const notesValue = String(getValueByAliases(row, ["notes", "note", "remarks"]) || "");
  const notesMatch = notesValue.match(/activity\s*id\s*[:#-]?\s*([a-z0-9._-]+)/i);
  return notesMatch?.[1] ? String(notesMatch[1]).trim() : "";
};

const extractCostRowsFromPayload = (payload = {}) => {
  if (Array.isArray(payload?.costs)) return payload.costs;
  if (Array.isArray(payload?.rows)) return payload.rows;
  if (Array.isArray(payload?.data)) return payload.data;
  if (Array.isArray(payload?.items)) return payload.items;
  return [];
};

const loadRemoteCostMetadata = async (projectFilter = {}) => {
  const dataSourceUrl = window.DataBridge?.DEFAULT_DATA_SOURCE_URL;
  if (!dataSourceUrl) return [];
  const lookups = buildProjectIdentityLookups(loadProjects());

  const parseCostRows = (payload = {}) => extractCostRowsFromPayload(payload)
    .map((row) => ({
      projectId: resolveProjectIdFromDailyCost({
        projectId: getValueByAliases(row, ["projectId", "project_id", "project id"]),
        projectName: getValueByAliases(row, ["project", "projectName", "project_name", "project name"]),
      }, lookups),
      activityRefId: extractActivityRefIdFromCostRow(row),
      activityName: String(getValueByAliases(row, ["activity", "activityName", "activity_name", "name"]) || "").trim(),
      costId: String(getValueByAliases(row, ["costId", "cost_id", "cost id", "costCode", "cost_code", "cost code"]) || "").trim(),
      plannedCost: parseBudgetValue(getValueByAliases(row, ["plannedCost", "planned_cost", "planned cost", "plannedValue", "planned_value", "planned value", "budget"])),
      date: String(getValueByAliases(row, ["date", "createdAt", "created_at"]) || "").trim(),
    }))
    .filter((row) => row.projectId && (row.activityRefId || row.activityName));

  try {
    const fetchRows = async (includeProjectId) => {
      const url = new URL(dataSourceUrl);
      url.searchParams.set("resource", "costs");
      if (includeProjectId && projectFilter?.projectId) url.searchParams.set("projectId", String(projectFilter.projectId));
      url.searchParams.set("_ts", String(Date.now()));
      const response = await fetch(url.toString(), { cache: "no-store" });
      if (!response.ok) return [];
      const payload = await response.json();
      return parseCostRows(payload);
    };

    const filteredRows = await fetchRows(true);
    if (filteredRows.length || !projectFilter?.projectId) return filteredRows;

    const allRows = await fetchRows(false);
    const selectedProjectId = normalizeLookup(projectFilter?.projectId || "");
    const selectedProjectName = normalizeLookup(projectFilter?.projectName || "");
    return allRows.filter((row) => {
      const rowProjectIdentity = String(row.projectId || "").trim();
      const rowProjectLookup = normalizeLookup(rowProjectIdentity);
      if (!rowProjectLookup) return false;
      if (selectedProjectId && rowProjectLookup === selectedProjectId) return true;
      if (selectedProjectName && rowProjectLookup === selectedProjectName) return true;

      const canonicalProjectId = normalizeLookup(
        resolveProjectIdFromDailyCost({ projectId: rowProjectIdentity, projectName: rowProjectIdentity }, lookups)
      );
      if (selectedProjectId && canonicalProjectId === selectedProjectId) return true;
      if (selectedProjectName && canonicalProjectId === selectedProjectName) return true;
      return false;
    });
  } catch (error) {
    console.warn("Unable to load cost metadata from resource endpoint:", error);
    return [];
  }
};

const normalizeLookup = (value) => String(value || "").trim().toLowerCase();

const buildProjectIdentityLookups = (projects = []) => {
  const byId = new Map();
  const byName = new Map();

  projects.forEach((project) => {
    const normalized = normalizeProject(project);
    const normalizedId = normalizeLookup(normalized.id);
    const normalizedName = normalizeLookup(normalized.name);
    if (normalizedId) byId.set(normalizedId, normalized.id);
    if (normalizedName) byName.set(normalizedName, normalized.id);
  });

  return { byId, byName };
};

const resolveActivityProjectId = (activity = {}, lookups = buildProjectIdentityLookups(loadProjects())) => {
  const directProjectId = String(activity.projectId || "").trim();
  const directProjectName = String(activity.projectName || "").trim();
  const parsedId = splitProjectIdentityLabel(directProjectId);
  const parsedName = splitProjectIdentityLabel(directProjectName);

  const candidates = [
    normalizeLookup(directProjectId),
    normalizeLookup(directProjectName),
    parsedId.id,
    parsedId.name,
    parsedName.id,
    parsedName.name,
  ].filter(Boolean);

  for (const candidate of candidates) {
    if (lookups.byId.has(candidate)) return lookups.byId.get(candidate) || directProjectId;
    if (lookups.byName.has(candidate)) return lookups.byName.get(candidate) || directProjectId;
  }

  return directProjectId;
};

const splitProjectIdentityLabel = (value = "") => {
  const raw = String(value || "").trim();
  if (!raw) return { id: "", name: "" };

  const separatorMatch = raw.match(/\s[-–—]\s/);
  if (!separatorMatch) return { id: "", name: normalizeLookup(raw) };

  const separatorIndex = separatorMatch.index || 0;
  const separatorLength = separatorMatch[0].length;
  const idPart = raw.slice(0, separatorIndex).trim();
  const namePart = raw.slice(separatorIndex + separatorLength).trim();
  if (!idPart || !namePart) return { id: "", name: normalizeLookup(raw) };

  return {
    id: normalizeLookup(idPart),
    name: normalizeLookup(namePart),
  };
};

const isDailyCostForProject = (entry, projectId, projectName = "") => {
  const entryProjectId = normalizeLookup(entry?.projectId);
  const projectIdLookup = normalizeLookup(projectId);
  const projectNameLookup = normalizeLookup(projectName);
  const parsed = splitProjectIdentityLabel(entry?.projectId);

  if (entryProjectId && projectIdLookup && entryProjectId === projectIdLookup) return true;
  if (entryProjectId && projectNameLookup && entryProjectId === projectNameLookup) return true;
  if (parsed.id && projectIdLookup && parsed.id === projectIdLookup) return true;
  if (parsed.name && projectNameLookup && parsed.name === projectNameLookup) return true;
  if (parsed.name && projectIdLookup && parsed.name === projectIdLookup) return true;
  if (parsed.id && projectNameLookup && parsed.id === projectNameLookup) return true;

  return false;
};

const isActivityForProject = (activity, projectId, projectName = "") => {
  const activityProjectId = normalizeLookup(activity?.projectId);
  const activityProjectName = normalizeLookup(activity?.projectName);
  const projectIdLookup = normalizeLookup(projectId);
  const projectNameLookup = normalizeLookup(projectName);
  const parsedFromProjectId = splitProjectIdentityLabel(activity?.projectId);
  const parsedFromProjectName = splitProjectIdentityLabel(activity?.projectName);

  if (activityProjectId && activityProjectId === projectIdLookup) return true;
  if (projectNameLookup && activityProjectName && activityProjectName === projectNameLookup) return true;
  if (projectNameLookup && activityProjectId && activityProjectId === projectNameLookup) return true;
  if (projectIdLookup && activityProjectName && activityProjectName === projectIdLookup) return true;

  if (projectIdLookup && (parsedFromProjectId.id === projectIdLookup || parsedFromProjectName.id === projectIdLookup)) return true;
  if (projectNameLookup && (parsedFromProjectId.name === projectNameLookup || parsedFromProjectName.name === projectNameLookup)) return true;
  if (projectIdLookup && (parsedFromProjectId.name === projectIdLookup || parsedFromProjectName.name === projectIdLookup)) return true;
  if (projectNameLookup && (parsedFromProjectId.id === projectNameLookup || parsedFromProjectName.id === projectNameLookup)) return true;
  return false;
};

const getProjectCostData = (projectId, allActivities = loadCostActivities()) => {
  const project = loadProjects().map(normalizeProject).find((item) => item.id === projectId);
  const projectName = String(project?.name || "").trim().toLowerCase();
  const compareActivitiesByStartPriority = (a, b) => {
    const aStart = new Date(a?.startDate || "").getTime();
    const bStart = new Date(b?.startDate || "").getTime();
    const aStartSafe = Number.isFinite(aStart) ? aStart : Number.POSITIVE_INFINITY;
    const bStartSafe = Number.isFinite(bStart) ? bStart : Number.POSITIVE_INFINITY;
    if (aStartSafe !== bStartSafe) return aStartSafe - bStartSafe;

    const aFinish = new Date(a?.finishDate || "").getTime();
    const bFinish = new Date(b?.finishDate || "").getTime();
    const aFinishSafe = Number.isFinite(aFinish) ? aFinish : Number.POSITIVE_INFINITY;
    const bFinishSafe = Number.isFinite(bFinish) ? bFinish : Number.POSITIVE_INFINITY;
    if (aFinishSafe !== bFinishSafe) return aFinishSafe - bFinishSafe;

    return String(getActivityRefId(a)).localeCompare(String(getActivityRefId(b)));
  };

  const activities = allActivities
    .filter((item) => isActivityForProject(item, projectId, projectName))
    .sort(compareActivitiesByStartPriority);
  const daily = loadDailyCosts().filter((item) => isDailyCostForProject(item, projectId, projectName));
  const rawRows = activities.map((activity) => {
    const refId = getActivityRefId(activity);
    const rowCostId = String(activity.costId || "").trim();
    const dailyItems = daily.filter((entry) => {
      const entryActivityId = String(entry.activityId || "").trim();
      const entryCostId = String(entry.costId || "").trim();
      return entryActivityId === refId || (rowCostId && entryCostId === rowCostId);
    });
    const actualCost = dailyItems.reduce((sum, entry) => sum + parseBudgetValue(entry.actualCost), 0);
    return { ...activity, actualCost, dailyItems };
  });

  // Consolidate duplicate activity rows (same costId/activity) so daily-cost updates
  // always appear on a single costing row instead of creating visual duplicates.
  const consolidated = new Map();
  rawRows.forEach((row) => {
    const refId = String(getActivityRefId(row) || "").trim();
    const costId = String(row.costId || "").trim();
    const key = costId || refId || `${String(row.name || "").trim().toLowerCase()}::${String(row.startDate || "")}::${String(row.finishDate || "")}`;
    if (!key) return;

    if (!consolidated.has(key)) {
      consolidated.set(key, row);
      return;
    }

    const existing = consolidated.get(key) || {};
    consolidated.set(key, {
      ...existing,
      ...row,
      costId: String(existing.costId || row.costId || "").trim(),
      activityRefId: String(getActivityRefId(existing) || getActivityRefId(row) || "").trim(),
      plannedCost: Math.max(parseBudgetValue(existing.plannedCost), parseBudgetValue(row.plannedCost)),
      actualCost: parseBudgetValue(existing.actualCost) + parseBudgetValue(row.actualCost),
      durationDays: Math.max(Number(existing.durationDays) || 0, Number(row.durationDays) || 0),
      dailyItems: [...(existing.dailyItems || []), ...(row.dailyItems || [])],
      name: existing.name || row.name || "Untitled Activity",
    });
  });

  return { rows: Array.from(consolidated.values()), activities, daily };
};

const buildSelectedProjectBannerMarkup = (project) => `<section class="selected-project-banner"><div><p class="selected-project-label">Selected Project</p><h3>${escapeHtml(formatProjectIdentityLabel(project))}</h3></div><a href="cost-management.html" class="ghost-btn">← Back to Projects</a></section>`;

const buildDetailsMarkup = (project, rows) => {
  const plannedCost = rows.reduce((sum, row) => sum + parseBudgetValue(row.plannedCost), 0);
  const actualCost = rows.reduce((sum, row) => sum + row.actualCost, 0);
  const variance = plannedCost - actualCost;
  const variancePercent = plannedCost ? (variance / plannedCost) * 100 : 0;
  const totalDuration = rows.reduce((sum, row) => sum + (Number(row.durationDays) || 0), 0);
  const avgActualPerDay = totalDuration > 0 ? actualCost / totalDuration : 0;
  const underBudgetCount = rows.filter((row) => row.actualCost > 0 && row.actualCost < row.plannedCost).length;
  const overBudgetCount = rows.filter((row) => row.actualCost > row.plannedCost).length;
  const noActualCostCount = rows.filter((row) => row.actualCost === 0).length;
  const activityTotal = Math.max(rows.length, 1);
  const underBudgetPct = (underBudgetCount / activityTotal) * 100;
  const overBudgetPct = (overBudgetCount / activityTotal) * 100;

  const tableRows = rows.length
    ? rows.map((row) => {
      const hasPlannedCost = parseBudgetValue(row.plannedCost) > 0;
      const hasActualCost = parseBudgetValue(row.actualCost) > 0;
      const plannedCostPerDay = hasPlannedCost && Number(row.durationDays) > 0
        ? parseBudgetValue(row.plannedCost) / Number(row.durationDays)
        : 0;
      const costIdCell = row.costId ? escapeHtml(row.costId) : "";
      const plannedCostCell = hasPlannedCost ? formatBudget(row.plannedCost) : "";
      const plannedCostPerDayCell = hasPlannedCost ? formatBudget(plannedCostPerDay) : "";
      const actualCostCell = hasActualCost ? formatBudget(row.actualCost) : "";

      const durationCell = Number(row.durationDays) > 0 ? `${row.durationDays} days` : "";

      const activityId = escapeHtml(getActivityRefId(row));
      return `<tr><td>${costIdCell}</td><td>${escapeHtml(row.name)}</td><td>${durationCell}</td><td>${plannedCostCell}</td><td>${plannedCostPerDayCell}</td><td>${actualCostCell}</td><td class="actions-col"><button type="button" class="action-menu-trigger" data-cost-actions="${activityId}" aria-label="Open cost actions" aria-expanded="false">⋮</button><div class="project-actions-menu hidden" data-cost-menu="${activityId}" role="menu" aria-label="Cost actions"><button type="button" class="project-action-btn edit-cost-meta-btn" data-activity-id="${activityId}" role="menuitem">Add / Edit Cost Details</button><button type="button" class="project-action-btn view-daily-cost-btn" data-activity-id="${activityId}" role="menuitem">View / Add Daily Cost</button></div></td></tr>`;
    }).join("")
    : '<tr><td colspan="7" class="empty-cell">No costing records yet. Add activities to start tracking costs.</td></tr>';

  const maxCost = Math.max(plannedCost, actualCost, 1);
  const plannedHeight = Math.max(10, Math.round((plannedCost / maxCost) * 100));
  const actualHeight = Math.max(10, Math.round((actualCost / maxCost) * 100));
  const varianceLabel = variance >= 0 ? "Under budget" : "Over budget";
  const varianceClass = variance >= 0 ? "good" : "bad";
  const topRows = rows
    .slice()
    .sort((a, b) => (b.actualCost - b.plannedCost) - (a.actualCost - a.plannedCost))
    .filter((row) => row.actualCost > row.plannedCost)
    .slice(0, 5)
    .map((row) => `<tr><td>${escapeHtml(row.costId || "-")}</td><td>${escapeHtml(row.name)}</td><td>${formatBudget(row.plannedCost)}</td><td>${formatBudget(row.actualCost)}</td><td class="bad">-${formatBudget(row.actualCost - row.plannedCost)}</td></tr>`)
    .join("") || '<tr><td colspan="5" class="empty-cell">No over budget activities.</td></tr>';

  return `<nav class="details-tabs"><button class="tab-btn active" data-tab="overview" type="button">Overview</button><button class="tab-btn" data-tab="costing" type="button">Costing Record</button></nav>
  <section class="details-tab-panel" data-panel="overview"><section class="details-kpis">
  <article class="kpi-card"><h4>Total Planned Cost</h4><p>${formatBudget(plannedCost)}</p></article>
  <article class="kpi-card"><h4>Total Actual Cost</h4><p>${formatBudget(actualCost)}</p></article>
  <article class="kpi-card"><h4>Variance</h4><p class="${varianceClass}">${formatBudget(variance)}</p><small>${varianceLabel}</small></article>
  <article class="kpi-card"><h4>Total Duration</h4><p>${totalDuration} days</p></article></section>
  <section class="overview-grid"><article class="panel chart-panel"><h3>Budget vs Actual</h3><div class="bars"><div class="bars-grid"><span>${formatBudget(maxCost)}</span><span>${formatBudget(maxCost * 0.75)}</span><span>${formatBudget(maxCost * 0.5)}</span><span>${formatBudget(maxCost * 0.25)}</span><span>0</span></div><div class="bars-track"><div class="bar-wrap"><strong>${formatBudget(plannedCost)}</strong><div class="bar planned" style="height:${plannedHeight}%"></div><p>Total Planned Cost</p></div><div class="bar-wrap"><strong>${formatBudget(actualCost)}</strong><div class="bar actual" style="height:${actualHeight}%"></div><p>Total Actual Cost</p></div></div><ul class="bars-legend"><li><span class="legend-dot planned"></span> Planned Cost</li><li><span class="legend-dot actual"></span> Actual Cost</li></ul></div></article>
  <article class="panel summary-panel"><h3>Cost Summary</h3><ul><li><span>Total Planned Cost</span><strong>${formatBudget(plannedCost)}</strong></li><li><span>Total Actual Cost</span><strong>${formatBudget(actualCost)}</strong></li><li><span>Variance</span><strong class="${varianceClass}">${formatBudget(variance)}</strong></li><li><span>Variance Percent</span><strong class="${varianceClass}">${variancePercent.toFixed(2)}%</strong></li><li><span>Total Duration</span><strong>${totalDuration} days</strong></li><li><span>Average Cost per Day (Actual)</span><strong>${formatBudget(avgActualPerDay)}</strong></li></ul></article>
  <article class="panel donut-panel"><h3>Cost Status (by Actual vs Planned)</h3><div class="donut-wrap"><div class="donut" style="background: conic-gradient(#34b567 0 ${underBudgetPct.toFixed(2)}%, #ef5050 ${underBudgetPct.toFixed(2)}% ${(underBudgetPct + overBudgetPct).toFixed(2)}%, #f0c23f ${(underBudgetPct + overBudgetPct).toFixed(2)}% 100%);"></div><ul class="status-list"><li><span class="legend-dot actual"></span>Under Budget <strong>${underBudgetCount} activities</strong></li><li><span class="legend-dot bad-dot"></span>Over Budget <strong>${overBudgetCount} activities</strong></li><li><span class="legend-dot neutral-dot"></span>No Actual Cost <strong>${noActualCostCount} activities</strong></li></ul></div></article>
  <article class="panel table-panel"><h3>Top Over Budget Activities</h3><table><thead><tr><th>Cost ID</th><th>Activity</th><th>Planned Cost</th><th>Actual Cost</th><th>Variance</th></tr></thead><tbody>${topRows}</tbody></table></article></section></section>
  <section class="details-tab-panel hidden" data-panel="costing">
  <section class="panel"><table><thead><tr><th>Cost ID</th><th>Activity</th><th>Duration</th><th>Planned Cost</th><th>Planned Cost/Day</th><th>Actual Cost</th><th>Actions</th></tr></thead><tbody>${tableRows}</tbody></table></section><div class="info-banner"><p>Tip: add daily actual costs by date via “View / Add Daily Cost”.</p></div></section>
  <section class="daily-cost-modal hidden" id="dailyCostModal"></section>
  <section class="cost-meta-modal hidden" id="costMetaModal"></section>`;
};

const renderDailyCostModal = (projectId, activityId, allActivities = loadCostActivities()) => {
  const modal = detailsView.querySelector("#dailyCostModal");
  const project = loadProjects().map(normalizeProject).find((item) => item.id === projectId);
  const projectName = String(project?.name || "").trim().toLowerCase();
  const activities = allActivities;
  const dailyCosts = loadDailyCosts();
  const activity = activities.find((item) => getActivityRefId(item) === activityId && isActivityForProject(item, projectId, projectName));
  if (!modal || !activity) return;

  const projectRows = getProjectCostData(projectId, allActivities).rows;
  const rowFallback = projectRows.find((row) => String(getActivityRefId(row) || "").trim() === activityId)
    || projectRows.find((row) => String(row.costId || "").trim() && String(row.name || "").trim().toLowerCase() === String(activity.name || "").trim().toLowerCase());

  const activityCostId = String(activity.costId || rowFallback?.costId || "").trim();
  const activityName = String(activity.name || rowFallback?.name || "").trim();
  const activityPlannedCost = Math.max(parseBudgetValue(activity.plannedCost), parseBudgetValue(rowFallback?.plannedCost));
  const activityDurationDays = Math.max(Number(activity.durationDays) || 0, Number(rowFallback?.durationDays) || 0);
  const activityPlannedCostPerDay = activityDurationDays > 0 ? (activityPlannedCost / activityDurationDays) : 0;
  if (!activityCostId || activityPlannedCost <= 0) {
    alert("Please add Cost ID and Planned Cost first before adding daily costs for this activity.");
    editCostMetadata(projectId, activityId, allActivities);
    return;
  }
  const availableDates = buildDateRangeOptions(activity.startDate, activity.finishDate);
  const hasAvailableDates = availableDates.length > 0;
  const dateOptions = hasAvailableDates
    ? availableDates.map((dateValue) => `<option value="${dateValue}">${formatLongHumanDate(dateValue)}</option>`).join("")
    : `<option value="" selected disabled>No working days in this date range.</option>`;

  const entries = dailyCosts
    .filter((item) => isDailyCostForProject(item, projectId, projectName) && String(item.activityId || "").trim() === activityId)
    .sort((a, b) => String(a.date || "").localeCompare(String(b.date || "")));
  const rows = entries.length
    ? entries.map((entry) => `<tr><td>${formatHumanDate(entry.date)}</td><td>${formatBudget(entry.actualCost)}</td><td><button type="button" class="daily-cost-delete-btn" data-delete-date="${entry.date}">Delete</button></td></tr>`).join("")
    : '<tr><td colspan="3" class="empty-cell">No daily costs recorded yet.</td></tr>';
  modal.classList.remove("hidden");
  modal.innerHTML = `<div class="daily-cost-dialog panel" role="dialog" aria-modal="true" aria-labelledby="dailyCostTitle"><div class="daily-cost-head"><h3 id="dailyCostTitle">${escapeHtml(activity.name)} Daily Cost</h3><button type="button" class="daily-cost-close" id="closeDailyModalBtn" aria-label="Close">×</button></div><p class="daily-cost-range">📅 ${escapeHtml(formatLongHumanDate(activity.startDate))} to ${escapeHtml(formatLongHumanDate(activity.finishDate))}</p>
    <section class="daily-cost-section"><h4>Add Daily Cost</h4><form id="dailyCostForm" class="daily-cost-form"><label><span>Select Date</span><select name="date" required ${hasAvailableDates ? "" : "disabled"}>${dateOptions}</select></label><label><span>Daily Cost (₱)</span><input name="actualCost" type="number" min="0" step="0.01" placeholder="Enter amount" required ${hasAvailableDates ? "" : "disabled"}></label><button class="primary-btn" type="submit" ${hasAvailableDates ? "" : "disabled"}>Add</button></form></section>
    <section class="daily-cost-section"><h4>Daily Cost Records</h4><div class="daily-cost-table-wrap"><table><thead><tr><th>Date</th><th>Amount (₱)</th><th>Action</th></tr></thead><tbody>${rows}</tbody></table></div></section>
    <div class="daily-cost-footer"><button type="button" class="ghost-btn" id="closeDailyModalBtnFooter">Close</button></div></div>`;

  modal.querySelector("#closeDailyModalBtn")?.addEventListener("click", () => modal.classList.add("hidden"));
  modal.querySelector("#closeDailyModalBtnFooter")?.addEventListener("click", () => modal.classList.add("hidden"));
  modal.querySelectorAll("[data-delete-date]").forEach((button) => button.addEventListener("click", () => {
    const date = String(button.dataset.deleteDate || "");
    const nextDailyCosts = loadDailyCosts().filter((item) => !(isDailyCostForProject(item, projectId, projectName)
      && String(item.activityId || "").trim() === activityId
      && String(item.date || "") === date));
    saveDailyCosts(nextDailyCosts);
    const resolvedProjectId = String(projectId || activity.projectId || "").trim();
    if (!resolvedProjectId) {
      alert("Unable to delete daily cost because Project ID is missing.");
      return;
    }
    postToDataSource("daily_costs", "delete", { dailyCost: { projectId: resolvedProjectId, costId: activityCostId, activityId, date } });
    const activeTab = detailsView.querySelector(".tab-btn.active")?.dataset.tab || "overview";
    const nextActivities = loadCostActivities();
    showProjectDetails(projectId, activeTab, nextActivities);
    renderDailyCostModal(projectId, activityId);
  }));
  const dateSelect = modal.querySelector("#dailyCostForm select[name=\"date\"]");
  if (dateSelect) {
    const todayIso = new Date().toISOString().slice(0, 10);
    if ([...dateSelect.options].some((option) => option.value === todayIso)) dateSelect.value = todayIso;
  }

  modal.querySelector("#dailyCostForm")?.addEventListener("submit", async (event) => {
    event.preventDefault();
    const formData = new FormData(event.currentTarget);
    const date = String(formData.get("date") || "");
    const actualCost = parseBudgetValue(formData.get("actualCost"));
    if (!hasAvailableDates) {
      alert("No valid working dates are available for this activity range.");
      return;
    }
    if (!date) {
      alert("Please select a date.");
      return;
    }
    if (actualCost <= 0) {
      alert("Daily cost must be greater than 0.");
      return;
    }
    const activityStartDate = String(activity.startDate || "");
    const activityFinishDate = String(activity.finishDate || "");
    if (activityStartDate && date < activityStartDate) {
      alert(`Date must be on or after ${activityStartDate}.`);
      return;
    }
    if (activityFinishDate && date > activityFinishDate) {
      alert(`Date must be on or before ${activityFinishDate}.`);
      return;
    }
    const selectedDate = new Date(date);
    if (Number.isNaN(selectedDate.getTime()) || !isWorkingDate(selectedDate)) {
      alert("Selected date must be a working day (Monday to Friday and not a holiday).");
      return;
    }
    const existingIndex = dailyCosts.findIndex((item) =>
      isDailyCostForProject(item, projectId, projectName)
      && String(item.activityId || "").trim() === activityId
      && String(item.date || "") === date
    );
    const resolvedProjectId = String(projectId || activity.projectId || "").trim();
    if (!resolvedProjectId) {
      alert("Unable to save daily cost because Project ID is missing.");
      return;
    }
    const payload = {
      projectId: resolvedProjectId,
      costId: activityCostId,
      activityId,
      activity: activityName,
      plannedCost: activityPlannedCost,
      plannedCostPerDay: activityPlannedCostPerDay,
      date,
      actualCost,
    };
    if (existingIndex >= 0) dailyCosts[existingIndex] = payload;
    else dailyCosts.push(payload);
    saveDailyCosts(dailyCosts);
    try {
      const dailyCostAction = existingIndex >= 0 ? "update" : "create";
      await postToDataSource("daily_costs", dailyCostAction, {
        dailyCost: {
          projectId: resolvedProjectId,
          costId: activityCostId,
          activityId,
          activity: activityName,
          plannedCost: activityPlannedCost,
          plannedCostPerDay: activityPlannedCostPerDay,
          date,
          actualCost,
        },
      });
    } catch (error) {
      console.warn("Unable to save daily cost to Google Sheets:", error);
      const resetDailyCosts = loadDailyCosts().filter((item) => !(
        isDailyCostForProject(item, projectId, projectName)
        && String(item.activityId || "").trim() === activityId
        && String(item.date || "") === date
      ));
      if (existingIndex >= 0) resetDailyCosts.push(dailyCosts[existingIndex]);
      saveDailyCosts(resetDailyCosts);
      alert(`Unable to save to Google Sheets. ${error?.message || "Please check Apps Script deployment permissions and try again."}`);
      return;
    }

    const activeTab = detailsView.querySelector(".tab-btn.active")?.dataset.tab || "overview";
    const nextActivities = loadCostActivities();
    showProjectDetails(projectId, activeTab, nextActivities);
    renderDailyCostModal(projectId, activityId);
  });
};

const showProjectDetails = (projectId, activeTab = "overview", allActivities = loadCostActivities()) => {
  const project = loadProjects().map(normalizeProject).find((item) => item.id === projectId);
  if (!project || !selectionView || !detailsView || !selectedProjectBannerHost) return false;
  selectionView.classList.add("hidden");
  selectedProjectBannerHost.classList.remove("hidden");
  selectedProjectBannerHost.innerHTML = buildSelectedProjectBannerMarkup(project);
  detailsView.classList.remove("hidden");
  const { rows } = getProjectCostData(projectId, allActivities);
  detailsView.innerHTML = buildDetailsMarkup(project, rows);

  const applyActiveTab = (target = "overview") => {
    detailsView.querySelectorAll(".tab-btn").forEach((item) => item.classList.toggle("active", item.dataset.tab === target));
    detailsView.querySelectorAll(".details-tab-panel").forEach((panel) => panel.classList.toggle("hidden", panel.dataset.panel !== target));
  };

  applyActiveTab(activeTab);

  detailsView.querySelectorAll(".tab-btn").forEach((btn) => btn.addEventListener("click", () => {
    const target = btn.dataset.tab || "overview";
    const nextUrl = new URL(window.location.href);
    nextUrl.searchParams.set("tab", target);
    window.history.replaceState({}, "", nextUrl.toString());
    applyActiveTab(target);
  }));

  const closeCostActionMenus = () => {
    detailsView.querySelectorAll("[data-cost-menu]").forEach((menu) => menu.classList.add("hidden"));
    detailsView.querySelectorAll("[data-cost-actions]").forEach((trigger) => trigger.setAttribute("aria-expanded", "false"));
  };

  detailsView.querySelectorAll("[data-cost-actions]").forEach((trigger) => {
    trigger.addEventListener("click", (event) => {
      const activityId = trigger.dataset.costActions;
      const menu = detailsView.querySelector(`[data-cost-menu="${CSS.escape(activityId || "")}"]`);
      if (!menu) return;
      const isOpen = !menu.classList.contains("hidden");
      closeCostActionMenus();
      if (!isOpen) {
        menu.classList.remove("hidden");
        trigger.setAttribute("aria-expanded", "true");
      }
      event.stopPropagation();
    });
  });

  detailsView.addEventListener("click", (event) => {
    const actionBtn = event.target.closest(".view-daily-cost-btn, .edit-cost-meta-btn");
    if (actionBtn) {
      const activityId = actionBtn.dataset.activityId;
      closeCostActionMenus();
      if (actionBtn.classList.contains("view-daily-cost-btn")) {
        renderDailyCostModal(projectId, activityId, allActivities);
      } else {
        editCostMetadata(projectId, activityId, allActivities);
      }
      return;
    }
    if (!event.target.closest(".actions-col")) closeCostActionMenus();
  });

  return true;
};

const saveCostActivityOverrides = (items = []) => {
  costActivitiesState = Array.isArray(items) ? items.slice() : [];
};
const renderCostMetadataModal = (projectId, activityRefId, target) => {
  const modal = detailsView.querySelector("#costMetaModal");
  if (!modal) return;

  modal.classList.remove("hidden");
  modal.innerHTML = `<div class="cost-meta-dialog panel" role="dialog" aria-modal="true" aria-labelledby="costMetaTitle"><div class="cost-meta-head"><h3 id="costMetaTitle">Add Cost</h3><button type="button" class="cost-meta-close" id="closeCostMetaModalBtn" aria-label="Close">×</button></div><form id="costMetaForm" class="cost-meta-form"><label for="costMetaIdInput">Cost ID <span aria-hidden="true">*</span></label><input id="costMetaIdInput" name="costId" type="text" placeholder="e.g., C001" value="${escapeHtml(target.costId || "")}" required><label for="costMetaPlannedInput">Planned Cost (₱) <span aria-hidden="true">*</span></label><input id="costMetaPlannedInput" name="plannedCost" type="number" min="0" step="0.01" placeholder="Enter planned cost" value="${Number(target.plannedCost) || 0}" required><div class="cost-meta-actions"><button type="button" class="ghost-btn" id="cancelCostMetaModalBtn">Cancel</button><button type="submit" class="primary-btn">Save Cost</button></div></form></div>`;

  const closeModal = () => modal.classList.add("hidden");
  modal.querySelector("#closeCostMetaModalBtn")?.addEventListener("click", closeModal);
  modal.querySelector("#cancelCostMetaModalBtn")?.addEventListener("click", closeModal);
  modal.addEventListener("click", (event) => {
    if (event.target === modal) closeModal();
  });

  modal.querySelector("#costMetaForm")?.addEventListener("submit", async (event) => {
    event.preventDefault();
    const formData = new FormData(event.currentTarget);
    const resolvedProjectId = String(projectId || target.projectId || "").trim();
    if (!resolvedProjectId) {
      alert("Unable to save cost because Project ID is missing.");
      return;
    }
    const nextCostId = String(formData.get("costId") || "").trim();
    const nextPlannedCost = parseBudgetValue(formData.get("plannedCost"));
    const existingOverrides = loadCostActivities().map(normalizeCostActivity);
    const nextOverrides = existingOverrides.filter((item) => !(String(item.projectId || "").trim() === String(projectId).trim() && getActivityRefId(item) === activityRefId));
    nextOverrides.push(normalizeCostActivity({
      ...target,
      costId: nextCostId,
      plannedCost: nextPlannedCost,
      activityRefId,
    }));
    saveCostActivityOverrides(nextOverrides);
    try {
      const durationDays = Number(target.durationDays) || 0;
      const plannedCostPerDay = durationDays > 0 ? nextPlannedCost / durationDays : 0;
      await postToDataSource("costs", "create", {
        cost: {
          costId: nextCostId,
          projectId: resolvedProjectId,
          project: target.projectName || "",
          activityId: activityRefId,
          activity: target.name || "",
          duration: durationDays,
          category: "Planned Cost",
          date: new Date().toISOString().slice(0, 10),
          plannedCost: nextPlannedCost,
          plannedCostPerDay,
          actualCost: 0,
          notes: `Activity ID: ${activityRefId}`,
        },
      });
    } catch (error) {
      console.warn("Unable to save cost record to Google Sheets:", error);
      saveCostActivityOverrides(existingOverrides);
      alert(`Unable to save cost record to Google Sheets. ${error?.message || "Please verify your Apps Script deployment settings and try again."}`);
      return;
    }
    closeModal();
    showProjectDetails(projectId, "costing", loadCostActivities());
  });
};

const editCostMetadata = (projectId, activityRefId, allActivities = loadCostActivities()) => {
  const target = allActivities.find((item) => String(item.projectId || "").trim() === String(projectId).trim() && getActivityRefId(item) === activityRefId);
  if (!target) return;
  renderCostMetadataModal(projectId, activityRefId, target);
};

const renderProjects = (query = "") => {
  const normalizedQuery = query.trim().toLowerCase();
  const projects = loadProjects().map(normalizeProject).filter((project) => !normalizedQuery || [project.name, project.code, project.status].some((value) => value.toLowerCase().includes(normalizedQuery)));
  projectsList.innerHTML = "";
  if (!projects.length) return projectsEmpty.classList.remove("hidden");
  projectsEmpty.classList.add("hidden");
  projects.forEach((project) => {
    const row = document.createElement("button");
    row.type = "button";
    row.className = "project-row";
    row.innerHTML = `<div class="project-meta"><strong>${escapeHtml(formatProjectIdentityLabel(project))}</strong><p>Status: ${escapeHtml(project.status)}</p></div><strong>${formatBudget(project.budget)}</strong>`;
    row.addEventListener("click", () => {
      const nextUrl = new URL(window.location.href);
      nextUrl.searchParams.set("projectId", project.id);
      window.location.href = nextUrl.toString();
    });
    projectsList.append(row);
  });
};

const syncSearches = (value) => { topSearch.value = value; listSearch.value = value; renderProjects(value); };
topSearch?.addEventListener("input", (event) => syncSearches(event.target.value));
listSearch?.addEventListener("input", (event) => syncSearches(event.target.value));

const params = new URLSearchParams(window.location.search);
const selectedProjectId = params.get("projectId") || "";
const selectedProjectName = params.get("project") || "";
const selectedTab = params.get("tab") === "costing" ? "costing" : "overview";
const getSelectedProjectFilter = (project = null) => ({
  projectId: project?.id || selectedProjectId,
  projectName: project?.name || selectedProjectName,
});
const bootstrapCostManagement = async () => {
  projectsState = (await loadRemoteProjects()).map(normalizeProject).filter((project) => project.id);

  const selectedProjectAfterBootstrap = loadProjects().map(normalizeProject).find((project) =>
    (selectedProjectId && project.id === selectedProjectId)
    || (selectedProjectName && project.name === selectedProjectName)
  );
  const projectFilter = getSelectedProjectFilter(selectedProjectAfterBootstrap);
  const [remoteActivities, remoteDailyCosts, remoteCostMetadataRows] = await Promise.all([
    loadRemoteCostActivities(projectFilter),
    loadRemoteDailyCosts(projectFilter),
    loadRemoteCostMetadata(projectFilter),
  ]);
  const merged = [...remoteActivities];
  const deduped = new Map();
  merged.forEach((item) => {
    const key = `${String(item.projectId).trim()}::${String(getActivityRefId(item)).trim()}`;
    if (!key || key === "::") return;

    if (!deduped.has(key)) {
      deduped.set(key, item);
      return;
    }

    const existing = deduped.get(key) || {};
    deduped.set(key, {
      ...existing,
      ...item,
      // Keep user-maintained cost metadata from local overrides when available.
      costId: String(existing.costId || item.costId || "").trim(),
      plannedCost: parseBudgetValue(existing.plannedCost) || parseBudgetValue(item.plannedCost),
      // Keep schedule fields aligned with the freshest merged source row when available.
      startDate: item.startDate || existing.startDate || "",
      finishDate: item.finishDate || existing.finishDate || "",
      durationDays: computeDurationDays(
        item.startDate || existing.startDate || "",
        item.finishDate || existing.finishDate || "",
        Number(item.durationDays) || Number(existing.durationDays) || 0
      ),
      name: existing.name || item.name || "Untitled Activity",
    });
  });
  const allActivities = Array.from(deduped.values());
  costActivitiesState = allActivities.slice();

  const mergedDailyCosts = [...remoteDailyCosts];
  const dedupedDailyCosts = new Map();
  mergedDailyCosts.forEach((item) => {
    const key = `${String(item.projectId || "").trim()}::${String(item.activityId || "").trim()}::${normalizeDateKey(item.date)}`;
    if (!key || key === "::::") return;
    dedupedDailyCosts.set(key, {
      projectId: String(item.projectId || "").trim(),
      activityId: String(item.activityId || "").trim(),
      costId: String(item.costId || "").trim(),
      date: normalizeDateKey(item.date),
      actualCost: parseBudgetValue(item.actualCost),
    });
  });
  saveDailyCosts(Array.from(dedupedDailyCosts.values()));

  if (remoteCostMetadataRows.length) {
    const metadataByActivityId = new Map();
    const metadataByActivityName = new Map();
    const metadataByActivityIdFallback = new Map();
    const metadataByActivityNameFallback = new Map();
    const pickLatest = (existing, incoming) => {
      if (!existing) return incoming;
      const existingDate = new Date(existing.date || "").getTime();
      const incomingDate = new Date(incoming.date || "").getTime();
      const shouldReplace = Number.isFinite(incomingDate) && (!Number.isFinite(existingDate) || incomingDate >= existingDate);
      return shouldReplace ? incoming : existing;
    };

    remoteCostMetadataRows.forEach((row) => {
      const projectKey = String(row.projectId || "").trim();
      const activityRefKey = String(row.activityRefId || "").trim();
      const activityNameKey = normalizeLookup(row.activityName);
      if (activityRefKey) {
        const byIdKey = `${projectKey}::${activityRefKey}`;
        metadataByActivityId.set(byIdKey, pickLatest(metadataByActivityId.get(byIdKey), row));
        metadataByActivityIdFallback.set(activityRefKey, pickLatest(metadataByActivityIdFallback.get(activityRefKey), row));
      }
      if (activityNameKey) {
        const byNameKey = `${projectKey}::${activityNameKey}`;
        metadataByActivityName.set(byNameKey, pickLatest(metadataByActivityName.get(byNameKey), row));
        metadataByActivityNameFallback.set(activityNameKey, pickLatest(metadataByActivityNameFallback.get(activityNameKey), row));
      }
    });

    costActivitiesState = costActivitiesState.map((activity) => {
      const projectKey = String(activity.projectId || "").trim();
      const activityKey = `${projectKey}::${String(getActivityRefId(activity) || "").trim()}`;
      const activityNameKey = `${projectKey}::${normalizeLookup(activity.name)}`;
      const fallbackActivityRefKey = String(getActivityRefId(activity) || "").trim();
      const fallbackActivityNameKey = normalizeLookup(activity.name);
      const metadata = metadataByActivityId.get(activityKey)
        || metadataByActivityName.get(activityNameKey)
        || metadataByActivityIdFallback.get(fallbackActivityRefKey)
        || metadataByActivityNameFallback.get(fallbackActivityNameKey);
      if (!metadata) return activity;
      return normalizeCostActivity({
        ...activity,
        costId: String(metadata.costId || activity.costId || "").trim(),
        plannedCost: parseBudgetValue(metadata.plannedCost) || parseBudgetValue(activity.plannedCost),
      });
    });
  }

  const activitiesForDisplay = loadCostActivities();

  if (!selectedProjectAfterBootstrap || !showProjectDetails(selectedProjectAfterBootstrap.id, selectedTab, activitiesForDisplay)) renderProjects();
};

let isCostManagementSyncInFlight = false;

const hasOpenCostDialog = () => Boolean(
  detailsView?.querySelector("#dailyCostModal:not(.hidden)")
  || detailsView?.querySelector("#costMetaModal:not(.hidden)")
);

const refreshSelectedProjectCostView = async ({ force = false } = {}) => {
  if (isCostManagementSyncInFlight) return;
  if (!force && document.visibilityState === "hidden") return;
  if (hasOpenCostDialog()) return;

  isCostManagementSyncInFlight = true;
  try {
    await bootstrapCostManagement();
  } finally {
    isCostManagementSyncInFlight = false;
    window.dispatchEvent(new CustomEvent("cost-management:data-loaded"));
  }
};

window.addEventListener("focus", () => refreshSelectedProjectCostView({ force: true }));
window.addEventListener("pageshow", () => refreshSelectedProjectCostView({ force: true }));
document.addEventListener("visibilitychange", () => {
  if (!document.hidden) refreshSelectedProjectCostView({ force: true });
});

refreshSelectedProjectCostView({ force: true });
