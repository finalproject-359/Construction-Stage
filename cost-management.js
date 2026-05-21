
const costPageHero = document.querySelector(".cost-hero.page-hero");
const topSearch = document.getElementById("costTopSearch");
const projectsList = document.getElementById("costProjectsList");
const projectTypeFilter = document.getElementById("costProjectTypeFilter");
const projectStatusFilter = document.getElementById("costProjectStatusFilter");
const costDateFilterWrap = document.querySelector(".cost-selection-date-filter");
const costDateFilterBtn = document.getElementById("costDateFilterBtn");
const costDateFilterLabel = document.getElementById("costDateFilterLabel");
const costDateRangePanel = document.getElementById("costDateRangePanel");
const dateStartInput = document.getElementById("costDateStart");
const dateEndInput = document.getElementById("costDateEnd");
const costDateClearBtn = document.getElementById("costDateClearBtn");
const costDateApplyBtn = document.getElementById("costDateApplyBtn");
const topFiltersSection = document.querySelector(".cost-selection-filters-top");
const projectsEmpty = document.getElementById("costProjectsEmpty");
const selectionView = document.getElementById("costSelectionView");
const detailsView = document.getElementById("costDetailsView");
const selectedProjectBannerHost = document.getElementById("selectedProjectBannerHost");
const clearCostFiltersBtn = document.getElementById("clearCostFilters");
const visibleProjectCount = document.getElementById("visibleProjectCount");
const visibleProjectBudget = document.getElementById("visibleProjectBudget");

const hasProjectSelectionInUrl = (() => {
  const query = new URLSearchParams(window.location.search);
  return Boolean(String(query.get("projectId") || "").trim() || String(query.get("project") || "").trim());
})();

if (hasProjectSelectionInUrl) {
  selectionView?.classList.add("hidden");
  detailsView?.classList.remove("hidden");
  costPageHero?.classList.add("hidden");
  topFiltersSection?.classList.add("hidden");
} else {
  costPageHero?.classList.remove("hidden");
  topFiltersSection?.classList.remove("hidden");
}

const safeJsonParse = (raw, fallback = []) => {
  try {
    const parsed = JSON.parse(raw || "[]");
    return Array.isArray(parsed) ? parsed : fallback;
  } catch {
    return fallback;
  }
};

const PROJECTS_LOCAL_STORAGE_KEY = "constructionStageProjects";
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

const COST_SYNC_TIMEOUT_MS = 15000;

const fetchWithTimeout = async (url, options = {}, timeoutMs = COST_SYNC_TIMEOUT_MS) => {
  if (typeof AbortController !== "function") return fetch(url, options);
  const controller = new AbortController();
  const timeoutId = window.setTimeout(() => controller.abort(), timeoutMs);

  try {
    return await fetch(url, { ...options, signal: controller.signal });
  } catch (error) {
    if (error?.name === "AbortError") {
      throw new Error("The cost data request timed out. Please check the Google Sheets connection and try again.");
    }
    throw error;
  } finally {
    window.clearTimeout(timeoutId);
  }
};

const getCssEscapedValue = (value = "") => {
  const raw = String(value || "");
  if (window.CSS?.escape) return window.CSS.escape(raw);
  return raw.replace(/[^a-zA-Z0-9_-]/g, "\\$&");
};

let projectsState = loadFromLocalStorageArray(PROJECTS_LOCAL_STORAGE_KEY);
let costActivitiesState = loadFromLocalStorageArray(COST_ACTIVITIES_LOCAL_STORAGE_KEY)
  .concat(loadFromLocalStorageArray(LEGACY_COST_ACTIVITIES_LOCAL_STORAGE_KEY));
let dailyCostsState = loadFromLocalStorageArray(DAILY_COSTS_LOCAL_STORAGE_KEY);
let hasWarnedAboutCachedProjects = false;
let hasWarnedAboutCachedDailyCosts = false;

const loadProjects = () => projectsState.slice();
const loadDailyCosts = () => dailyCostsState.slice();
const saveDailyCosts = (items) => {
  dailyCostsState = Array.isArray(items) ? items.slice() : [];
  persistToLocalStorage(DAILY_COSTS_LOCAL_STORAGE_KEY, dailyCostsState);
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
    fetchWithTimeout(endpoint, {
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
    response = await fetchWithTimeout(url.toString(), { cache: "no-store" });
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
  window.DataBridge?.pollRealtimeSync?.();
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
const hasPositiveActualCost = (item = {}) => parseBudgetValue(item.actualCost) > 0;
const normalizeProjectDateValue = (value) => {
  const d = new Date(value);
  if (Number.isNaN(d.getTime())) return "";
  return d.toISOString().slice(0, 10);
};
const normalizeProject = (project = {}) => ({
  id: String(getValueByAliases(project, ["id", "projectId", "project_id"]) || "").trim(),
  name: String(getValueByAliases(project, ["name", "project", "projectName", "project_name"]) || "Untitled Project").trim(),
  code: String(getValueByAliases(project, ["code", "projectCode", "project_code"]) || "-").trim(),
  type: String(getValueByAliases(project, ["type", "projectType", "project_type"]) || "General").trim(),
  status: String(getValueByAliases(project, ["status", "projectStatus", "project_status"]) || "Not Started").trim(),
  startDate: normalizeProjectDateValue(getValueByAliases(project, ["startDate", "start_date", "plannedStart", "planned_start"])),
  finishDate: normalizeProjectDateValue(getValueByAliases(project, ["finishDate", "targetFinish", "target_finish", "endDate", "end_date", "plannedFinish", "planned_finish"])),
  budget: parseBudgetValue(getValueByAliases(project, ["budget", "plannedCost", "planned_value"])),
});
const escapeHtml = (value) => String(value).replaceAll("&", "&amp;").replaceAll("<", "&lt;").replaceAll(">", "&gt;").replaceAll('"', "&quot;").replaceAll("'", "&#39;");
const formatBudget = (value) => new Intl.NumberFormat("en-PH", { style: "currency", currency: "PHP", minimumFractionDigits: 2 }).format(value || 0);
const ACTIVITY_PLANNED_COST_ALIASES = ["plannedCost", "planned_cost", "planned cost", "plannedValue", "planned_value", "planned value"];
const COST_METADATA_PLANNED_COST_ALIASES = ["plannedCost", "planned_cost", "planned cost", "plannedValue", "planned_value", "planned value"];
const COST_METADATA_COST_ID_ALIASES = ["costId", "cost_id", "cost id", "costCode", "cost_code", "cost code"];
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
const formatProjectTimeline = (project) => {
  const start = project.startDate ? formatHumanDate(project.startDate) : "Start not set";
  const finish = project.finishDate ? formatHumanDate(project.finishDate) : "Finish not set";
  return `${start} — ${finish}`;
};
const setFormSavingState = (form, isSaving, savingText = "Saving…") => {
  if (!(form instanceof HTMLFormElement)) return;
  const submitButton = form.querySelector('button[type="submit"]');
  if (isSaving) {
    if (submitButton instanceof HTMLButtonElement && !submitButton.dataset.defaultLabel) {
      submitButton.dataset.defaultLabel = submitButton.textContent.trim();
    }
    form.classList.add("is-saving");
    form.setAttribute("aria-busy", "true");
  } else {
    form.classList.remove("is-saving");
    form.removeAttribute("aria-busy");
  }
  form.querySelectorAll("input, select, textarea, button").forEach((control) => {
    if (control instanceof HTMLInputElement || control instanceof HTMLSelectElement || control instanceof HTMLTextAreaElement || control instanceof HTMLButtonElement) {
      control.disabled = isSaving;
    }
  });
  if (submitButton instanceof HTMLButtonElement) {
    submitButton.innerHTML = isSaving
      ? `<span class="btn-spinner" aria-hidden="true"></span><span>${escapeHtml(savingText)}</span>`
      : escapeHtml(submitButton.dataset.defaultLabel || "Save");
  }
};

const getStatusTone = (status = "") => {
  const normalized = String(status).toLowerCase();
  if (/complete|done|finished/.test(normalized)) return "complete";
  if (/progress|ongoing|active|started/.test(normalized)) return "active";
  if (/hold|delay|risk|issue/.test(normalized)) return "risk";
  return "neutral";
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
const countWorkingDaysInclusive = (startDate, finishDate) => {
  const start = new Date(startDate);
  const finish = new Date(finishDate);
  if (Number.isNaN(start.getTime()) || Number.isNaN(finish.getTime()) || finish < start) return 0;

  const cursor = new Date(start);
  cursor.setHours(0, 0, 0, 0);
  const end = new Date(finish);
  end.setHours(0, 0, 0, 0);

  let workingDays = 0;
  while (cursor <= end) {
    if (isWorkingDate(cursor)) workingDays += 1;
    cursor.setDate(cursor.getDate() + 1);
  }
  return workingDays;
};

const computeDurationDays = (startDate, finishDate, fallback = 0) => {
  const fallbackValue = Number(fallback) || 0;
  if (fallbackValue > 0) return fallbackValue;

  return countWorkingDaysInclusive(startDate, finishDate);
};
const computeEffectiveDurationDays = (activity = {}, startDate = "", finishDate = "", fallback = 0) => {
  const baseDuration = computeDurationDays(startDate, finishDate, fallback);
  const delayedDayCount = Number(String(getValueByAliases(activity, ["delayedDayCount", "delayed_day_count"]) || "0").replace(/[^\d.-]/g, ""));
  if (Number.isFinite(delayedDayCount) && delayedDayCount > 0) return baseDuration + Math.round(delayedDayCount);

  const actualFinishDate = toDateInputValue(getValueByAliases(activity, ["actualFinish", "actual_finish", "actualFinishDate", "actual_finish_date"]));
  if (!startDate || !actualFinishDate) return baseDuration;
  return Math.max(baseDuration, computeDurationDays(startDate, actualFinishDate, 0));
};
const clampPercent = (value) => {
  const parsed = Number(String(value ?? "").replace(/[^\d.-]/g, ""));
  if (!Number.isFinite(parsed)) return 0;
  return Math.max(0, Math.min(100, parsed));
};
const computeEarnedValue = (plannedCostPerDay, dailyRecordCount, progressPercent, explicitEarnedValue) => {
  const normalizedProgress = clampPercent(progressPercent);
  const normalizedPlannedCostPerDay = parseBudgetValue(plannedCostPerDay);
  const normalizedDailyRecordCount = Math.max(0, Number(dailyRecordCount) || 0);
  if (normalizedPlannedCostPerDay > 0 && normalizedDailyRecordCount > 0 && normalizedProgress > 0) {
    return normalizedPlannedCostPerDay * normalizedDailyRecordCount * (normalizedProgress / 100);
  }

  const normalizedExplicitEv = parseBudgetValue(explicitEarnedValue);
  if (normalizedExplicitEv > 0) return normalizedExplicitEv;
  if (normalizedPlannedCostPerDay <= 0 || normalizedDailyRecordCount <= 0) return 0;
  return normalizedPlannedCostPerDay * normalizedDailyRecordCount * (normalizedProgress / 100);
};
const normalizeCostActivity = (activity = {}) => {
  const startDate = toDateInputValue(getValueByAliases(activity, ["startDate", "plannedStart", "planned_start"]));
  const finishDate = toDateInputValue(getValueByAliases(activity, ["finishDate", "plannedFinish", "planned_finish"]));
  const explicitDuration = Number(String(getValueByAliases(activity, ["durationDays", "duration_days", "duration"]) || "0").replace(/[^\d.-]/g, "")) || 0;
  const progressPercent = clampPercent(getValueByAliases(activity, ["percentComplete", "percent_complete", "progressPercent", "progress_percent", "% complete", "percent complete", "completion", "progress"]));
  const plannedCost = parseBudgetValue(getValueByAliases(activity, ACTIVITY_PLANNED_COST_ALIASES));
  const effectiveDurationDays = computeEffectiveDurationDays(activity, startDate, finishDate, explicitDuration);
  const plannedCostPerDay = plannedCost > 0 && Number(effectiveDurationDays) > 0
    ? plannedCost / Number(effectiveDurationDays)
    : 0;
  const earnedValue = computeEarnedValue(
    plannedCostPerDay,
    0,
    progressPercent,
    getValueByAliases(activity, ["earnedValue", "earned_value", "earned value", "earned value (ev)", "ev"]),
  );

  return {
    id: String(getValueByAliases(activity, ["activityId", "activity_id", "activity id", "sourceActivityId", "source_activity_id", "source activity id", "code", "id"]) || "").trim(),
    costId: String(getValueByAliases(activity, COST_METADATA_COST_ID_ALIASES) || "").trim(),
    activityRefId: String(getValueByAliases(activity, ["activityRefId", "activity_ref_id", "activity ref id", "sourceActivityId", "source_activity_id", "source activity id", "activityId", "activity_id", "activity id", "id", "code"]) || "").trim(),
    projectId: String(getValueByAliases(activity, ["projectId", "project_id", "project id", "project", "projectName", "project_name", "project name"]) || "").trim(),
    projectName: String(getValueByAliases(activity, ["project", "projectName", "project_name", "project name"]) || "").trim(),
    name: String(getValueByAliases(activity, ["name", "activity", "activityName", "activity_name"]) || "Untitled Activity").trim(),
    startDate,
    finishDate,
    durationDays: effectiveDurationDays,
    progressPercent,
    plannedCost,
    earnedValue,
  };
};

const getActivityRefId = (activity = {}) => String(activity.activityRefId || activity.id || "").trim();
const getCostActivityProjectKey = (activity = {}) => String(activity.projectId || activity.projectName || "").trim();
const getCostActivityKey = (activity = {}) => `${getCostActivityProjectKey(activity)}::${getActivityRefId(activity)}`;

const dedupeCostActivities = (items = []) => {
  const deduped = new Map();
  (Array.isArray(items) ? items : []).map(normalizeCostActivity).forEach((item) => {
    const scopedActivityKey = getCostActivityKey(item);
    const fallbackNameKey = `${String(item.projectId || item.projectName || "").trim()}::${String(item.name || "").trim().toLowerCase()}`;
    const key = scopedActivityKey && scopedActivityKey !== "::" ? scopedActivityKey : fallbackNameKey;
    if (!key || key === "::") return;
    const existing = deduped.get(key);
    deduped.set(key, {
      ...(existing || {}),
      ...item,
      costId: item.costId || existing?.costId || "",
      plannedCost: parseBudgetValue(item.plannedCost) || parseBudgetValue(existing?.plannedCost),
      earnedValue: parseBudgetValue(item.earnedValue) || parseBudgetValue(existing?.earnedValue),
    });
  });
  return Array.from(deduped.values());
};

const cleanupOrphanedDailyCosts = (activities = loadCostActivities()) => {
  const normalizedActivities = activities.map(normalizeCostActivity);
  if (!normalizedActivities.length) return;

  const validActivityIds = new Set();
  const validCostIds = new Set();
  const validNames = new Set();
  normalizedActivities.forEach((activity) => {
    const projectKey = normalizeLookup(resolveActivityProjectId(activity));
    const activityId = normalizeLookup(getActivityRefId(activity));
    const costId = normalizeLookup(activity.costId);
    const activityName = normalizeLookup(activity.name);
    if (projectKey && activityId) validActivityIds.add(`${projectKey}::${activityId}`);
    if (projectKey && costId) validCostIds.add(`${projectKey}::${costId}`);
    if (projectKey && activityName) validNames.add(`${projectKey}::${activityName}`);
  });

  const lookups = buildProjectIdentityLookups(loadProjects());
  const cleaned = loadDailyCosts()
    .map((item) => normalizeDailyCostRecord(item, lookups))
    .filter((item) => {
      const projectKey = normalizeLookup(item.projectId);
      if (!projectKey) return false;
      const activityId = normalizeLookup(item.activityId);
      const costId = normalizeLookup(item.costId);
      const activityName = normalizeLookup(item.activity);
      return (activityId && validActivityIds.has(`${projectKey}::${activityId}`))
        || (costId && validCostIds.has(`${projectKey}::${costId}`))
        || (activityName && validNames.has(`${projectKey}::${activityName}`));
    });

  if (JSON.stringify(cleaned) !== JSON.stringify(dailyCostsState)) saveDailyCosts(cleaned);
};

const loadCostActivities = () => dedupeCostActivities(costActivitiesState).slice();

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
    costId: "",
    startDate,
    finishDate,
    durationDays: computeEffectiveDurationDays(row, startDate, finishDate, explicitDuration),
    plannedCost: 0,
    progressPercent: getValueByAliases(row, ["percentComplete", "percent_complete", "progressPercent", "progress_percent", "% complete", "percent complete", "completion", "progress"]),
    earnedValue: 0,
  });
};
const loadRemoteCostActivities = async (projectFilter = {}) => {
  const resourceRows = await loadActivitiesFromResourceEndpoint(projectFilter);
  if (Array.isArray(resourceRows)) {
    return { rows: resourceRows, authoritative: true };
  }

  if (!window.DataBridge?.fetchRowsFromSource) return { rows: [], authoritative: false };
  try {
    const { rows } = await window.DataBridge.fetchRowsFromSource(window.DataBridge.DEFAULT_DATA_SOURCE_URL);
    return {
      rows: (rows || [])
        .map(normalizeRemoteActivity)
        .filter((item) => getCostActivityProjectKey(item) && item.id),
      authoritative: true,
    };
  } catch (error) {
    console.warn("Unable to load cost activities from Google Sheets:", error);
    return { rows: [], authoritative: false };
  }
};

const isArchivedProject = (project = {}) => String(project.status || '').trim().toLowerCase() === 'archived';

const normalizeRemoteProject = (row = {}) => normalizeProject({
  id: getValueByAliases(row, ["id", "projectId", "project_id", "project id"]),
  name: getValueByAliases(row, ["name", "project", "projectName", "project_name", "project name"]),
  code: getValueByAliases(row, ["code", "projectCode", "project_code", "project code"]),
  status: getValueByAliases(row, ["status", "projectStatus", "project_status", "project status"]),
  budget: getValueByAliases(row, ["budget", "plannedCost", "planned_cost", "plannedValue", "planned value"]),
});

const loadRemoteProjectsFromBundle = async () => {
  if (!window.DataBridge?.fetchDashboardBundleFromSource || !window.DataBridge?.DEFAULT_DATA_SOURCE_URL) return null;
  try {
    const bundle = await window.DataBridge.fetchDashboardBundleFromSource(window.DataBridge.DEFAULT_DATA_SOURCE_URL, { bypassCache: true });
    if (!Array.isArray(bundle?.projects)) return null;
    return bundle.projects.map(normalizeRemoteProject).filter((item) => item.id);
  } catch (error) {
    console.warn("Unable to load projects from bundled Google Sheets payload:", error);
    return null;
  }
};

const loadRemoteProjects = async () => {
  if (!window.DataBridge?.DEFAULT_DATA_SOURCE_URL) return null;
  try {
    const url = new URL(window.DataBridge.DEFAULT_DATA_SOURCE_URL);
    url.searchParams.set("resource", "projects");
    url.searchParams.set("_ts", String(Date.now()));
    const response = await fetchWithTimeout(url.toString(), { cache: "no-store" });
    if (!response.ok) throw new Error(`Projects fetch failed (HTTP ${response.status}).`);
    const payload = await response.json();
    if (payload?.ok === false || payload?.error) throw new Error(String(payload.error || "Projects fetch failed."));
    const rows = Array.isArray(payload?.projects) ? payload.projects : [];
    return rows.map(normalizeRemoteProject).filter((item) => item.id);
  } catch (error) {
    console.warn("Unable to load projects from resource endpoint:", error);
    return loadRemoteProjectsFromBundle();
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
    const response = await fetchWithTimeout(url.toString(), { cache: "no-store" });
    if (!response.ok) return null;
    const payload = await response.json();
    if (payload?.ok === false) return null;
    const rows = Array.isArray(payload?.activities) ? payload.activities : [];
    return rows
      .map(normalizeRemoteActivity)
      .filter((item) => getCostActivityProjectKey(item) && item.id);
  } catch (error) {
    console.warn("Unable to load cost activities from activities resource endpoint:", error);
    return null;
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

const deriveDailyCostStatus = (item = {}, activity = null) => {
  const explicitStatus = String(
    getValueByAliases(item, ["status", "dailyStatus", "daily_status", "scheduleStatus", "schedule_status"]) || "",
  ).trim();

  const delayedValue = getValueByAliases(item, ["isDelayed", "is_delayed", "delayed"]);
  const delayedText = String(delayedValue || "").trim().toLowerCase();
  const isExplicitlyDelayed = delayedValue === true || ["true", "delayed", "yes", "1"].includes(delayedText);
  const isDateDelayed = Boolean(
    activity &&
      item?.date &&
      ((activity.startDate && item.date < activity.startDate) ||
        (activity.finishDate && item.date > activity.finishDate)),
  );
  const normalizedActivityStatus = String(activity?.status || "").trim().toLowerCase();
  const activityActualFinish = toDateInputValue(
    getValueByAliases(activity || {}, ["actualFinish", "actual_finish", "actualFinishDate", "actual_finish_date"]),
  );
  const isOperationallyAlignedCompletion =
    Boolean(activity) &&
    normalizedActivityStatus === "completed" &&
    Boolean(activity.finishDate) &&
    Boolean(activityActualFinish) &&
    activityActualFinish <= activity.finishDate;

  if (isOperationallyAlignedCompletion && activity && item?.date) return "On Schedule";

  if (activity && item?.date) return isDateDelayed ? "Delayed" : "On Schedule";
  if (isExplicitlyDelayed) return "Delayed";
  if (explicitStatus) return explicitStatus;
  return "On Schedule";
};

const normalizeDailyCostRecord = (item = {}, lookups = buildProjectIdentityLookups(loadProjects())) => ({
  projectId: resolveProjectIdFromDailyCost(item, lookups),
  activityId: String(getValueByAliases(item, ["activityId", "activity_id", "activity id"]) || item.activityId || "").trim(),
  activity: String(getValueByAliases(item, ["activity", "activityName", "activity_name", "name"]) || item.activity || "").trim(),
  costId: String(getValueByAliases(item, ["costId", "cost_id", "cost id"]) || item.costId || "").trim(),
  date: normalizeDateKey(getValueByAliases(item, ["date"]) || item.date),
  status: deriveDailyCostStatus(item),
  actualCost: parseBudgetValue(getValueByAliases(item, ["actualCost", "actual_cost", "amount"]) ?? item.actualCost),
  progress: clampPercent(getValueByAliases(item, ["progress", "percentComplete", "percent_complete", "% complete", "percent complete"]) ?? item.progress),
  earnedValue: parseBudgetValue(getValueByAliases(item, ["earnedValue", "earned_value", "earned value", "ev"]) ?? item.earnedValue),
});

const getDailyCostRecordKey = (item = {}) => {
  const projectId = String(item.projectId || "").trim();
  const activityId = String(item.activityId || "").trim();
  const costId = String(item.costId || "").trim();
  const date = normalizeDateKey(item.date);
  if (!projectId || !date || (!activityId && !costId)) return "";
  return `${projectId}::${activityId}::${costId}::${date}`;
};

const isSavedDailyCostMatch = (item = {}, expected = {}) => {
  const normalizedItem = normalizeDailyCostRecord(item);
  const expectedProjectId = String(expected.projectId || "").trim();
  const expectedActivityId = String(expected.activityId || "").trim();
  const expectedCostId = String(expected.costId || "").trim();
  const expectedDate = normalizeDateKey(expected.date);
  const sameProject = !expectedProjectId || normalizedItem.projectId === expectedProjectId;
  const sameDate = normalizedItem.date === expectedDate;
  const sameActivity =
    (expectedActivityId && normalizedItem.activityId === expectedActivityId) ||
    (expectedCostId && normalizedItem.costId === expectedCostId);
  const sameProgress = Math.abs((Number(normalizedItem.progress) || 0) - (Number(expected.progress) || 0)) < 0.0001;
  const sameActualCost = Math.abs((Number(normalizedItem.actualCost) || 0) - (Number(expected.actualCost) || 0)) < 0.01;
  return sameProject && sameDate && sameActivity && sameProgress && sameActualCost;
};

const dailyCostMatchesProjectFilter = (item = {}, projectFilter = {}) => {
  const projectId = String(projectFilter?.projectId || "").trim();
  const projectName = String(projectFilter?.projectName || "").trim().toLowerCase();
  if (!projectId && !projectName) return true;
  return isDailyCostForProject(item, projectId, projectName);
};

const loadRemoteDailyCosts = async (projectFilter = {}) => {
  const dataSourceUrl = window.DataBridge?.DEFAULT_DATA_SOURCE_URL;
  if (!dataSourceUrl) return null;
  const lookups = buildProjectIdentityLookups(loadProjects());
  try {
    const url = new URL(dataSourceUrl);
    url.searchParams.set("resource", "daily_costs");
    if (projectFilter?.projectId) url.searchParams.set("projectId", String(projectFilter.projectId));
    url.searchParams.set("_ts", String(Date.now()));
    const response = await fetchWithTimeout(url.toString(), { cache: "no-store" });
    if (!response.ok) return null;
    const payload = await response.json();
    if (payload?.ok === false) return null;
    const rows = Array.isArray(payload?.dailyCosts) ? payload.dailyCosts : [];
    return rows.map((row) => ({
      projectId: resolveProjectIdFromDailyCost(row, lookups),
      activityId: String(getValueByAliases(row, ["activityId", "activity_id", "activity id"]) || "").trim(),
      activity: String(getValueByAliases(row, ["activity", "activityName", "activity_name", "name"]) || "").trim(),
      costId: String(getValueByAliases(row, ["costId", "cost_id", "cost id"]) || "").trim(),
      date: normalizeDateKey(getValueByAliases(row, ["date"])),
      status: deriveDailyCostStatus(row),
      actualCost: parseBudgetValue(getValueByAliases(row, ["actualCost", "actual_cost", "amount"])),
      progress: clampPercent(getValueByAliases(row, ["progress", "percentComplete", "percent_complete", "% complete", "percent complete"])),
      earnedValue: parseBudgetValue(getValueByAliases(row, ["earnedValue", "earned_value", "earned value", "ev"])),
    })).filter((r) => r.projectId && r.date && hasPositiveActualCost(r));
  } catch (error) {
    console.warn("Unable to load daily costs from resource endpoint:", error);
    return null;
  }
};


const syncDailyCostsFromSheet = async (projectFilter = {}, prefetchedDailyCosts = null) => {
  const remoteDailyCosts = Array.isArray(prefetchedDailyCosts)
    ? prefetchedDailyCosts
    : await loadRemoteDailyCosts(projectFilter);

  if (!Array.isArray(remoteDailyCosts)) {
    if (!hasWarnedAboutCachedDailyCosts && typeof window.notify === "function") {
      hasWarnedAboutCachedDailyCosts = true;
      window.notify("Using cached daily costs because Google Sheets could not be reached.", "warning");
    }
    return false;
  }
  hasWarnedAboutCachedDailyCosts = false;

  const lookups = buildProjectIdentityLookups(loadProjects());
  const dedupedDailyCosts = new Map();

  // When syncing a single selected project, replace only that project slice with
  // remote rows. This keeps unrelated locally cached projects available while
  // still letting remote deletes remove stale rows for the active project.
  loadDailyCosts()
    .map((item) => normalizeDailyCostRecord(item, lookups))
    .filter((item) => getDailyCostRecordKey(item))
    .filter((item) => !dailyCostMatchesProjectFilter(item, projectFilter))
    .forEach((item) => dedupedDailyCosts.set(getDailyCostRecordKey(item), item));

  remoteDailyCosts
    .map((item) => normalizeDailyCostRecord(item, lookups))
    .filter((item) => getDailyCostRecordKey(item))
    .forEach((item) => dedupedDailyCosts.set(getDailyCostRecordKey(item), item));

  saveDailyCosts(Array.from(dedupedDailyCosts.values()));
  return true;
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
  if (!dataSourceUrl) return null;
  const lookups = buildProjectIdentityLookups(loadProjects());

  const parseCostRows = (payload = {}) => extractCostRowsFromPayload(payload)
    .map((row) => ({
      projectId: resolveProjectIdFromDailyCost({
        projectId: getValueByAliases(row, ["projectId", "project_id", "project id"]),
        projectName: getValueByAliases(row, ["project", "projectName", "project_name", "project name"]),
      }, lookups),
      activityRefId: extractActivityRefIdFromCostRow(row),
      activityName: String(getValueByAliases(row, ["activity", "activityName", "activity_name", "name"]) || "").trim(),
      costId: String(getValueByAliases(row, COST_METADATA_COST_ID_ALIASES) || "").trim(),
      plannedCost: parseBudgetValue(getValueByAliases(row, COST_METADATA_PLANNED_COST_ALIASES)),
      actualCost: parseBudgetValue(getValueByAliases(row, ["actualCost", "actual_cost", "amount"])),
      progressPercent: clampPercent(getValueByAliases(row, ["progress", "progressPercent", "progress_percent", "percentComplete", "percent_complete", "% complete", "percent complete"])),
      earnedValue: parseBudgetValue(getValueByAliases(row, ["earnedValue", "earned_value", "earned value", "ev"])),
      date: String(getValueByAliases(row, ["date", "createdAt", "created_at"]) || "").trim(),
    }))
    .filter((row) => row.projectId && (row.activityRefId || row.activityName) && (row.costId || parseBudgetValue(row.plannedCost) > 0));

  try {
    const fetchRows = async (includeProjectId) => {
      const url = new URL(dataSourceUrl);
      url.searchParams.set("resource", "costs");
      if (includeProjectId && projectFilter?.projectId) url.searchParams.set("projectId", String(projectFilter.projectId));
      url.searchParams.set("_ts", String(Date.now()));
      const response = await fetchWithTimeout(url.toString(), { cache: "no-store" });
      if (!response.ok) return null;
      const payload = await response.json();
      if (payload?.ok === false) return null;
      return parseCostRows(payload);
    };

    const filteredRows = await fetchRows(true);
    if (!Array.isArray(filteredRows)) return null;
    if (filteredRows.length || !projectFilter?.projectId) return filteredRows;

    const allRows = await fetchRows(false);
    if (!Array.isArray(allRows)) return null;
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
    return null;
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

const findResolvedCostActivity = (projectId, activityRefId, allActivities = loadCostActivities()) => {
  const project = loadProjects().map(normalizeProject).find((item) => item.id === projectId);
  const projectName = String(project?.name || "").trim();
  const normalizedProjectName = projectName.toLowerCase();
  const requestedActivityRefId = String(activityRefId || "").trim();
  const requestedLookup = normalizeLookup(requestedActivityRefId);
  const sourceActivities = dedupeCostActivities([
    ...(Array.isArray(allActivities) ? allActivities : []),
    ...loadCostActivities(),
  ]);
  const projectActivities = sourceActivities.filter((item) => isActivityForProject(item, projectId, normalizedProjectName));
  const activityMatchesIdentifier = (item = {}) => {
    const itemRefId = String(getActivityRefId(item) || "").trim();
    const itemCostId = String(item.costId || "").trim();
    const itemNameLookup = normalizeLookup(item.name);
    return itemRefId === requestedActivityRefId
      || (requestedLookup && normalizeLookup(itemRefId) === requestedLookup)
      || (requestedLookup && normalizeLookup(itemCostId) === requestedLookup)
      || (requestedLookup && itemNameLookup === requestedLookup);
  };
  const baseActivity = projectActivities.find(activityMatchesIdentifier);
  const normalizedActivityName = String(baseActivity?.name || "").trim().toLowerCase();
  const projectRows = getProjectCostData(projectId, sourceActivities).rows;
  const rowFallback = projectRows.find(activityMatchesIdentifier)
    || (normalizedActivityName
      ? projectRows.find((row) => String(row.name || "").trim().toLowerCase() === normalizedActivityName)
      : null)
    || projectRows.find((row) => String(row.costId || "").trim() && String(getActivityRefId(row) || "").trim() === requestedActivityRefId);
  const resolvedName = String(baseActivity?.name || rowFallback?.name || "").trim();
  const resolvedNameLookup = resolvedName.toLowerCase();
  const resolvedCostIdLookup = normalizeLookup(baseActivity?.costId || rowFallback?.costId || "");
  const relatedActivities = projectActivities.filter((item) => {
    const itemRefId = String(getActivityRefId(item) || "").trim();
    const itemName = String(item.name || "").trim().toLowerCase();
    const itemCostIdLookup = normalizeLookup(item.costId);
    return itemRefId === requestedActivityRefId
      || (requestedLookup && normalizeLookup(itemRefId) === requestedLookup)
      || (resolvedCostIdLookup && itemCostIdLookup === resolvedCostIdLookup)
      || (resolvedNameLookup && itemName === resolvedNameLookup);
  });
  const relatedRows = projectRows.filter((row) => {
    const rowRefId = String(getActivityRefId(row) || "").trim();
    const rowName = String(row.name || "").trim().toLowerCase();
    const rowCostIdLookup = normalizeLookup(row.costId);
    return rowRefId === requestedActivityRefId
      || (requestedLookup && normalizeLookup(rowRefId) === requestedLookup)
      || (resolvedCostIdLookup && rowCostIdLookup === resolvedCostIdLookup)
      || (resolvedNameLookup && rowName === resolvedNameLookup);
  });
  const costId = String(
    baseActivity?.costId
    || relatedActivities.find((item) => String(item.costId || "").trim())?.costId
    || rowFallback?.costId
    || relatedRows.find((row) => String(row.costId || "").trim())?.costId
    || "",
  ).trim();
  const plannedCost = Math.max(
    parseBudgetValue(baseActivity?.plannedCost),
    ...relatedActivities.map((item) => parseBudgetValue(item.plannedCost)),
    parseBudgetValue(rowFallback?.plannedCost),
    ...relatedRows.map((row) => parseBudgetValue(row.plannedCost)),
  );
  const durationDays = Math.max(
    Number(baseActivity?.durationDays) || 0,
    Number(rowFallback?.durationDays) || 0,
    ...relatedActivities.map((item) => Number(item.durationDays) || 0),
    ...relatedRows.map((row) => Number(row.durationDays) || 0),
  );

  if (!baseActivity && !rowFallback) return null;

  return normalizeCostActivity({
    ...(baseActivity || {}),
    ...(rowFallback || {}),
    id: getActivityRefId(baseActivity || rowFallback || {}) || requestedActivityRefId,
    activityRefId: getActivityRefId(baseActivity || rowFallback || {}) || requestedActivityRefId,
    projectId: projectId || baseActivity?.projectId || rowFallback?.projectId || "",
    projectName: projectName || baseActivity?.projectName || rowFallback?.projectName || "",
    name: resolvedName || baseActivity?.name || rowFallback?.name || "Untitled Activity",
    costId,
    plannedCost,
    durationDays,
    startDate: baseActivity?.startDate || rowFallback?.startDate || "",
    finishDate: baseActivity?.finishDate || rowFallback?.finishDate || "",
  });
};

const getProjectCostData = (projectId, allActivities = loadCostActivities()) => {
  const project = loadProjects().map(normalizeProject).find((item) => item.id === projectId);
  const projectName = String(project?.name || "").trim();
  const normalizedProjectName = projectName.toLowerCase();
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
    .filter((item) => isActivityForProject(item, projectId, normalizedProjectName))
    .sort(compareActivitiesByStartPriority);
  const daily = loadDailyCosts()
    .filter((item) => isDailyCostForProject(item, projectId, normalizedProjectName));
  const rawRows = activities.map((activity) => {
    const refId = getActivityRefId(activity);
    const rowCostId = String(activity.costId || "").trim();
    const dailyItems = daily.filter((entry) => {
      const entryActivityId = String(entry.activityId || "").trim();
      const entryCostId = String(entry.costId || "").trim();
      return entryActivityId === refId || (rowCostId && entryCostId === rowCostId);
    });
    const actualCost = dailyItems.reduce((sum, entry) => sum + parseBudgetValue(entry.actualCost), 0);
    const accumulatedProgress = dailyItems.reduce((sum, entry) => sum + clampPercent(entry.progress), 0);
    const progressPercent = dailyItems.length ? clampPercent(accumulatedProgress) : clampPercent(activity.progressPercent);
    const earnedValueFromDaily = dailyItems.reduce((sum, entry) => sum + parseBudgetValue(entry.earnedValue), 0);
    const plannedCostPerDay = Number(activity.durationDays) > 0
      ? parseBudgetValue(activity.plannedCost) / Number(activity.durationDays)
      : 0;
    const earnedValue = earnedValueFromDaily > 0
      ? earnedValueFromDaily
      : computeEarnedValue(plannedCostPerDay, dailyItems.length, progressPercent, activity.earnedValue);
    return { ...activity, progressPercent, actualCost, dailyItems, earnedValue };
  });

  // Consolidate duplicate activity rows (same costId/activity) so daily-cost updates
  // always appear on a single costing row instead of creating visual duplicates.
  const consolidated = new Map();
  rawRows.forEach((row) => {
    const refId = String(getActivityRefId(row) || "").trim();
    const costId = String(row.costId || "").trim();
    const projectKey = String(getCostActivityProjectKey(row) || "").trim();
    const key = [
      projectKey,
      costId || "-",
      refId || "-",
      String(row.name || "").trim().toLowerCase(),
      String(row.startDate || ""),
      String(row.finishDate || ""),
    ].join("::");
    if (!key) return;

    if (!consolidated.has(key)) {
      consolidated.set(key, row);
      return;
    }

    const existing = consolidated.get(key) || {};
    const mergedDailyItems = [...(existing.dailyItems || []), ...(row.dailyItems || [])];
    const uniqueDailyItems = Array.from(new Map(
      mergedDailyItems.map((entry) => {
        const dedupeKey = [
          String(entry.projectId || "").trim(),
          String(entry.activityId || "").trim(),
          String(entry.costId || "").trim(),
          normalizeDateKey(entry.date),
        ].join("::");
        return [dedupeKey, entry];
      }),
    ).values());

    consolidated.set(key, {
      ...existing,
      ...row,
      costId: String(existing.costId || row.costId || "").trim(),
      activityRefId: String(getActivityRefId(existing) || getActivityRefId(row) || "").trim(),
      plannedCost: parseBudgetValue(row.plannedCost) || parseBudgetValue(existing.plannedCost),
      earnedValue: (() => {
        const mergedProgress = uniqueDailyItems.reduce((sum, entry) => sum + clampPercent(entry.progress), 0);
        const mergedProgressPercent = uniqueDailyItems.length
          ? clampPercent(mergedProgress)
          : clampPercent(Number(row.progressPercent) || Number(existing.progressPercent) || 0);
        const mergedEarnedValueFromDaily = uniqueDailyItems.reduce((sum, entry) => sum + parseBudgetValue(entry.earnedValue), 0);
        if (mergedEarnedValueFromDaily > 0) return mergedEarnedValueFromDaily;
        const mergedPlannedCost = parseBudgetValue(row.plannedCost) || parseBudgetValue(existing.plannedCost);
        const mergedDurationDays = Math.max(Number(existing.durationDays) || 0, Number(row.durationDays) || 0);
        const mergedPlannedCostPerDay = mergedPlannedCost > 0 && mergedDurationDays > 0
          ? mergedPlannedCost / mergedDurationDays
          : 0;
        return computeEarnedValue(
          mergedPlannedCostPerDay,
          uniqueDailyItems.length,
          mergedProgressPercent,
          parseBudgetValue(row.earnedValue) || parseBudgetValue(existing.earnedValue),
        );
      })(),
      progressPercent: (() => {
        const mergedProgress = uniqueDailyItems.reduce((sum, entry) => sum + clampPercent(entry.progress), 0);
        if (uniqueDailyItems.length) return clampPercent(mergedProgress);
        return clampPercent(Number(existing.progressPercent) || Number(row.progressPercent) || 0);
      })(),
      actualCost: uniqueDailyItems.reduce((sum, entry) => sum + parseBudgetValue(entry.actualCost), 0),
      durationDays: Math.max(Number(existing.durationDays) || 0, Number(row.durationDays) || 0),
      dailyItems: uniqueDailyItems,
      name: existing.name || row.name || "Untitled Activity",
    });
  });

  return { rows: Array.from(consolidated.values()), activities, daily };
};

const buildSelectedProjectBannerMarkup = (project) => `<section class="selected-project-banner"><div><p class="selected-project-label">Selected Project</p><h3>${escapeHtml(formatProjectIdentityLabel(project))}</h3></div><a href="cost-management.html" class="ghost-btn">← Back to Projects</a></section>`;

const buildDetailsMarkup = (project, rows) => {
  const plannedCost = rows.reduce((sum, row) => sum + parseBudgetValue(row.plannedCost), 0);
  const actualCost = rows.reduce((sum, row) => sum + row.actualCost, 0);
  const earnedValue = rows.reduce((sum, row) => sum + parseBudgetValue(row.earnedValue), 0);
  const comparisonItems = [
    { key: "planned", label: "PC", value: plannedCost, fullLabel: "Total Planned Cost" },
    { key: "actual", label: "AC", value: actualCost, fullLabel: "Total Actual Cost" },
    { key: "earned", label: "EV", value: earnedValue, fullLabel: "Total Earned Value" },
  ];
  const variance = earnedValue - actualCost;
  const variancePercent = earnedValue ? (variance / earnedValue) * 100 : 0;
  const totalDuration = rows.reduce((sum, row) => sum + (Number(row.durationDays) || 0), 0);
  const avgActualPerDay = totalDuration > 0 ? actualCost / totalDuration : 0;

  const tableRows = rows.length
    ? rows.map((row) => {
      const hasPlannedCost = parseBudgetValue(row.plannedCost) > 0;
      const hasLoggedActualCost = row.actualCost !== null
        && row.actualCost !== undefined
        && String(row.actualCost).trim() !== "";
      const hasActualCost = parseBudgetValue(row.actualCost) > 0;
      const costIdCell = row.costId ? escapeHtml(row.costId) : "";
      const plannedCostCell = hasPlannedCost ? formatBudget(row.plannedCost) : "";
      const progressValue = clampPercent(row.progressPercent);
      const progressCell = Number.isFinite(progressValue) ? `${progressValue.toFixed(2)}%` : "";
      const actualCostCell = hasLoggedActualCost ? formatBudget(row.actualCost) : "";
      const hasEarnedValue = parseBudgetValue(row.earnedValue) > 0;
      const earnedValueCell = hasEarnedValue ? formatBudget(row.earnedValue) : "Not logged";
      const varianceValue = parseBudgetValue(row.earnedValue) - parseBudgetValue(row.actualCost);
      const varianceTone = varianceValue >= 0 ? "good" : "bad";
      const normalizedActivityStatus = String(row.status || "").trim().toLowerCase();
      let statusLabel = "Ready";
      let statusTone = "ready";
      if (!hasPlannedCost) {
        statusLabel = "Setup needed";
        statusTone = "needs-setup";
      } else if (normalizedActivityStatus === "completed" || progressValue >= 100) {
        statusLabel = "Completed";
        statusTone = "complete";
      } else if (normalizedActivityStatus === "delayed") {
        statusLabel = "Delayed";
        statusTone = "risk";
      } else if (normalizedActivityStatus === "in progress" || hasLoggedActualCost) {
        statusLabel = "In progress";
        statusTone = "active";
      } else if (normalizedActivityStatus === "not started" || progressValue <= 0) {
        statusLabel = "Not started";
        statusTone = "ready";
      }
      const durationCell = Number(row.durationDays) > 0 ? `${row.durationDays} days` : "Not set";
      const activityIdValue = getActivityRefId(row);
      const actionIdentifierValue = activityIdValue || String(row.costId || "").trim() || String(row.name || "").trim();
      const activityId = escapeHtml(actionIdentifierValue);
      const progressWidth = Math.max(0, Math.min(100, progressValue || 0));
      return `<tr class="cost-record-row"><td><span class="cost-id-pill ${costIdCell ? "" : "is-missing"}">${costIdCell || "Add ID"}</span></td><td><div class="cost-activity-cell"><strong>${escapeHtml(row.name)}</strong><span>Activity ID: ${escapeHtml(activityIdValue || "—")}</span></div></td><td><span class="cost-muted-value">${durationCell}</span></td><td><div class="cost-progress-cell"><span>${progressCell || "0.00%"}</span><div class="cost-progress-track" aria-hidden="true"><i style="width:${progressWidth}%"></i></div></div></td><td class="planned-cost-cell"><span class="planned-cost-text cost-money">${plannedCostCell || "Not set"}</span></td><td><span class="cost-money">${actualCostCell || "Not logged"}</span></td><td><span class="cost-money">${earnedValueCell}</span><small class="cost-variance-note ${varianceTone}">${formatBudget(varianceValue)}</small></td><td><span class="cost-status-badge ${statusTone}">${statusLabel}</span></td><td class="actions-col"><button type="button" class="action-menu-trigger cost-actions-button" data-cost-actions="${activityId}" aria-label="Open cost actions for ${escapeHtml(row.name)}" aria-expanded="false">⋮</button><div class="project-actions-menu hidden" data-cost-menu="${activityId}" role="menu" aria-label="Cost actions"><button type="button" class="project-action-btn edit-cost-meta-btn" data-activity-id="${activityId}" role="menuitem">Add / Edit Cost Details</button><button type="button" class="project-action-btn view-daily-cost-btn" data-activity-id="${activityId}" role="menuitem">View / Add Daily Cost</button></div></td></tr>`;
    }).join("")
    : '<tr><td colspan="9" class="empty-cell"><strong>No costing records yet.</strong><span>Add activities to start tracking costs, then use Add / Edit Cost Details to assign budgets.</span></td></tr>';

  const maxCost = Math.max(...comparisonItems.map((item) => item.value), 1);
  const comparisonBarsMarkup = comparisonItems
    .map((item) => {
      const height = item.value > 0 ? Math.max(10, Math.round((item.value / maxCost) * 100)) : 0;
      return `<div class="bar-wrap"><strong>${formatBudget(item.value)}</strong><div class="bar ${item.key}" style="height:${height}%"></div><p title="${item.fullLabel}">${item.label}</p></div>`;
    })
    .join("");
  const comparisonLegendMarkup = comparisonItems
    .map((item) => `<li><span class="legend-dot ${item.key}"></span> ${item.label} <small>(${item.fullLabel})</small></li>`)
    .join("");
  const comparisonTotal = comparisonItems.reduce((sum, item) => sum + item.value, 0);
  const donutSegments = comparisonItems.map((item, index) => {
    const start = comparisonTotal > 0
      ? (comparisonItems.slice(0, index).reduce((sum, entry) => sum + entry.value, 0) / comparisonTotal) * 100
      : (index / comparisonItems.length) * 100;
    const end = comparisonTotal > 0
      ? ((comparisonItems.slice(0, index + 1).reduce((sum, entry) => sum + entry.value, 0) / comparisonTotal) * 100)
      : (((index + 1) / comparisonItems.length) * 100);
    const color = item.key === "budget"
      ? "#f0c23f"
      : item.key === "planned"
        ? "#2f67ff"
        : item.key === "earned"
          ? "#8a52ff"
          : "#38c26d";
    return `${color} ${start.toFixed(2)}% ${end.toFixed(2)}%`;
  }).join(", ");
  const donutLegendMarkup = comparisonItems
    .map((item) => `<li><span class="legend-dot ${item.key}"></span>${item.label} <small>(${item.fullLabel})</small> <strong>${formatBudget(item.value)}</strong></li>`)
    .join("");
  const varianceLabel = variance >= 0 ? "Under budget" : "Over budget";
  const varianceClass = variance >= 0 ? "good" : "bad";
  const topRows = rows
    .slice()
    .map((row) => ({ ...row, costVariance: parseBudgetValue(row.earnedValue) - parseBudgetValue(row.actualCost) }))
    .filter((row) => row.costVariance < 0)
    .sort((a, b) => a.costVariance - b.costVariance)
    .slice(0, 5)
    .map((row) => `<tr><td>${escapeHtml(row.costId || "-")}</td><td>${escapeHtml(row.name)}</td><td>${formatBudget(row.plannedCost)}</td><td>${formatBudget(row.actualCost)}</td><td>${formatBudget(row.earnedValue)}</td><td class="bad">${formatBudget(row.costVariance)}</td></tr>`)
    .join("") || '<tr><td colspan="6" class="empty-cell">No over budget activities.</td></tr>';

  const costRecordHealth = rows.filter((row) => parseBudgetValue(row.plannedCost) > 0).length;
  const costRecordSummaryMarkup = `<div class="cost-record-summary" aria-label="Costing record summary"><article><span>Cost items</span><strong>${rows.length}</strong></article><article><span>Budgeted items</span><strong>${costRecordHealth}/${rows.length || 0}</strong></article><article><span>Actual spend</span><strong>${formatBudget(actualCost)}</strong></article><article><span>Variance</span><strong class="${varianceClass}">${formatBudget(variance)}</strong></article></div>`;

  return `<nav class="details-tabs"><button class="tab-btn active" data-tab="overview" type="button">Overview</button><button class="tab-btn" data-tab="costing" type="button">Costing Record</button></nav>
  <section class="details-tab-panel" data-panel="overview"><section class="details-kpis">
  <article class="kpi-card"><h4><span class="kpi-icon" aria-hidden="true"><svg viewBox="0 0 24 24" fill="none" aria-hidden="true"><rect x="3" y="6.5" width="18" height="11" rx="2.5"/><path d="M3 10.5h18"/><path d="M7.5 14h3"/></svg></span>Total Planned Cost</h4><p>${formatBudget(plannedCost)}</p></article>
  <article class="kpi-card"><h4><span class="kpi-icon" aria-hidden="true"><svg viewBox="0 0 24 24" fill="none" aria-hidden="true"><path d="M7 3.5h10v17l-2-1.2-2 1.2-2-1.2-2 1.2-2-1.2-2 1.2v-17h2"/><path d="M9 8h6M9 12h6M9 16h4"/></svg></span>Total Actual Cost</h4><p>${formatBudget(actualCost)}</p></article>
  <article class="kpi-card"><h4><span class="kpi-icon" aria-hidden="true"><svg viewBox="0 0 24 24" fill="none" aria-hidden="true"><path d="M4 19h16"/><path d="M6.5 15.5 10 12l3 2.5 4.5-5"/><path d="M17.5 9.5H14"/></svg></span>Total Earned Value</h4><p>${formatBudget(earnedValue)}</p></article>
  <article class="kpi-card"><h4><span class="kpi-icon" aria-hidden="true"><svg viewBox="0 0 24 24" fill="none" aria-hidden="true"><path d="M12 4v16"/><path d="M7 8h10"/><path d="m5.5 8-2.5 4h5l-2.5-4Zm13 0-2.5 4h5l-2.5-4Z"/><path d="M8 12a3 3 0 0 1-6 0M22 12a3 3 0 0 1-6 0"/></svg></span>Total Cost Variance</h4><p class="${varianceClass}">${formatBudget(variance)}</p><small>${varianceLabel}</small></article>
  <article class="kpi-card"><h4><span class="kpi-icon" aria-hidden="true"><svg viewBox="0 0 24 24" fill="none" aria-hidden="true"><circle cx="12" cy="13" r="7"/><path d="M12 13V9.5"/><path d="m12 13 2.8 1.6"/><path d="M9.5 3h5"/></svg></span>Total Duration</h4><p>${totalDuration} days</p></article></section>
  <section class="overview-grid"><article class="panel chart-panel"><h3><span class="panel-icon" aria-hidden="true"><svg viewBox="0 0 24 24" fill="none" aria-hidden="true"><path d="M4 19h16"/><rect x="6" y="11" width="3" height="6" rx="1"/><rect x="11" y="8" width="3" height="9" rx="1"/><rect x="16" y="5" width="3" height="12" rx="1"/></svg></span>Budget vs Actual</h3><div class="bars"><div class="bars-grid"><span>${formatBudget(maxCost)}</span><span>${formatBudget(maxCost * 0.75)}</span><span>${formatBudget(maxCost * 0.5)}</span><span>${formatBudget(maxCost * 0.25)}</span><span>0</span></div><div class="bars-track">${comparisonBarsMarkup}</div><ul class="bars-legend">${comparisonLegendMarkup}</ul></div></article>
  <article class="panel summary-panel"><h3><span class="panel-icon" aria-hidden="true"><svg viewBox="0 0 24 24" fill="none" aria-hidden="true"><rect x="5" y="4" width="14" height="16" rx="2"/><path d="M9 4.5h6v3H9z"/><path d="M8.5 11h7M8.5 15h7"/></svg></span>Cost Summary</h3><ul><li><span>Total Planned Cost</span><strong>${formatBudget(plannedCost)}</strong></li><li><span>Total Actual Cost</span><strong>${formatBudget(actualCost)}</strong></li><li><span>Total Earned Value</span><strong>${formatBudget(earnedValue)}</strong></li><li><span>Total Cost Variance</span><strong class="${varianceClass}">${formatBudget(variance)}</strong></li><li><span>Variance Percent</span><strong class="${varianceClass}">${variancePercent.toFixed(2)}%</strong></li><li><span>Total Duration</span><strong>${totalDuration} days</strong></li><li><span>Average Cost per Day (Actual)</span><strong>${formatBudget(avgActualPerDay)}</strong></li></ul></article>
  <article class="panel donut-panel"><h3><span class="panel-icon" aria-hidden="true"><svg viewBox="0 0 24 24" fill="none" aria-hidden="true"><path d="M12 4a8 8 0 1 1-8 8"/><path d="M12 12V4"/><path d="M12 12h8"/></svg></span>Cost Status</h3><div class="donut-wrap"><div class="donut" style="background: conic-gradient(${donutSegments});"></div><ul class="status-list">${donutLegendMarkup}</ul></div></article>
  <article class="panel table-panel"><h3><span class="panel-icon" aria-hidden="true"><svg viewBox="0 0 24 24" fill="none" aria-hidden="true"><path d="m14 4 6 6-3 1-3 7-3-3-7 3 3-7-1-3 6-6z"/></svg></span>Top Over Budget Activities</h3><table class="top-over-budget-table"><colgroup><col class="col-cost-id"><col class="col-activity"><col class="col-planned"><col class="col-actual"><col class="col-earned"><col class="col-variance"></colgroup><thead><tr><th>Cost ID</th><th>Activity</th><th>Planned Cost</th><th>Actual Cost</th><th>Earned Value</th><th>Cost Variance</th></tr></thead><tbody>${topRows}</tbody></table></article></section></section>
  <section class="details-tab-panel hidden" data-panel="costing">
  <div class="cost-record-workspace"><div class="cost-record-head cost-record-head-enhanced"><div><p class="eyebrow">Cost ledger</p><h3>Activity costing record</h3><p>Track each activity from planned budget setup through daily spend and earned value updates.</p></div><div class="cost-record-actions"><span class="cost-record-hint">Use the actions menu to add budgets or daily costs.</span></div></div>${costRecordSummaryMarkup}<section class="panel cost-record-table-panel"><table class="cost-table"><thead><tr><th>Cost ID</th><th>Activity</th><th>Duration</th><th>Progress</th><th>Planned Cost</th><th>Actual Cost</th><th>Earned Value</th><th>Status</th><th>Actions</th></tr></thead><tbody>${tableRows}</tbody></table></section><div class="info-banner cost-record-tip"><p><strong>Tip:</strong> Planned Cost is view-only in this table and can be changed only via “Add / Edit Cost Details”. Daily cost entries refresh Actual Cost and Earned Value after saving.</p></div></div></section>
  <section class="daily-cost-modal hidden" id="dailyCostModal"></section>
  <section class="cost-meta-modal hidden" id="costMetaModal"></section>`;
};

const renderDailyCostModal = (projectId, activityId, allActivities = loadCostActivities()) => {
  const modal = detailsView.querySelector("#dailyCostModal");
  const project = loadProjects().map(normalizeProject).find((item) => item.id === projectId);
  const projectName = String(project?.name || "").trim();
  const normalizedProjectName = projectName.toLowerCase();
  const activities = allActivities;
  const dailyCosts = loadDailyCosts();
  const activity = findResolvedCostActivity(projectId, activityId, activities);
  if (!modal || !activity) return;

  const requestedActivityIdentifier = String(activityId || "").trim();
  const resolvedActivityRefId = String(getActivityRefId(activity) || requestedActivityIdentifier).trim();
  const normalizedActivityName = String(activity.name || "").trim().toLowerCase();
  const activityCostId = String(activity.costId || "").trim();
  const activityName = String(activity.name || "").trim();
  const activityPlannedCost = parseBudgetValue(activity.plannedCost);
  const activityDurationDays = Number(activity.durationDays) || 0;
  const activityPlannedCostPerDay = activityDurationDays > 0 ? (activityPlannedCost / activityDurationDays) : 0;
  if (!activityCostId || activityPlannedCost <= 0) {
    alert("Please add Cost ID and Planned Cost first before adding daily costs for this activity.");
    editCostMetadata(projectId, activityId, allActivities);
    return;
  }
  const availableDates = buildDateRangeOptions(activity.startDate, activity.finishDate);
  const hasAvailableDates = availableDates.length > 0;
  const todayIso = new Date().toISOString().slice(0, 10);
  const fallbackDate = hasAvailableDates ? (availableDates.includes(todayIso) ? todayIso : availableDates[0]) : todayIso;
  const dateOptions = availableDates
    .map((dateValue) => `<option value="${dateValue}">${formatLongHumanDate(dateValue)}</option>`)
    .join("");

  const entries = dailyCosts
    .filter((item) => {
      if (!isDailyCostForProject(item, projectId, normalizedProjectName)) return false;
      const entryActivityId = String(item.activityId || "").trim();
      const entryCostId = String(item.costId || "").trim();
      const entryActivityName = String(item.activity || "").trim().toLowerCase();
      const nameMatches = entryActivityName && entryActivityName === normalizedActivityName;
      return entryActivityId === resolvedActivityRefId || (activityCostId && entryCostId === activityCostId) || nameMatches;
    })
    .filter((item) => parseBudgetValue(item.actualCost) > 0)
    .sort((a, b) => String(a.date || "").localeCompare(String(b.date || "")));
  const rows = entries.length
    ? entries.map((entry) => {
      const progressValue = Number(entry.progress);
      const hasManualProgress = Number.isFinite(progressValue);
      const earnedValue = hasManualProgress ? activityPlannedCostPerDay * (progressValue / 100) : Number.NaN;
      const progressLabel = hasManualProgress ? `${progressValue.toFixed(2)}%` : "—";
      const earnedValueLabel = Number.isFinite(earnedValue) ? formatBudget(earnedValue) : "—";
      const status = deriveDailyCostStatus(entry, activity);
      const isDelayed = status.toLowerCase() === "delayed";
      const statusMarkup = isDelayed ? '<span class="cost-status-badge risk">Delayed</span>' : `<span class="cost-status-badge complete">${escapeHtml(status)}</span>`;
      return `<tr><td>${formatHumanDate(entry.date)}</td><td>${statusMarkup}</td><td>${progressLabel}</td><td>${formatBudget(activityPlannedCostPerDay)}</td><td>${formatBudget(entry.actualCost)}</td><td>${earnedValueLabel}</td><td><button type="button" class="daily-cost-delete-btn" data-delete-date="${entry.date}">Delete</button></td></tr>`;
    }).join("")
    : '<tr><td colspan="7" class="empty-cell">No daily costs recorded yet.</td></tr>';
  modal.classList.remove("hidden");
  modal.innerHTML = `<div class="daily-cost-dialog panel" role="dialog" aria-modal="true" aria-labelledby="dailyCostTitle"><div class="daily-cost-head"><h3 id="dailyCostTitle">${escapeHtml(activity.name)} Daily Cost</h3><button type="button" class="daily-cost-close" id="closeDailyModalBtn" aria-label="Close">×</button></div><p class="daily-cost-range">📅 ${escapeHtml(formatLongHumanDate(activity.startDate))} to ${escapeHtml(formatLongHumanDate(activity.finishDate))}</p>
    <section class="daily-cost-section"><h4>Add Daily Cost</h4><p class="daily-cost-note">Flow: 1) Add entry, 2) Save to Google Sheet, 3) Daily Cost Records refresh from sheet.</p><form id="dailyCostForm" class="daily-cost-form"><label><span>Select Date</span><select name="datePreset" required><option value="" disabled ${hasAvailableDates ? "" : "selected"}>${hasAvailableDates ? "Choose date" : "No in-range working dates available"}</option>${dateOptions}<option value="__custom__">+ Add Date</option></select></label><label id="customDateField" class="hidden"><span>Add Date</span><input name="customDate" type="date" value="${escapeHtml(fallbackDate)}"></label><label><span>Progress/Day (%)</span><input name="progress" type="number" min="0" max="100" step="0.01" placeholder="Enter progress" required></label><label><span>Daily Cost (₱)</span><input name="actualCost" type="number" min="0" step="0.01" placeholder="Enter amount" required></label><button class="primary-btn" type="submit">Add Daily Cost</button></form><p class="daily-cost-duplicate-hint" id="dailyCostDuplicateHint" hidden></p></section>
    <section class="daily-cost-section"><h4>Daily Cost Records</h4><div class="daily-cost-table-wrap"><table><thead><tr><th>Date</th><th>Status</th><th>Progress/Day</th><th>Planned Cost/Day (₱)</th><th>Actual Cost/Day (₱)</th><th>Earned Value/Day (₱)</th><th>Action</th></tr></thead><tbody>${rows}</tbody></table></div></section>
    <div class="daily-cost-footer"><button type="button" class="ghost-btn" id="closeDailyModalBtnFooter">Close</button></div></div>`;

  modal.querySelector("#closeDailyModalBtn")?.addEventListener("click", () => modal.classList.add("hidden"));
  modal.querySelector("#closeDailyModalBtnFooter")?.addEventListener("click", () => modal.classList.add("hidden"));
  modal.querySelectorAll("[data-delete-date]").forEach((button) => button.addEventListener("click", async () => {
    const date = String(button.dataset.deleteDate || "");
    const resolvedProjectId = String(projectId || activity.projectId || "").trim();
    const resolvedProjectName = String(activity.projectName || activity.project || projectName || "").trim();
    if (!resolvedProjectId) {
      alert("Unable to delete daily cost because Project ID is missing.");
      return;
    }
    try {
      await postToDataSource("daily_costs", "delete", { dailyCost: { projectId: resolvedProjectId, costId: activityCostId, activityId: resolvedActivityRefId, date } });
    } catch (error) {
      console.warn("Unable to delete daily cost from Google Sheets:", error);
      alert(`Unable to delete in strict mode. ${error?.message || "Missing project/cost parent record."}`);
      return;
    }
    await syncDailyCostsFromSheet({ projectId, projectName: normalizedProjectName });
    const activeTab = detailsView.querySelector(".tab-btn.active")?.dataset.tab || "overview";
    const nextActivities = loadCostActivities();
    showProjectDetails(projectId, activeTab, nextActivities);
    renderDailyCostModal(projectId, resolvedActivityRefId);
  }));
  const datePresetSelect = modal.querySelector("#dailyCostForm select[name=\"datePreset\"]");
  const customDateField = modal.querySelector("#customDateField");
  const customDateInput = modal.querySelector("#dailyCostForm input[name=\"customDate\"]");
  const progressInput = modal.querySelector("#dailyCostForm input[name=\"progress\"]");
  const actualCostInput = modal.querySelector("#dailyCostForm input[name=\"actualCost\"]");
  const dailyCostForm = modal.querySelector("#dailyCostForm");
  const resolveSelectedDate = () => {
    if (!datePresetSelect) return "";
    if (datePresetSelect.value === "__custom__") return normalizeDateKey(customDateInput?.value || "");
    return normalizeDateKey(datePresetSelect.value || "");
  };
  const updateDailyCostSubmitMode = () => {
    const submitButton = dailyCostForm?.querySelector('button[type="submit"]');
    const hint = modal.querySelector("#dailyCostDuplicateHint");
    if (!(submitButton instanceof HTMLButtonElement)) return;
    const selectedDate = resolveSelectedDate();
    const hasExistingEntry = entries.some((entry) => normalizeDateKey(entry.date) === selectedDate);
    submitButton.textContent = hasExistingEntry ? "Update Daily Cost" : "Add Daily Cost";
    submitButton.dataset.defaultLabel = submitButton.textContent;
    if (hint instanceof HTMLElement) {
      hint.hidden = !hasExistingEntry;
      hint.textContent = hasExistingEntry
        ? "A record already exists for this date, so saving will update it instead of creating a duplicate."
        : "";
    }
  };
  const applySelectedDateDefaults = () => {
    if (!(progressInput instanceof HTMLInputElement) || !(actualCostInput instanceof HTMLInputElement)) return;
    const selectedDate = resolveSelectedDate();
    const existingEntry = entries.find((entry) => normalizeDateKey(entry.date) === selectedDate);
    if (existingEntry) {
      const existingProgress = Number(existingEntry.progress);
      const existingActualCost = Number(existingEntry.actualCost);
      progressInput.value = Number.isFinite(existingProgress) ? existingProgress.toFixed(2) : "";
      actualCostInput.value = Number.isFinite(existingActualCost) ? existingActualCost.toFixed(2) : "";
      return;
    }
    progressInput.value = "";
    actualCostInput.value = "";
  };
  if (datePresetSelect) {
    if (hasAvailableDates) {
      datePresetSelect.value = fallbackDate;
    } else {
      datePresetSelect.value = "__custom__";
    }
    datePresetSelect.addEventListener("change", () => {
      const isCustom = datePresetSelect.value === "__custom__";
      customDateField?.classList.toggle("hidden", !isCustom);
      if (customDateInput) customDateInput.required = isCustom;
      updateDailyCostSubmitMode();
      applySelectedDateDefaults();
    });
    customDateField?.classList.toggle("hidden", datePresetSelect.value !== "__custom__");
  }
  customDateInput?.addEventListener("change", () => {
    updateDailyCostSubmitMode();
    applySelectedDateDefaults();
  });
  updateDailyCostSubmitMode();
  applySelectedDateDefaults();

  dailyCostForm?.addEventListener("submit", async (event) => {
    event.preventDefault();
    const form = event.currentTarget;
    if (form.classList.contains("is-saving")) return;
    const formData = new FormData(event.currentTarget);
    const date = resolveSelectedDate();
    const rawProgress = String(formData.get("progress") || "").trim();
    const actualCost = parseBudgetValue(formData.get("actualCost"));
    const progress = rawProgress === "" ? Number.NaN : Number(rawProgress);
    const currentDailyCosts = loadDailyCosts();
    if (!date) {
      alert("Please select a date.");
      return;
    }
    if (Number.isNaN(progress) || progress < 0 || progress > 100) {
      alert("Progress must be between 0 and 100.");
      return;
    }
    if (actualCost <= 0) {
      alert("Daily cost must be greater than 0.");
      return;
    }
    const activityStartDate = String(activity.startDate || "");
    const activityFinishDate = String(activity.finishDate || "");
    const selectedDate = new Date(`${date}T00:00:00`);
    if (Number.isNaN(selectedDate.getTime())) {
      alert("Please provide a valid date.");
      return;
    }
    const isDelayedByDate = Boolean((activityStartDate && date < activityStartDate) || (activityFinishDate && date > activityFinishDate));
    const normalizedActivityStatus = String(activity.status || "").trim().toLowerCase();
    const activityActualFinishDate = toDateInputValue(
      getValueByAliases(activity || {}, ["actualFinish", "actual_finish", "actualFinishDate", "actual_finish_date"]),
    );
    const isOperationallyAlignedCompletion =
      normalizedActivityStatus === "completed" &&
      Boolean(activityFinishDate) &&
      Boolean(activityActualFinishDate) &&
      activityActualFinishDate <= activityFinishDate;
    const isDelayed = isOperationallyAlignedCompletion ? false : isDelayedByDate;
    const status = isDelayed ? "Delayed" : "On Schedule";
    const existingIndex = currentDailyCosts.findIndex((item) => {
      if (!isDailyCostForProject(item, projectId, normalizedProjectName)) return false;
      const itemDate = String(item.date || "");
      const itemActivityId = String(item.activityId || "").trim();
      const itemCostId = String(item.costId || "").trim();
      const sameActivity = itemActivityId === resolvedActivityRefId || (activityCostId && itemCostId === activityCostId);
      return sameActivity && itemDate === date;
    });
    const existingDailyCost = existingIndex >= 0 ? currentDailyCosts[existingIndex] : null;
    const resolvedProjectId = String(projectId || activity.projectId || "").trim();
    const resolvedProjectName = String(activity.projectName || activity.project || projectName || "").trim();
    if (!resolvedProjectId) {
      alert("Unable to save daily cost because Project ID is missing.");
      return;
    }
    if (!activityCostId) {
      alert("Strict mode: please add a valid Cost ID first before saving daily costs.");
      return;
    }
    const earnedValue = Number((activityPlannedCostPerDay * (progress / 100)).toFixed(2));
    if (existingDailyCost) {
      const existingProgress = Number(existingDailyCost.progress);
      const existingActualCost = Number(existingDailyCost.actualCost);
      const existingEarnedValue = Number(existingDailyCost.earnedValue);
      const existingStatus = deriveDailyCostStatus(existingDailyCost, activity);
      const isSameProgress = Number.isFinite(existingProgress) && Math.abs(existingProgress - progress) < 0.0001;
      const isSameActualCost = Number.isFinite(existingActualCost) && Math.abs(existingActualCost - actualCost) < 0.0001;
      const isSameEarnedValue = Number.isFinite(existingEarnedValue) && Math.abs(existingEarnedValue - earnedValue) < 0.0001;
      const isSameStatus = existingStatus === status;
      if (isSameProgress && isSameActualCost && isSameEarnedValue && isSameStatus) {
        alert("No changes detected for this date. Existing daily cost record is already up to date.");
        return;
      }
    }
    const dailyCostPayload = {
      projectId: resolvedProjectId,
      project: resolvedProjectName,
      costId: activityCostId,
      activityId: resolvedActivityRefId,
      activity: activityName,
      plannedCost: activityPlannedCost,
      plannedCostPerDay: activityPlannedCostPerDay,
      progress,
      date,
      status,
      actualCost,
      earnedValue,
      isDelayed,
    };
    setFormSavingState(form, true, existingIndex >= 0 ? "Updating…" : "Adding…");
    let saveVerifiedAfterTransportError = false;
    try {
      try {
        const dailyCostAction = existingIndex >= 0 ? "update" : "create";
        await postToDataSource("daily_costs", dailyCostAction, { dailyCost: dailyCostPayload });
      } catch (error) {
        console.warn("Daily cost POST returned an error; checking whether Google Sheets saved it anyway:", error);
        const syncedAfterError = await syncDailyCostsFromSheet({ projectId, projectName: normalizedProjectName });
        saveVerifiedAfterTransportError = Boolean(
          syncedAfterError && loadDailyCosts().some((item) => isSavedDailyCostMatch(item, dailyCostPayload)),
        );
        if (!saveVerifiedAfterTransportError) {
          alert(`Unable to save to Google Sheets. ${error?.message || "Please check Apps Script deployment permissions and try again."}`);
          return;
        }
      }

      try {
        await syncDailyCostsFromSheet({ projectId, projectName: normalizedProjectName });
        await syncCostSummaryToSheet({ projectId, projectName, activity });
        const refreshedMetadataRows = await loadRemoteCostMetadata({ projectId, projectName: normalizedProjectName });
        applyCostMetadataRows(refreshedMetadataRows);
      } catch (error) {
        console.warn("Daily cost was saved, but the follow-up refresh/summary sync failed:", error);
        if (typeof window.notify === "function") {
          window.notify("Daily cost saved. Refresh the page if the latest totals do not appear yet.", "warning");
        }
      }

      if (typeof window.notify === "function") {
        window.notify(
          saveVerifiedAfterTransportError
            ? "Daily cost saved in Google Sheets. The Apps Script response could not be read, but the saved row was verified."
            : existingIndex >= 0 ? "Daily cost updated successfully." : "Daily cost added successfully.",
          "success",
        );
      }
      const activeTab = detailsView.querySelector(".tab-btn.active")?.dataset.tab || "overview";
      const nextActivities = loadCostActivities();
      showProjectDetails(projectId, activeTab, nextActivities);
      renderDailyCostModal(projectId, resolvedActivityRefId);
    } finally {
      if (form.isConnected) setFormSavingState(form, false);
    }
  });
};

let detailViewListenerController = null;

const syncCostSummaryToSheet = async ({ projectId, projectName, activity }) => {
  const resolvedProjectId = String(projectId || activity?.projectId || "").trim();
  const resolvedProjectName = String(projectName || activity?.projectName || activity?.project || "").trim();
  const activityId = String(getActivityRefId(activity) || "").trim();
  const costId = String(activity?.costId || "").trim();
  if (!resolvedProjectId || !activityId || !costId) return;

  const refreshed = getProjectCostData(resolvedProjectId, loadCostActivities()).rows
    .find((row) => String(getActivityRefId(row) || "").trim() === activityId || String(row.costId || "").trim() === costId);
  if (!refreshed) return;

  const durationDays = Number(refreshed.durationDays) || 0;
  const plannedCost = parseBudgetValue(refreshed.plannedCost);
  const plannedCostPerDay = durationDays > 0 ? plannedCost / durationDays : 0;
  await postToDataSource("costs", "update", {
    skipDailyCostSync: true,
    cost: {
      costId,
      projectId: resolvedProjectId,
      project: resolvedProjectName,
      activityId,
      activity: refreshed.name || activity?.name || "",
      duration: durationDays,
      category: "Planned Cost",
      date: refreshed.startDate || activity?.startDate || "",
      plannedCost,
      plannedCostPerDay,
      progress: clampPercent(refreshed.progressPercent),
      progressPercent: clampPercent(refreshed.progressPercent),
      percentComplete: clampPercent(refreshed.progressPercent),
      actualCost: parseBudgetValue(refreshed.actualCost),
      earnedValue: parseBudgetValue(refreshed.earnedValue),
      notes: `Activity ID: ${activityId}`,
    },
  });
};

const showProjectDetails = (projectId, activeTab = "overview", allActivities = loadCostActivities()) => {
  const project = loadProjects().map(normalizeProject).find((item) => item.id === projectId);
  if (!project || isArchivedProject(project) || !selectionView || !detailsView || !selectedProjectBannerHost) return false;
  selectionView.classList.add("hidden");
  costPageHero?.classList.add("hidden");
  selectedProjectBannerHost.classList.remove("hidden");
  selectedProjectBannerHost.innerHTML = buildSelectedProjectBannerMarkup(project);
  detailsView.classList.remove("hidden");
  const { rows } = getProjectCostData(projectId, allActivities);
  detailsView.innerHTML = buildDetailsMarkup(project, rows);
  detailViewListenerController?.abort();
  detailViewListenerController = typeof AbortController === "function" ? new AbortController() : null;
  const listenerOptions = detailViewListenerController ? { signal: detailViewListenerController.signal } : undefined;

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
  }, listenerOptions));

  const closeCostActionMenus = () => {
    detailsView.querySelectorAll("[data-cost-menu]").forEach((menu) => {
      menu.classList.add("hidden");
      menu.classList.remove("open-up");
    });
    detailsView.querySelectorAll(".cost-record-row.menu-open").forEach((row) => row.classList.remove("menu-open"));
    detailsView.querySelectorAll("[data-cost-actions]").forEach((trigger) => trigger.setAttribute("aria-expanded", "false"));
  };

  const positionCostActionMenu = (menu) => {
    if (!(menu instanceof HTMLElement)) return;

    menu.classList.remove("open-up");

    const menuRect = menu.getBoundingClientRect();
    const tablePanelRect = menu.closest(".cost-record-table-panel")?.getBoundingClientRect();
    const lowerBoundary = Math.min(
      window.innerHeight - 16,
      tablePanelRect ? tablePanelRect.bottom - 8 : window.innerHeight - 16,
    );

    if (menuRect.bottom > lowerBoundary) {
      menu.classList.add("open-up");
    }
  };

  detailsView.querySelectorAll("[data-cost-actions]").forEach((trigger) => {
    trigger.addEventListener("click", (event) => {
      const activityId = trigger.dataset.costActions;
      const menu = detailsView.querySelector(`[data-cost-menu="${getCssEscapedValue(activityId)}"]`);
      if (!menu) return;
      const isOpen = !menu.classList.contains("hidden");
      closeCostActionMenus();
      if (!isOpen) {
        menu.classList.remove("hidden");
        trigger.closest(".cost-record-row")?.classList.add("menu-open");
        trigger.setAttribute("aria-expanded", "true");
        positionCostActionMenu(menu);
      }
      event.stopPropagation();
    }, listenerOptions);
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
  }, listenerOptions);

  return true;
};

const saveCostActivityOverrides = (items = []) => {
  costActivitiesState = dedupeCostActivities(items);
  persistToLocalStorage(COST_ACTIVITIES_LOCAL_STORAGE_KEY, costActivitiesState);
};
const renderCostMetadataModal = (projectId, activityRefId, target) => {
  const modal = detailsView.querySelector("#costMetaModal");
  if (!modal) return;

  modal.classList.remove("hidden");
  const hasExistingCost = Boolean(String(target.costId || "").trim()) || parseBudgetValue(target.plannedCost) > 0;
  const plannedPreview = formatBudget(parseBudgetValue(target.plannedCost));
  modal.innerHTML = `<div class="cost-meta-dialog panel" role="dialog" aria-modal="true" aria-labelledby="costMetaTitle"><div class="cost-meta-head cost-meta-head-enhanced"><div><span class="eyebrow">${hasExistingCost ? "Update cost setup" : "New cost setup"}</span><h3 id="costMetaTitle">${hasExistingCost ? "Edit Cost Details" : "Add Cost Details"}</h3><p>Assign a cost code and planned budget before recording daily field costs.</p></div><button type="button" class="cost-meta-close" id="closeCostMetaModalBtn" aria-label="Close">×</button></div><div class="cost-meta-activity-card"><span>Activity</span><strong>${escapeHtml(target.name || "Untitled Activity")}</strong><p>${Number(target.durationDays) || 0} day duration · Current planned cost ${plannedPreview}</p></div><form id="costMetaForm" class="cost-meta-form"><label class="cost-meta-field" for="costMetaIdInput"><span>Cost ID <strong aria-hidden="true">*</strong></span><input id="costMetaIdInput" name="costId" type="text" placeholder="e.g., C001" value="${escapeHtml(target.costId || "")}" autocomplete="off" required><small>Use a short unique code that matches your cost sheet.</small></label><label class="cost-meta-field" for="costMetaPlannedInput"><span>Planned Cost (₱) <strong aria-hidden="true">*</strong></span><input id="costMetaPlannedInput" name="plannedCost" type="number" min="0" step="0.01" placeholder="Enter planned cost" value="${Number(target.plannedCost) || 0}" required><small>This amount is used to calculate planned cost per day and earned value.</small></label><div class="cost-meta-save-note"><strong>What happens next?</strong><span>The record is saved to Google Sheets, then the costing table refreshes with the updated budget.</span></div><div class="cost-meta-actions"><button type="button" class="ghost-btn" id="cancelCostMetaModalBtn">Cancel</button><button type="submit" class="primary-btn">Save Cost Details</button></div></form></div>`;

  const closeModal = () => modal.classList.add("hidden");
  modal.querySelector("#closeCostMetaModalBtn")?.addEventListener("click", closeModal);
  modal.querySelector("#cancelCostMetaModalBtn")?.addEventListener("click", closeModal);
  modal.addEventListener("click", (event) => {
    if (event.target === modal) closeModal();
  });

  modal.querySelector("#costMetaForm")?.addEventListener("submit", async (event) => {
    event.preventDefault();
    const form = event.currentTarget;
    if (form.classList.contains("is-saving")) return;
    const formData = new FormData(form);
    const resolvedProjectId = String(projectId || target.projectId || "").trim();
    if (!resolvedProjectId) {
      alert("Unable to save cost because Project ID is missing.");
      return;
    }
    const nextCostId = String(formData.get("costId") || "").trim();
    const nextPlannedCost = parseBudgetValue(formData.get("plannedCost"));
    if (!nextCostId || nextPlannedCost <= 0) {
      alert("Please enter a Cost ID and a planned cost greater than 0.");
      return;
    }
    const existingOverrides = loadCostActivities().map(normalizeCostActivity);
    const nextOverrides = existingOverrides.filter((item) => !(String(item.projectId || "").trim() === String(projectId).trim() && getActivityRefId(item) === activityRefId));
    nextOverrides.push(normalizeCostActivity({
      ...target,
      costId: nextCostId,
      plannedCost: nextPlannedCost,
      activityRefId,
    }));
    saveCostActivityOverrides(nextOverrides);
    setFormSavingState(form, true, hasExistingCost ? "Updating…" : "Saving…");
    try {
      const durationDays = Number(target.durationDays) || 0;
      const plannedCostPerDay = durationDays > 0 ? nextPlannedCost / durationDays : 0;
      const payload = {
        cost: {
          costId: nextCostId,
          projectId: resolvedProjectId,
          project: target.projectName || "",
          activityId: activityRefId,
          activity: target.name || "",
          duration: durationDays,
          category: "Planned Cost",
          date: target.startDate || "",
          plannedCost: nextPlannedCost,
          plannedCostPerDay,
          actualCost: 0,
          notes: `Activity ID: ${activityRefId}`,
        },
      };
      try {
        await postToDataSource("costs", "create", { ...payload, skipDailyCostSync: true });
      } catch (error) {
        const shouldUpdateExisting = /already exists|use update instead of create/i.test(String(error?.message || ""));
        if (!shouldUpdateExisting) throw error;
        await postToDataSource("costs", "update", { ...payload, skipDailyCostSync: true });
      }
    } catch (error) {
      console.warn("Unable to save cost record to Google Sheets:", error);
      saveCostActivityOverrides(existingOverrides);
      setFormSavingState(form, false);
      alert(`Unable to save cost record to Google Sheets. ${error?.message || "Please verify your Apps Script deployment settings and try again."}`);
      return;
    }
    if (typeof window.notify === "function") {
      window.notify(hasExistingCost ? "Cost details updated successfully." : "Cost details saved successfully.", "success");
    }
    closeModal();
    showProjectDetails(projectId, "costing", loadCostActivities());
  });
};

const editCostMetadata = (projectId, activityRefId, allActivities = loadCostActivities()) => {
  const target = findResolvedCostActivity(projectId, activityRefId, allActivities);
  if (!target) return;
  renderCostMetadataModal(projectId, activityRefId, target);
};

const uniqueSorted = (values) =>
  Array.from(new Set(values.map((value) => String(value || "").trim()).filter(Boolean)))
    .sort((a, b) => a.localeCompare(b));

const populateSelect = (selectEl, values, defaultLabel) => {
  if (!selectEl) return;
  const previousValue = selectEl.value || defaultLabel;
  selectEl.innerHTML = [defaultLabel, ...values]
    .map((value) => `<option>${escapeHtml(value)}</option>`)
    .join("");
  selectEl.value = Array.from(selectEl.options).some((option) => option.value === previousValue)
    ? previousValue
    : defaultLabel;
};

const refreshProjectFilterOptions = (projects) => {
  populateSelect(projectTypeFilter, uniqueSorted(projects.map((project) => project.type)), "All Project Types");
  populateSelect(projectStatusFilter, uniqueSorted(projects.map((project) => project.status)), "All Statuses");
};

const parseDateFilterValue = (value) => {
  if (!value) return null;
  const parsed = new Date(value);
  if (Number.isNaN(parsed.getTime())) return null;
  parsed.setHours(0, 0, 0, 0);
  return parsed;
};

const formatCostDateFilterValue = (value) => {
  const parsed = parseDateFilterValue(value);
  if (!parsed) return "";
  return parsed.toLocaleDateString("en-US", { month: "short", day: "numeric", year: "numeric" });
};

const formatCostDateFilterLabel = () => {
  const start = String(dateStartInput?.value || "").trim();
  const end = String(dateEndInput?.value || "").trim();
  if (!start && !end) return "All Dates";
  if (start && end) return `${formatCostDateFilterValue(start)} - ${formatCostDateFilterValue(end)}`;
  if (start) return `${formatCostDateFilterValue(start)} onwards`;
  return `Until ${formatCostDateFilterValue(end)}`;
};

const syncCostDateFilterLabel = () => {
  if (costDateFilterLabel) costDateFilterLabel.textContent = formatCostDateFilterLabel();
};

const closeCostDateRangePanel = () => {
  if (!costDateRangePanel) return;
  costDateRangePanel.classList.add("hidden");
  costDateFilterWrap?.classList.remove("is-open");
  costDateFilterBtn?.setAttribute("aria-expanded", "false");
};

const openCostDateRangePanel = () => {
  if (!costDateRangePanel) return;
  costDateRangePanel.classList.remove("hidden");
  costDateFilterWrap?.classList.add("is-open");
  costDateFilterBtn?.setAttribute("aria-expanded", "true");
  window.setTimeout(() => dateStartInput?.focus(), 0);
};

const projectMatchesDateRange = (project, startDateFilter, endDateFilter) => {
  if (!startDateFilter && !endDateFilter) return true;
  const projectStart = project.startDate ? new Date(project.startDate) : null;
  const projectEnd = project.finishDate ? new Date(project.finishDate) : null;
  const filterStart = startDateFilter ? new Date(startDateFilter) : null;
  const filterEnd = endDateFilter ? new Date(endDateFilter) : null;
  if (filterStart && filterEnd && filterStart > filterEnd) return false;
  if (!projectStart && !projectEnd) return false;
  const effectiveProjectStart = projectStart && !Number.isNaN(projectStart.getTime()) ? projectStart : projectEnd;
  const effectiveProjectEnd = projectEnd && !Number.isNaN(projectEnd.getTime()) ? projectEnd : projectStart;
  if (!effectiveProjectStart || !effectiveProjectEnd) return false;
  if (filterEnd && effectiveProjectStart > filterEnd) return false;
  if (filterStart && effectiveProjectEnd < filterStart) return false;
  return true;
};

const renderProjects = (query = "") => {
  costPageHero?.classList.remove("hidden");
  const normalizedQuery = query.trim().toLowerCase();
  const selectedType = projectTypeFilter?.value || "All Project Types";
  const selectedStatus = projectStatusFilter?.value || "All Statuses";
  const startDateFilter = String(dateStartInput?.value || "").trim();
  const endDateFilter = String(dateEndInput?.value || "").trim();
  const allProjects = loadProjects().map(normalizeProject).filter((project) => !isArchivedProject(project));
  refreshProjectFilterOptions(allProjects);
  if (projectTypeFilter) projectTypeFilter.value = Array.from(projectTypeFilter.options).some((option) => option.value === selectedType) ? selectedType : "All Project Types";
  if (projectStatusFilter) projectStatusFilter.value = Array.from(projectStatusFilter.options).some((option) => option.value === selectedStatus) ? selectedStatus : "All Statuses";
  const activeType = projectTypeFilter?.value || "All Project Types";
  const activeStatus = projectStatusFilter?.value || "All Statuses";
  const projects = allProjects
    .filter((project) => !normalizedQuery || [project.name, project.code, project.status, project.type].some((value) => String(value || "").toLowerCase().includes(normalizedQuery)))
    .filter((project) => activeType === "All Project Types" || project.type === activeType)
    .filter((project) => activeStatus === "All Statuses" || project.status === activeStatus)
    .filter((project) => projectMatchesDateRange(project, startDateFilter, endDateFilter));
  projectsList.innerHTML = "";
  if (visibleProjectCount) visibleProjectCount.textContent = String(projects.length);
  if (visibleProjectBudget) visibleProjectBudget.textContent = formatBudget(projects.reduce((sum, project) => sum + parseBudgetValue(project.budget), 0));
  if (!projects.length) return projectsEmpty.classList.remove("hidden");
  projectsEmpty.classList.add("hidden");
  projects.forEach((project) => {
    const row = document.createElement("button");
    row.type = "button";
    row.className = "project-row";
    const statusTone = getStatusTone(project.status);
    row.innerHTML = `<div class="project-row-main"><span class="project-status-badge ${statusTone}">${escapeHtml(project.status || "Not Started")}</span><div class="project-meta"><strong>${escapeHtml(formatProjectIdentityLabel(project))}</strong><p>${escapeHtml(formatProjectTimeline(project))}</p></div></div><div class="project-cost-preview"><span>Approved budget</span><strong>${formatBudget(project.budget)}</strong></div>`;
    row.addEventListener("click", () => {
      const nextUrl = new URL(window.location.href);
      nextUrl.searchParams.set("projectId", project.id);
      window.location.href = nextUrl.toString();
    });
    projectsList.append(row);
  });
};

const syncSearches = (value) => {
  if (topSearch) topSearch.value = value;
  renderProjects(value);
};
topSearch?.addEventListener("input", (event) => syncSearches(event.target.value));
[projectTypeFilter, projectStatusFilter]
  .filter(Boolean)
  .forEach((el) => el.addEventListener("change", () => renderProjects(topSearch?.value || "")));

costDateFilterBtn?.addEventListener("click", () => {
  const isOpen = costDateRangePanel && !costDateRangePanel.classList.contains("hidden");
  if (isOpen) {
    closeCostDateRangePanel();
    return;
  }
  openCostDateRangePanel();
});

dateStartInput?.addEventListener("change", () => {
  if (dateEndInput) dateEndInput.min = dateStartInput.value || "";
});

costDateApplyBtn?.addEventListener("click", () => {
  const start = parseDateFilterValue(dateStartInput?.value);
  const end = parseDateFilterValue(dateEndInput?.value);
  if (start && end && start > end) {
    window.alert("End date must be on or after start date.");
    return;
  }
  syncCostDateFilterLabel();
  closeCostDateRangePanel();
  renderProjects(topSearch?.value || "");
});

costDateClearBtn?.addEventListener("click", () => {
  if (dateStartInput) dateStartInput.value = "";
  if (dateEndInput) {
    dateEndInput.value = "";
    dateEndInput.min = "";
  }
  syncCostDateFilterLabel();
  closeCostDateRangePanel();
  renderProjects(topSearch?.value || "");
});

clearCostFiltersBtn?.addEventListener("click", () => {
  if (topSearch) topSearch.value = "";
  if (projectTypeFilter) projectTypeFilter.value = "All Project Types";
  if (projectStatusFilter) projectStatusFilter.value = "All Statuses";
  if (dateStartInput) dateStartInput.value = "";
  if (dateEndInput) {
    dateEndInput.value = "";
    dateEndInput.min = "";
  }
  syncCostDateFilterLabel();
  closeCostDateRangePanel();
  renderProjects("");
  topSearch?.focus();
});

document.addEventListener("click", (event) => {
  if (
    costDateRangePanel &&
    !costDateRangePanel.classList.contains("hidden") &&
    costDateFilterWrap &&
    !costDateFilterWrap.contains(event.target)
  ) {
    closeCostDateRangePanel();
  }
});

document.addEventListener("keydown", (event) => {
  if (event.key === "Escape") closeCostDateRangePanel();
});

syncCostDateFilterLabel();

const getSelectedViewParams = () => {
  const params = new URLSearchParams(window.location.search);
  return {
    selectedProjectId: params.get("projectId") || "",
    selectedProjectName: params.get("project") || "",
    selectedTab: params.get("tab") === "costing" ? "costing" : "overview",
  };
};

const getSelectedProjectFilter = (project = null) => {
  const { selectedProjectId, selectedProjectName } = getSelectedViewParams();
  return {
    projectId: project?.id || selectedProjectId,
    projectName: project?.name || selectedProjectName,
  };
};
const getActiveDetailsTabFromUi = () => {
  const activeTab = detailsView?.querySelector(".tab-btn.active")?.dataset.tab;
  return activeTab === "costing" ? "costing" : "overview";
};

const removeInferredProjectBudgetCostMetadata = (activity = {}) => {
  const normalized = normalizeCostActivity(activity);
  const plannedCost = parseBudgetValue(normalized.plannedCost);
  if (plannedCost <= 0) return normalized;

  const activityProjectId = normalizeLookup(resolveActivityProjectId(normalized));
  const activityProjectName = normalizeLookup(normalized.projectName);
  const project = loadProjects().map(normalizeProject).find((item) => {
    const projectId = normalizeLookup(item.id);
    const projectName = normalizeLookup(item.name);
    return (activityProjectId && projectId === activityProjectId)
      || (activityProjectName && projectName === activityProjectName);
  });
  const projectBudget = parseBudgetValue(project?.budget);
  if (projectBudget <= 0 || plannedCost !== projectBudget) return normalized;

  return {
    ...normalized,
    costId: "",
    plannedCost: 0,
    earnedValue: 0,
  };
};

const applyCostMetadataRows = (rows = [], projectFilter = {}) => {
  if (!Array.isArray(rows)) return;
  const selectedProjectId = String(projectFilter?.projectId || "").trim();
  const selectedProjectName = String(projectFilter?.projectName || "").trim();
  const shouldResetProjectMetadata = Boolean(selectedProjectId || selectedProjectName);

  // Treat a successful Costs resource response as authoritative for the selected
  // project. If the sheet has no rows, clear cached Cost ID / planned-cost values
  // that may be left in localStorage from earlier sessions so blank Sheets stay
  // blank in the costing table until the user explicitly adds cost details.
  if (shouldResetProjectMetadata) {
    saveCostActivityOverrides(costActivitiesState.map((activity) => {
      if (!isActivityForProject(activity, selectedProjectId, selectedProjectName)) return activity;
      return normalizeCostActivity({
        ...activity,
        costId: "",
        plannedCost: 0,
        earnedValue: 0,
      });
    }));
  }

  if (!rows.length) return;
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

  rows.forEach((row) => {
    const projectKey = normalizeLookup(resolveActivityProjectId({
      projectId: row.projectId,
      projectName: row.projectName || selectedProjectName,
    }));
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

  saveCostActivityOverrides(costActivitiesState.map((activity) => {
    const projectKey = normalizeLookup(resolveActivityProjectId(activity));
    const activityKey = `${projectKey}::${String(getActivityRefId(activity) || "").trim()}`;
    const activityNameKey = `${projectKey}::${normalizeLookup(activity.name)}`;
    const fallbackActivityRefKey = String(getActivityRefId(activity) || "").trim();
    const fallbackActivityNameKey = normalizeLookup(activity.name);
    const scopedMetadata = metadataByActivityId.get(activityKey)
      || metadataByActivityName.get(activityNameKey);
    const fallbackMetadata = shouldResetProjectMetadata
      ? null
      : (metadataByActivityIdFallback.get(fallbackActivityRefKey)
        || metadataByActivityNameFallback.get(fallbackActivityNameKey));
    const metadata = scopedMetadata || fallbackMetadata;
    if (!metadata) return activity;
    return normalizeCostActivity({
      ...activity,
      costId: String(metadata.costId || "").trim(),
      plannedCost: parseBudgetValue(metadata.plannedCost),
      progressPercent: clampPercent(metadata.progressPercent),
      earnedValue: parseBudgetValue(metadata.earnedValue),
    });
  }));
};

const bootstrapCostManagement = async ({ preferredTab = null } = {}) => {
  const remoteProjectsResult = await loadRemoteProjects();
  const hasAuthoritativeProjectSheet = Array.isArray(remoteProjectsResult);
  const remoteProjects = hasAuthoritativeProjectSheet
    ? remoteProjectsResult.map(normalizeProject).filter((project) => project.id)
    : [];
  projectsState = hasAuthoritativeProjectSheet
    ? remoteProjects
    : loadFromLocalStorageArray(PROJECTS_LOCAL_STORAGE_KEY).map(normalizeProject).filter((project) => project.id);
  // Local storage is now a warm-start cache only. An authoritative empty remote
  // project list clears stale browser data instead of resurrecting deleted rows.
  persistToLocalStorage(PROJECTS_LOCAL_STORAGE_KEY, projectsState);
  if (!hasAuthoritativeProjectSheet && !hasWarnedAboutCachedProjects && typeof window.notify === "function") {
    hasWarnedAboutCachedProjects = true;
    window.notify("Using cached projects because Google Sheets could not be reached.", "warning");
  } else if (hasAuthoritativeProjectSheet) {
    hasWarnedAboutCachedProjects = false;
  }

  const { selectedProjectId, selectedProjectName, selectedTab } = getSelectedViewParams();
  const resolvedSelectedTab = preferredTab === "costing" ? "costing" : selectedTab;
  const selectedProjectAfterBootstrap = loadProjects().map(normalizeProject).filter((project) => !isArchivedProject(project)).find((project) =>
    (selectedProjectId && project.id === selectedProjectId)
    || (selectedProjectName && project.name === selectedProjectName)
  );
  const projectFilter = getSelectedProjectFilter(selectedProjectAfterBootstrap);
  const [remoteActivityResult, remoteDailyCosts, remoteCostMetadataRows] = await Promise.all([
    loadRemoteCostActivities(projectFilter),
    loadRemoteDailyCosts(projectFilter),
    loadRemoteCostMetadata(projectFilter),
  ]);
  const remoteActivities = Array.isArray(remoteActivityResult?.rows) ? remoteActivityResult.rows : [];
  const hasAuthoritativeActivitySheet = Boolean(remoteActivityResult?.authoritative);
  const selectedProjectHasFilter = Boolean(projectFilter?.projectId || projectFilter?.projectName);
  const remoteActivityKeys = new Set(remoteActivities.map((item) => getCostActivityKey(item)));
  // Merge cached local overrides first so user-maintained Cost ID / planned-cost
  // metadata survives for matching sheet activities, then let fresh remote rows
  // refresh schedule fields. When the sheet was read successfully for the
  // selected project, do not keep local-only activities; otherwise deleted or
  // unrelated browser-cache rows can appear as costing records that are no
  // longer present in Google Sheets.
  const localActivities = loadFromLocalStorageArray(COST_ACTIVITIES_LOCAL_STORAGE_KEY)
    .concat(loadFromLocalStorageArray(LEGACY_COST_ACTIVITIES_LOCAL_STORAGE_KEY))
    .map(removeInferredProjectBudgetCostMetadata)
    .filter((item) => getCostActivityProjectKey(item) && getActivityRefId(item))
    .filter((item) => {
      if (!hasAuthoritativeActivitySheet || !selectedProjectHasFilter) return true;
      if (!isActivityForProject(item, projectFilter.projectId, projectFilter.projectName)) return true;
      return remoteActivityKeys.has(getCostActivityKey(item));
    });
  const merged = [...localActivities, ...remoteActivities];
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
  saveCostActivityOverrides(allActivities);

  await syncDailyCostsFromSheet(projectFilter, remoteDailyCosts);
  cleanupOrphanedDailyCosts(allActivities);

  if (Array.isArray(remoteCostMetadataRows)) applyCostMetadataRows(remoteCostMetadataRows, projectFilter);

  const activitiesForDisplay = loadCostActivities();

  if (!selectedProjectAfterBootstrap || !showProjectDetails(selectedProjectAfterBootstrap.id, resolvedSelectedTab, activitiesForDisplay)) renderProjects();
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
  if (document.activeElement instanceof HTMLElement) {
    const isTyping =
      document.activeElement.matches("input, textarea, select")
      || document.activeElement.isContentEditable;
    if (isTyping) return;
  }

  isCostManagementSyncInFlight = true;
  try {
    await bootstrapCostManagement({ preferredTab: getActiveDetailsTabFromUi() });
  } catch (error) {
    console.warn("Unable to refresh cost management data:", error);
    if (typeof window.notify === "function") {
      window.notify(error?.message || "Unable to refresh cost data. Showing the latest cached values.", "warning");
    }
    const { selectedProjectId, selectedProjectName, selectedTab } = getSelectedViewParams();
    const fallbackProject = loadProjects().map(normalizeProject).find((project) =>
      (selectedProjectId && project.id === selectedProjectId)
      || (selectedProjectName && project.name === selectedProjectName)
    );
    if (fallbackProject) {
      showProjectDetails(fallbackProject.id, getActiveDetailsTabFromUi() || selectedTab, loadCostActivities());
    } else {
      renderProjects(topSearch?.value || "");
    }
  } finally {
    isCostManagementSyncInFlight = false;
    window.dispatchEvent(new CustomEvent("cost-management:data-loaded"));
  }
};

window.addEventListener("focus", () => refreshSelectedProjectCostView({ force: true }));
window.addEventListener("pageshow", () => refreshSelectedProjectCostView({ force: true }));
window.addEventListener("google-sheet:changed", () => refreshSelectedProjectCostView({ force: true }));
document.addEventListener("visibilitychange", () => {
  if (!document.hidden) refreshSelectedProjectCostView({ force: true });
});

refreshSelectedProjectCostView({ force: true });
