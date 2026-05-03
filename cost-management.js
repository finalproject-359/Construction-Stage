const LOCAL_STORAGE_KEY = "constructionStageProjects";
const COST_ACTIVITY_KEY = "constructionStageCostActivities";
const ACTIVITIES_LOCAL_STORAGE_KEY = "constructionStageActivities";
const COST_DAILY_KEY = "constructionStageDailyCosts";

const topSearch = document.getElementById("costTopSearch");
const listSearch = document.getElementById("projectListSearch");
const projectsList = document.getElementById("costProjectsList");
const projectsEmpty = document.getElementById("costProjectsEmpty");
const selectionView = document.getElementById("costSelectionView");
const detailsView = document.getElementById("costDetailsView");
const selectedProjectBannerHost = document.getElementById("selectedProjectBannerHost");

const safeJsonParse = (raw, fallback = []) => {
  try {
    const parsed = JSON.parse(raw || "[]");
    return Array.isArray(parsed) ? parsed : fallback;
  } catch {
    return fallback;
  }
};

const loadProjects = () => safeJsonParse(localStorage.getItem(LOCAL_STORAGE_KEY), []);
const loadDailyCosts = () => safeJsonParse(localStorage.getItem(COST_DAILY_KEY), []);
const saveDailyCosts = (items) => localStorage.setItem(COST_DAILY_KEY, JSON.stringify(items));

const getValueByAliases = (source, aliases = []) => {
  if (!source || typeof source !== "object") return undefined;
  for (const alias of aliases) if (Object.prototype.hasOwnProperty.call(source, alias)) return source[alias];
  const normalizedEntries = Object.keys(source).map((key) => ({ key, normalized: String(key).toLowerCase().replace(/[^a-z0-9]/g, "") }));
  for (const alias of aliases) {
    const normalizedAlias = String(alias).toLowerCase().replace(/[^a-z0-9]/g, "");
    const matched = normalizedEntries.find((entry) => entry.normalized === normalizedAlias);
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
  const start = new Date(startDate);
  const finish = new Date(finishDate);
  if (!Number.isNaN(start.getTime()) && !Number.isNaN(finish.getTime()) && finish >= start) {
    return Math.max(1, Math.round((finish.getTime() - start.getTime()) / 86400000) + 1);
  }
  return Number(fallback) || 0;
};
const normalizeCostActivity = (activity = {}) => {
  const startDate = toDateInputValue(getValueByAliases(activity, ["startDate", "plannedStart", "planned_start"]));
  const finishDate = toDateInputValue(getValueByAliases(activity, ["finishDate", "plannedFinish", "planned_finish"]));
  const explicitDuration = Number(String(getValueByAliases(activity, ["durationDays", "duration_days", "duration"]) || "0").replace(/[^\d.-]/g, "")) || 0;

  return {
    id: String(getValueByAliases(activity, ["activityId", "activity_id", "activity id", "sourceActivityId", "source_activity_id", "source activity id", "code", "id"]) || "").trim(),
    costId: String(getValueByAliases(activity, ["costId", "cost_id", "costCode", "cost_code"]) || "").trim(),
    activityRefId: String(getValueByAliases(activity, ["activityRefId", "activity_ref_id", "activity ref id", "sourceActivityId", "source_activity_id", "source activity id", "activityId", "activity_id", "activity id", "id", "code"]) || "").trim(),
    projectId: String(getValueByAliases(activity, ["projectId", "project_id", "project id", "project", "projectName", "project_name", "project name"]) || "").trim(),
    projectName: String(getValueByAliases(activity, ["project", "projectName", "project_name", "project name"]) || "").trim(),
    name: String(getValueByAliases(activity, ["name", "activity", "activityName", "activity_name"]) || "Untitled Activity").trim(),
    startDate,
    finishDate,
    durationDays: computeDurationDays(startDate, finishDate, explicitDuration),
    plannedCost: parseBudgetValue(getValueByAliases(activity, ["plannedCost", "planned_cost", "plannedValue", "planned_value", "budget"])),
  };
};

const getActivityRefId = (activity = {}) => String(activity.activityRefId || activity.id || "").trim();
const getCostActivityProjectKey = (activity = {}) => String(activity.projectId || activity.projectName || "").trim();
const getCostActivityKey = (activity = {}) => `${getCostActivityProjectKey(activity)}::${getActivityRefId(activity)}`;

const loadCostActivities = () => {
  const projectLookups = buildProjectIdentityLookups(loadProjects());
  const activitiesSource = safeJsonParse(localStorage.getItem(ACTIVITIES_LOCAL_STORAGE_KEY), [])
    .map(normalizeCostActivity)
    .map((activity) => ({ ...activity, projectId: resolveActivityProjectId(activity, projectLookups) || activity.projectId }));
  const costSource = safeJsonParse(localStorage.getItem(COST_ACTIVITY_KEY), [])
    .map(normalizeCostActivity)
    .map((activity) => ({ ...activity, projectId: resolveActivityProjectId(activity, projectLookups) || activity.projectId }));

  // Prefer the Activities page source because it is the actively maintained dataset.
  // Keep legacy cost entries only as metadata overrides when there is an existing activity match.
  // If there are no activities yet, use legacy data as a fallback for backward compatibility.
  const activitiesByKey = new Map();
  activitiesSource.forEach((item) => {
    const key = getCostActivityKey(item);
    if (!key || key === "::") return;
    activitiesByKey.set(key, { ...item, costId: "", plannedCost: 0 });
  });

  if (!activitiesByKey.size) {
    return costSource;
  }

  costSource.forEach((item) => {
    const key = getCostActivityKey(item);
    if (!activitiesByKey.has(key)) return;

    const baseActivity = activitiesByKey.get(key) || {};
    activitiesByKey.set(key, {
      ...baseActivity,
      // Keep canonical activity identity/schedule from Activities data,
      // and only apply cost-specific metadata overrides from legacy cost entries.
      // Cost-specific overrides are user-maintained values.
      costId: String(item.costId || item.id || baseActivity.costId || "").trim(),
      plannedCost: Number(item.plannedCost) || 0,
    });
  });

  return Array.from(activitiesByKey.values());
};

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
    startDate,
    finishDate,
    durationDays: computeDurationDays(startDate, finishDate, explicitDuration),
    plannedCost: 0,
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
  const [idPart, ...nameParts] = raw.split("-");
  if (!nameParts.length) return { id: "", name: raw };
  return {
    id: normalizeLookup(idPart),
    name: normalizeLookup(nameParts.join("-").trim()),
  };
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
  const activities = allActivities.filter((item) => isActivityForProject(item, projectId, projectName));
  const daily = loadDailyCosts().filter((item) => String(item.projectId || "").trim() === projectId);
  const rows = activities.map((activity) => {
    const refId = getActivityRefId(activity);
    const dailyItems = daily.filter((entry) => entry.activityId === refId);
    const actualCost = dailyItems.reduce((sum, entry) => sum + parseBudgetValue(entry.actualCost), 0);
    return { ...activity, actualCost, dailyItems };
  });
  return { rows, activities, daily };
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
  const availableDates = buildDateRangeOptions(activity.startDate, activity.finishDate);
  const hasAvailableDates = availableDates.length > 0;
  const dateOptions = hasAvailableDates
    ? availableDates.map((dateValue) => `<option value="${dateValue}">${formatLongHumanDate(dateValue)}</option>`).join("")
    : `<option value="" selected disabled>No working days in this date range.</option>`;

  const entries = dailyCosts
    .filter((item) => String(item.projectId || "").trim() === projectId && String(item.activityId || "").trim() === activityId)
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
    const nextDailyCosts = loadDailyCosts().filter((item) => !(String(item.projectId || "").trim() === projectId
      && String(item.activityId || "").trim() === activityId
      && String(item.date || "") === date));
    saveDailyCosts(nextDailyCosts);
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

  modal.querySelector("#dailyCostForm")?.addEventListener("submit", (event) => {
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
      String(item.projectId || "").trim() === projectId
      && String(item.activityId || "").trim() === activityId
      && String(item.date || "") === date
    );
    const payload = { projectId, activityId, date, actualCost };
    if (existingIndex >= 0) dailyCosts[existingIndex] = payload;
    else dailyCosts.push(payload);
    saveDailyCosts(dailyCosts);
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

const saveCostActivityOverrides = (items = []) => localStorage.setItem(COST_ACTIVITY_KEY, JSON.stringify(items));
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

  modal.querySelector("#costMetaForm")?.addEventListener("submit", (event) => {
    event.preventDefault();
    const formData = new FormData(event.currentTarget);
    const nextCostId = String(formData.get("costId") || "").trim();
    const nextPlannedCost = parseBudgetValue(formData.get("plannedCost"));
    const existingOverrides = safeJsonParse(localStorage.getItem(COST_ACTIVITY_KEY), []).map(normalizeCostActivity);
    const nextOverrides = existingOverrides.filter((item) => !(String(item.projectId || "").trim() === String(projectId).trim() && getActivityRefId(item) === activityRefId));
    nextOverrides.push(normalizeCostActivity({
      ...target,
      costId: nextCostId,
      plannedCost: nextPlannedCost,
      activityRefId,
    }));
    saveCostActivityOverrides(nextOverrides);
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
const bootstrapCostManagement = async () => {
  const localProjects = loadProjects().map(normalizeProject).filter((project) => project.id);
  if (!localProjects.length) {
    const remoteProjects = await loadRemoteProjects();
    if (remoteProjects.length) {
      localStorage.setItem(LOCAL_STORAGE_KEY, JSON.stringify(remoteProjects));
    }
  }

  const selectedProjectAfterBootstrap = loadProjects().map(normalizeProject).find((project) =>
    (selectedProjectId && project.id === selectedProjectId)
    || (selectedProjectName && project.name === selectedProjectName)
  );
  const localActivities = loadCostActivities();
  const remoteActivities = await loadRemoteCostActivities({
    projectId: selectedProjectAfterBootstrap?.id || selectedProjectId,
    projectName: selectedProjectAfterBootstrap?.name || selectedProjectName,
  });
  // Prefer local activities first so unsynced edits remain visible on this device.
  // Remote rows are still used as a fallback when local rows are missing.
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

  if (!selectedProjectAfterBootstrap || !showProjectDetails(selectedProjectAfterBootstrap.id, selectedTab, allActivities)) renderProjects();
};

const refreshSelectedProjectCostView = () => {
  const selectedProject = loadProjects().map(normalizeProject).find((project) =>
    (selectedProjectId && project.id === selectedProjectId)
    || (selectedProjectName && project.name === selectedProjectName)
  );
  if (!selectedProject) return;
  showProjectDetails(selectedProject.id, selectedTab, loadCostActivities());
};

window.addEventListener("storage", (event) => {
  if (![ACTIVITIES_LOCAL_STORAGE_KEY, COST_ACTIVITY_KEY, COST_DAILY_KEY, LOCAL_STORAGE_KEY].includes(event.key)) return;
  refreshSelectedProjectCostView();
});
window.addEventListener("focus", refreshSelectedProjectCostView);
window.addEventListener("pageshow", refreshSelectedProjectCostView);
document.addEventListener("visibilitychange", () => {
  if (!document.hidden) refreshSelectedProjectCostView();
});

bootstrapCostManagement();
