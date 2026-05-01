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
const normalizeCostActivity = (activity = {}) => ({
  id: String(getValueByAliases(activity, ["id", "activityId", "activity_id", "code"]) || "").trim(),
  projectId: String(getValueByAliases(activity, ["projectId", "project_id"]) || "").trim(),
  projectName: String(getValueByAliases(activity, ["project", "projectName", "project_name"]) || "").trim(),
  name: String(getValueByAliases(activity, ["name", "activity", "activityName", "activity_name"]) || "Untitled Activity").trim(),
  startDate: toDateInputValue(getValueByAliases(activity, ["startDate", "plannedStart", "planned_start"])),
  finishDate: toDateInputValue(getValueByAliases(activity, ["finishDate", "plannedFinish", "planned_finish"])),
  durationDays: Number(String(getValueByAliases(activity, ["durationDays", "duration_days", "duration"]) || "0").replace(/[^\d.-]/g, "")) || 0,
  plannedCost: parseBudgetValue(getValueByAliases(activity, ["plannedCost", "planned_cost", "plannedValue", "planned_value", "budget"])),
});

const loadCostActivities = () => {
  const activitiesSource = safeJsonParse(localStorage.getItem(ACTIVITIES_LOCAL_STORAGE_KEY), []).map(normalizeCostActivity);
  const costSource = safeJsonParse(localStorage.getItem(COST_ACTIVITY_KEY), []).map(normalizeCostActivity);

  // Prefer the Activities page source because it is the actively maintained dataset.
  // Keep a fallback to legacy cost-specific storage to avoid losing historical entries.
  const merged = [...activitiesSource, ...costSource].filter((item) => item.projectId);
  const dedupedByProjectAndActivity = new Map();
  merged.forEach((item) => {
    const key = `${String(item.projectId).trim()}::${String(item.id).trim()}`;
    if (!key || key === '::') return;
    if (!dedupedByProjectAndActivity.has(key)) dedupedByProjectAndActivity.set(key, item);
  });

  return Array.from(dedupedByProjectAndActivity.values());
};

const normalizeRemoteActivity = (row = {}) => {
  const projectId = String(getValueByAliases(row, ["projectId", "project_id", "project id"]) || "").trim();
  const startDate = toDateInputValue(getValueByAliases(row, ["startDate", "plannedStart", "planned_start", "planned start"]));
  const finishDate = toDateInputValue(getValueByAliases(row, ["finishDate", "plannedFinish", "planned_finish", "planned finish"]));
  const explicitDuration = Number(String(getValueByAliases(row, ["durationDays", "duration_days", "duration", "duration day"]) || "0").replace(/[^\d.-]/g, "")) || 0;
  const computedDuration = startDate && finishDate
    ? Math.max(1, Math.round((new Date(finishDate).getTime() - new Date(startDate).getTime()) / 86400000) + 1)
    : 0;

  return normalizeCostActivity({
    id: getValueByAliases(row, ["id", "activityId", "activity_id", "activity id", "code"]),
    projectId,
    projectName: getValueByAliases(row, ["project", "projectName", "project_name", "project name"]),
    name: getValueByAliases(row, ["name", "activity", "activityName", "activity_name"]),
    startDate,
    finishDate,
    durationDays: explicitDuration || computedDuration,
    plannedCost: getValueByAliases(row, ["plannedCost", "planned_cost", "plannedValue", "planned value", "budget"]),
  });
};

const loadRemoteCostActivities = async () => {
  const resourceRows = await loadActivitiesFromResourceEndpoint();
  if (resourceRows.length) return resourceRows;

  if (!window.DataBridge?.fetchRowsFromSource) return [];
  try {
    const { rows } = await window.DataBridge.fetchRowsFromSource();
    return (rows || []).map(normalizeRemoteActivity).filter((item) => item.projectId && item.id);
  } catch (error) {
    console.warn("Unable to load cost activities from Google Sheets:", error);
    return [];
  }
};


const loadActivitiesFromResourceEndpoint = async () => {
  const dataSourceUrl = window.DataBridge?.DEFAULT_DATA_SOURCE_URL;
  if (!dataSourceUrl) return [];
  try {
    const url = new URL(dataSourceUrl);
    url.searchParams.set("resource", "activities");
    url.searchParams.set("_ts", String(Date.now()));
    const response = await fetch(url.toString(), { cache: "no-store" });
    if (!response.ok) return [];
    const payload = await response.json();
    const rows = Array.isArray(payload?.activities) ? payload.activities : [];
    return rows.map(normalizeRemoteActivity).filter((item) => item.projectId && item.id);
  } catch (error) {
    console.warn("Unable to load cost activities from activities resource endpoint:", error);
    return [];
  }
};

const normalizeLookup = (value) => String(value || "").trim().toLowerCase();

const isActivityForProject = (activity, projectId, projectName = "") => {
  const activityProjectId = normalizeLookup(activity?.projectId);
  const activityProjectName = normalizeLookup(activity?.projectName);
  const projectIdLookup = normalizeLookup(projectId);
  const projectNameLookup = normalizeLookup(projectName);

  if (activityProjectId && activityProjectId === projectIdLookup) return true;
  if (projectNameLookup && activityProjectName && activityProjectName === projectNameLookup) return true;
  if (projectNameLookup && activityProjectId && activityProjectId === projectNameLookup) return true;
  if (projectIdLookup && activityProjectName && activityProjectName === projectIdLookup) return true;
  return false;
};

const getProjectCostData = (projectId, allActivities = loadCostActivities()) => {
  const project = loadProjects().map(normalizeProject).find((item) => item.id === projectId);
  const projectName = String(project?.name || "").trim().toLowerCase();
  const activities = allActivities.filter((item) => isActivityForProject(item, projectId, projectName));
  const daily = loadDailyCosts().filter((item) => String(item.projectId || "").trim() === projectId);
  const rows = activities.map((activity) => {
    const dailyItems = daily.filter((entry) => entry.activityId === activity.id);
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
    ? rows.map((row) => `<tr><td>${escapeHtml(row.id)}</td><td>${escapeHtml(row.name)}</td><td>${row.durationDays || "-"} days</td><td>${formatBudget(row.plannedCost)}</td><td>${formatBudget((row.plannedCost || 0) / (row.durationDays || 1))}</td><td>${formatBudget(row.actualCost)}</td><td><button type="button" class="ghost-btn view-daily-cost-btn" data-activity-id="${escapeHtml(row.id)}">View / Add Daily Cost</button></td></tr>`).join("")
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
    .map((row) => `<tr><td>${escapeHtml(row.id)}</td><td>${escapeHtml(row.name)}</td><td>${formatBudget(row.plannedCost)}</td><td>${formatBudget(row.actualCost)}</td><td class="bad">-${formatBudget(row.actualCost - row.plannedCost)}</td></tr>`)
    .join("") || '<tr><td colspan="5" class="empty-cell">No over budget activities.</td></tr>';

  return `<header class="details-header"><h2>Cost Management</h2><p>Project: <strong>${escapeHtml(formatProjectIdentityLabel(project))}</strong></p></header>
  <nav class="details-tabs"><button class="tab-btn active" data-tab="overview" type="button">Overview</button><button class="tab-btn" data-tab="costing" type="button">Costing Record</button></nav>
  <section class="details-tab-panel" data-panel="overview"><section class="details-kpis">
  <article class="kpi-card"><h4>Total Planned Cost</h4><p>${formatBudget(plannedCost)}</p></article>
  <article class="kpi-card"><h4>Total Actual Cost</h4><p>${formatBudget(actualCost)}</p></article>
  <article class="kpi-card"><h4>Variance</h4><p class="${varianceClass}">${formatBudget(variance)}</p><small>${varianceLabel}</small></article>
  <article class="kpi-card"><h4>Total Duration</h4><p>${totalDuration} days</p></article></section>
  <section class="overview-grid"><article class="panel chart-panel"><h3>Budget vs Actual</h3><div class="bars"><div class="bars-grid"><span>${formatBudget(maxCost)}</span><span>${formatBudget(maxCost * 0.75)}</span><span>${formatBudget(maxCost * 0.5)}</span><span>${formatBudget(maxCost * 0.25)}</span><span>0</span></div><div class="bars-track"><div class="bar-wrap"><strong>${formatBudget(plannedCost)}</strong><div class="bar planned" style="height:${plannedHeight}%"></div><p>Total Planned Cost</p></div><div class="bar-wrap"><strong>${formatBudget(actualCost)}</strong><div class="bar actual" style="height:${actualHeight}%"></div><p>Total Actual Cost</p></div></div><ul class="bars-legend"><li><span class="legend-dot planned"></span> Planned Cost</li><li><span class="legend-dot actual"></span> Actual Cost</li></ul></div></article>
  <article class="panel summary-panel"><h3>Cost Summary</h3><ul><li><span>Total Planned Cost</span><strong>${formatBudget(plannedCost)}</strong></li><li><span>Total Actual Cost</span><strong>${formatBudget(actualCost)}</strong></li><li><span>Variance</span><strong class="${varianceClass}">${formatBudget(variance)}</strong></li><li><span>Variance Percent</span><strong class="${varianceClass}">${variancePercent.toFixed(2)}%</strong></li><li><span>Total Duration</span><strong>${totalDuration} days</strong></li><li><span>Average Cost per Day (Actual)</span><strong>${formatBudget(avgActualPerDay)}</strong></li></ul></article>
  <article class="panel donut-panel"><h3>Cost Status (by Actual vs Planned)</h3><div class="donut-wrap"><div class="donut" style="background: conic-gradient(#34b567 0 ${underBudgetPct.toFixed(2)}%, #ef5050 ${underBudgetPct.toFixed(2)}% ${(underBudgetPct + overBudgetPct).toFixed(2)}%, #f0c23f ${(underBudgetPct + overBudgetPct).toFixed(2)}% 100%);"></div><ul class="status-list"><li><span class="legend-dot actual"></span>Under Budget <strong>${underBudgetCount} activities</strong></li><li><span class="legend-dot bad-dot"></span>Over Budget <strong>${overBudgetCount} activities</strong></li><li><span class="legend-dot neutral-dot"></span>No Actual Cost <strong>${noActualCostCount} activities</strong></li></ul></div></article>
  <article class="panel table-panel"><h3>Top Over Budget Activities</h3><table><thead><tr><th>Cost ID</th><th>Activity</th><th>Planned Cost</th><th>Actual Cost</th><th>Variance</th></tr></thead><tbody>${topRows}</tbody></table></article></section></section>
  <section class="details-tab-panel hidden" data-panel="costing"><section class="cost-record-head panel"><div><h3>Costing Record</h3><p>Manage planned and actual costs for all project activities.</p></div><div class="cost-record-actions"><p class="muted-note">Activities are managed in the Activities page.</p></div></section>
  <section class="panel"><table><thead><tr><th>Cost ID</th><th>Activity</th><th>Duration</th><th>Planned Cost</th><th>Planned Cost/Day</th><th>Actual Cost</th><th>Actions</th></tr></thead><tbody>${tableRows}</tbody></table></section><div class="info-banner"><p>Tip: add daily actual costs by date via “View / Add Daily Cost”.</p></div></section>
  <section class="daily-cost-modal hidden" id="dailyCostModal"></section>`;
};

const renderDailyCostModal = (projectId, activityId) => {
  const modal = detailsView.querySelector("#dailyCostModal");
  const project = loadProjects().map(normalizeProject).find((item) => item.id === projectId);
  const projectName = String(project?.name || "").trim().toLowerCase();
  const activities = loadCostActivities();
  const dailyCosts = loadDailyCosts();
  const activity = activities.find((item) => item.id === activityId && isActivityForProject(item, projectId, projectName));
  if (!modal || !activity) return;
  const entries = dailyCosts
    .filter((item) => String(item.projectId || "").trim() === projectId && String(item.activityId || "").trim() === activityId)
    .sort((a, b) => String(a.date || "").localeCompare(String(b.date || "")));
  const rows = entries.length ? entries.map((entry) => `<tr><td>${entry.date}</td><td>${formatBudget(entry.actualCost)}</td></tr>`).join("") : '<tr><td colspan="2" class="empty-cell">No daily costs yet.</td></tr>';
  modal.classList.remove("hidden");
  modal.innerHTML = `<div class="daily-cost-dialog panel"><h3>${escapeHtml(activity.name)} Daily Cost</h3><p>${escapeHtml(activity.startDate)} to ${escapeHtml(activity.finishDate)}</p>
    <form id="dailyCostForm" class="daily-cost-form"><input name="date" type="date" min="${activity.startDate}" max="${activity.finishDate}" required><input name="actualCost" type="number" min="0" step="0.01" placeholder="Actual cost" required><button class="primary-btn" type="submit">Save</button><button type="button" class="ghost-btn" id="closeDailyModalBtn">Close</button></form>
    <table><thead><tr><th>Date</th><th>Actual Cost</th></tr></thead><tbody>${rows}</tbody></table></div>`;

  modal.querySelector("#closeDailyModalBtn")?.addEventListener("click", () => modal.classList.add("hidden"));
  modal.querySelector("#dailyCostForm")?.addEventListener("submit", (event) => {
    event.preventDefault();
    const formData = new FormData(event.currentTarget);
    const date = String(formData.get("date") || "");
    const actualCost = parseBudgetValue(formData.get("actualCost"));
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
    showProjectDetails(projectId, activeTab, allActivities);
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

  detailsView.querySelectorAll(".view-daily-cost-btn").forEach((btn) => btn.addEventListener("click", () => renderDailyCostModal(projectId, btn.dataset.activityId)));
  return true;
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
const selectedProject = loadProjects().map(normalizeProject).find((project) =>
  (selectedProjectId && project.id === selectedProjectId)
  || (selectedProjectName && project.name === selectedProjectName)
);
const bootstrapCostManagement = async () => {
  const localActivities = loadCostActivities();
  const remoteActivities = await loadRemoteCostActivities();
  const merged = [...remoteActivities, ...localActivities];
  const deduped = new Map();
  merged.forEach((item) => {
    const key = `${String(item.projectId).trim()}::${String(item.id).trim()}`;
    if (!key || key === "::") return;
    if (!deduped.has(key)) deduped.set(key, item);
  });
  const allActivities = Array.from(deduped.values());

  if (!selectedProject || !showProjectDetails(selectedProject.id, selectedTab, allActivities)) renderProjects();
};

bootstrapCostManagement();
