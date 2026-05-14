const projectModal = document.getElementById("projectModal");
const projectForm = document.getElementById("projectForm");
const projectsPageHero = document.querySelector(".projects-topbar.page-hero");
const projectModalClose = document.getElementById("projectModalClose");
const projectModalBackdrop = document.getElementById("projectModalBackdrop");
const projectFormCancel = document.getElementById("projectFormCancel");
const openAddProjectModalBtn = document.getElementById("openAddProjectModalBtn");
const openAddProjectModalEmptyBtn = document.getElementById("openAddProjectModalEmptyBtn");
const projectsTableBody = document.getElementById("projectsTableBody");
const projectsEmptyState = document.querySelector(".projects-empty-state");
const projectsTableSummary = document.getElementById("projectsTableSummary");
const projectsSearchInput = document.getElementById("projectsSearchInput");
const projectsStatusFilter = document.getElementById("projectsStatusFilter");
const projectsTypeFilter = document.getElementById("projectsTypeFilter");
const projectModalTitle = document.getElementById("projectModalTitle");
const projectModalSubtitle = document.querySelector(".project-modal-header p");
const projectSubmitButton = projectForm?.querySelector('button[type="submit"]');
const openArchivedProjectsModalBtn = document.getElementById("openArchivedProjectsModalBtn");
const archivedProjectsModal = document.getElementById("archivedProjectsModal");
const archivedProjectsBackdrop = document.getElementById("archivedProjectsBackdrop");
const archivedProjectsClose = document.getElementById("archivedProjectsClose");
const archivedProjectsFooterClose = document.getElementById("archivedProjectsFooterClose");
const archivedProjectsSearch = document.getElementById("archivedProjectsSearch");
const archivedProjectsTypeFilter = document.getElementById("archivedProjectsTypeFilter");
const archivedProjectsTimeFilter = document.getElementById("archivedProjectsTimeFilter");
const archivedProjectsSort = document.getElementById("archivedProjectsSort");
const archivedProjectsTableBody = document.getElementById("archivedProjectsTableBody");
const archivedProjectsEmpty = document.getElementById("archivedProjectsEmpty");
const archivedProjectsSummary = document.getElementById("archivedProjectsSummary");

const kpiEls = {
  total: document.getElementById("kpiTotalProjects"),
  active: document.getElementById("kpiActiveProjects"),
  completed: document.getElementById("kpiCompletedProjects"),
  hold: document.getElementById("kpiOnHoldProjects"),
  archived: document.getElementById("kpiArchivedProjects"),
};

if (!projectModal || !projectForm || !projectsTableBody) {
  throw new Error("Projects page is missing required elements.");
}

const LOCAL_STORAGE_KEY = "constructionStageProjects";
const DATA_SOURCE_URL = window.DataBridge?.DEFAULT_DATA_SOURCE_URL || "";
const PROJECTS_REFRESH_INTERVAL_MS = 30 * 1000;

const RELATED_LOCAL_STORAGE_KEYS = {
  activities: "constructionStageActivities",
  costActivities: "constructionStageCostActivities",
  dailyCosts: "constructionStageDailyCosts",
};

const budgetInput = projectForm.elements.namedItem("budget");
const pesoBudgetFormatter = new Intl.NumberFormat("en-PH", {
  style: "currency",
  currency: "PHP",
  minimumFractionDigits: 2,
  maximumFractionDigits: 2,
});

const state = {
  allProjects: [],
  filteredProjects: [],
  editingProjectId: null,
  selectedProjectId: (() => {
    const query = new URLSearchParams(window.location.search);
    return query.get("projectId") || query.get("project") || "";
  })(),
};

let projectsRefreshTimer = null;
let isProjectsSyncInFlight = false;
let lastProjectsSignature = "";
let isProjectFormSubmitting = false;

const syncProjectsHeroState = () => {
  if (projectsPageHero) {
    projectsPageHero.hidden = Boolean(state.selectedProjectId);
  }
};

const selectProjectRecord = (projectId) => {
  const normalizedProjectId = String(projectId || "").trim();
  if (!normalizedProjectId) return;
  state.selectedProjectId = normalizedProjectId;
  const nextUrl = new URL(window.location.href);
  nextUrl.searchParams.set("projectId", normalizedProjectId);
  window.history.replaceState({}, "", nextUrl.toString());
  syncProjectsHeroState();
};


const getValueByAliases = (source, aliases = []) => {
  if (!source || typeof source !== "object") return undefined;

  for (const alias of aliases) {
    if (Object.prototype.hasOwnProperty.call(source, alias)) {
      return source[alias];
    }
  }

  const normalizedEntries = Object.keys(source).map((key) => ({
    key,
    normalized: String(key).toLowerCase().replace(/[^a-z0-9]/g, ""),
  }));

  for (const alias of aliases) {
    const normalizedAlias = String(alias).toLowerCase().replace(/[^a-z0-9]/g, "");
    const matched = normalizedEntries.find((entry) => entry.normalized === normalizedAlias);
    if (matched) return source[matched.key];
  }

  return undefined;
};

const toNumericBudgetValue = (value) => {
  const sanitized = String(value || "").replace(/[^\d.]/g, "");
  const [integerPart, ...fractionParts] = sanitized.split(".");
  const normalized = fractionParts.length
    ? `${integerPart}.${fractionParts.join("")}`
    : integerPart;

  return normalized.replace(/^0+(\d)/, "$1");
};

const parseBudgetValue = (value) => {
  if (value === null || value === undefined || value === "") return 0;
  if (typeof value === "number") return Number.isFinite(value) ? value : 0;
  const raw = String(value).trim();
  if (/^\d{4}-\d{2}-\d{2}$/.test(raw) || /^\d{1,2}\/\d{1,2}\/\d{2,4}/.test(raw)) return 0;
  const cleaned = raw.replace(/[^\d.-]/g, "");
  const parsed = Number(cleaned);
  return Number.isFinite(parsed) ? parsed : 0;
};

const formatBudgetAsPeso = (value) => {
  const normalized = toNumericBudgetValue(String(value || ""));

  if (!normalized) {
    return "";
  }

  const numericValue = Number.parseFloat(normalized);
  if (!Number.isFinite(numericValue)) {
    return "";
  }

  return pesoBudgetFormatter.format(numericValue);
};

const formatDate = (value) => {
  if (!value) return "-";
  const parsed = new Date(value);
  if (Number.isNaN(parsed.getTime())) return String(value);
  return parsed.toLocaleDateString("en-US", { month: "short", day: "numeric", year: "numeric" });
};

const formatTime = (value) => {
  if (!value) return "";
  const parsed = new Date(value);
  if (Number.isNaN(parsed.getTime())) return "";
  return parsed.toLocaleTimeString("en-US", { hour: "numeric", minute: "2-digit" });
};

const normalizeProject = (project = {}) => {
  const idRaw = getValueByAliases(project, ["id", "projectId", "project_id"]);
  const codeRaw = getValueByAliases(project, ["code", "projectCode", "project_code"]);
  const nameRaw = getValueByAliases(project, ["name", "project", "projectName", "project_name"]);
  const typeRaw = getValueByAliases(project, ["type", "projectType", "project_type"]);
  let statusRaw = getValueByAliases(project, ["status", "projectStatus", "project_status"]);
  let locationRaw = getValueByAliases(project, ["location", "projectLocation", "project_location", "site", "address"]);
  let startDateRaw = getValueByAliases(project, ["startDate", "start_date", "plannedStart", "planned_start"]);
  let finishDateRaw = getValueByAliases(project, ["finishDate", "targetFinish", "target_finish", "endDate", "end_date", "plannedFinish", "planned_finish"]);
  let budgetRaw = getValueByAliases(project, ["budget", "plannedValue", "planned_value", "plannedCost", "planned_cost"]);
  let descriptionRaw = getValueByAliases(project, ["description", "notes"]);
  const archivedDateRaw = getValueByAliases(project, ["archivedDate", "archived_date", "dateArchived", "date_archived"]);
  const archiveReasonRaw = getValueByAliases(project, ["archiveReason", "archive_reason", "reason", "archivedReason", "archived_reason"]);
  const createdAtRaw = getValueByAliases(project, ["createdAt", "created_at", "timestamp", "dateCreated"]);

  const isKnownStatus = (value) =>
    ["not started", "in progress", "on hold", "completed", "archived"].includes(
      String(value || "").trim().toLowerCase()
    );
  const isDateLike = (value) => {
    if (!value) return false;
    const parsed = new Date(value);
    return !Number.isNaN(parsed.getTime());
  };

  // Guard against shifted source data where values are offset by one column:
  // status <- type, location <- status, startDate <- location,
  // finishDate <- startDate, budget <- finishDate, description <- budget,
  // and Created At can hold the original budget value.
  const looksShifted =
    !isKnownStatus(statusRaw) &&
    isKnownStatus(locationRaw) &&
    !isDateLike(startDateRaw) &&
    isDateLike(finishDateRaw);

  if (looksShifted) {
    statusRaw = locationRaw;
    locationRaw = startDateRaw;
    startDateRaw = finishDateRaw;
    finishDateRaw = budgetRaw;
    budgetRaw = descriptionRaw || createdAtRaw;
    descriptionRaw = "";
  }

  const startDate = startDateRaw || "";
  const finishDate = finishDateRaw || "";
  const normalizedStatus = String(statusRaw || "Not Started").trim() || "Not Started";

  const progressByStatus = {
    Completed: 100,
    "In Progress": 55,
    "On Hold": 25,
    Archived: 0,
    "Not Started": 0,
  };

  return {
    id: String(idRaw || "").trim(),
    name: String(nameRaw || "Untitled Project").trim(),
    code: String(codeRaw || idRaw || "").trim(),
    type: String(typeRaw || "General").trim() || "General",
    status: normalizedStatus,
    location: String(locationRaw || "").trim(),
    startDate,
    finishDate,
    budget: parseBudgetValue(budgetRaw),
    progress:
      Number.isFinite(Number(getValueByAliases(project, ["progress", "percentComplete", "percent_complete"])))
        ? Math.max(0, Math.min(100, Number(getValueByAliases(project, ["progress", "percentComplete", "percent_complete"]))))
        : progressByStatus[normalizedStatus] ?? 0,
    description: String(descriptionRaw || "").trim(),
    archivedDate: archivedDateRaw || finishDate || createdAtRaw || "",
    archiveReason: String(archiveReasonRaw || descriptionRaw || "Project completed").trim(),
  };
};

const escapeHtml = (value) =>
  String(value)
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#39;");

const isProjectArchived = (project = {}) =>
  String(project.status || "").trim().toLowerCase() === "archived";

const closeArchivedProjectsModal = () => {
  archivedProjectsModal?.classList.add("hidden");
  document.body.style.overflow = "";
};

const openArchivedProjectsModal = () => {
  hydrateArchivedProjectFilters();
  renderArchivedProjects();
  archivedProjectsModal?.classList.remove("hidden");
  document.body.style.overflow = "hidden";
  archivedProjectsSearch?.focus({ preventScroll: true });
};

const closeProjectModal = () => {
  projectModal.classList.add("hidden");
  document.body.style.overflow = "";
  state.editingProjectId = null;

  if (projectModalTitle) projectModalTitle.textContent = "Add New Project";
  if (projectModalSubtitle) projectModalSubtitle.textContent = "Create a new project and define its details";
  if (projectSubmitButton) projectSubmitButton.innerHTML = `
    <svg viewBox="0 0 24 24" fill="none"><path d="M12 5v14M5 12h14"/></svg>
    Create Project
  `;
};

const openProjectModal = () => {
  projectModal.classList.remove("hidden");
  document.body.style.overflow = "hidden";
};

const openEditProjectModal = (project) => {
  if (!project) return;

  state.editingProjectId = project.id;
  projectForm.elements.projectName.value = project.name;
  projectForm.elements.projectCode.value = project.code;
  projectForm.elements.projectType.value = project.type;
  projectForm.elements.location.value = project.location || "";
  projectForm.elements.startDate.value = project.startDate;
  projectForm.elements.targetFinish.value = project.finishDate;
  projectForm.elements.status.value = project.status;
  if (projectForm.elements.budget) projectForm.elements.budget.value = formatBudgetAsPeso(project.budget);

  if (projectModalTitle) projectModalTitle.textContent = "Edit Project";
  if (projectModalSubtitle) projectModalSubtitle.textContent = "Update your project details";
  if (projectSubmitButton) projectSubmitButton.innerHTML = "Save Changes";

  openProjectModal();
};

const updateKpis = (projects) => {
  const countByStatus = projects.reduce((acc, project) => {
    const key = project.status || "Not Started";
    acc[key] = (acc[key] || 0) + 1;
    return acc;
  }, {});

  if (kpiEls.total) kpiEls.total.textContent = String(projects.length);
  if (kpiEls.active) kpiEls.active.textContent = String(countByStatus["In Progress"] || 0);
  if (kpiEls.completed) kpiEls.completed.textContent = String(countByStatus.Completed || 0);
  if (kpiEls.hold) kpiEls.hold.textContent = String(countByStatus["On Hold"] || 0);
  if (kpiEls.archived) kpiEls.archived.textContent = String(countByStatus.Archived || 0);
};

const populateFilterSelect = (selectEl, values, defaultLabel) => {
  if (!selectEl) return;

  const previousValue = selectEl.value;
  selectEl.innerHTML = `<option>${defaultLabel}</option>`;
  Array.from(new Set(values.filter(Boolean))).sort((a, b) => a.localeCompare(b)).forEach((value) => {
    const option = document.createElement("option");
    option.value = value;
    option.textContent = value;
    selectEl.append(option);
  });

  if ([...selectEl.options].some((opt) => opt.value === previousValue)) {
    selectEl.value = previousValue;
  }
};

const getStatusBadgeClass = (status) => {
  const normalized = String(status || "").toLowerCase();
  if (normalized === "completed") return "badge-completed";
  if (normalized === "in progress") return "badge-in-progress";
  if (normalized === "on hold") return "badge-delayed";
  if (normalized === "archived") return "badge-not-started";
  return "badge-not-started";
};

const getProgressFillClass = (status) => {
  const normalized = String(status || "").toLowerCase();
  if (normalized === "completed") return "progress-completed";
  if (normalized === "on hold") return "progress-delayed";
  return "";
};

const getPlannedCostByProject = () => {
  const raw = localStorage.getItem(RELATED_LOCAL_STORAGE_KEYS.costActivities);
  if (!raw) return new Map();

  let parsed;
  try {
    parsed = JSON.parse(raw);
  } catch {
    return new Map();
  }

  if (!Array.isArray(parsed)) return new Map();

  const totals = new Map();
  parsed.forEach((activity) => {
    const projectId = String(getValueByAliases(activity, ["projectId", "project_id", "project code", "projectCode"]) || "").trim();
    const projectName = String(getValueByAliases(activity, ["project", "projectName", "project_name"]) || "").trim().toLowerCase();
    const projectCode = String(getValueByAliases(activity, ["code", "projectCode", "project_code", "project code"]) || "").trim();
    const plannedCost = parseBudgetValue(getValueByAliases(activity, ["plannedCost", "planned_cost", "planned cost", "plannedValue", "planned_value", "planned value", "budget"]));

    if (!projectId && !projectName) return;
    const keys = [projectId, projectCode, projectName].filter(Boolean);
    keys.forEach((key) => {
      totals.set(key, (totals.get(key) || 0) + plannedCost);
    });
  });

  return totals;
};

const getDerivedBudgetForProject = (project = {}) => {
  const plannedCostByProject = getPlannedCostByProject();
  const projectId = String(project.id || "").trim();
  const projectCode = String(project.code || "").trim();
  const projectName = String(project.name || "").trim().toLowerCase();

  return (
    plannedCostByProject.get(projectId) ||
    plannedCostByProject.get(projectCode) ||
    plannedCostByProject.get(projectName) ||
    parseBudgetValue(project.budget)
  );
};

const renderProjects = (projects) => {
  if (!projects.length) {
    projectsTableBody.innerHTML = "";
    projectsEmptyState?.classList.remove("hidden");
    if (projectsTableSummary) {
      projectsTableSummary.textContent = "Showing 0 to 0 of 0 projects";
    }
    return;
  }

  projectsEmptyState?.classList.add("hidden");
  const plannedCostByProject = getPlannedCostByProject();

  projectsTableBody.innerHTML = projects
    .map(
      (project) => `
      <tr class="project-record-row" tabindex="0" data-project-row="${escapeHtml(project.id)}" aria-label="Select ${escapeHtml(project.name)}">
        <td>${escapeHtml(project.code || "-")}</td>
        <td>${escapeHtml(project.name)}</td>
        <td>${escapeHtml(project.type)}</td>
        <td><span class="badge ${getStatusBadgeClass(project.status)}">${escapeHtml(project.status)}</span></td>
        <td>${escapeHtml(project.location || "-")}</td>
        <td>${escapeHtml(formatDate(project.startDate))}</td>
        <td>${escapeHtml(formatDate(project.finishDate))}</td>
        <td>
          <div class="progress-cell">
            <div class="progress-track"><div class="progress-fill ${getProgressFillClass(project.status)}" style="width: ${Math.round(project.progress)}%;"></div></div>
            <span>${Math.round(project.progress)}%</span>
          </div>
        </td>
        <td class="project-budget-cell">${escapeHtml(pesoBudgetFormatter.format(
          plannedCostByProject.get(String(project.id || "").trim()) ||
          plannedCostByProject.get(String(project.code || "").trim()) ||
          plannedCostByProject.get(String(project.name || "").trim().toLowerCase()) ||
          project.budget ||
          0
        ))}</td>
        <td class="actions-col">
          <button type="button" class="action-menu-trigger" data-project-actions="${escapeHtml(project.id)}" aria-label="Open project actions" aria-expanded="false">⋮</button>
          <div class="project-actions-menu hidden" data-project-menu="${escapeHtml(project.id)}" role="menu" aria-label="Project actions">
            <button type="button" class="project-action-btn" data-project-edit="${escapeHtml(project.id)}" role="menuitem">Edit</button>
            <button type="button" class="project-action-btn danger" data-project-delete="${escapeHtml(project.id)}" role="menuitem">Delete</button>
          </div>
        </td>
      </tr>
    `
    )
    .join("");

  if (projectsTableSummary) {
    projectsTableSummary.textContent = `Showing 1 to ${projects.length} of ${projects.length} projects`;
  }
};

const saveToLocalStorage = (projects) => {
  localStorage.setItem(LOCAL_STORAGE_KEY, JSON.stringify(projects));
};

const readFromLocalStorage = () => {
  try {
    const raw = localStorage.getItem(LOCAL_STORAGE_KEY);
    if (!raw) return [];
    const parsed = JSON.parse(raw);
    if (!Array.isArray(parsed)) return [];
    return parsed.map(normalizeProject);
  } catch {
    return [];
  }
};

const resolveProjectBudgetForSync = (project = {}) => {
  if (typeof getDerivedBudgetForProject === "function") {
    return getDerivedBudgetForProject(project);
  }

  return parseBudgetValue(project.budget);
};

const syncProjectWithGoogleSheet = async ({ action, project, projectId }) => {
  if (!DATA_SOURCE_URL) return;

  const requestPayload = { resource: "projects", action };
  if (project) {
    requestPayload.project = {
      ...project,
      budget: resolveProjectBudgetForSync(project),
    };
  }
  if (projectId) requestPayload.projectId = projectId;

  const postWithFormat = async (format) =>
    fetch(DATA_SOURCE_URL, {
      method: "POST",
      headers: format === "json" ? { "Content-Type": "application/json" } : undefined,
      body:
        format === "json"
          ? JSON.stringify(requestPayload)
          : new URLSearchParams({ payload: JSON.stringify(requestPayload) }),
    });

  const parseResponsePayload = async (response) => {
    try {
      return await response.json();
    } catch {
      return null;
    }
  };

  let response;
  let payload;
  const sendViaGet = async () => {
    const url = new URL(DATA_SOURCE_URL);
    url.searchParams.set("payload", JSON.stringify(requestPayload));
    response = await fetch(url.toString(), { cache: "no-store" });
    payload = await parseResponsePayload(response);
  };

  try {
    response = await postWithFormat("form");
    payload = await parseResponsePayload(response);

    const needsJsonFallback =
      !response.ok ||
      (payload?.ok === false &&
        /invalid payload|invalid payload parameter json|invalid json payload/i.test(String(payload.error)));

    if (needsJsonFallback) {
      response = await postWithFormat("json");
      payload = await parseResponsePayload(response);
    }
  } catch (error) {
    const maybeCorsIssue =
      DATA_SOURCE_URL.includes("script.google.com/macros/s/") &&
      /failed to fetch|networkerror|cors/i.test(String(error?.message || ""));
    if (maybeCorsIssue) {
      try {
        await sendViaGet();
      } catch {
        // Use the shared error messaging below.
      }
    }

    if (!response || payload?.ok === false) {
      const guidance = maybeCorsIssue
        ? "CORS check failed for POST. Verify your Google Apps Script Web App is deployed to Anyone and use the latest /exec deployment URL."
        : "Unable to reach the Google Sheet endpoint.";
      throw new Error(
        `${guidance} If this endpoint was recently changed, update DATA_SOURCE_URL in data-service.js.`
      );
    }
  }

  if (!response.ok) {
    throw new Error(`Unable to save to Google Sheet (HTTP ${response.status}).`);
  }

  if (payload?.ok === false) {
    throw new Error(payload.error || "Unable to save to Google Sheet.");
  }

  window.DataBridge?.pollRealtimeSync?.();
};

const showProjectPersistenceError = (actionLabel, error) => {
  const reason = error?.message ? `\nReason: ${error.message}` : "";
  window.alert(
    `Unable to ${actionLabel} project in Google Sheets. Changes were not saved locally so the interface stays aligned with real-time data.${reason}`
  );
};

const loadProjectsFromSource = async () => {
  const localProjects = readFromLocalStorage();

  if (!DATA_SOURCE_URL) {
    return localProjects;
  }

  try {
    const response = await fetch(`${DATA_SOURCE_URL}?resource=projects`, { cache: "no-store" });
    if (!response.ok) {
      return localProjects;
    }

    const payload = await response.json();
    const remoteProjects = Array.isArray(payload?.projects)
      ? payload.projects.map(normalizeProject)
      : [];

    saveToLocalStorage(remoteProjects);
    return remoteProjects;
  } catch {
    return localProjects;
  }
};

const applyFilters = () => {
  const searchTerm = String(projectsSearchInput?.value || "").toLowerCase().trim();
  const statusValue = projectsStatusFilter?.value || "All Statuses";
  const typeValue = projectsTypeFilter?.value || "All Project Types";

  state.filteredProjects = state.allProjects.filter((project) => {
    if (isProjectArchived(project)) return false;

    const matchesSearch =
      !searchTerm ||
      project.name.toLowerCase().includes(searchTerm) ||
      project.code.toLowerCase().includes(searchTerm) ||
      project.type.toLowerCase().includes(searchTerm) ||
      project.location.toLowerCase().includes(searchTerm);

    const matchesStatus = statusValue === "All Statuses" || project.status === statusValue;
    const matchesType = typeValue === "All Project Types" || project.type === typeValue;

    return matchesSearch && matchesStatus && matchesType;
  });

  renderProjects(state.filteredProjects);
  updateKpis(state.allProjects);
};

const hydrateFilters = () => {
  populateFilterSelect(
    projectsStatusFilter,
    state.allProjects.filter((project) => !isProjectArchived(project)).map((project) => project.status),
    "All Statuses"
  );
  populateFilterSelect(
    projectsTypeFilter,
    state.allProjects.filter((project) => !isProjectArchived(project)).map((project) => project.type),
    "All Project Types"
  );
};

const getArchivedProjectTypeClass = (type) => {
  const normalized = String(type || "").toLowerCase();
  if (normalized.includes("residential")) return "residential";
  if (normalized.includes("infrastructure")) return "infrastructure";
  if (normalized.includes("industrial")) return "industrial";
  if (normalized.includes("institutional") || normalized.includes("school")) return "institutional";
  return "commercial";
};

const getArchivedProjects = () =>
  state.allProjects
    .filter(isProjectArchived)
    .slice()
    .sort((a, b) => new Date(b.archivedDate || b.finishDate || 0) - new Date(a.archivedDate || a.finishDate || 0));

const hydrateArchivedProjectFilters = () => {
  populateFilterSelect(
    archivedProjectsTypeFilter,
    getArchivedProjects().map((project) => project.type),
    "All Project Types"
  );
};

const matchesArchivedTimeRange = (project, timeRange) => {
  if (!timeRange || timeRange === "All Time") return true;
  const archivedDate = new Date(project.archivedDate || project.finishDate || "");
  if (Number.isNaN(archivedDate.getTime())) return true;

  const now = new Date();
  if (timeRange === "Last 30 Days") {
    const threshold = new Date(now);
    threshold.setDate(now.getDate() - 30);
    return archivedDate >= threshold;
  }
  if (timeRange === "Last 90 Days") {
    const threshold = new Date(now);
    threshold.setDate(now.getDate() - 90);
    return archivedDate >= threshold;
  }
  if (timeRange === "This Year") {
    return archivedDate.getFullYear() === now.getFullYear();
  }
  return true;
};

const getFilteredArchivedProjects = () => {
  const searchTerm = String(archivedProjectsSearch?.value || "").toLowerCase().trim();
  const typeValue = archivedProjectsTypeFilter?.value || "All Project Types";
  const timeRange = archivedProjectsTimeFilter?.value || "All Time";

  return getArchivedProjects().filter((project) => {
    const reason = project.archiveReason || project.description || "Project completed";
    const matchesSearch =
      !searchTerm ||
      project.name.toLowerCase().includes(searchTerm) ||
      project.code.toLowerCase().includes(searchTerm) ||
      project.type.toLowerCase().includes(searchTerm) ||
      reason.toLowerCase().includes(searchTerm);
    const matchesType = typeValue === "All Project Types" || project.type === typeValue;
    return matchesSearch && matchesType && matchesArchivedTimeRange(project, timeRange);
  });
};

const renderArchivedProjects = () => {
  if (!archivedProjectsTableBody) return;

  const archivedProjects = getFilteredArchivedProjects();
  archivedProjectsTableBody.innerHTML = archivedProjects
    .map((project) => {
      const typeClass = getArchivedProjectTypeClass(project.type);
      const archivedDate = project.archivedDate || project.finishDate || project.startDate;
      const reason = project.archiveReason || project.description || "Project completed";
      return `
        <tr>
          <td>
            <div class="archived-project-identity">
              <span class="archived-project-avatar ${typeClass}" aria-hidden="true">
                <svg viewBox="0 0 24 24" fill="none"><path d="M4 21V8l8-5 8 5v13M9 21v-7h6v7M8 10h.01M16 10h.01"/></svg>
              </span>
              <span>
                <span class="archived-project-name">${escapeHtml(project.name)}</span>
                <span class="archived-project-id">${escapeHtml(project.code || project.id || "—")}</span>
              </span>
            </div>
          </td>
          <td>${escapeHtml(project.type)}</td>
          <td><span class="archived-date-main">${escapeHtml(formatDate(archivedDate))}</span><span class="archived-date-time">${escapeHtml(formatTime(archivedDate) || "09:00 AM")}</span></td>
          <td>${escapeHtml(reason)}</td>
          <td>
            <div class="archived-actions">
              <button type="button" class="archived-restore-btn" data-archived-restore="${escapeHtml(project.id)}" aria-label="Restore ${escapeHtml(project.name)}" title="Restore">
                <svg viewBox="0 0 24 24" fill="none" aria-hidden="true"><path d="M4 12a8 8 0 1 0 2.34-5.66M4 4v6h6"/></svg>
              </button>
              <button type="button" class="archived-delete-btn" data-archived-delete="${escapeHtml(project.id)}" aria-label="Delete ${escapeHtml(project.name)}" title="Delete">
                <svg viewBox="0 0 24 24" fill="none" aria-hidden="true"><path d="M4 7h16M10 11v6M14 11v6M6 7l1 14h10l1-14M9 7V4h6v3"/></svg>
              </button>
            </div>
          </td>
        </tr>`;
    })
    .join("");

  archivedProjectsEmpty?.classList.toggle("hidden", archivedProjects.length > 0);
  if (archivedProjectsSummary) {
    archivedProjectsSummary.textContent = archivedProjects.length
      ? `Showing 1 to ${Math.min(5, archivedProjects.length)} of ${archivedProjects.length} archived projects`
      : "Showing 0 to 0 of 0 archived projects";
  }
};



const addProject = async (project) => {
  await syncProjectWithGoogleSheet({ action: "create", project });
  state.allProjects = [project, ...state.allProjects];
  saveToLocalStorage(state.allProjects);
  hydrateFilters();
  applyFilters();
};

const addProjectLocally = (project) => {
  state.allProjects = [project, ...state.allProjects];
  saveToLocalStorage(state.allProjects);
  hydrateFilters();
  applyFilters();
};

const updateProject = async (updatedProject) => {
  await syncProjectWithGoogleSheet({ action: "update", project: updatedProject });
  state.allProjects = state.allProjects.map((project) =>
    project.id === updatedProject.id ? updatedProject : project
  );
  saveToLocalStorage(state.allProjects);
  hydrateFilters();
  hydrateArchivedProjectFilters();
  applyFilters();
  renderArchivedProjects();
};

const restoreArchivedProject = async (projectId) => {
  const projectToRestore = state.allProjects.find((project) => project.id === projectId);
  if (!projectToRestore) return;
  await updateProject({
    ...projectToRestore,
    status: "Not Started",
    progress: 0,
    archiveReason: "",
  });
};

const removeRelatedProjectDataFromLocalStorage = (projectToDelete) => {
  if (!projectToDelete) return;

  const projectId = String(projectToDelete.id || "").trim();
  const projectName = String(projectToDelete.name || "").trim().toLowerCase();

  const safeParseArray = (rawValue) => {
    try {
      const parsed = JSON.parse(rawValue || "[]");
      return Array.isArray(parsed) ? parsed : [];
    } catch {
      return [];
    }
  };

  const normalizeProjectIdentity = (value) => String(value || "").trim().toLowerCase();

  const shouldDeleteByProject = (entry) => {
    const entryProjectId = normalizeProjectIdentity(entry?.projectId || entry?.project_id || entry?.project);
    const entryProjectName = normalizeProjectIdentity(entry?.projectName || entry?.project_name || entry?.project);

    if (projectId && entryProjectId === projectId.toLowerCase()) return true;
    if (projectId && entryProjectName === projectId.toLowerCase()) return true;
    if (projectName && entryProjectName === projectName) return true;
    if (projectName && entryProjectId === projectName) return true;
    return false;
  };

  const activities = safeParseArray(localStorage.getItem(RELATED_LOCAL_STORAGE_KEYS.activities));
  const nextActivities = activities.filter((activity) => !shouldDeleteByProject(activity));
  localStorage.setItem(RELATED_LOCAL_STORAGE_KEYS.activities, JSON.stringify(nextActivities));

  const costActivities = safeParseArray(localStorage.getItem(RELATED_LOCAL_STORAGE_KEYS.costActivities));
  const nextCostActivities = costActivities.filter((activity) => !shouldDeleteByProject(activity));
  localStorage.setItem(RELATED_LOCAL_STORAGE_KEYS.costActivities, JSON.stringify(nextCostActivities));

  const dailyCosts = safeParseArray(localStorage.getItem(RELATED_LOCAL_STORAGE_KEYS.dailyCosts));
  const nextDailyCosts = dailyCosts.filter((dailyCost) => !shouldDeleteByProject(dailyCost));
  localStorage.setItem(RELATED_LOCAL_STORAGE_KEYS.dailyCosts, JSON.stringify(nextDailyCosts));
};

const deleteProject = async (projectId) => {
  const projectToDelete = state.allProjects.find((project) => project.id === projectId);
  await syncProjectWithGoogleSheet({ action: "delete", projectId });
  removeRelatedProjectDataFromLocalStorage(projectToDelete);
  state.allProjects = state.allProjects.filter((project) => project.id !== projectId);
  saveToLocalStorage(state.allProjects);
  hydrateFilters();
  hydrateArchivedProjectFilters();
  applyFilters();
  renderArchivedProjects();
};

const closeAllActionMenus = () => {
  document.querySelectorAll(".project-actions-menu").forEach((menu) => {
    menu.classList.add("hidden");
    menu.classList.remove("open-up");
  });
  document.querySelectorAll(".action-menu-trigger").forEach((button) => {
    button.setAttribute("aria-expanded", "false");
  });
};

const toggleActionMenu = (projectId, triggerBtn) => {
  if (!projectId || !(triggerBtn instanceof HTMLElement)) return;

  const targetMenu = projectsTableBody.querySelector(`[data-project-menu="${projectId}"]`);
  if (!(targetMenu instanceof HTMLElement)) return;

  const isClosed = targetMenu.classList.contains("hidden");
  closeAllActionMenus();
  if (!isClosed) {
    return;
  }

  targetMenu.classList.remove("hidden");
  targetMenu.classList.remove("open-up");
  triggerBtn.setAttribute("aria-expanded", "true");

  const menuRect = targetMenu.getBoundingClientRect();
  const viewportBottom = window.innerHeight;
  const desiredBottomPadding = 16;
  if (menuRect.bottom > viewportBottom - desiredBottomPadding) {
    targetMenu.classList.add("open-up");
  }
};

projectsTableBody.addEventListener("click", async (event) => {
  const row = event.target.closest("[data-project-row]");
  if (row instanceof HTMLElement && !event.target.closest(".actions-col")) {
    selectProjectRecord(row.dataset.projectRow);
    return;
  }

  const triggerBtn = event.target.closest("[data-project-actions]");
  if (triggerBtn instanceof HTMLElement) {
    const projectId = triggerBtn.dataset.projectActions;
    toggleActionMenu(projectId, triggerBtn);
    return;
  }

  const editBtn = event.target.closest("[data-project-edit]");
  if (editBtn instanceof HTMLElement) {
    closeAllActionMenus();
    const projectToEdit = state.allProjects.find((project) => project.id === editBtn.dataset.projectEdit);
    openEditProjectModal(projectToEdit);
    return;
  }

  const deleteBtn = event.target.closest("[data-project-delete]");
  if (deleteBtn instanceof HTMLElement) {
    const projectId = deleteBtn.dataset.projectDelete;
    const projectToDelete = state.allProjects.find((project) => project.id === projectId);
    if (!projectToDelete) return;
    closeAllActionMenus();
    const isConfirmed = window.confirm(`Delete "${projectToDelete.name}"? This will also delete all related activities and costing records. This action cannot be undone.`);
    if (!isConfirmed) return;
    try {
      await deleteProject(projectId);
    } catch (error) {
      console.warn(error);
      showProjectPersistenceError("delete", error);
    }
  }
});

projectsTableBody.addEventListener("keydown", (event) => {
  if (event.key !== "Enter" && event.key !== " ") return;
  const row = event.target.closest("[data-project-row]");
  if (!(row instanceof HTMLElement) || event.target.closest(".actions-col")) return;
  event.preventDefault();
  selectProjectRecord(row.dataset.projectRow);
});

syncProjectsHeroState();

document.addEventListener("click", (event) => {
  if (!(event.target instanceof Element)) return;
  if (event.target.closest(".actions-col")) return;
  closeAllActionMenus();
});

if (budgetInput instanceof HTMLInputElement) {
  budgetInput.addEventListener("focus", () => {
    budgetInput.value = toNumericBudgetValue(budgetInput.value);
  });

  budgetInput.addEventListener("input", () => {
    budgetInput.value = toNumericBudgetValue(budgetInput.value);

    if (budgetInput.value) {
      budgetInput.setCustomValidity("");
      return;
    }

    budgetInput.setCustomValidity("Please enter a budget amount.");
  });

  budgetInput.addEventListener("blur", () => {
    if (!budgetInput.value) {
      return;
    }

    const formattedValue = formatBudgetAsPeso(budgetInput.value);

    if (!formattedValue) {
      budgetInput.setCustomValidity("Please enter a valid budget amount.");
      return;
    }

    budgetInput.setCustomValidity("");
    budgetInput.value = formattedValue;
  });
}

openAddProjectModalBtn?.addEventListener("click", openProjectModal);
openAddProjectModalEmptyBtn?.addEventListener("click", openProjectModal);
projectModalClose?.addEventListener("click", closeProjectModal);
projectFormCancel?.addEventListener("click", closeProjectModal);
projectModalBackdrop?.addEventListener("click", closeProjectModal);
openArchivedProjectsModalBtn?.addEventListener("click", openArchivedProjectsModal);
archivedProjectsClose?.addEventListener("click", closeArchivedProjectsModal);
archivedProjectsFooterClose?.addEventListener("click", closeArchivedProjectsModal);
archivedProjectsBackdrop?.addEventListener("click", closeArchivedProjectsModal);
archivedProjectsSearch?.addEventListener("input", renderArchivedProjects);
archivedProjectsTypeFilter?.addEventListener("change", renderArchivedProjects);
archivedProjectsTimeFilter?.addEventListener("change", renderArchivedProjects);
archivedProjectsSort?.addEventListener("click", renderArchivedProjects);
archivedProjectsTableBody?.addEventListener("click", async (event) => {
  if (!(event.target instanceof Element)) return;
  const restoreBtn = event.target.closest("[data-archived-restore]");
  if (restoreBtn instanceof HTMLElement) {
    try {
      await restoreArchivedProject(restoreBtn.dataset.archivedRestore);
    } catch (error) {
      console.warn(error);
      showProjectPersistenceError("restore", error);
    }
    return;
  }

  const deleteBtn = event.target.closest("[data-archived-delete]");
  if (deleteBtn instanceof HTMLElement) {
    const projectId = deleteBtn.dataset.archivedDelete;
    const projectToDelete = state.allProjects.find((project) => project.id === projectId);
    if (!projectToDelete) return;
    const isConfirmed = window.confirm(`Permanently delete "${projectToDelete.name}" from archived projects? This action cannot be undone.`);
    if (!isConfirmed) return;
    try {
      await deleteProject(projectId);
    } catch (error) {
      console.warn(error);
      showProjectPersistenceError("delete", error);
    }
  }
});

projectModal.addEventListener("click", (event) => {
  if (event.target === projectModal) {
    closeProjectModal();
  }
});

document.addEventListener("keydown", (event) => {
  if (event.key === "Escape") {
    closeAllActionMenus();
  }

  if (event.key === "Escape" && !projectModal.classList.contains("hidden")) {
    closeProjectModal();
  }

  if (event.key === "Escape" && archivedProjectsModal && !archivedProjectsModal.classList.contains("hidden")) {
    closeArchivedProjectsModal();
  }
});

projectForm.addEventListener("submit", async (event) => {
  event.preventDefault();
  if (isProjectFormSubmitting) return;

  const formData = new FormData(projectForm);
  const projectName = String(formData.get("projectName") || "").trim();
  const projectCode = String(formData.get("projectCode") || "").trim();
  const projectType = String(formData.get("projectType") || "").trim();
  const location = String(formData.get("location") || "").trim();
  const startDate = String(formData.get("startDate") || "").trim();
  const targetFinish = String(formData.get("targetFinish") || "").trim();
  const status = String(formData.get("status") || "Not Started").trim();
  const budget = parseBudgetValue(formData.get("budget"));
  const editingProject = state.allProjects.find((project) => project.id === state.editingProjectId);
  const description = editingProject?.description || "";

  if (!projectName || !projectCode || !projectType || !location || !startDate || !targetFinish) {
    return;
  }

  const projectId = state.editingProjectId || projectCode;

  const project = normalizeProject({
    id: projectId,
    name: projectName,
    code: projectCode,
    type: projectType,
    status,
    location,
    startDate,
    finishDate: targetFinish,
    budget,
    description,
  });

  let syncError = null;

  isProjectFormSubmitting = true;
  if (projectSubmitButton) {
    projectSubmitButton.disabled = true;
    projectSubmitButton.dataset.originalLabel = projectSubmitButton.textContent || "Save";
    projectSubmitButton.textContent = "Saving...";
  }

  try {
    if (state.editingProjectId) {
      await updateProject(project);
    } else {
      await addProject(project);
    }
  } catch (error) {
    console.warn(error);

    if (state.editingProjectId) {
      showProjectPersistenceError("update", error);
      isProjectFormSubmitting = false;
      if (projectSubmitButton) {
        projectSubmitButton.disabled = false;
        projectSubmitButton.textContent = projectSubmitButton.dataset.originalLabel || "Save Project";
      }
      return;
    }

    addProjectLocally(project);
    syncError = error;
    window.alert(
      `Project was saved locally but could not sync to Google Sheets.\nReason: ${error?.message || "Unknown error"}`
    );
  }

  closeProjectModal();
  projectForm.reset();

  if (syncError) {
    refreshProjectsIfVisible({ force: true });
  }

  isProjectFormSubmitting = false;
  if (projectSubmitButton) {
    projectSubmitButton.disabled = false;
    projectSubmitButton.textContent = projectSubmitButton.dataset.originalLabel || "Save Project";
  }
});

projectsSearchInput?.addEventListener("input", applyFilters);
projectsStatusFilter?.addEventListener("change", applyFilters);
projectsTypeFilter?.addEventListener("change", applyFilters);

const bootstrapProjectsPage = async () => {
  const projects = await loadProjectsFromSource();
  const normalizedProjects = Array.isArray(projects) ? projects : [];
  const nextSignature = JSON.stringify(normalizedProjects);
  if (nextSignature === lastProjectsSignature) return;
  lastProjectsSignature = nextSignature;

  state.allProjects = normalizedProjects;
  hydrateFilters();
  applyFilters();
};

const refreshProjectsIfVisible = async ({ force = false } = {}) => {
  if (isProjectsSyncInFlight) return;
  if (!force && document.visibilityState === "hidden") return;
  if (projectModal && !projectModal.classList.contains("hidden")) return;
  if (archivedProjectsModal && !archivedProjectsModal.classList.contains("hidden")) return;
  if (document.activeElement instanceof HTMLElement) {
    const isTyping =
      document.activeElement.matches("input, textarea, select")
      || document.activeElement.isContentEditable;
    if (isTyping) return;
  }

  isProjectsSyncInFlight = true;
  try {
    await bootstrapProjectsPage();
  } finally {
    isProjectsSyncInFlight = false;
  }
};

const setupProjectsRealtimeSync = () => {
  if (projectsRefreshTimer) {
    clearInterval(projectsRefreshTimer);
  }

  projectsRefreshTimer = setInterval(() => {
    refreshProjectsIfVisible();
  }, PROJECTS_REFRESH_INTERVAL_MS);

  window.addEventListener("focus", () => refreshProjectsIfVisible({ force: true }));
  window.addEventListener("online", () => refreshProjectsIfVisible({ force: true }));
  window.addEventListener("google-sheet:changed", () => refreshProjectsIfVisible({ force: true }));
  document.addEventListener("visibilitychange", () => {
    if (document.visibilityState === "visible") {
      refreshProjectsIfVisible({ force: true });
    }
  });
};

refreshProjectsIfVisible({ force: true });
setupProjectsRealtimeSync();
