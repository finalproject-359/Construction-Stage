const activitiesTableBody = document.getElementById("activitiesTableBody");
const activitiesSearchInput = document.getElementById("activitiesSearchInput");
const activitiesStatusFilter = document.getElementById("activitiesStatusFilter");
const activitiesTypeFilter = document.getElementById("activitiesTypeFilter");
const activitiesSelectedProjectName = document.getElementById("activitiesSelectedProjectName");
const activitiesBackToProjectsBtn = document.getElementById("activitiesBackToProjectsBtn");
const activitiesDateFilterWrap = document.querySelector(".activities-date-filter");
const activitiesDateFilterBtn = document.getElementById("activitiesDateFilterBtn");
const activitiesDateFilterLabel = document.getElementById("activitiesDateFilterLabel");
const activitiesDateRangePanel = document.getElementById("activitiesDateRangePanel");
const activitiesFilterStartDate = document.getElementById("activitiesFilterStartDate");
const activitiesFilterEndDate = document.getElementById("activitiesFilterEndDate");
const activitiesDateClearBtn = document.getElementById("activitiesDateClearBtn");
const activitiesDateApplyBtn = document.getElementById("activitiesDateApplyBtn");
const activitiesTableSummary = document.getElementById("activitiesTableSummary");
const activitiesProjectSelectionFilters = document.getElementById("activitiesProjectSelectionFilters");
const activitiesProjectSelection = document.getElementById("activitiesProjectSelection");
const activitiesViewShell = document.getElementById("activitiesViewShell");
const activitiesProjectPickerSearch = document.getElementById("activitiesProjectPickerSearch");
const activitiesProjectPickerGrid = document.getElementById("activitiesProjectPickerGrid");
const activitiesProjectSelectionFooter = document.getElementById("activitiesProjectSelectionFooter");
const activitiesProjectTypeFilter = document.getElementById("activitiesProjectTypeFilter");
const activitiesProjectStatusFilter = document.getElementById("activitiesProjectStatusFilter");
const activitiesProjectDateFilterWrap = document.querySelector(".activities-selection-date-filter");
const activitiesProjectDateFilterBtn = document.getElementById("activitiesProjectDateFilterBtn");
const activitiesProjectDateFilterLabel = document.getElementById("activitiesProjectDateFilterLabel");
const activitiesProjectDateRangePanel = document.getElementById("activitiesProjectDateRangePanel");
const activitiesProjectFilterStartDate = document.getElementById("activitiesProjectFilterStartDate");
const activitiesProjectFilterEndDate = document.getElementById("activitiesProjectFilterEndDate");
const activitiesProjectDateClearBtn = document.getElementById("activitiesProjectDateClearBtn");
const activitiesProjectDateApplyBtn = document.getElementById("activitiesProjectDateApplyBtn");
const activitiesHowItWorksBtn = document.getElementById("activitiesHowItWorksBtn");
const activitiesAddButton = document.getElementById("activitiesAddButton");
const activitiesPagination = document.querySelector(".activities-pagination");
const activityModal = document.getElementById("activityModal");
const activityModalBackdrop = document.getElementById("activityModalBackdrop");
const activityModalCloseBtn = document.getElementById("activityModalCloseBtn");
const activityModalCancelBtn = document.getElementById("activityModalCancelBtn");
const activityModalForm = document.getElementById("activityModalForm");
const activityModalProjectInput = document.getElementById("activityModalProjectInput");
const activityModalProjectDisplay = document.getElementById("activityModalProjectDisplay");
const activityModalStartDateInput = document.getElementById("activityModalStartDateInput");
const activityModalFinishDateInput = document.getElementById("activityModalFinishDateInput");
const activityModalTitle = document.getElementById("activityModalTitle");
const activityModalSubtitle = document.querySelector(".activity-modal-header p");
const activityModalSubmitBtn = activityModalForm?.querySelector('button[type="submit"]');
const DATA_SOURCE_URL = window.DataBridge?.DEFAULT_DATA_SOURCE_URL || "";
const LOCAL_STORAGE_KEY = "constructionStageActivities";
const PROJECTS_LOCAL_STORAGE_KEY = "constructionStageProjects";

if (!activitiesTableBody) {
  throw new Error("Activities page is missing the table body element.");
}

const kpiEls = {
  total: document.getElementById("kpiTotalActivities"),
  completed: document.getElementById("kpiCompletedActivities"),
  inProgress: document.getElementById("kpiInProgressActivities"),
  notStarted: document.getElementById("kpiNotStartedActivities"),
  delayed: document.getElementById("kpiDelayedActivities"),
};

const statusTextToKey = {
  Completed: "completed",
  "In Progress": "inProgress",
  "Not Started": "notStarted",
  Delayed: "delayed",
};

const BADGE_CLASS_BY_STATUS = {
  Completed: "badge-completed",
  "In Progress": "badge-in-progress",
  "Not Started": "badge-not-started",
  Delayed: "badge-delayed",
};

const PROGRESS_CLASS_BY_STATUS = {
  Completed: "progress-completed",
  Delayed: "progress-delayed",
};

const PAGE_SIZE = 8;
const ACTIVITIES_REFRESH_INTERVAL_MS = 20 * 1000;

let activitiesRefreshTimer = null;
let isActivitiesSyncInFlight = false;
let lastActivitiesSignature = "";

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

const uniqueSorted = (values) =>
  Array.from(new Set(values.filter(Boolean))).sort((a, b) => a.localeCompare(b));

const populateSelect = (selectEl, values, defaultLabel) => {
  if (!selectEl) return;

  selectEl.innerHTML = `<option>${defaultLabel}</option>`;
  values.forEach((value) => {
    const option = document.createElement("option");
    option.value = value;
    option.textContent = value;
    selectEl.append(option);
  });
};

const parseDateStringAsLocalDate = (value) => {
  if (typeof value !== "string") return null;
  const trimmed = value.trim();
  const match = trimmed.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (!match) return null;

  const year = Number(match[1]);
  const monthIndex = Number(match[2]) - 1;
  const day = Number(match[3]);
  const localDate = new Date(year, monthIndex, day);
  localDate.setHours(0, 0, 0, 0);
  return Number.isNaN(localDate.getTime()) ? null : localDate;
};

const toDisplayDate = (value) => {
  if (!value) return "-";
  if (value instanceof Date && !Number.isNaN(value.getTime())) {
    return value.toLocaleDateString("en-US", { month: "short", day: "numeric", year: "numeric" });
  }

  const parsed = parseDateStringAsLocalDate(value) || new Date(value);
  if (Number.isNaN(parsed.getTime())) return String(value);
  return parsed.toLocaleDateString("en-US", { month: "short", day: "numeric", year: "numeric" });
};

const parseDateValue = (value) => {
  if (!value) return null;
  const parsed = value instanceof Date ? value : parseDateStringAsLocalDate(value) || new Date(value);
  if (Number.isNaN(parsed.getTime())) return null;
  parsed.setHours(0, 0, 0, 0);
  return parsed;
};

const normalizeProject = (project = {}) => {
  const name = getValueByAliases(project, ["name", "project", "projectName", "project_name"]);
  const code = getValueByAliases(project, ["code", "projectCode", "project_code"]);
  const location = getValueByAliases(project, ["location", "site", "address"]);
  const type = getValueByAliases(project, ["type", "projectType", "project_type"]);
  const status = getValueByAliases(project, ["status", "projectStatus", "project_status"]);
  const startDateRaw = getValueByAliases(project, ["startDate", "plannedStart", "planned_start"]);
  const endDateRaw = getValueByAliases(project, ["endDate", "plannedFinish", "planned_finish"]);

  return {
    name: name || "Untitled Project",
    code: code || "",
    location: location || "",
    type: type || "General",
    status: status || "Not Started",
    startDate: parseDateValue(startDateRaw),
    endDate: parseDateValue(endDateRaw),
  };
};

const toPercent = (value) => {
  const n = Number(value);
  if (!Number.isFinite(n)) return 0;
  return Math.max(0, Math.min(100, Math.round(n)));
};

const escapeHtml = (value) =>
  String(value)
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#39;");

const normalizeActivity = (activity = {}) => {
  const id = getValueByAliases(activity, ["id", "activityId", "activity_id", "code", "activityCode", "activity_code"]);
  const name = getValueByAliases(activity, ["name", "activity", "activityName", "activity_name"]);
  const project = getValueByAliases(activity, ["project", "projectName", "project_name", "projectId", "project_id"]);
  let type = getValueByAliases(activity, ["type", "activityType", "activity_type"]);
  let status = getValueByAliases(activity, ["status"]);
  let plannedStartRaw = getValueByAliases(activity, ["plannedStart", "planned_start", "startDate", "plannedStartDate"]);
  let plannedFinishRaw = getValueByAliases(activity, ["plannedFinish", "planned_finish", "finishDate", "plannedFinishDate"]);
  let progressRaw = getValueByAliases(activity, ["progress", "percentComplete", "percent_complete"]);
  let costStatus = getValueByAliases(activity, ["costStatus", "cost_status", "budgetStatus"]);

  const isDateLike = (value) => Boolean(parseDateValue(value));
  const isKnownStatus = (value) => Object.prototype.hasOwnProperty.call(BADGE_CLASS_BY_STATUS, String(value || "").trim());
  const isNumericLike = (value) => /^-?\d+(\.\d+)?%?$/.test(String(value || "").trim());
  const isKnownCostStatus = (value) => ["Under Budget", "On Budget", "Over Budget"].includes(String(value || "").trim());
  const normalizedType = String(type || "").trim();

  // Guard against shifted source data where values are offset by one column:
  // type <- status, status <- plannedStart, plannedStart <- plannedFinish,
  // plannedFinish <- progress, progress <- costStatus.
  const looksShifted =
    isKnownStatus(normalizedType) &&
    isDateLike(status) &&
    isDateLike(plannedStartRaw) &&
    (
      (isNumericLike(plannedFinishRaw) && isKnownCostStatus(progressRaw)) ||
      (isNumericLike(progressRaw) && isKnownCostStatus(costStatus)) ||
      (!plannedFinishRaw && isNumericLike(progressRaw))
    );

  if (looksShifted) {
    const shiftedStatus = type;
    const shiftedPlannedStart = status;
    const shiftedPlannedFinish = plannedStartRaw;
    const shiftedProgress = isNumericLike(plannedFinishRaw) ? plannedFinishRaw : progressRaw;
    const shiftedCostStatus = isKnownCostStatus(progressRaw) ? progressRaw : costStatus;

    type = "-";
    status = shiftedStatus;
    plannedStartRaw = shiftedPlannedStart;
    plannedFinishRaw = shiftedPlannedFinish;
    progressRaw = shiftedProgress;
    costStatus = shiftedCostStatus;
  }

  if (!status && Object.prototype.hasOwnProperty.call(BADGE_CLASS_BY_STATUS, normalizedType)) {
    status = normalizedType;
    type = "-";
  }

  status = status || "Not Started";
  costStatus = costStatus || "On Budget";
  const progress = progressRaw ?? (status === "Completed" ? 100 : 0);

  return {
    id: id || "-",
    name: name || "Untitled Activity",
    project: project || "-",
    type: type || "-",
    status,
    plannedStart: toDisplayDate(plannedStartRaw),
    plannedFinish: toDisplayDate(plannedFinishRaw),
    plannedStartDate: parseDateValue(plannedStartRaw),
    plannedFinishDate: parseDateValue(plannedFinishRaw),
    progress: toPercent(progress),
    costStatus,
  };
};

const createActivityKey = (activity = {}) =>
  [
    String(activity.id || "").trim().toLowerCase(),
    String(activity.project || "").trim().toLowerCase(),
    String(activity.name || "").trim().toLowerCase(),
    String(activity.type || "").trim().toLowerCase(),
    String(activity.plannedStart || "").trim().toLowerCase(),
    String(activity.plannedFinish || "").trim().toLowerCase(),
  ].join("::");

const toInputDate = (value) => {
  const dateValue = parseDateValue(value);
  if (!dateValue) return "";
  const year = dateValue.getFullYear();
  const month = String(dateValue.getMonth() + 1).padStart(2, "0");
  const day = String(dateValue.getDate()).padStart(2, "0");
  return `${year}-${month}-${day}`;
};

const buildActivityRowHtml = (activity) => {
  const progressClass = PROGRESS_CLASS_BY_STATUS[activity.status] || "";
  const statusClass = BADGE_CLASS_BY_STATUS[activity.status] || "badge-on";
  const rowKey = createActivityKey(activity);
  return `
    <tr>
      <td>${escapeHtml(activity.id)}</td>
      <td>${escapeHtml(activity.name)}</td>
      <td>${escapeHtml(activity.type)}</td>
      <td><span class="badge ${statusClass}">${escapeHtml(activity.status)}</span></td>
      <td>${escapeHtml(activity.plannedStart)}</td>
      <td>${escapeHtml(activity.plannedFinish)}</td>
      <td>
        <div class="progress-cell"><div class="progress-track"><div class="progress-fill ${progressClass}" style="width:${activity.progress}%"></div></div><span>${activity.progress}%</span></div>
      </td>
      <td class="actions-col">
        <div class="activity-row-actions">
          <button type="button" class="project-action-btn" data-action="edit" data-activity-key="${escapeHtml(rowKey)}">Edit</button>
          <button type="button" class="project-action-btn danger" data-action="delete" data-activity-key="${escapeHtml(rowKey)}">Delete</button>
        </div>
      </td>
    </tr>
  `;
};

const EMPTY_STATE_HTML = `
  <div class="activities-empty-state">
    <div class="activities-empty-illustration" aria-hidden="true">
      <svg viewBox="0 0 120 120" fill="none">
        <rect x="30" y="22" width="60" height="74" rx="8" />
        <rect x="51" y="14" width="18" height="14" rx="6" />
        <path d="m46 46 6 6 10-10" />
        <path d="m46 62 6 6 10-10" />
        <path d="m46 78 6 6 10-10" />
        <path d="M66 48h14M66 64h14M66 80h14" />
      </svg>
    </div>
    <h3>No activities yet</h3>
    <p>Get started by adding your first activity</p>
    <button id="activitiesAddButtonEmpty" class="activities-add-btn" type="button">
      <svg viewBox="0 0 24 24" fill="none" aria-hidden="true"><path d="M12 5v14M5 12h14"/></svg>
      Add Activity
    </button>
  </div>
`;

const renderEmptyState = (message = "Get started by adding your first activity") => {
  activitiesTableBody.innerHTML = `
    <tr class="activities-empty-row">
      <td colspan="8">${message === "Get started by adding your first activity" ? EMPTY_STATE_HTML : escapeHtml(message)}</td>
    </tr>
  `;
};

const initialActivities = Array.isArray(window.activitiesData)
  ? window.activitiesData.map(normalizeActivity)
  : [];
const initialProjectCatalog = Array.isArray(window.activitiesProjectCatalog)
  ? window.activitiesProjectCatalog.map(normalizeProject)
  : [];

const state = {
  allActivities: initialActivities,
  filteredActivities: initialActivities,
  currentPage: 1,
  selectedProject: null,
  projectSearch: "",
  projectDateRange: {
    start: null,
    end: null,
  },
  dateRange: {
    start: null,
    end: null,
  },
  didHydrateProjectFromUrl: false,
  editingActivityKey: null,
};

const updateActivitiesUrlParams = ({ project = state.selectedProject, keepAddedFlag = false } = {}) => {
  const nextUrl = new URL(window.location.href);
  if (project) {
    nextUrl.searchParams.set("project", project);
  } else {
    nextUrl.searchParams.delete("project");
  }

  if (!keepAddedFlag) {
    nextUrl.searchParams.delete("added");
  }

  const nextHref = `${nextUrl.pathname}${nextUrl.search ? `?${nextUrl.searchParams.toString()}` : ""}${nextUrl.hash}`;
  window.history.replaceState({}, "", nextHref);
};

const readActivitiesFromLocalStorage = () => {
  try {
    const raw = localStorage.getItem(LOCAL_STORAGE_KEY);
    if (!raw) return [];
    const parsed = JSON.parse(raw);
    if (!Array.isArray(parsed)) return [];
    return parsed.map(normalizeActivity);
  } catch {
    return [];
  }
};

const saveActivitiesToLocalStorage = (activities) => {
  try {
    localStorage.setItem(LOCAL_STORAGE_KEY, JSON.stringify(activities));
  } catch {
    // Ignore storage quota/access errors and keep in-memory state.
  }
};

const readProjectsFromLocalStorage = () => {
  try {
    const raw = localStorage.getItem(PROJECTS_LOCAL_STORAGE_KEY);
    if (!raw) return [];
    const parsed = JSON.parse(raw);
    if (!Array.isArray(parsed)) return [];
    return parsed.map(normalizeProject);
  } catch {
    return [];
  }
};

const mergeActivities = (primary, secondary) => {
  const merged = [];
  const seen = new Set();

  [primary, secondary].forEach((list) => {
    (list || []).forEach((item) => {
      const normalized = normalizeActivity(item);
      const key = `${normalized.id}::${normalized.project}::${normalized.name}`.toLowerCase();
      if (seen.has(key)) return;
      seen.add(key);
      merged.push(normalized);
    });
  });

  return merged;
};

const mergeProjects = (primary, secondary) => {
  const merged = [];
  const seen = new Set();

  [primary, secondary].forEach((list) => {
    (list || []).forEach((item) => {
      const normalized = normalizeProject(item);
      const key = `${normalized.code}::${normalized.name}`.toLowerCase();
      if (seen.has(key)) return;
      seen.add(key);
      merged.push(normalized);
    });
  });

  return merged;
};

const fetchSourcePayload = async (resource) => {
  if (!DATA_SOURCE_URL) return null;

  const response = await fetch(`${DATA_SOURCE_URL}?resource=${encodeURIComponent(resource)}`, {
    cache: "no-store",
  });
  if (!response.ok) {
    throw new Error(`Unable to load ${resource} data (HTTP ${response.status})`);
  }

  const payload = await response.json();
  if (payload?.ok === false) {
    throw new Error(payload.error || `Unable to load ${resource} data`);
  }
  return payload;
};

const loadActivitiesAndProjectsFromSource = async () => {
  const localActivities = readActivitiesFromLocalStorage();
  const localProjects = readProjectsFromLocalStorage();

  try {
    const [activitiesPayload, projectsPayload] = await Promise.all([
      fetchSourcePayload("activities"),
      fetchSourcePayload("projects"),
    ]);

    const remoteActivities = Array.isArray(activitiesPayload?.activities)
      ? activitiesPayload.activities.map(normalizeActivity)
      : [];
    const mergedActivities = mergeActivities(remoteActivities, localActivities);
    const remoteProjects = Array.isArray(projectsPayload?.projects)
      ? projectsPayload.projects.map(normalizeProject)
      : [];
    const mergedProjects = mergeProjects(remoteProjects, localProjects);

    saveActivitiesToLocalStorage(mergedActivities);

    const computedKpis = {
      completed: 0,
      inProgress: 0,
      notStarted: 0,
      delayed: 0,
    };

    mergedActivities.forEach((activity) => {
      const key = statusTextToKey[activity.status];
      if (key && Object.prototype.hasOwnProperty.call(computedKpis, key)) {
        computedKpis[key] += 1;
      }
    });

    const activityMeta = {
      totalCount: mergedActivities.length,
      kpi: computedKpis,
    };

    return {
      activities: mergedActivities,
      projects: mergedProjects,
      meta: activityMeta,
    };
  } catch (error) {
    console.warn("Falling back to embedded activities data:", error);
    const fallbackActivities = mergeActivities(localActivities, initialActivities);
    saveActivitiesToLocalStorage(fallbackActivities);
    return {
      activities: fallbackActivities,
      projects: mergeProjects(localProjects, initialProjectCatalog),
      meta: window.activitiesMeta || {
        totalCount: fallbackActivities.length,
        kpi: { completed: 0, inProgress: 0, notStarted: 0, delayed: 0 },
      },
    };
  }
};

const hasSelectedProject = () => Boolean(state.selectedProject);

const formatDateRangeLabel = () => {
  const { start, end } = state.dateRange;
  if (!start && !end) return "All Dates";
  if (start && end) return `${toDisplayDate(start)} - ${toDisplayDate(end)}`;
  if (start) return `${toDisplayDate(start)} onwards`;
  return `Until ${toDisplayDate(end)}`;
};

const syncDateFilterLabel = () => {
  if (!activitiesDateFilterLabel) return;
  activitiesDateFilterLabel.textContent = formatDateRangeLabel();
};

const formatProjectDateRangeLabel = () => {
  const { start, end } = state.projectDateRange;
  if (!start && !end) return "All Dates";
  if (start && end) return `${toDisplayDate(start)} - ${toDisplayDate(end)}`;
  if (start) return `${toDisplayDate(start)} onwards`;
  return `Until ${toDisplayDate(end)}`;
};

const syncProjectDateFilterLabel = () => {
  if (!activitiesProjectDateFilterLabel) return;
  activitiesProjectDateFilterLabel.textContent = formatProjectDateRangeLabel();
};

const closeProjectDateRangePanel = () => {
  if (!activitiesProjectDateRangePanel) return;
  activitiesProjectDateRangePanel.classList.add("hidden");
  activitiesProjectDateFilterWrap?.classList.remove("is-open");
  activitiesProjectDateFilterBtn?.setAttribute("aria-expanded", "false");
};

const openProjectDateRangePanel = () => {
  if (!activitiesProjectDateRangePanel) return;
  activitiesProjectDateRangePanel.classList.remove("hidden");
  activitiesProjectDateFilterWrap?.classList.add("is-open");
  activitiesProjectDateFilterBtn?.setAttribute("aria-expanded", "true");
};

const deriveProjectStatus = (projectActivities = []) => {
  if (!projectActivities.length) return "Not Started";
  const statuses = projectActivities.map((item) => item.status);
  if (statuses.includes("Delayed")) return "Delayed";
  if (statuses.every((status) => status === "Completed")) return "Completed";
  if (statuses.includes("In Progress")) return "In Progress";
  return "Not Started";
};

const deriveProjectType = (projectActivities = []) => {
  const firstType = projectActivities.find((item) => item.type && item.type !== "-")?.type;
  return firstType || "General";
};

const buildProjectSummaries = () => {
  const projectMap = new Map(initialProjectCatalog.map((project) => [project.name, { ...project }]));

  state.allActivities.forEach((activity) => {
    if (!activity.project || activity.project === "-") return;
    if (!projectMap.has(activity.project)) {
      projectMap.set(activity.project, {
        name: activity.project,
        code: "",
        location: "",
        type: "General",
        status: "Not Started",
        startDate: null,
        endDate: null,
      });
    }
    const summary = projectMap.get(activity.project);
    if (!summary.activities) summary.activities = [];
    summary.activities.push(activity);
  });

  return Array.from(projectMap.values())
    .map((summary) => {
      const projectActivities = summary.activities || [];
      const starts = projectActivities.map((item) => item.plannedStartDate).filter(Boolean);
      const finishes = projectActivities.map((item) => item.plannedFinishDate).filter(Boolean);
      const activityStartDate = starts.length ? new Date(Math.min(...starts.map((date) => date.getTime()))) : null;
      const activityEndDate = finishes.length ? new Date(Math.max(...finishes.map((date) => date.getTime()))) : null;

      return {
        name: summary.name,
        code: summary.code,
        location: summary.location,
        type: projectActivities.length ? deriveProjectType(projectActivities) : summary.type,
        status: projectActivities.length ? deriveProjectStatus(projectActivities) : summary.status,
        startDate: activityStartDate || summary.startDate,
        endDate: activityEndDate || summary.endDate,
      };
    })
    .sort((a, b) => a.name.localeCompare(b.name));
};

const closeDateRangePanel = () => {
  if (!activitiesDateRangePanel) return;
  activitiesDateRangePanel.classList.add("hidden");
  activitiesDateFilterWrap?.classList.remove("is-open");
  activitiesDateFilterBtn?.setAttribute("aria-expanded", "false");
};

const openDateRangePanel = () => {
  if (!activitiesDateRangePanel) return;
  activitiesDateRangePanel.classList.remove("hidden");
  activitiesDateFilterWrap?.classList.add("is-open");
  activitiesDateFilterBtn?.setAttribute("aria-expanded", "true");
};

const updateKpis = (sourceActivities) => {
  const hasActiveFilters =
    Boolean(activitiesSearchInput?.value.trim()) ||
    (activitiesStatusFilter?.value && activitiesStatusFilter.value !== "All Statuses") ||
    (activitiesTypeFilter?.value && activitiesTypeFilter.value !== "All Activity Types") ||
    Boolean(state.dateRange.start || state.dateRange.end);

  if (!hasActiveFilters && window.activitiesMeta?.kpi) {
    const metaKpis = window.activitiesMeta.kpi;
    if (kpiEls.total) kpiEls.total.textContent = window.activitiesMeta.totalCount ?? sourceActivities.length;
    if (kpiEls.completed) kpiEls.completed.textContent = Number(metaKpis.completed) || 0;
    if (kpiEls.inProgress) kpiEls.inProgress.textContent = Number(metaKpis.inProgress) || 0;
    if (kpiEls.notStarted) kpiEls.notStarted.textContent = Number(metaKpis.notStarted) || 0;
    if (kpiEls.delayed) kpiEls.delayed.textContent = Number(metaKpis.delayed) || 0;
    return;
  }

  const counts = {
    total: sourceActivities.length,
    completed: 0,
    inProgress: 0,
    notStarted: 0,
    delayed: 0,
  };

  sourceActivities.forEach((item) => {
    const key = statusTextToKey[item.status];
    if (key) counts[key] += 1;
  });

  Object.entries(kpiEls).forEach(([key, el]) => {
    if (el) el.textContent = counts[key];
  });
};

const updateSummary = () => {
  if (!activitiesTableSummary) return;

  const totalCount = Number(window.activitiesMeta?.totalCount) || state.allActivities.length;
  const filteredCount = state.filteredActivities.length;

  if (!filteredCount) {
    activitiesTableSummary.textContent = `Showing 0 to 0 of ${totalCount} activities`;
    return;
  }

  const start = (state.currentPage - 1) * PAGE_SIZE + 1;
  const end = Math.min(state.currentPage * PAGE_SIZE, filteredCount);
  activitiesTableSummary.textContent = `Showing ${start} to ${end} of ${filteredCount} activities`;
};

const getSelectableProjects = () =>
  uniqueSorted(state.allActivities.map((row) => row.project)).filter((project) => project !== "-");

const renderProjectPicker = () => {
  if (!activitiesProjectPickerGrid) return;

  const searchValue = state.projectSearch.trim().toLowerCase();
  const selectedType = activitiesProjectTypeFilter?.value || "All Project Types";
  const selectedStatus = activitiesProjectStatusFilter?.value || "All Statuses";
  const selectedStartDate = state.projectDateRange.start;
  const selectedEndDate = state.projectDateRange.end;
  const allProjects = buildProjectSummaries();

  const projects = allProjects.filter((project) => {
    const searchable = `${project.name} ${project.code} ${project.location}`.toLowerCase();
    const textMatch = !searchValue || searchable.includes(searchValue);
    const typeMatch = selectedType === "All Project Types" || project.type === selectedType;
    const statusMatch = selectedStatus === "All Statuses" || project.status === selectedStatus;
    const hasDateFilter = Boolean(selectedStartDate || selectedEndDate);
    const projectAnchorDate = project.startDate || project.endDate;
    const dateMatch =
      !hasDateFilter ||
      (projectAnchorDate &&
        (!selectedStartDate || projectAnchorDate >= selectedStartDate) &&
        (!selectedEndDate || projectAnchorDate <= selectedEndDate));

    return textMatch && typeMatch && statusMatch && dateMatch;
  });

  const hasNoProjects = !allProjects.length;
  if (activitiesProjectSelectionFooter) {
    activitiesProjectSelectionFooter.hidden = hasNoProjects;
  }

  if (hasNoProjects) {
    activitiesProjectPickerGrid.innerHTML = `
      <article class="activities-project-empty-advisory" role="status" aria-live="polite">
        <div class="activities-project-empty-advisory-icon" aria-hidden="true">
          <svg viewBox="0 0 24 24" fill="none"><path d="M3 7a2 2 0 0 1 2-2h5l2 2h7a2 2 0 0 1 2 2v8a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2V7Z"/></svg>
        </div>
        <h4>No Projects</h4>
        <p>There are currently no projects in the system.</p>
        <p>Create a project first to start adding activities.</p>
      </article>
    `;
    return;
  }

  if (!projects.length) {
    activitiesProjectPickerGrid.innerHTML = `<p class="activities-project-picker-empty">No projects found.</p>`;
    return;
  }

  activitiesProjectPickerGrid.innerHTML = projects
    .map(
      (project) => `
        <button type="button" class="activities-project-picker-card" data-project="${encodeURIComponent(project.name)}">
          <span>${escapeHtml(project.code ? `${project.code} · ${project.name}` : project.name)}</span>
          <svg viewBox="0 0 24 24" fill="none" aria-hidden="true"><path d="m9 6 6 6-6 6"/></svg>
        </button>
      `
    )
    .join("");
};

const syncWorkflowState = () => {
  const shouldShowProjectSelection = !hasSelectedProject();
  if (activitiesProjectSelectionFilters) {
    activitiesProjectSelectionFilters.hidden = !shouldShowProjectSelection;
  }
  if (activitiesProjectSelection) {
    activitiesProjectSelection.hidden = !shouldShowProjectSelection;
  }
  if (activitiesViewShell) {
    activitiesViewShell.hidden = shouldShowProjectSelection;
  }
  if (activitiesSelectedProjectName) {
    activitiesSelectedProjectName.textContent = state.selectedProject || "—";
  }
};

const renderPagination = () => {
  if (!activitiesPagination) return;

  if (!hasSelectedProject() || !state.filteredActivities.length) {
    activitiesPagination.innerHTML = "";
    return;
  }

  const totalPages = Math.ceil(state.filteredActivities.length / PAGE_SIZE);
  if (state.currentPage > totalPages) state.currentPage = totalPages;

  const prevDisabled = state.currentPage <= 1;
  const nextDisabled = state.currentPage >= totalPages;

  const pageButtons = Array.from({ length: totalPages }, (_, idx) => {
    const page = idx + 1;
    const isActive = page === state.currentPage;
    return `<button type="button" class="page-number ${isActive ? "active" : ""}" data-page="${page}" ${isActive ? 'aria-current="page"' : ""}>${page}</button>`;
  }).join("");

  activitiesPagination.innerHTML = `
    <button type="button" class="page-arrow" data-dir="prev" aria-label="Previous page" ${prevDisabled ? "disabled" : ""}>←</button>
    ${pageButtons}
    <button type="button" class="page-arrow" data-dir="next" aria-label="Next page" ${nextDisabled ? "disabled" : ""}>→</button>
  `;
};

const renderTable = () => {
  if (!hasSelectedProject()) {
    renderEmptyState("Select a project to view activities.");
    return;
  }

  if (!state.filteredActivities.length) {
    if (!state.allActivities.length) {
      renderEmptyState();
    } else {
      renderEmptyState("No activities found for the selected filters.");
    }
    return;
  }

  const startIdx = (state.currentPage - 1) * PAGE_SIZE;
  const pageActivities = state.filteredActivities.slice(startIdx, startIdx + PAGE_SIZE);
  activitiesTableBody.innerHTML = pageActivities.map(buildActivityRowHtml).join("");
};

const applyFilters = () => {
  const searchValue = activitiesSearchInput?.value.trim().toLowerCase() || "";
  const statusValue = activitiesStatusFilter?.value || "All Statuses";
  const typeValue = activitiesTypeFilter?.value || "All Activity Types";
  const dateStartValue = state.dateRange.start;
  const dateEndValue = state.dateRange.end;

  if (!hasSelectedProject()) {
    state.filteredActivities = [];
  } else {
    state.filteredActivities = state.allActivities.filter((item) => {
    const projectMatch = item.project === state.selectedProject;
    const statusMatch = statusValue === "All Statuses" || item.status === statusValue;
    const typeMatch = typeValue === "All Activity Types" || item.type === typeValue;
    const textMatch =
      !searchValue ||
      `${item.name} ${item.type} ${item.status}`.toLowerCase().includes(searchValue);
    const dateValue = item.plannedStartDate;
    const hasDateFilter = Boolean(dateStartValue || dateEndValue);
    const dateMatch =
      !hasDateFilter ||
      (dateValue &&
        (!dateStartValue || dateValue >= dateStartValue) &&
        (!dateEndValue || dateValue <= dateEndValue));

    return projectMatch && statusMatch && typeMatch && textMatch && dateMatch;
  });
  }

  state.currentPage = 1;
  syncWorkflowState();
  renderPagination();
  renderTable();
  updateKpis(state.filteredActivities);
  updateSummary();
};


const closeAddActivityModal = ({ resetForm = true } = {}) => {
  if (!activityModal) return;
  activityModal.classList.add("hidden");
  document.body.style.overflow = "";
  if (resetForm && activityModalForm) {
    activityModalForm.reset();
    state.editingActivityKey = null;
    if (activityModalFinishDateInput) {
      activityModalFinishDateInput.min = "";
    }
  }
};

const setActivityModalMode = (mode = "add") => {
  const isEdit = mode === "edit";
  if (activityModalTitle) {
    activityModalTitle.textContent = isEdit ? "Edit Activity" : "Add New Activity";
  }
  if (activityModalSubtitle) {
    activityModalSubtitle.textContent = isEdit
      ? "Update the activity details for your selected project."
      : "Create an activity and attach it to your selected project.";
  }
  if (activityModalSubmitBtn) {
    activityModalSubmitBtn.textContent = isEdit ? "Update Activity" : "Save Activity";
  }
};

const updateActivityModalProjectDetails = (projectName) => {
  const projectSummaries = buildProjectSummaries();
  const selectedProjectSummary = projectSummaries.find((project) => project.name === projectName);
  const formattedProject = selectedProjectSummary?.code
    ? `${selectedProjectSummary.code} — ${selectedProjectSummary.name}`
    : projectName || "—";

  if (activityModalProjectInput) {
    activityModalProjectInput.value = projectName || "";
  }
  if (activityModalProjectDisplay) {
    activityModalProjectDisplay.textContent = formattedProject;
  }
};

const openAddActivityModal = () => {
  if (!hasSelectedProject()) {
    window.alert("Please select a project first.");
    activitiesProjectPickerSearch?.focus();
    return;
  }

  if (!activityModal) return;
  state.editingActivityKey = null;
  setActivityModalMode("add");
  updateActivityModalProjectDetails(state.selectedProject);
  activityModal.classList.remove("hidden");
  document.body.style.overflow = "hidden";
  window.setTimeout(() => {
    document.getElementById("activityModalIdInput")?.focus();
  }, 0);
};

const findActivityIndexByKey = (activityKey) =>
  state.allActivities.findIndex((item) => createActivityKey(item) === activityKey);

const openEditActivityModal = (activityKey) => {
  if (!hasSelectedProject()) return;

  const index = findActivityIndexByKey(activityKey);
  if (index < 0) {
    window.alert("Activity not found.");
    return;
  }

  const activity = state.allActivities[index];
  const idInput = document.getElementById("activityModalIdInput");
  const nameInput = document.getElementById("activityModalNameInput");
  const typeInput = document.getElementById("activityModalTypeInput");

  if (!activityModal || !idInput || !nameInput || !typeInput) return;

  setActivityModalMode("edit");
  state.editingActivityKey = activityKey;
  idInput.value = activity.id || "";
  nameInput.value = activity.name || "";
  typeInput.value = activity.type || "";
  updateActivityModalProjectDetails(activity.project || state.selectedProject || "");
  if (activityModalStartDateInput) activityModalStartDateInput.value = toInputDate(activity.plannedStartDate || activity.plannedStart);
  if (activityModalFinishDateInput) {
    activityModalFinishDateInput.value = toInputDate(activity.plannedFinishDate || activity.plannedFinish);
    activityModalFinishDateInput.min = activityModalStartDateInput?.value || "";
  }

  activityModal.classList.remove("hidden");
  document.body.style.overflow = "hidden";
  window.setTimeout(() => {
    idInput.focus();
  }, 0);
};

const handleDeleteActivity = (activityKey) => {
  const index = findActivityIndexByKey(activityKey);
  if (index < 0) {
    window.alert("Activity not found.");
    return;
  }

  const activity = state.allActivities[index];
  const shouldDelete = window.confirm(`Delete activity "${activity.name}"?`);
  if (!shouldDelete) return;

  state.allActivities.splice(index, 1);
  saveActivitiesToLocalStorage(state.allActivities);
  refreshFilterOptions();
  applyFilters();
  window.alert("Activity deleted successfully.");
};

const refreshFilterOptions = () => {
  const projectSummaries = buildProjectSummaries();
  const previousActivityStatus = activitiesStatusFilter?.value || "All Statuses";
  const previousActivityType = activitiesTypeFilter?.value || "All Activity Types";
  const previousProjectType = activitiesProjectTypeFilter?.value || "All Project Types";
  const previousProjectStatus = activitiesProjectStatusFilter?.value || "All Statuses";

  populateSelect(activitiesStatusFilter, uniqueSorted(state.allActivities.map((row) => row.status)), "All Statuses");
  populateSelect(activitiesTypeFilter, uniqueSorted(state.allActivities.map((row) => row.type)), "All Activity Types");
  populateSelect(activitiesProjectTypeFilter, uniqueSorted(projectSummaries.map((row) => row.type)), "All Project Types");
  populateSelect(activitiesProjectStatusFilter, uniqueSorted(projectSummaries.map((row) => row.status)), "All Statuses");

  if (activitiesStatusFilter) {
    activitiesStatusFilter.value = Array.from(activitiesStatusFilter.options).some((option) => option.value === previousActivityStatus)
      ? previousActivityStatus
      : "All Statuses";
  }

  if (activitiesTypeFilter) {
    activitiesTypeFilter.value = Array.from(activitiesTypeFilter.options).some((option) => option.value === previousActivityType)
      ? previousActivityType
      : "All Activity Types";
  }

  if (activitiesProjectTypeFilter) {
    activitiesProjectTypeFilter.value = Array.from(activitiesProjectTypeFilter.options).some((option) => option.value === previousProjectType)
      ? previousProjectType
      : "All Project Types";
  }

  if (activitiesProjectStatusFilter) {
    activitiesProjectStatusFilter.value = Array.from(activitiesProjectStatusFilter.options).some((option) => option.value === previousProjectStatus)
      ? previousProjectStatus
      : "All Statuses";
  }

  if (state.selectedProject) {
    const projectStillExists = projectSummaries.some((project) => project.name === state.selectedProject);
    if (!projectStillExists) {
      state.selectedProject = null;
    }
  }
  renderProjectPicker();
};

const onPaginationClick = (event) => {
  const button = event.target.closest("button");
  if (!button) return;

  const { page, dir } = button.dataset;
  const totalPages = Math.max(1, Math.ceil(state.filteredActivities.length / PAGE_SIZE));

  if (dir === "prev" && state.currentPage > 1) {
    state.currentPage -= 1;
  } else if (dir === "next" && state.currentPage < totalPages) {
    state.currentPage += 1;
  } else if (page) {
    const parsedPage = Number(page);
    if (Number.isInteger(parsedPage) && parsedPage >= 1 && parsedPage <= totalPages) {
      state.currentPage = parsedPage;
    }
  }

  renderPagination();
  renderTable();
  updateSummary();
};

refreshFilterOptions();
syncWorkflowState();

[activitiesSearchInput, activitiesStatusFilter, activitiesTypeFilter]
  .filter(Boolean)
  .forEach((el) => {
    el.addEventListener("input", applyFilters);
    el.addEventListener("change", applyFilters);
  });

if (activitiesPagination) {
  activitiesPagination.addEventListener("click", onPaginationClick);
}

if (activitiesAddButton) {
  activitiesAddButton.addEventListener("click", openAddActivityModal);
}

if (activitiesProjectPickerSearch) {
  activitiesProjectPickerSearch.addEventListener("input", () => {
    state.projectSearch = activitiesProjectPickerSearch.value || "";
    renderProjectPicker();
  });
}

[activitiesProjectTypeFilter, activitiesProjectStatusFilter]
  .filter(Boolean)
  .forEach((el) => el.addEventListener("change", renderProjectPicker));

if (activitiesProjectDateFilterBtn) {
  activitiesProjectDateFilterBtn.addEventListener("click", () => {
    const isOpen = activitiesProjectDateRangePanel && !activitiesProjectDateRangePanel.classList.contains("hidden");
    if (isOpen) {
      closeProjectDateRangePanel();
      return;
    }
    openProjectDateRangePanel();
  });
}

if (activitiesProjectFilterStartDate && activitiesProjectFilterEndDate) {
  activitiesProjectFilterStartDate.addEventListener("change", () => {
    activitiesProjectFilterEndDate.min = activitiesProjectFilterStartDate.value || "";
  });
}

if (activitiesProjectDateApplyBtn) {
  activitiesProjectDateApplyBtn.addEventListener("click", () => {
    const start = parseDateValue(activitiesProjectFilterStartDate?.value);
    const end = parseDateValue(activitiesProjectFilterEndDate?.value);

    if (start && end && start > end) {
      window.alert("End date must be on or after start date.");
      return;
    }

    state.projectDateRange.start = start;
    state.projectDateRange.end = end;
    syncProjectDateFilterLabel();
    closeProjectDateRangePanel();
    renderProjectPicker();
  });
}

if (activitiesProjectDateClearBtn) {
  activitiesProjectDateClearBtn.addEventListener("click", () => {
    state.projectDateRange.start = null;
    state.projectDateRange.end = null;
    if (activitiesProjectFilterStartDate) activitiesProjectFilterStartDate.value = "";
    if (activitiesProjectFilterEndDate) {
      activitiesProjectFilterEndDate.value = "";
      activitiesProjectFilterEndDate.min = "";
    }
    syncProjectDateFilterLabel();
    closeProjectDateRangePanel();
    renderProjectPicker();
  });
}


if (activitiesHowItWorksBtn) {
  activitiesHowItWorksBtn.addEventListener("click", () => {
    window.alert("Use the project filters to narrow by project type, status, or planned date, then select a project card to view activities.");
  });
}

if (activitiesProjectPickerGrid) {
  activitiesProjectPickerGrid.addEventListener("click", (event) => {
    const button = event.target.closest("[data-project]");
    if (!button) return;
    const encodedProject = button.dataset.project || "";
    const project = encodedProject ? decodeURIComponent(encodedProject) : "All Projects";
    state.selectedProject = project;
    syncWorkflowState();
    closeProjectDateRangePanel();
    updateActivitiesUrlParams({ project, keepAddedFlag: true });
    applyFilters();
    window.scrollTo({ top: 0, behavior: "smooth" });
  });
}

const handleBackToProjects = () => {
  state.selectedProject = null;
  state.currentPage = 1;
  if (activitiesSearchInput) activitiesSearchInput.value = "";
  if (activitiesStatusFilter) activitiesStatusFilter.value = "All Statuses";
  if (activitiesTypeFilter) activitiesTypeFilter.value = "All Activity Types";
  state.dateRange.start = null;
  state.dateRange.end = null;
  if (activitiesFilterStartDate) activitiesFilterStartDate.value = "";
  if (activitiesFilterEndDate) {
    activitiesFilterEndDate.value = "";
    activitiesFilterEndDate.min = "";
  }
  syncDateFilterLabel();
  updateActivitiesUrlParams({ project: null });
  applyFilters();
  activitiesProjectSelection?.scrollIntoView({ behavior: "smooth", block: "start" });
};

if (activitiesBackToProjectsBtn) {
  activitiesBackToProjectsBtn.addEventListener("click", handleBackToProjects);
}

activitiesTableBody.addEventListener("click", (event) => {
  const addButton = event.target.closest("#activitiesAddButtonEmpty");
  if (addButton) {
    openAddActivityModal();
    return;
  }

  const actionButton = event.target.closest("[data-action][data-activity-key]");
  if (!actionButton) return;

  const action = actionButton.dataset.action;
  const activityKey = actionButton.dataset.activityKey || "";
  if (!activityKey) return;

  if (action === "edit") {
    openEditActivityModal(activityKey);
    return;
  }

  if (action === "delete") {
    handleDeleteActivity(activityKey);
  }
});


document.addEventListener("keydown", (event) => {
  if (event.key === "Escape" && activityModal && !activityModal.classList.contains("hidden")) {
    closeAddActivityModal();
    return;
  }

  if (event.key === "Escape" && state.openActivityMenuKey) {
    closeActivityActionMenus();
    return;
  }

  if (event.key === "Escape" && activitiesDateRangePanel && !activitiesDateRangePanel.classList.contains("hidden")) {
    closeDateRangePanel();
    return;
  }

  if (event.key === "Escape" && activitiesProjectDateRangePanel && !activitiesProjectDateRangePanel.classList.contains("hidden")) {
    closeProjectDateRangePanel();
    return;
  }

});

if (activitiesDateFilterBtn) {
  activitiesDateFilterBtn.addEventListener("click", () => {
    const isOpen = activitiesDateRangePanel && !activitiesDateRangePanel.classList.contains("hidden");
    if (isOpen) {
      closeDateRangePanel();
      return;
    }
    openDateRangePanel();
  });
}

if (activitiesFilterStartDate && activitiesFilterEndDate) {
  activitiesFilterStartDate.addEventListener("change", () => {
    activitiesFilterEndDate.min = activitiesFilterStartDate.value || "";
  });
}

if (activitiesDateApplyBtn) {
  activitiesDateApplyBtn.addEventListener("click", () => {
    const start = parseDateValue(activitiesFilterStartDate?.value);
    const end = parseDateValue(activitiesFilterEndDate?.value);

    if (start && end && start > end) {
      window.alert("End date must be on or after start date.");
      return;
    }

    state.dateRange.start = start;
    state.dateRange.end = end;
    syncDateFilterLabel();
    closeDateRangePanel();
    applyFilters();
  });
}

if (activitiesDateClearBtn) {
  activitiesDateClearBtn.addEventListener("click", () => {
    state.dateRange.start = null;
    state.dateRange.end = null;
    if (activitiesFilterStartDate) activitiesFilterStartDate.value = "";
    if (activitiesFilterEndDate) {
      activitiesFilterEndDate.value = "";
      activitiesFilterEndDate.min = "";
    }
    syncDateFilterLabel();
    closeDateRangePanel();
    applyFilters();
  });
}

document.addEventListener("click", (event) => {
  if (state.openActivityMenuKey && !event.target.closest(".actions-col")) {
    closeActivityActionMenus();
  }

  if (
    activitiesDateFilterWrap &&
    activitiesDateRangePanel &&
    !activitiesDateRangePanel.classList.contains("hidden") &&
    !activitiesDateFilterWrap.contains(event.target)
  ) {
    closeDateRangePanel();
  }

  if (
    activitiesProjectDateFilterWrap &&
    activitiesProjectDateRangePanel &&
    !activitiesProjectDateRangePanel.classList.contains("hidden") &&
    !activitiesProjectDateFilterWrap.contains(event.target)
  ) {
    closeProjectDateRangePanel();
  }
});

if (activityModalStartDateInput && activityModalFinishDateInput) {
  activityModalStartDateInput.addEventListener("change", () => {
    activityModalFinishDateInput.min = activityModalStartDateInput.value || "";
  });
}

if (activityModalCloseBtn) {
  activityModalCloseBtn.addEventListener("click", () => closeAddActivityModal());
}

if (activityModalCancelBtn) {
  activityModalCancelBtn.addEventListener("click", () => closeAddActivityModal());
}

if (activityModalBackdrop) {
  activityModalBackdrop.addEventListener("click", () => closeAddActivityModal());
}

if (activityModalForm) {
  activityModalForm.addEventListener("submit", (event) => {
    event.preventDefault();
    if (!hasSelectedProject()) {
      window.alert("Please select a project first.");
      closeAddActivityModal();
      return;
    }

    const formData = new FormData(activityModalForm);
    const plannedStart = String(formData.get("plannedStart") || "");
    const plannedFinish = String(formData.get("plannedFinish") || "");
    if (plannedStart && plannedFinish && new Date(plannedStart) > new Date(plannedFinish)) {
      window.alert("Planned finish date must be on or after planned start date.");
      return;
    }

    const nextActivity = normalizeActivity({
      id: String(formData.get("activityId") || "").trim(),
      name: String(formData.get("activityName") || "").trim(),
      project: state.selectedProject,
      type: String(formData.get("activityType") || "").trim(),
      status: "Not Started",
      plannedStart,
      plannedFinish,
      progress: 0,
      costStatus: "On Budget",
    });

    const isEditing = Boolean(state.editingActivityKey);
    if (isEditing) {
      const index = findActivityIndexByKey(state.editingActivityKey);
      if (index < 0) {
        window.alert("Unable to update. Activity no longer exists.");
        closeAddActivityModal();
        return;
      }
      state.allActivities[index] = nextActivity;
    } else {
      state.allActivities.unshift(nextActivity);
    }

    saveActivitiesToLocalStorage(state.allActivities);
    refreshFilterOptions();
    applyFilters();
    closeAddActivityModal();
    window.alert(isEditing ? "Activity updated successfully." : "Activity saved successfully.");
  });
}



const hydrateSelectedProjectFromUrl = () => {
  if (state.didHydrateProjectFromUrl) return;
  state.didHydrateProjectFromUrl = true;

  const query = new URLSearchParams(window.location.search);
  const projectFromUrl = query.get("project");
  if (!projectFromUrl) return;

  const projectSummaries = buildProjectSummaries();
  const matchedProject = projectSummaries.find((project) => project.name === projectFromUrl);
  if (matchedProject) {
    state.selectedProject = matchedProject.name;
    updateActivitiesUrlParams({ project: matchedProject.name, keepAddedFlag: true });
  }
};

const handleAddedActivityNotice = () => {
  const query = new URLSearchParams(window.location.search);
  if (query.get("added") !== "1") return;
  window.alert("Activity saved successfully.");
  updateActivitiesUrlParams({ keepAddedFlag: false });
};

const bootstrapActivitiesPage = async () => {
  const source = await loadActivitiesAndProjectsFromSource();
  const nextSignature = JSON.stringify({
    activities: source.activities,
    projects: source.projects,
    meta: source.meta,
  });

  if (nextSignature === lastActivitiesSignature) {
    return;
  }

  lastActivitiesSignature = nextSignature;

  state.allActivities = source.activities;
  state.filteredActivities = source.activities;
  window.activitiesMeta = source.meta;

  initialProjectCatalog.length = 0;
  source.projects.forEach((project) => {
    initialProjectCatalog.push(project);
  });

  syncDateFilterLabel();
  syncProjectDateFilterLabel();
  refreshFilterOptions();
  renderProjectPicker();
  hydrateSelectedProjectFromUrl();
  handleAddedActivityNotice();
  applyFilters();
};

const refreshActivitiesIfVisible = async ({ force = false } = {}) => {
  if (isActivitiesSyncInFlight) return;
  if (!force && document.visibilityState === "hidden") return;

  isActivitiesSyncInFlight = true;
  try {
    await bootstrapActivitiesPage();
  } finally {
    isActivitiesSyncInFlight = false;
  }
};

const setupActivitiesRealtimeSync = () => {
  if (activitiesRefreshTimer) {
    clearInterval(activitiesRefreshTimer);
  }

  activitiesRefreshTimer = setInterval(() => {
    refreshActivitiesIfVisible();
  }, ACTIVITIES_REFRESH_INTERVAL_MS);

  window.addEventListener("focus", () => refreshActivitiesIfVisible({ force: true }));
  window.addEventListener("online", () => refreshActivitiesIfVisible({ force: true }));
  document.addEventListener("visibilitychange", () => {
    if (document.visibilityState === "visible") {
      refreshActivitiesIfVisible({ force: true });
    }
  });
};

refreshActivitiesIfVisible({ force: true });
setupActivitiesRealtimeSync();
