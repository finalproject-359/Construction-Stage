const activitiesTableBody = document.getElementById("activitiesTableBody");
const activitiesSearchInput = document.getElementById("activitiesSearchInput");
const activitiesProjectFilter = document.getElementById("activitiesProjectFilter");
const activitiesStatusFilter = document.getElementById("activitiesStatusFilter");
const activitiesTypeFilter = document.getElementById("activitiesTypeFilter");
const activitiesDateFilterWrap = document.querySelector(".activities-date-filter");
const activitiesDateFilterBtn = document.getElementById("activitiesDateFilterBtn");
const activitiesDateFilterLabel = document.getElementById("activitiesDateFilterLabel");
const activitiesDateRangePanel = document.getElementById("activitiesDateRangePanel");
const activitiesFilterStartDate = document.getElementById("activitiesFilterStartDate");
const activitiesFilterEndDate = document.getElementById("activitiesFilterEndDate");
const activitiesDateClearBtn = document.getElementById("activitiesDateClearBtn");
const activitiesDateApplyBtn = document.getElementById("activitiesDateApplyBtn");
const activitiesTableSummary = document.getElementById("activitiesTableSummary");
const activitiesProjectSelection = document.getElementById("activitiesProjectSelection");
const activitiesViewShell = document.getElementById("activitiesViewShell");
const activitiesProjectPickerSearch = document.getElementById("activitiesProjectPickerSearch");
const activitiesProjectPickerGrid = document.getElementById("activitiesProjectPickerGrid");
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
const activitiesProjectFiltersReset = document.getElementById("activitiesProjectFiltersReset");
const activitiesHowItWorksBtn = document.getElementById("activitiesHowItWorksBtn");
const activitiesAddButton = document.getElementById("activitiesAddButton");
const activitiesPagination = document.querySelector(".activities-pagination");
const activityModal = document.getElementById("activityModal");
const activityModalBackdrop = document.getElementById("activityModalBackdrop");
const activityModalClose = document.getElementById("activityModalClose");
const activityForm = document.getElementById("activityForm");
const activityFormCancel = document.getElementById("activityFormCancel");
const activityStartDateInput = document.getElementById("activityStartDateInput");
const activityFinishDateInput = document.getElementById("activityFinishDateInput");
const activityProjectInput = document.getElementById("activityProjectInput");

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

const BADGE_CLASS_BY_COST = {
  "Under Budget": "badge-under",
  "On Budget": "badge-on",
  "Over Budget": "badge-over",
};

const PROGRESS_CLASS_BY_STATUS = {
  Completed: "progress-completed",
  Delayed: "progress-delayed",
};

const PAGE_SIZE = 8;

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

const toDisplayDate = (value) => {
  if (!value) return "-";
  if (value instanceof Date && !Number.isNaN(value.getTime())) {
    return value.toLocaleDateString("en-US", { month: "short", day: "numeric", year: "numeric" });
  }

  const parsed = new Date(value);
  if (Number.isNaN(parsed.getTime())) return String(value);
  return parsed.toLocaleDateString("en-US", { month: "short", day: "numeric", year: "numeric" });
};

const parseDateValue = (value) => {
  if (!value) return null;
  const parsed = value instanceof Date ? value : new Date(value);
  if (Number.isNaN(parsed.getTime())) return null;
  parsed.setHours(0, 0, 0, 0);
  return parsed;
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
  const project = getValueByAliases(activity, ["project", "projectName", "project_name"]);
  let type = getValueByAliases(activity, ["type", "activityType", "activity_type"]);
  let status = getValueByAliases(activity, ["status"]);
  let plannedStartRaw = getValueByAliases(activity, ["plannedStart", "planned_start", "startDate", "plannedStartDate"]);
  let plannedFinishRaw = getValueByAliases(activity, ["plannedFinish", "planned_finish", "finishDate", "plannedFinishDate"]);
  let progressRaw = getValueByAliases(activity, ["progress", "percentComplete", "percent_complete"]);
  let costStatus = getValueByAliases(activity, ["costStatus", "cost_status", "budgetStatus"]);

  const isDateLike = (value) => Boolean(parseDateValue(value));
  const isKnownStatus = (value) => Object.prototype.hasOwnProperty.call(BADGE_CLASS_BY_STATUS, String(value || "").trim());
  const isNumericLike = (value) => /^-?\d+(\.\d+)?%?$/.test(String(value || "").trim());
  const isKnownCostStatus = (value) => Object.prototype.hasOwnProperty.call(BADGE_CLASS_BY_COST, String(value || "").trim());
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

const buildActivityRowHtml = (activity) => {
  const progressClass = PROGRESS_CLASS_BY_STATUS[activity.status] || "";
  const statusClass = BADGE_CLASS_BY_STATUS[activity.status] || "badge-on";
  const costClass = BADGE_CLASS_BY_COST[activity.costStatus] || "badge-on";

  return `
    <tr>
      <td>${escapeHtml(activity.id)}</td>
      <td>${escapeHtml(activity.name)}</td>
      <td>${escapeHtml(activity.project)}</td>
      <td>${escapeHtml(activity.type)}</td>
      <td><span class="badge ${statusClass}">${escapeHtml(activity.status)}</span></td>
      <td>${escapeHtml(activity.plannedStart)}</td>
      <td>${escapeHtml(activity.plannedFinish)}</td>
      <td>
        <div class="progress-cell"><div class="progress-track"><div class="progress-fill ${progressClass}" style="width:${activity.progress}%"></div></div><span>${activity.progress}%</span></div>
      </td>
      <td><span class="badge ${costClass}">${escapeHtml(activity.costStatus)}</span></td>
      <td class="actions-col">⋮</td>
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
      <td colspan="10">${message === "Get started by adding your first activity" ? EMPTY_STATE_HTML : escapeHtml(message)}</td>
    </tr>
  `;
};

const initialActivities = Array.isArray(window.activitiesData)
  ? window.activitiesData.map(normalizeActivity)
  : [];

const state = {
  allActivities: initialActivities,
  filteredActivities: initialActivities,
  currentPage: 1,
  selectedProject: "All Projects",
  projectSearch: "",
  projectDateRange: {
    start: null,
    end: null,
  },
  dateRange: {
    start: null,
    end: null,
  },
};

const hasSelectedProject = () => state.selectedProject && state.selectedProject !== "All Projects";

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
  const projectMap = new Map();

  state.allActivities.forEach((activity) => {
    if (!activity.project || activity.project === "-") return;
    if (!projectMap.has(activity.project)) {
      projectMap.set(activity.project, []);
    }
    projectMap.get(activity.project).push(activity);
  });

  return Array.from(projectMap.entries())
    .map(([name, projectActivities]) => {
      const starts = projectActivities.map((item) => item.plannedStartDate).filter(Boolean);
      const finishes = projectActivities.map((item) => item.plannedFinishDate).filter(Boolean);
      const startDate = starts.length ? new Date(Math.min(...starts.map((date) => date.getTime()))) : null;
      const endDate = finishes.length ? new Date(Math.max(...finishes.map((date) => date.getTime()))) : null;

      return {
        name,
        type: deriveProjectType(projectActivities),
        status: deriveProjectStatus(projectActivities),
        startDate,
        endDate,
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
    (activitiesProjectFilter?.value && activitiesProjectFilter.value !== "All Projects") ||
    (activitiesStatusFilter?.value && activitiesStatusFilter.value !== "All Statuses") ||
    (activitiesTypeFilter?.value && activitiesTypeFilter.value !== "All Activity Types");

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

  const projects = buildProjectSummaries().filter((project) => {
    const textMatch = !searchValue || project.name.toLowerCase().includes(searchValue);
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

  if (!projects.length) {
    activitiesProjectPickerGrid.innerHTML = `<p class="activities-project-picker-empty">No projects found.</p>`;
    return;
  }

  activitiesProjectPickerGrid.innerHTML = projects
    .map(
      (project) => `
        <button type="button" class="activities-project-picker-card" data-project="${escapeHtml(project.name)}">
          <span>${escapeHtml(project.name)}</span>
          <svg viewBox="0 0 24 24" fill="none" aria-hidden="true"><path d="m9 6 6 6-6 6"/></svg>
        </button>
      `
    )
    .join("");
};

const syncWorkflowState = () => {
  const shouldShowProjectSelection = !hasSelectedProject();
  if (activitiesProjectSelection) {
    activitiesProjectSelection.hidden = !shouldShowProjectSelection;
  }
  if (activitiesViewShell) {
    activitiesViewShell.hidden = shouldShowProjectSelection;
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
  const projectValue = activitiesProjectFilter?.value || "All Projects";
  const statusValue = activitiesStatusFilter?.value || "All Statuses";
  const typeValue = activitiesTypeFilter?.value || "All Activity Types";
  const dateStartValue = state.dateRange.start;
  const dateEndValue = state.dateRange.end;
  state.selectedProject = projectValue;

  if (!hasSelectedProject()) {
    state.filteredActivities = [];
  } else {
    state.filteredActivities = state.allActivities.filter((item) => {
    const projectMatch = projectValue === "All Projects" || item.project === projectValue;
    const statusMatch = statusValue === "All Statuses" || item.status === statusValue;
    const typeMatch = typeValue === "All Activity Types" || item.type === typeValue;
    const textMatch =
      !searchValue ||
      `${item.name} ${item.project} ${item.type} ${item.status} ${item.costStatus}`.toLowerCase().includes(searchValue);
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


const resetFilters = () => {
  if (activitiesSearchInput) activitiesSearchInput.value = "";
  if (activitiesProjectFilter) activitiesProjectFilter.value = "All Projects";
  if (activitiesStatusFilter) activitiesStatusFilter.value = "All Statuses";
  if (activitiesTypeFilter) activitiesTypeFilter.value = "All Activity Types";
  state.dateRange.start = null;
  state.dateRange.end = null;
  state.projectDateRange.start = null;
  state.projectDateRange.end = null;
  if (activitiesProjectTypeFilter) activitiesProjectTypeFilter.value = "All Project Types";
  if (activitiesProjectStatusFilter) activitiesProjectStatusFilter.value = "All Statuses";
  if (activitiesFilterStartDate) activitiesFilterStartDate.value = "";
  if (activitiesFilterEndDate) activitiesFilterEndDate.value = "";
  if (activitiesProjectFilterStartDate) activitiesProjectFilterStartDate.value = "";
  if (activitiesProjectFilterEndDate) {
    activitiesProjectFilterEndDate.value = "";
    activitiesProjectFilterEndDate.min = "";
  }
  syncDateFilterLabel();
  syncProjectDateFilterLabel();
  renderProjectPicker();
  applyFilters();
};

const openActivityModal = () => {
  if (!activityModal || !activityForm) return;
  if (!hasSelectedProject()) {
    window.alert("Please select a project first.");
    activitiesProjectFilter?.focus();
    return;
  }

  activityModal.classList.remove("hidden");
  document.body.style.overflow = "hidden";
  activityForm.reset();
  if (activityProjectInput) {
    activityProjectInput.value = state.selectedProject;
    activityProjectInput.readOnly = true;
  }
};

const closeActivityModal = () => {
  if (!activityModal) return;
  activityModal.classList.add("hidden");
  document.body.style.overflow = "";
};

const refreshFilterOptions = () => {
  const previousProjectSelection = state.selectedProject;
  const projectSummaries = buildProjectSummaries();
  populateSelect(activitiesProjectFilter, uniqueSorted(projectSummaries.map((row) => row.name)), "All Projects");
  populateSelect(activitiesStatusFilter, uniqueSorted(state.allActivities.map((row) => row.status)), "All Statuses");
  populateSelect(activitiesTypeFilter, uniqueSorted(state.allActivities.map((row) => row.type)), "All Activity Types");
  populateSelect(activitiesProjectTypeFilter, uniqueSorted(projectSummaries.map((row) => row.type)), "All Project Types");
  populateSelect(activitiesProjectStatusFilter, uniqueSorted(projectSummaries.map((row) => row.status)), "All Statuses");
  if (activitiesProjectFilter && previousProjectSelection !== "All Projects") {
    const stillExists = Array.from(activitiesProjectFilter.options).some(
      (option) => option.value === previousProjectSelection
    );
    activitiesProjectFilter.value = stillExists ? previousProjectSelection : "All Projects";
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

[activitiesSearchInput, activitiesProjectFilter, activitiesStatusFilter, activitiesTypeFilter]
  .filter(Boolean)
  .forEach((el) => {
    el.addEventListener("input", applyFilters);
    el.addEventListener("change", applyFilters);
  });

if (activitiesPagination) {
  activitiesPagination.addEventListener("click", onPaginationClick);
}

if (activitiesAddButton) {
  activitiesAddButton.addEventListener("click", openActivityModal);
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

if (activitiesProjectFiltersReset) {
  activitiesProjectFiltersReset.addEventListener("click", resetFilters);
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
    const project = button.dataset.project || "All Projects";
    if (activitiesProjectFilter) {
      activitiesProjectFilter.value = project;
    }
    state.selectedProject = project;
    applyFilters();
  });
}

activitiesTableBody.addEventListener("click", (event) => {
  const button = event.target.closest("#activitiesAddButtonEmpty");
  if (!button) return;
  openActivityModal();
});

[activityModalBackdrop, activityModalClose, activityFormCancel]
  .filter(Boolean)
  .forEach((el) => el.addEventListener("click", closeActivityModal));

document.addEventListener("keydown", (event) => {
  if (event.key === "Escape" && activitiesDateRangePanel && !activitiesDateRangePanel.classList.contains("hidden")) {
    closeDateRangePanel();
    return;
  }

  if (event.key === "Escape" && activitiesProjectDateRangePanel && !activitiesProjectDateRangePanel.classList.contains("hidden")) {
    closeProjectDateRangePanel();
    return;
  }

  if (event.key === "Escape" && activityModal && !activityModal.classList.contains("hidden")) {
    closeActivityModal();
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

if (activityForm) {
  activityForm.addEventListener("submit", (event) => {
    event.preventDefault();
    if (!hasSelectedProject()) {
      window.alert("Please select a project first.");
      activitiesProjectFilter?.focus();
      return;
    }

    const formData = new FormData(activityForm);
    const plannedStart = formData.get("plannedStart");
    const plannedFinish = formData.get("plannedFinish");
    if (plannedStart && plannedFinish && new Date(plannedStart) > new Date(plannedFinish)) {
      window.alert("Planned finish date must be on or after planned start date.");
      return;
    }

    const newActivity = normalizeActivity({
      activityId: formData.get("activityId"),
      activityName: formData.get("activityName"),
      project: state.selectedProject,
      activityType: formData.get("activityType"),
      status: "Not Started",
      plannedStart,
      plannedFinish,
      progress: 0,
      costStatus: "On Budget",
    });

    state.allActivities.unshift(newActivity);
    refreshFilterOptions();
    applyFilters();
    closeActivityModal();
  });
}

if (activityStartDateInput && activityFinishDateInput) {
  activityStartDateInput.addEventListener("change", () => {
    activityFinishDateInput.min = activityStartDateInput.value;
  });
}

syncDateFilterLabel();
syncProjectDateFilterLabel();
renderProjectPicker();
applyFilters();
