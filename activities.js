const activitiesTableBody = document.getElementById("activitiesTableBody");
const activitiesSearchInput = document.getElementById("activitiesSearchInput");
const activitiesProjectFilter = document.getElementById("activitiesProjectFilter");
const activitiesStatusFilter = document.getElementById("activitiesStatusFilter");
const activitiesTypeFilter = document.getElementById("activitiesTypeFilter");
const activitiesTableSummary = document.getElementById("activitiesTableSummary");
const activitiesAddButton = document.getElementById("activitiesAddButton");
const activitiesPagination = document.querySelector(".activities-pagination");

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
  const name = getValueByAliases(activity, ["name", "activity", "activityName", "activity_name"]);
  const project = getValueByAliases(activity, ["project", "projectName", "project_name"]);
  const type = getValueByAliases(activity, ["type", "activityType", "activity_type"]);
  const status = getValueByAliases(activity, ["status"]) || "Not Started";
  const plannedStartRaw = getValueByAliases(activity, ["plannedStart", "planned_start", "startDate", "plannedStartDate"]);
  const plannedFinishRaw = getValueByAliases(activity, ["plannedFinish", "planned_finish", "finishDate", "plannedFinishDate"]);
  const progressRaw = getValueByAliases(activity, ["progress", "percentComplete", "percent_complete"]);
  const costStatus = getValueByAliases(activity, ["costStatus", "cost_status", "budgetStatus"]) || "On Budget";
  const progress = progressRaw ?? (status === "Completed" ? 100 : 0);

  return {
    name: name || "Untitled Activity",
    project: project || "-",
    type: type || "-",
    status,
    plannedStart: toDisplayDate(plannedStartRaw),
    plannedFinish: toDisplayDate(plannedFinishRaw),
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
      <td colspan="9">${message === "Get started by adding your first activity" ? EMPTY_STATE_HTML : escapeHtml(message)}</td>
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

const renderPagination = () => {
  if (!activitiesPagination) return;

  const totalPages = Math.max(1, Math.ceil(state.filteredActivities.length / PAGE_SIZE));
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

  state.filteredActivities = state.allActivities.filter((item) => {
    const projectMatch = projectValue === "All Projects" || item.project === projectValue;
    const statusMatch = statusValue === "All Statuses" || item.status === statusValue;
    const typeMatch = typeValue === "All Activity Types" || item.type === typeValue;
    const textMatch =
      !searchValue ||
      `${item.name} ${item.project} ${item.type} ${item.status} ${item.costStatus}`.toLowerCase().includes(searchValue);

    return projectMatch && statusMatch && typeMatch && textMatch;
  });

  state.currentPage = 1;
  renderPagination();
  renderTable();
  updateKpis(state.filteredActivities);
  updateSummary();
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

populateSelect(activitiesProjectFilter, uniqueSorted(state.allActivities.map((row) => row.project)), "All Projects");
populateSelect(activitiesStatusFilter, uniqueSorted(state.allActivities.map((row) => row.status)), "All Statuses");
populateSelect(activitiesTypeFilter, uniqueSorted(state.allActivities.map((row) => row.type)), "All Activity Types");

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
  activitiesAddButton.addEventListener("click", () => {
    window.alert("Add Activity form is not connected yet.");
  });
}

activitiesTableBody.addEventListener("click", (event) => {
  const button = event.target.closest("#activitiesAddButtonEmpty");
  if (!button) return;
  window.alert("Add Activity form is not connected yet.");
});

renderPagination();
renderTable();
updateKpis(state.filteredActivities);
updateSummary();
