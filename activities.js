const activitiesTableBody = document.getElementById("activitiesTableBody");
const activitiesSearchInput = document.getElementById("activitiesSearchInput");
const activitiesProjectFilter = document.getElementById("activitiesProjectFilter");
const activitiesStatusFilter = document.getElementById("activitiesStatusFilter");
const activitiesTypeFilter = document.getElementById("activitiesTypeFilter");
const activitiesTableSummary = document.getElementById("activitiesTableSummary");
const activitiesAddButton = document.getElementById("activitiesAddButton");
const activityMeta = window.activitiesMeta || {};

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

const uniqueSorted = (values) =>
  Array.from(new Set(values.filter(Boolean))).sort((a, b) => a.localeCompare(b));

const populateSelect = (selectEl, values, defaultLabel) => {
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
  const status = activity.status || "Not Started";
  const progress = activity.progress ?? (status === "Completed" ? 100 : 0);

  return {
    name: activity.name || "Untitled Activity",
    project: activity.project || "-",
    type: activity.type || "-",
    status,
    plannedStart: toDisplayDate(activity.plannedStart),
    plannedFinish: toDisplayDate(activity.plannedFinish),
    progress: toPercent(progress),
    costStatus: activity.costStatus || "On Budget",
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

const renderEmptyState = () => {
  activitiesTableBody.innerHTML = `
    <tr class="activities-empty-row">
      <td colspan="9">No activities yet. Connect your backend or load activities data to display records.</td>
    </tr>
  `;
};

const initialActivities = Array.isArray(window.activitiesData)
  ? window.activitiesData.map(normalizeActivity)
  : [];

if (initialActivities.length) {
  activitiesTableBody.innerHTML = initialActivities.map(buildActivityRowHtml).join("");
} else {
  renderEmptyState();
}

const activityRows = Array.from(activitiesTableBody.querySelectorAll("tr"))
  .filter((row) => !row.classList.contains("activities-empty-row"))
  .map((row) => {
    const cells = row.querySelectorAll("td");
    const status = cells[3]?.textContent.trim() || "";

    return {
      row,
      searchableText: row.textContent.toLowerCase(),
      project: cells[1]?.textContent.trim() || "",
      type: cells[2]?.textContent.trim() || "",
      status,
    };
  });

const updateKpis = (visibleRows) => {
  if (activityMeta.kpi) {
    const total = Number(activityMeta.totalCount);
    if (kpiEls.total) kpiEls.total.textContent = Number.isFinite(total) ? total : visibleRows.length;
    if (kpiEls.completed) kpiEls.completed.textContent = activityMeta.kpi.completed ?? 0;
    if (kpiEls.inProgress) kpiEls.inProgress.textContent = activityMeta.kpi.inProgress ?? 0;
    if (kpiEls.notStarted) kpiEls.notStarted.textContent = activityMeta.kpi.notStarted ?? 0;
    if (kpiEls.delayed) kpiEls.delayed.textContent = activityMeta.kpi.delayed ?? 0;
    return;
  }

  const counts = {
    total: visibleRows.length,
    completed: 0,
    inProgress: 0,
    notStarted: 0,
    delayed: 0,
  };

  visibleRows.forEach((item) => {
    const key = statusTextToKey[item.status];
    if (key) counts[key] += 1;
  });

  Object.entries(kpiEls).forEach(([key, el]) => {
    if (el) el.textContent = counts[key];
  });
};

const updateSummary = (visibleCount) => {
  const configuredTotal = Number(activityMeta.totalCount);
  const totalCount = Number.isFinite(configuredTotal) ? configuredTotal : activityRows.length;
  if (!visibleCount) {
    activitiesTableSummary.textContent = `Showing 0 of ${totalCount} activities`;
    return;
  }
  activitiesTableSummary.textContent = `Showing 1 to ${visibleCount} of ${totalCount} activities`;
};

const applyFilters = () => {
  if (!activityRows.length) {
    updateKpis([]);
    updateSummary(0);
    return;
  }

  const searchValue = activitiesSearchInput.value.trim().toLowerCase();
  const projectValue = activitiesProjectFilter.value;
  const statusValue = activitiesStatusFilter.value;
  const typeValue = activitiesTypeFilter.value;

  const visibleRows = activityRows.filter((item) => {
    const projectMatch = projectValue === "All Projects" || item.project === projectValue;
    const statusMatch = statusValue === "All Statuses" || item.status === statusValue;
    const typeMatch = typeValue === "All Activity Types" || item.type === typeValue;
    const textMatch = !searchValue || item.searchableText.includes(searchValue);

    const shouldShow = projectMatch && statusMatch && typeMatch && textMatch;
    item.row.hidden = !shouldShow;
    return shouldShow;
  });

  updateKpis(visibleRows);
  updateSummary(visibleRows.length);
};

populateSelect(activitiesProjectFilter, uniqueSorted(activityRows.map((row) => row.project)), "All Projects");
populateSelect(activitiesStatusFilter, uniqueSorted(activityRows.map((row) => row.status)), "All Statuses");
populateSelect(activitiesTypeFilter, uniqueSorted(activityRows.map((row) => row.type)), "All Activity Types");

[activitiesSearchInput, activitiesProjectFilter, activitiesStatusFilter, activitiesTypeFilter].forEach((el) => {
  el.addEventListener("input", applyFilters);
  el.addEventListener("change", applyFilters);
});

activitiesAddButton.addEventListener("click", () => {
  window.alert("Add Activity form is not connected yet.");
});

applyFilters();
