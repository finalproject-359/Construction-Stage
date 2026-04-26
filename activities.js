const activitiesTableBody = document.getElementById("activitiesTableBody");
const activitiesSearchInput = document.getElementById("activitiesSearchInput");
const activitiesProjectFilter = document.getElementById("activitiesProjectFilter");
const activitiesStatusFilter = document.getElementById("activitiesStatusFilter");
const activitiesTypeFilter = document.getElementById("activitiesTypeFilter");
const activitiesTableSummary = document.getElementById("activitiesTableSummary");
const activitiesAddButton = document.getElementById("activitiesAddButton");

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

const activityRows = Array.from(activitiesTableBody.querySelectorAll("tr")).map((row) => {
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

const uniqueSorted = (values) => Array.from(new Set(values.filter(Boolean))).sort((a, b) => a.localeCompare(b));

const populateSelect = (selectEl, values, defaultLabel) => {
  selectEl.innerHTML = `<option>${defaultLabel}</option>`;
  values.forEach((value) => {
    const option = document.createElement("option");
    option.value = value;
    option.textContent = value;
    selectEl.append(option);
  });
};

const updateKpis = (visibleRows) => {
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
  const totalCount = activityRows.length;
  if (!visibleCount) {
    activitiesTableSummary.textContent = `Showing 0 of ${totalCount} activities`;
    return;
  }
  activitiesTableSummary.textContent = `Showing 1 to ${visibleCount} of ${totalCount} activities`;
};

const applyFilters = () => {
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
