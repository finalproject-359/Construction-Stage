const projectModal = document.getElementById("projectModal");
const projectForm = document.getElementById("projectForm");
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

const kpiTotalProjects = document.getElementById("kpiTotalProjects");
const kpiActiveProjects = document.getElementById("kpiActiveProjects");
const kpiCompletedProjects = document.getElementById("kpiCompletedProjects");
const kpiOnHoldProjects = document.getElementById("kpiOnHoldProjects");
const kpiArchivedProjects = document.getElementById("kpiArchivedProjects");

if (!projectModal || !projectForm || !projectsTableBody) {
  throw new Error("Projects page is missing required elements.");
}

const state = {
  projects: [],
  filters: {
    query: "",
    status: "All Statuses",
    type: "All Project Types",
  },
};

const escapeHtml = (value) =>
  String(value)
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#39;");

const formatDate = (value) => {
  if (!value) return "-";
  const parsed = new Date(value);
  if (Number.isNaN(parsed.getTime())) return "-";
  return parsed.toLocaleDateString("en-US", { month: "short", day: "numeric", year: "numeric" });
};

const formatBudget = (value) => {
  const n = Number(value);
  if (!Number.isFinite(n) || n <= 0) return "-";
  return new Intl.NumberFormat("en-PH", {
    style: "currency",
    currency: "PHP",
    minimumFractionDigits: 0,
    maximumFractionDigits: 0,
  }).format(n);
};

const normalizeForSearch = (text) => String(text || "").trim().toLowerCase();

const updateKpis = () => {
  const byStatus = (status) => state.projects.filter((project) => project.status === status).length;

  kpiTotalProjects.textContent = String(state.projects.length);
  kpiActiveProjects.textContent = String(byStatus("In Progress"));
  kpiCompletedProjects.textContent = String(byStatus("Completed"));
  kpiOnHoldProjects.textContent = String(byStatus("On Hold"));
  kpiArchivedProjects.textContent = String(byStatus("Archived"));
};

const populateFilterOptions = () => {
  const allStatuses = Array.from(new Set(state.projects.map((project) => project.status))).sort();
  const allTypes = Array.from(new Set(state.projects.map((project) => project.type))).sort();

  projectsStatusFilter.innerHTML = '<option>All Statuses</option>';
  allStatuses.forEach((status) => {
    const option = document.createElement("option");
    option.value = status;
    option.textContent = status;
    projectsStatusFilter.append(option);
  });
  projectsStatusFilter.value = state.filters.status;

  projectsTypeFilter.innerHTML = '<option>All Project Types</option>';
  allTypes.forEach((type) => {
    const option = document.createElement("option");
    option.value = type;
    option.textContent = type;
    projectsTypeFilter.append(option);
  });
  projectsTypeFilter.value = state.filters.type;
};

const getFilteredProjects = () => {
  const query = normalizeForSearch(state.filters.query);

  return state.projects.filter((project) => {
    const matchesQuery =
      !query ||
      [project.name, project.code, project.type, project.status].some((field) =>
        normalizeForSearch(field).includes(query)
      );

    const matchesStatus =
      state.filters.status === "All Statuses" || project.status === state.filters.status;

    const matchesType =
      state.filters.type === "All Project Types" || project.type === state.filters.type;

    return matchesQuery && matchesStatus && matchesType;
  });
};

const renderProjectsTable = () => {
  const visibleProjects = getFilteredProjects();

  if (!visibleProjects.length) {
    projectsTableBody.innerHTML = "";
    if (projectsEmptyState) {
      projectsEmptyState.style.display = "grid";
      if (state.projects.length) {
        projectsEmptyState.querySelector("h2").textContent = "No matching projects";
        projectsEmptyState.querySelector("p").textContent = "Try adjusting your search or filters";
      } else {
        projectsEmptyState.querySelector("h2").textContent = "No projects yet";
        projectsEmptyState.querySelector("p").textContent = "Get started by adding your first project";
      }
    }
    projectsTableSummary.textContent = `Showing 0 to 0 of ${state.projects.length} projects`;
    return;
  }

  projectsTableBody.innerHTML = visibleProjects
    .map(
      (project) => `
      <tr>
        <td>${escapeHtml(project.name)}</td>
        <td>${escapeHtml(project.code)}</td>
        <td>${escapeHtml(project.type)}</td>
        <td>${escapeHtml(project.status)}</td>
        <td>${escapeHtml(formatDate(project.startDate))}</td>
        <td>${escapeHtml(formatDate(project.targetFinish))}</td>
        <td>${escapeHtml(project.progress)}%</td>
        <td>${escapeHtml(formatBudget(project.budget))}</td>
        <td>⋮</td>
      </tr>
    `
    )
    .join("");

  if (projectsEmptyState) projectsEmptyState.style.display = "none";

  projectsTableSummary.textContent = `Showing 1 to ${visibleProjects.length} of ${state.projects.length} projects`;
};

const syncUi = () => {
  updateKpis();
  populateFilterOptions();
  renderProjectsTable();
};

const openProjectModal = () => {
  projectModal.classList.remove("hidden");
  document.body.style.overflow = "hidden";
};

const closeProjectModal = () => {
  projectModal.classList.add("hidden");
  document.body.style.overflow = "";
};

const readProjectFromForm = () => {
  const formData = new FormData(projectForm);

  return {
    name: String(formData.get("projectName") || "").trim(),
    code: String(formData.get("projectCode") || "").trim(),
    type: String(formData.get("projectType") || "").trim(),
    status: String(formData.get("status") || "Not Started").trim(),
    priority: String(formData.get("priority") || "Medium").trim(),
    startDate: String(formData.get("startDate") || "").trim(),
    targetFinish: String(formData.get("targetFinish") || "").trim(),
    budget: Number(formData.get("budget") || 0),
    description: String(formData.get("description") || "").trim(),
    progress: 0,
  };
};

openAddProjectModalBtn?.addEventListener("click", openProjectModal);
openAddProjectModalEmptyBtn?.addEventListener("click", openProjectModal);
projectModalClose?.addEventListener("click", closeProjectModal);
projectFormCancel?.addEventListener("click", closeProjectModal);
projectModalBackdrop?.addEventListener("click", closeProjectModal);

document.addEventListener("keydown", (event) => {
  if (event.key === "Escape" && !projectModal.classList.contains("hidden")) {
    closeProjectModal();
  }
});

projectForm.addEventListener("submit", (event) => {
  event.preventDefault();

  const project = readProjectFromForm();
  if (!project.name || !project.code || !project.type || !project.startDate || !project.targetFinish) {
    return;
  }

  state.projects.unshift(project);
  projectForm.reset();
  closeProjectModal();
  syncUi();
});

projectsSearchInput?.addEventListener("input", (event) => {
  state.filters.query = event.target.value;
  renderProjectsTable();
});

projectsStatusFilter?.addEventListener("change", (event) => {
  state.filters.status = event.target.value;
  renderProjectsTable();
});

projectsTypeFilter?.addEventListener("change", (event) => {
  state.filters.type = event.target.value;
  renderProjectsTable();
});

syncUi();
