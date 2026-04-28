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
const projectModalTitle = document.getElementById("projectModalTitle");
const projectModalSubtitle = document.querySelector(".project-modal-header p");
const projectSubmitButton = projectForm?.querySelector('button[type="submit"]');

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
const PROJECTS_REFRESH_INTERVAL_MS = 20 * 1000;

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
};

let projectsRefreshTimer = null;
let isProjectsSyncInFlight = false;
let lastProjectsSignature = "";

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
  };
};

const escapeHtml = (value) =>
  String(value)
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#39;");

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
  projectForm.elements.budget.value = formatBudgetAsPeso(project.budget);

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

  projectsTableBody.innerHTML = projects
    .map(
      (project) => `
      <tr>
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
        <td class="project-budget-cell">${escapeHtml(pesoBudgetFormatter.format(project.budget || 0))}</td>
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

const syncProjectWithGoogleSheet = async ({ action, project, projectId }) => {
  if (!DATA_SOURCE_URL) return;

  const requestPayload = { resource: "projects", action };
  if (project) requestPayload.project = project;
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
    state.allProjects.map((project) => project.status),
    "All Statuses"
  );
  populateFilterSelect(
    projectsTypeFilter,
    state.allProjects.map((project) => project.type),
    "All Project Types"
  );
};

const addProject = async (project) => {
  await syncProjectWithGoogleSheet({ action: "create", project });
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
  applyFilters();
};

const deleteProject = async (projectId) => {
  await syncProjectWithGoogleSheet({ action: "delete", projectId });
  state.allProjects = state.allProjects.filter((project) => project.id !== projectId);
  saveToLocalStorage(state.allProjects);
  hydrateFilters();
  applyFilters();
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
    const isConfirmed = window.confirm(`Delete "${projectToDelete.name}"? This action cannot be undone.`);
    if (!isConfirmed) return;
    try {
      await deleteProject(projectId);
    } catch (error) {
      console.warn(error);
      window.alert("Failed to delete project from Google Sheet. Please try again.");
    }
  }
});

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
});

projectForm.addEventListener("submit", async (event) => {
  event.preventDefault();

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

  if (!projectName || !projectCode || !projectType || !location || !startDate || !targetFinish || !budget) {
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

  try {
    if (state.editingProjectId) {
      await updateProject(project);
    } else {
      await addProject(project);
    }
  } catch (error) {
    console.warn(error);
    window.alert(
      error instanceof Error
        ? error.message
        : "Failed to sync project to Google Sheet. Please verify Apps Script deployment access and try again."
    );
    return;
  }

  closeProjectModal();
  projectForm.reset();
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
  document.addEventListener("visibilitychange", () => {
    if (document.visibilityState === "visible") {
      refreshProjectsIfVisible({ force: true });
    }
  });
};

refreshProjectsIfVisible({ force: true });
setupProjectsRealtimeSync();
