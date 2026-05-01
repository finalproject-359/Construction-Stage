const LOCAL_STORAGE_KEY = "constructionStageProjects";

const topSearch = document.getElementById("costTopSearch");
const listSearch = document.getElementById("projectListSearch");
const projectsList = document.getElementById("costProjectsList");
const projectsEmpty = document.getElementById("costProjectsEmpty");
const selectionView = document.getElementById("costSelectionView");
const detailsView = document.getElementById("costDetailsView");

const loadProjects = () => {
  try {
    const raw = localStorage.getItem(LOCAL_STORAGE_KEY);
    const projects = JSON.parse(raw || "[]");
    return Array.isArray(projects) ? projects : [];
  } catch {
    return [];
  }
};

const normalizeProject = (project = {}) => ({
  id: String(project.id || project.projectId || "").trim(),
  name: String(project.name || project.projectName || "Untitled Project").trim(),
  code: String(project.code || project.projectCode || "-").trim(),
  status: String(project.status || "Not Started").trim(),
  budget: Number(project.budget) || 0,
});

const formatBudget = (value) =>
  new Intl.NumberFormat("en-PH", {
    style: "currency",
    currency: "PHP",
    minimumFractionDigits: 2,
    maximumFractionDigits: 2,
  }).format(value || 0);

const buildDetailsMarkup = (project) => {
  const plannedCost = project.budget || 0;
  const actualCost = plannedCost * 0.661;
  const variance = actualCost - plannedCost;
  const totalDuration = 128;

  return `
    <a href="cost-management.html" class="back-link">← Back to Projects Selection</a>
    <header class="details-header">
      <div>
        <h2>Cost Management</h2>
        <p>Project: <strong>${project.name}</strong></p>
      </div>
      <button class="ghost-btn" type="button">How it works</button>
    </header>
    <section class="details-kpis">
      <article class="kpi-card"><h4>Total Planned Cost</h4><p>${formatBudget(plannedCost)}</p></article>
      <article class="kpi-card"><h4>Total Actual Cost</h4><p>${formatBudget(actualCost)}</p></article>
      <article class="kpi-card"><h4>Variance</h4><p>${formatBudget(variance)}</p></article>
      <article class="kpi-card"><h4>Total Duration</h4><p>${totalDuration} days</p></article>
    </section>
    <div class="info-banner"><p>Detailed analytics and the full costing record are shown here for <strong>${project.code}</strong>.</p></div>
  `;
};

const showProjectDetails = (projectId) => {
  const project = loadProjects().map(normalizeProject).find((item) => item.id === projectId);
  if (!project || !selectionView || !detailsView) return false;
  selectionView.classList.add("hidden");
  detailsView.classList.remove("hidden");
  detailsView.innerHTML = buildDetailsMarkup(project);
  return true;
};

const renderProjects = (query = "") => {
  const normalizedQuery = query.trim().toLowerCase();
  const projects = loadProjects().map(normalizeProject).filter((project) => {
    if (!normalizedQuery) return true;
    return [project.name, project.code, project.status].some((value) =>
      value.toLowerCase().includes(normalizedQuery)
    );
  });

  projectsList.innerHTML = "";

  if (!projects.length) {
    projectsEmpty.classList.remove("hidden");
    return;
  }

  projectsEmpty.classList.add("hidden");
  projects.forEach((project) => {
    const row = document.createElement("button");
    row.type = "button";
    row.className = "project-row";
    row.innerHTML = `
      <div class="project-meta">
        <strong>${project.code} · ${project.name}</strong>
        <p>Status: ${project.status}</p>
      </div>
      <strong>${formatBudget(project.budget)}</strong>
    `;
    row.addEventListener("click", () => {
      const nextUrl = new URL(window.location.href);
      nextUrl.searchParams.set("projectId", project.id);
      window.location.href = nextUrl.toString();
    });
    projectsList.append(row);
  });
};

const syncSearches = (value) => {
  topSearch.value = value;
  listSearch.value = value;
  renderProjects(value);
};

topSearch?.addEventListener("input", (event) => syncSearches(event.target.value));
listSearch?.addEventListener("input", (event) => syncSearches(event.target.value));

const params = new URLSearchParams(window.location.search);
const selectedProjectId = params.get("projectId");

if (!selectedProjectId || !showProjectDetails(selectedProjectId)) {
  renderProjects();
}
