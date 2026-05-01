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

const parseBudgetValue = (value) => {
  if (value === null || value === undefined || value === "") return 0;
  if (typeof value === "number") return Number.isFinite(value) ? value : 0;
  const cleaned = String(value).replace(/[^\d.-]/g, "");
  const parsed = Number(cleaned);
  return Number.isFinite(parsed) ? parsed : 0;
};

const normalizeProject = (project = {}) => ({
  id: String(getValueByAliases(project, ["id", "projectId", "project_id"]) || "").trim(),
  name: String(getValueByAliases(project, ["name", "project", "projectName", "project_name"]) || "Untitled Project").trim(),
  code: String(getValueByAliases(project, ["code", "projectCode", "project_code"]) || "-").trim(),
  status: String(getValueByAliases(project, ["status", "projectStatus", "project_status"]) || "Not Started").trim(),
  budget: parseBudgetValue(getValueByAliases(project, ["budget", "plannedCost", "planned_cost", "plannedValue", "planned_value"])),
});

const escapeHtml = (value) =>
  String(value)
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#39;");

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
  const variancePercent = plannedCost ? ((Math.abs(variance) / plannedCost) * 100) : 0;
  const totalDuration = 128;
  const avgCostPerDay = totalDuration ? actualCost / totalDuration : 0;

  const spendPercent = plannedCost > 0 ? (actualCost / plannedCost) * 100 : 0;
  const normalizedBarHeight = Math.max(20, Math.min(100, spendPercent));

  return `
    <header class="details-header">
      <h2>Cost Management</h2>
      <p>Track and manage project costs efficiently.</p>
    </header>
    <section class="selected-project-banner" aria-label="Selected project">
      <div>
        <p class="selected-project-label">Selected Project</p>
        <h3>${escapeHtml(project.name)}</h3>
      </div>
      <a href="cost-management.html" class="ghost-btn">← Back to Projects</a>
    </section>
    <nav class="details-tabs" aria-label="Cost tabs">
      <button class="tab-btn active" type="button">Overview</button>
      <button class="tab-btn" type="button">Costing Record</button>
    </nav>
    <section class="details-kpis">
      <article class="kpi-card"><h4>Total Planned Cost</h4><p>${formatBudget(plannedCost)}</p><small>Across all activities</small></article>
      <article class="kpi-card"><h4>Total Actual Cost</h4><p>${formatBudget(actualCost)}</p><small>66.10% of planned cost</small></article>
      <article class="kpi-card"><h4>Variance</h4><p>${formatBudget(variance)}</p><small class="good">${variancePercent.toFixed(2)}% under budget</small></article>
      <article class="kpi-card"><h4>Total Duration</h4><p>${totalDuration} days</p><small>Across all activities</small></article>
    </section>
    <section class="overview-grid">
      <article class="panel chart-panel">
        <h3>Budget vs Actual</h3>
        <div class="bars" role="img" aria-label="Comparison of planned and actual project costs">
          <div class="bars-grid" aria-hidden="true">
            <span>100%</span><span>75%</span><span>50%</span><span>25%</span><span>0%</span>
          </div>
          <div class="bars-track">
            <div class="bar-wrap"><div class="bar planned" style="height:100%"></div><strong>${formatBudget(plannedCost)}</strong><p>Total Planned Cost</p></div>
            <div class="bar-wrap"><div class="bar actual" style="height:${normalizedBarHeight}%"></div><strong>${formatBudget(actualCost)}</strong><p>Total Actual Cost</p></div>
          </div>
        </div>
        <p class="chart-caption">Actual spending is ${spendPercent.toFixed(2)}% of the planned cost.</p>
      </article>
      <article class="panel summary-panel">
        <h3>Cost Summary</h3>
        <ul>
          <li><span>Total Planned Cost</span><strong>${formatBudget(plannedCost)}</strong></li>
          <li><span>Total Actual Cost</span><strong>${formatBudget(actualCost)}</strong></li>
          <li><span>Variance</span><strong class="good">${formatBudget(variance)}</strong></li>
          <li><span>Variance Percent</span><strong class="good">${variancePercent.toFixed(2)}% under budget</strong></li>
          <li><span>Total Duration</span><strong>${totalDuration} days</strong></li>
          <li><span>Average Cost per Day (Actual)</span><strong>${formatBudget(avgCostPerDay)}</strong></li>
        </ul>
      </article>
      <article class="panel donut-panel">
        <h3>Cost Status (by Actual vs Planned)</h3>
        <div class="donut"></div>
        <p class="legend">Under Budget 71.43% · Over Budget 14.29% · No Actual Cost 14.29%</p>
      </article>
      <article class="panel table-panel">
        <h3>Top Over Budget Activities</h3>
        <table>
          <thead><tr><th>Activity ID</th><th>Activity</th><th>Planned Cost</th><th>Actual Cost</th><th>Variance</th></tr></thead>
          <tbody><tr><td>ACT-002</td><td>Site Preparation</td><td>₱1,250,000.00</td><td>₱1,300,000.00</td><td class="bad">+₱50,000.00</td></tr></tbody>
        </table>
      </article>
    </section>
    <div class="info-banner"><p>Costs are based on activities. Manage actual costs in the <a href="#">Costing Record</a> tab.</p></div>
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
        <strong>${escapeHtml(project.code)} · ${escapeHtml(project.name)}</strong>
        <p>Status: ${escapeHtml(project.status)}</p>
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
