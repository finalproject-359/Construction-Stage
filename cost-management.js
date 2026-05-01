const LOCAL_STORAGE_KEY = "constructionStageProjects";

const topSearch = document.getElementById("costTopSearch");
const listSearch = document.getElementById("projectListSearch");
const projectsList = document.getElementById("costProjectsList");
const projectsEmpty = document.getElementById("costProjectsEmpty");

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

const renderProjects = (query = "") => {
  const normalizedQuery = query.trim().toLowerCase();
  const projects = loadProjects().map(normalizeProject).filter((project) => {
    if (!normalizedQuery) return true;
    return [project.name, project.code, project.status].some((value) =>
      value.toLowerCase().includes(normalizedQuery)
    );
  });

  projectsList.replaceChildren();

  if (!projects.length) {
    projectsEmpty.classList.remove("hidden");
    return;
  }

  projectsEmpty.classList.add("hidden");
  projects.forEach((project) => {
    const row = document.createElement("article");
    row.className = "project-row";

    const meta = document.createElement("div");
    meta.className = "project-meta";

    const title = document.createElement("strong");
    title.textContent = `${project.code} · ${project.name}`;

    const status = document.createElement("p");
    status.textContent = `Status: ${project.status}`;

    const budget = document.createElement("strong");
    budget.textContent = formatBudget(project.budget);

    meta.append(title, status);
    row.append(meta, budget);
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

renderProjects();
