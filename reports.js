const PROJECTS_KEY = "constructionStageProjects";

const projectSearchInput = document.getElementById("reportProjectSearch");
const reportDateStart = document.getElementById("reportDateStart");
const reportDateEnd = document.getElementById("reportDateEnd");
const selectionView = document.getElementById("reportSelectionView");
const workspace = document.getElementById("reportWorkspace");
const projectsList = document.getElementById("reportProjectsList");
const projectsEmpty = document.getElementById("reportProjectsEmpty");

const safeReadProjects = () => {
  try {
    const parsed = JSON.parse(localStorage.getItem(PROJECTS_KEY) || "[]");
    return Array.isArray(parsed) ? parsed : [];
  } catch {
    return [];
  }
};

const formatDate = (value) => {
  if (!value) return "—";
  const d = new Date(value);
  return Number.isNaN(d.getTime()) ? "—" : d.toLocaleDateString("en-US", { month: "short", day: "numeric", year: "numeric" });
};

const matchesDateRange = (project, start, end) => {
  if (!start && !end) return true;
  const pStart = project.startDate ? new Date(project.startDate) : null;
  const pEnd = project.finishDate ? new Date(project.finishDate) : pStart;
  if (!pStart && !pEnd) return false;
  const s = start ? new Date(start) : null;
  const e = end ? new Date(end) : null;
  if (s && pEnd && pEnd < s) return false;
  if (e && pStart && pStart > e) return false;
  return true;
};

const renderWorkspace = (project) => {
  const from = reportDateStart.value || project.startDate || "";
  const to = reportDateEnd.value || project.finishDate || "";
  selectionView.classList.add("hidden");
  workspace.classList.remove("hidden");
  workspace.innerHTML = `
    <div class="report-header">
      <div>
        <span class="report-tag">Formal Report Workspace</span>
        <h2>${project.name || "Untitled Project"}</h2>
        <p>Period: ${formatDate(from)} to ${formatDate(to)}</p>
      </div>
      <button id="backToProjectsBtn" type="button">← Back to Projects</button>
    </div>

    <div class="report-actions">
      <button type="button" id="reportPrintBtn">Print Report</button>
      <button type="button" id="reportExportPdfBtn">Export PDF</button>
      <button type="button" id="reportExportCsvBtn">Export CSV</button>
    </div>

    <div class="report-grid">
      <article class="report-card"><h4>Project Status</h4><p>${project.status || "Not Started"}</p></article>
      <article class="report-card"><h4>Start Date</h4><p>${formatDate(project.startDate)}</p></article>
      <article class="report-card"><h4>Target Finish</h4><p>${formatDate(project.finishDate)}</p></article>
      <article class="report-card"><h4>Budget</h4><p>₱${Number(project.budget || 0).toLocaleString("en-PH")}</p></article>
    </div>

    <div class="report-sections">
      <article>
        <h3>Progress Narrative</h3>
        <textarea placeholder="What happened during this period?"></textarea>
      </article>
      <article>
        <h3>Variance & Decision Trail</h3>
        <textarea placeholder="Why delays or cost variances happened, and what decisions were made."></textarea>
      </article>
      <article>
        <h3>Sign-off</h3>
        <p>Prepared by: ____________________</p>
        <p>Reviewed by: ____________________</p>
        <p>Date: ____________________</p>
      </article>
    </div>
  `;

  document.getElementById("backToProjectsBtn")?.addEventListener("click", () => {
    workspace.classList.add("hidden");
    selectionView.classList.remove("hidden");
  });
  document.getElementById("reportPrintBtn")?.addEventListener("click", () => window.print());
};

const renderProjects = () => {
  const query = String(projectSearchInput.value || "").toLowerCase().trim();
  const start = reportDateStart.value;
  const end = reportDateEnd.value;
  const filtered = safeReadProjects().filter((project) => {
    const hay = `${project.name || ""} ${project.code || ""} ${project.status || ""}`.toLowerCase();
    return (!query || hay.includes(query)) && matchesDateRange(project, start, end);
  });

  projectsList.innerHTML = "";
  projectsEmpty.classList.toggle("hidden", filtered.length > 0);

  filtered.forEach((project) => {
    const row = document.createElement("button");
    row.type = "button";
    row.className = "project-row";
    row.innerHTML = `<div class="project-meta"><strong>${project.name || "Untitled Project"}</strong><p>Status: ${project.status || "Not Started"}</p></div><strong>${project.code || "—"}</strong>`;
    row.addEventListener("click", () => renderWorkspace(project));
    projectsList.appendChild(row);
  });
};

projectSearchInput?.addEventListener("input", renderProjects);
reportDateStart?.addEventListener("change", renderProjects);
reportDateEnd?.addEventListener("change", renderProjects);

document.addEventListener("DOMContentLoaded", renderProjects);
