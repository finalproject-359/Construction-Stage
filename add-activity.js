const LOCAL_STORAGE_KEY = "constructionStageActivities";

const activityPageForm = document.getElementById("activityPageForm");
const activityProjectInput = document.getElementById("activityProjectInput");
const activityStartDateInput = document.getElementById("activityStartDateInput");
const activityFinishDateInput = document.getElementById("activityFinishDateInput");
const addActivityBackLink = document.getElementById("addActivityBackLink");
const activityFormCancelLink = document.getElementById("activityFormCancelLink");

const query = new URLSearchParams(window.location.search);
const selectedProject = query.get("project") || "";
const returnToActivities = selectedProject
  ? `activities.html?project=${encodeURIComponent(selectedProject)}`
  : "activities.html";

if (!selectedProject) {
  window.alert("Please select a project in Activities before opening the Add Activity page.");
  window.location.replace("activities.html");
}

if (addActivityBackLink) addActivityBackLink.href = returnToActivities;
if (activityFormCancelLink) activityFormCancelLink.href = returnToActivities;
if (activityProjectInput) activityProjectInput.value = selectedProject;

if (activityStartDateInput && activityFinishDateInput) {
  activityStartDateInput.addEventListener("change", () => {
    activityFinishDateInput.min = activityStartDateInput.value || "";
  });
}

const readActivitiesFromLocalStorage = () => {
  try {
    const raw = localStorage.getItem(LOCAL_STORAGE_KEY);
    if (!raw) return [];
    const parsed = JSON.parse(raw);
    return Array.isArray(parsed) ? parsed : [];
  } catch {
    return [];
  }
};

if (activityPageForm) {
  activityPageForm.addEventListener("submit", (event) => {
    event.preventDefault();

    if (!selectedProject) {
      window.alert("Please select a project in Activities before adding an activity.");
      window.location.href = "activities.html";
      return;
    }

    const formData = new FormData(activityPageForm);
    const plannedStart = String(formData.get("plannedStart") || "");
    const plannedFinish = String(formData.get("plannedFinish") || "");
    if (plannedStart && plannedFinish && new Date(plannedStart) > new Date(plannedFinish)) {
      window.alert("Planned finish date must be on or after planned start date.");
      return;
    }

    const nextActivity = {
      activityId: String(formData.get("activityId") || "").trim(),
      activityName: String(formData.get("activityName") || "").trim(),
      project: selectedProject,
      activityType: "-",
      status: "Not Started",
      plannedStart,
      plannedFinish,
      progress: 0,
      costStatus: "On Budget",
    };

    const existingActivities = readActivitiesFromLocalStorage();
    existingActivities.unshift(nextActivity);
    localStorage.setItem(LOCAL_STORAGE_KEY, JSON.stringify(existingActivities));

    window.location.href = selectedProject
      ? `activities.html?project=${encodeURIComponent(selectedProject)}&added=1`
      : "activities.html?added=1";
  });
}
