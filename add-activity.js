const DATA_SOURCE_URL =
  window.DataBridge?.DEFAULT_DATA_SOURCE_URL ||
  "https://script.google.com/macros/s/AKfycbzS5JmCF8kxtUybOAa5gtthqOeynoRRVIKFYuScLaVjb7Njp2oOYS2GwwkmnzGyDpBY/exec";

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

const createActivityInSource = async (activity) => {
  const response = await fetch(DATA_SOURCE_URL, {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
    },
    body: JSON.stringify({
      resource: "activities",
      action: "create",
      activity,
    }),
  });

  if (!response.ok) {
    throw new Error(`Unable to save activity (HTTP ${response.status})`);
  }

  const payload = await response.json();
  if (payload?.ok === false) {
    throw new Error(payload.error || "Unable to save activity");
  }

  return payload;
};

if (activityPageForm) {
  activityPageForm.addEventListener("submit", async (event) => {
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

    const activityPayload = {
      id: String(formData.get("activityId") || "").trim(),
      name: String(formData.get("activityName") || "").trim(),
      project: selectedProject,
      plannedStart,
      plannedFinish,
      status: "Not Started",
      percentComplete: 0,
      notes: "",
    };

    try {
      await createActivityInSource(activityPayload);
    } catch (error) {
      const reason = error?.message ? `\nReason: ${error.message}` : "";
      window.alert(`Unable to save activity to Google Sheets. No local copy was saved.${reason}`);
      return;
    }

    window.location.href = selectedProject
      ? `activities.html?project=${encodeURIComponent(selectedProject)}&added=1`
      : "activities.html?added=1";
  });
}
