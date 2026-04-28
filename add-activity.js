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
  const requestPayload = {
    resource: "activities",
    action: "create",
    activity,
  };

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
    const rawBody = await response.text();
    if (!rawBody) {
      return { payload: null, parseError: new Error("Empty response body") };
    }

    try {
      return { payload: JSON.parse(rawBody), parseError: null };
    } catch (error) {
      return { payload: null, parseError: error };
    }
  };

  let response;
  let payload;
  let parseError;
  const sendViaGet = async () => {
    const url = new URL(DATA_SOURCE_URL);
    url.searchParams.set("payload", JSON.stringify(requestPayload));
    response = await fetch(url.toString(), { cache: "no-store" });
    ({ payload, parseError } = await parseResponsePayload(response));
  };

  try {
    response = await postWithFormat("form");
    ({ payload, parseError } = await parseResponsePayload(response));

    const needsJsonFallback =
      !!parseError ||
      !response.ok ||
      (payload?.ok === false &&
        /invalid payload|invalid payload parameter json|invalid json payload/i.test(String(payload.error)));

    if (needsJsonFallback) {
      response = await postWithFormat("json");
      ({ payload, parseError } = await parseResponsePayload(response));
    }
  } catch (error) {
    const maybeCorsIssue =
      DATA_SOURCE_URL.includes("script.google.com/macros/s/") &&
      /failed to fetch|networkerror|cors/i.test(String(error?.message || ""));
    if (maybeCorsIssue) {
      try {
        await sendViaGet();
      } catch {
        // Fallback failed. Shared error path below will surface guidance.
      }
    }

    if (!response || payload?.ok === false || parseError) {
      const guidance = maybeCorsIssue
        ? "CORS check failed for POST. Verify your Google Apps Script Web App is deployed to Anyone and use the latest /exec deployment URL."
        : "Unable to reach the Google Sheet endpoint.";
      throw new Error(`${guidance} If this endpoint was recently changed, update DATA_SOURCE_URL in data-service.js.`);
    }
  }

  if (!response.ok) {
    throw new Error(`Unable to save activity (HTTP ${response.status})`);
  }

  if (payload?.ok === false) {
    throw new Error(payload.error || "Unable to save activity");
  }

  if (parseError || !payload || typeof payload !== "object") {
    throw new Error("Unable to save activity: endpoint returned a non-JSON response");
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
