/**
 * Google Apps Script Web App endpoint for Construction Stage data.
 *
 * Supported query params:
 * - resource: dashboard | projects | activities | costs | daily-costs | reports | all
 * - projectId / project: optional project filter
 *
 * Expected sheet tabs (default names can be customized below):
 * - Projects
 * - Activities
 * - Costs
 */
const CONFIG = {
  sheetNames: {
    projects: "Projects",
    activities: "Activities",
    costs: "Costs",
    dailyCosts: "DailyCosts",
  },
  headers: {
    projects: [
      "Project ID",
      "Project Name",
      "Project Type",
      "Status",
      "Location",
      "Start Date",
      "Finish Date",
      "Budget",
      "Progress",
      "Created At",
    ],
    activities: [
      "Project ID",
      "Project Name",
      "Activity ID",
      "Activity",
      "Planned Start",
      "Planned Finish",
      "Duration",
      "Status",
      "Progress",
      "Created At",
    ],
    costs: [
      "Project ID",
      "Project Name",
      "Cost ID",
      "Activity ID",
      "Activity",
      "Duration",
      "Progress",
      "Planned Cost",
      "Actual Cost",
      "Earned Value",
      "Created At",
    ],
    dailyCosts: [
      "Project ID",
      "Project Name",
      "Cost ID",
      "Activity ID",
      "Activity",
      "Progress/Day",
      "Planned Cost",
      "Planned Cost/Day",
      "Date",
      "Actual Cost/Day",
      "Earned Value/Day",
      "Created At",
    ],
  },
};

function getSpreadsheetTodayDate() {
  var timeZone = Session.getScriptTimeZone();
  var isoDate = Utilities.formatDate(new Date(), timeZone, "yyyy-MM-dd");
  return new Date(isoDate + "T12:00:00");
}

function doGet(e) {
  const params = (e && e.parameter) || {};
  const payloadParam = params.payload;
  let payload = payloadParam ? safeParseJson(payloadParam) : {};

  if (!payload || typeof payload !== "object") payload = {};
  payload.resource =
    payload.resource || params.resource || params.view || params.type;
  payload.action = payload.action || params.action;
  payload.projectId =
    payload.projectId || params.projectId || params.project_id;

  return handleRequest(payload);
}

function doPost(e) {
  return handleRequest(parsePostPayload(e));
}

function onEdit(e) {
  try {
    if (!e || !e.range) return;
    var range = e.range;
    var sheet = range.getSheet();
    if (!sheet) return;
    var sheetName = cleanText(sheet.getName());

    if (sheetName === cleanText(CONFIG.sheetNames.dailyCosts)) {
      handleDailyCostsSheetEdit(sheet, range);
      return;
    }

    if (sheetName === cleanText(CONFIG.sheetNames.activities)) {
      handleActivitiesSheetEdit(sheet, range);
      return;
    }

    if (sheetName === cleanText(CONFIG.sheetNames.costs)) {
      handleCostsSheetEdit(sheet, range);
    }
  } catch (error) {
    Logger.log(
      "onEdit error: " + (error && error.message ? error.message : error),
    );
  }
}

function handleDailyCostsSheetEdit(sheet, range) {
  ensureSheetHeaders(sheet, CONFIG.headers.dailyCosts);
  var columns = getDailyCostColumnMap(sheet);
  var editedColumn = range.getColumn();
  var relevant = [
    columns.projectId,
    columns.costId,
    columns.activityId,
    columns.activity,
    columns.plannedCost,
    columns.plannedCostPerDay,
    columns.progress,
    columns.actualCost,
    columns.date,
  ];
  if (relevant.indexOf(editedColumn) < 0) return;

  var row = range.getRow();
  if (row <= 1) return;
  var rowValues = sheet
    .getRange(row, 1, 1, Math.max(sheet.getLastColumn(), columns.maxColumn))
    .getValues()[0];
  var projectId = columns.projectId
    ? cleanText(rowValues[columns.projectId - 1])
    : "";
  var costId = columns.costId ? cleanText(rowValues[columns.costId - 1]) : "";
  var activityId = columns.activityId
    ? cleanText(rowValues[columns.activityId - 1])
    : "";
  if (!projectId || !costId) return;

  refreshDailyCostRowMetrics(sheet, row, columns);
  syncCostActualFromDailyCost(projectId, costId, { activityId: activityId });
}

function refreshDailyCostRowMetrics(dailySheet, rowNumber, columns) {
  var rowValues = dailySheet
    .getRange(
      rowNumber,
      1,
      1,
      Math.max(dailySheet.getLastColumn(), columns.maxColumn),
    )
    .getValues()[0];
  var projectId = columns.projectId
    ? cleanText(rowValues[columns.projectId - 1])
    : "";
  var costId = columns.costId ? cleanText(rowValues[columns.costId - 1]) : "";
  if (!projectId || !costId) return;

  var progress = columns.progress
    ? parseNumber(rowValues[columns.progress - 1])
    : -1;
  if (progress < 0) {
    var linkedCost = findCostRecord(projectId, costId);
    if (linkedCost) {
      progress = findActivityProgress(
        projectId,
        linkedCost.activityId,
        linkedCost.activity,
      );
    }
  }

  var plannedCost = columns.plannedCost
    ? parseNumber(rowValues[columns.plannedCost - 1])
    : 0;
  var plannedCostPerDay = columns.plannedCostPerDay
    ? parseNumber(rowValues[columns.plannedCostPerDay - 1])
    : 0;
  var earnedValue = 0;

  if (plannedCost > 0 && progress >= 0) {
    earnedValue = roundTo(
      plannedCost * (Math.max(0, Math.min(100, progress)) / 100),
      2,
    );
  } else if (plannedCostPerDay > 0 && progress >= 0) {
    earnedValue = roundTo(
      plannedCostPerDay * (Math.max(0, Math.min(100, progress)) / 100),
      2,
    );
  }

  if (columns.progress && progress >= 0) {
    dailySheet
      .getRange(rowNumber, columns.progress)
      .setValue(roundTo(progress, 2));
    dailySheet.getRange(rowNumber, columns.progress).setNumberFormat("0.00");
  }

  if (columns.earnedValue) {
    dailySheet.getRange(rowNumber, columns.earnedValue).setValue(earnedValue);
    dailySheet
      .getRange(rowNumber, columns.earnedValue)
      .setNumberFormat("#,##0.00");
  }
}

function findCostRecord(projectId, costId) {
  var costsSheet = getOrCreateSheet(CONFIG.sheetNames.costs);
  ensureSheetHeaders(costsSheet, CONFIG.headers.costs);
  var columns = getCostColumnMap(costsSheet);
  var values = costsSheet.getDataRange().getValues();
  var normalizedProjectId = cleanText(projectId);
  var normalizedCostId = cleanText(costId);

  for (var i = 1; i < values.length; i += 1) {
    if (
      cleanText(values[i][columns.projectId - 1]) === normalizedProjectId &&
      cleanText(values[i][columns.costId - 1]) === normalizedCostId
    ) {
      return {
        activityId: columns.activityId
          ? cleanText(values[i][columns.activityId - 1])
          : "",
        activity: columns.activity
          ? cleanText(values[i][columns.activity - 1])
          : "",
      };
    }
  }

  return null;
}

function findActivitySchedule(projectId, activityId, activityName) {
  var activitiesSheet = getOrCreateSheet(CONFIG.sheetNames.activities);
  ensureSheetHeaders(activitiesSheet, CONFIG.headers.activities);
  var columns = getActivityColumnMap(activitiesSheet);
  var values = activitiesSheet.getDataRange().getValues();
  var normalizedProjectId = cleanText(projectId);
  var normalizedActivityId = cleanText(activityId);
  var normalizedActivityName = cleanText(activityName);

  for (var i = 1; i < values.length; i += 1) {
    var rowProjectId = columns.projectId
      ? cleanText(values[i][columns.projectId - 1])
      : "";
    if (rowProjectId !== normalizedProjectId) continue;

    var rowActivityId = columns.id ? cleanText(values[i][columns.id - 1]) : "";
    var rowActivityName = columns.name
      ? cleanText(values[i][columns.name - 1])
      : "";
    var matches =
      (normalizedActivityId && rowActivityId === normalizedActivityId) ||
      (!normalizedActivityId &&
        normalizedActivityName &&
        rowActivityName === normalizedActivityName) ||
      (normalizedActivityId &&
        !rowActivityId &&
        normalizedActivityName &&
        rowActivityName === normalizedActivityName);
    if (!matches) continue;

    return {
      plannedStart: columns.plannedStart
        ? normalizeDate(values[i][columns.plannedStart - 1])
        : "",
      plannedFinish: columns.plannedFinish
        ? normalizeDate(values[i][columns.plannedFinish - 1])
        : "",
    };
  }

  return { plannedStart: "", plannedFinish: "" };
}

function findActivityProgress(projectId, activityId, activityName) {
  var activitiesSheet = getOrCreateSheet(CONFIG.sheetNames.activities);
  ensureSheetHeaders(activitiesSheet, CONFIG.headers.activities);
  var columns = getActivityColumnMap(activitiesSheet);
  var values = activitiesSheet.getDataRange().getValues();
  var normalizedProjectId = cleanText(projectId);
  var normalizedActivityId = cleanText(activityId);
  var normalizedActivityName = cleanText(activityName);

  for (var i = 1; i < values.length; i += 1) {
    var rowProjectId = columns.projectId
      ? cleanText(values[i][columns.projectId - 1])
      : "";
    if (rowProjectId !== normalizedProjectId) continue;

    var rowActivityId = columns.id ? cleanText(values[i][columns.id - 1]) : "";
    var rowActivityName = columns.name
      ? cleanText(values[i][columns.name - 1])
      : "";
    var matches =
      (normalizedActivityId && rowActivityId === normalizedActivityId) ||
      (!normalizedActivityId &&
        normalizedActivityName &&
        rowActivityName === normalizedActivityName) ||
      (normalizedActivityId &&
        !rowActivityId &&
        normalizedActivityName &&
        rowActivityName === normalizedActivityName);
    if (!matches) continue;

    return parseNumber(
      columns.percentComplete ? values[i][columns.percentComplete - 1] : 0,
    );
  }

  return -1;
}

function handleActivitiesSheetEdit(sheet, range) {
  var columns = getActivityColumnMap(sheet);
  var editedColumn = range.getColumn();
  var relevant = [
    columns.id,
    columns.name,
    columns.percentComplete,
    columns.projectId,
  ];
  if (relevant.indexOf(editedColumn) < 0) return;

  var row = range.getRow();
  if (row <= 1) return;
  var rowValues = sheet
    .getRange(row, 1, 1, Math.max(sheet.getLastColumn(), columns.maxColumn))
    .getValues()[0];
  var projectId = columns.projectId
    ? cleanText(rowValues[columns.projectId - 1])
    : "";
  var activityId = columns.id ? cleanText(rowValues[columns.id - 1]) : "";
  var activityName = columns.name ? cleanText(rowValues[columns.name - 1]) : "";
  if (!projectId || (!activityId && !activityName)) return;
  syncEarnedValueForActivity(projectId, activityId, activityName);
  syncProjectProgressFromActivities(projectId, "");
}

function handleCostsSheetEdit(sheet, range) {
  var columns = getCostColumnMap(sheet);
  var firstEditedColumn = range.getColumn();
  var lastEditedColumn = firstEditedColumn + range.getNumColumns() - 1;
  var relevant = [
    columns.projectId,
    columns.project,
    columns.costId,
    columns.activityId,
    columns.activity,
    columns.plannedCost,
    columns.actualCost,
    columns.progress,
  ].filter(function (column) {
    return column > 0;
  });
  var touchesRelevantColumn = relevant.some(function (column) {
    return column >= firstEditedColumn && column <= lastEditedColumn;
  });
  if (!touchesRelevantColumn) return;

  var firstRow = Math.max(range.getRow(), 2);
  var lastRow = range.getRow() + range.getNumRows() - 1;
  for (var row = firstRow; row <= lastRow; row += 1) {
    refreshCostRowMetrics(sheet, row, columns);
  }
}

function refreshCostRowMetrics(costsSheet, rowNumber, columns) {
  var rowValues = costsSheet
    .getRange(
      rowNumber,
      1,
      1,
      Math.max(costsSheet.getLastColumn(), columns.maxColumn),
    )
    .getValues()[0];
  var projectId = columns.projectId
    ? cleanText(rowValues[columns.projectId - 1])
    : "";
  var costId = columns.costId ? cleanText(rowValues[columns.costId - 1]) : "";
  if (!projectId || !costId) return;

  var cost = {
    projectId: projectId,
    project: columns.project ? cleanText(rowValues[columns.project - 1]) : "",
    costId: costId,
    activityId: columns.activityId
      ? cleanText(rowValues[columns.activityId - 1])
      : "",
    activity: columns.activity
      ? cleanText(rowValues[columns.activity - 1])
      : "",
    plannedCost: columns.plannedCost
      ? parseNumber(rowValues[columns.plannedCost - 1])
      : 0,
    earnedValue: columns.earnedValue
      ? parseNumber(rowValues[columns.earnedValue - 1])
      : 0,
    progress: columns.progress
      ? parseNumber(rowValues[columns.progress - 1])
      : 0,
  };

  syncActivityProgressFromCost(cost);

  if (columns.earnedValue) {
    var computedEarnedValue = Number(computeEarnedValue(cost)) || 0;
    costsSheet
      .getRange(rowNumber, columns.earnedValue)
      .setValue(computedEarnedValue);
    costsSheet
      .getRange(rowNumber, columns.earnedValue)
      .setNumberFormat("#,##0.00");
  }

  syncCostActualFromDailyCost(projectId, costId, { syncProgress: false });
}

function calculateProjectProgressFromActivities(projectId, projectName) {
  var activitiesSheet = getOrCreateSheet(CONFIG.sheetNames.activities);
  ensureSheetHeaders(activitiesSheet, CONFIG.headers.activities);
  var columns = getActivityColumnMap(activitiesSheet);
  if (!columns.percentComplete || (!columns.projectId && !columns.project)) {
    return null;
  }

  var targetProjectCandidates = getIdentityCandidates(projectId, projectName);
  if (!targetProjectCandidates.length) return null;

  var values = activitiesSheet.getDataRange().getValues();
  var totalProgress = 0;
  var activityCount = 0;
  for (var i = 1; i < values.length; i += 1) {
    var rowProjectId = columns.projectId
      ? cleanText(values[i][columns.projectId - 1])
      : "";
    var rowProjectName = columns.project
      ? cleanText(values[i][columns.project - 1])
      : "";
    var rowProjectCandidates = getIdentityCandidates(rowProjectId, rowProjectName);
    if (!identityCandidatesMatch(rowProjectCandidates, targetProjectCandidates)) {
      continue;
    }

    totalProgress += clampPercent(values[i][columns.percentComplete - 1]);
    activityCount += 1;
  }

  if (!activityCount) return 0;
  return roundTo(totalProgress / activityCount, 2);
}

function syncProjectProgressFromActivities(projectId, projectName) {
  var projectsSheet = getOrCreateSheet(CONFIG.sheetNames.projects);
  ensureSheetHeaders(projectsSheet, CONFIG.headers.projects);
  var projectColumns = getProjectColumnMap(projectsSheet);
  if (!projectColumns.progress || (!projectColumns.id && !projectColumns.name)) {
    return null;
  }

  var targetProjectCandidates = getIdentityCandidates(projectId, projectName);
  if (!targetProjectCandidates.length) return null;

  var values = projectsSheet.getDataRange().getValues();
  for (var i = 1; i < values.length; i += 1) {
    var rowProjectId = projectColumns.id
      ? cleanText(values[i][projectColumns.id - 1])
      : "";
    var rowProjectName = projectColumns.name
      ? cleanText(values[i][projectColumns.name - 1])
      : "";
    var rowProjectCandidates = getIdentityCandidates(rowProjectId, rowProjectName);
    if (!identityCandidatesMatch(rowProjectCandidates, targetProjectCandidates)) {
      continue;
    }

    var progress = calculateProjectProgressFromActivities(rowProjectId, rowProjectName);
    if (progress === null) return null;

    projectsSheet.getRange(i + 1, projectColumns.progress).setValue(progress);
    projectsSheet
      .getRange(i + 1, projectColumns.progress)
      .setNumberFormat("0.00");
    return progress;
  }

  return null;
}

function syncAllProjectProgressFromActivities() {
  var projectsSheet = getOrCreateSheet(CONFIG.sheetNames.projects);
  ensureSheetHeaders(projectsSheet, CONFIG.headers.projects);
  var columns = getProjectColumnMap(projectsSheet);
  if (!columns.progress || (!columns.id && !columns.name)) return;

  var values = projectsSheet.getDataRange().getValues();
  for (var i = 1; i < values.length; i += 1) {
    syncProjectProgressFromActivities(
      columns.id ? cleanText(values[i][columns.id - 1]) : "",
      columns.name ? cleanText(values[i][columns.name - 1]) : "",
    );
  }
}

function syncEarnedValueForActivity(projectId, activityId, activityName) {
  var normalizedProjectId = cleanText(projectId);
  var normalizedActivityId = cleanText(activityId);
  var normalizedActivityName = cleanText(activityName);
  if (
    !normalizedProjectId ||
    (!normalizedActivityId && !normalizedActivityName)
  )
    return;

  var costsSheet = getOrCreateSheet(CONFIG.sheetNames.costs);
  ensureSheetHeaders(costsSheet, CONFIG.headers.costs);
  var columns = getCostColumnMap(costsSheet);
  var values = costsSheet.getDataRange().getValues();

  for (var i = 1; i < values.length; i += 1) {
    var row = values[i];
    var rowProjectId = columns.projectId
      ? cleanText(row[columns.projectId - 1])
      : "";
    if (rowProjectId !== normalizedProjectId) continue;

    var rowActivityId = columns.activityId
      ? cleanText(row[columns.activityId - 1])
      : "";
    var rowActivityName = columns.activity
      ? cleanText(row[columns.activity - 1])
      : "";
    var matches =
      (normalizedActivityId && rowActivityId === normalizedActivityId) ||
      (!normalizedActivityId &&
        normalizedActivityName &&
        rowActivityName === normalizedActivityName) ||
      (normalizedActivityId &&
        !rowActivityId &&
        normalizedActivityName &&
        rowActivityName === normalizedActivityName);
    if (!matches) continue;

    refreshCostRowMetrics(costsSheet, i + 1, columns);
  }
}

function syncActivityProgressFromCost(cost) {
  var normalizedProjectId = cleanText(cost && cost.projectId);
  var normalizedProjectName = cleanText(
    cost && (cost.project || cost.projectName || cost.project_name),
  );
  var normalizedActivityId = cleanText(cost && cost.activityId);
  var normalizedActivityName = cleanText(cost && cost.activity);
  if (
    (!normalizedProjectId && !normalizedProjectName) ||
    (!normalizedActivityId && !normalizedActivityName)
  ) {
    return false;
  }

  var activitiesSheet = getOrCreateSheet(CONFIG.sheetNames.activities);
  ensureSheetHeaders(activitiesSheet, CONFIG.headers.activities);
  var columns = getActivityColumnMap(activitiesSheet);
  if (!columns.percentComplete) return false;

  var values = activitiesSheet.getDataRange().getValues();
  var progress = roundTo(clampPercent(cost && cost.progress), 2);
  var targetProjectCandidates = getIdentityCandidates(
    normalizedProjectId,
    normalizedProjectName,
  );
  var targetActivityCandidates = getIdentityCandidates(
    normalizedActivityId,
    normalizedActivityName,
  );

  for (var i = 1; i < values.length; i += 1) {
    var rowProjectId = columns.projectId
      ? cleanText(values[i][columns.projectId - 1])
      : "";
    var rowProjectName = columns.project
      ? cleanText(values[i][columns.project - 1])
      : "";
    var rowProjectCandidates = getIdentityCandidates(
      rowProjectId,
      rowProjectName,
    );
    var matchesProject = identityCandidatesMatch(
      rowProjectCandidates,
      targetProjectCandidates,
    );
    if (!matchesProject) continue;

    var rowActivityId = columns.id ? cleanText(values[i][columns.id - 1]) : "";
    var rowActivityName = columns.name
      ? cleanText(values[i][columns.name - 1])
      : "";
    var rowActivityCandidates = getIdentityCandidates(
      rowActivityId,
      rowActivityName,
    );
    var matchesActivity = identityCandidatesMatch(
      rowActivityCandidates,
      targetActivityCandidates,
    );
    if (!matchesActivity) continue;

    activitiesSheet
      .getRange(i + 1, columns.percentComplete)
      .setValue(progress);
    activitiesSheet
      .getRange(i + 1, columns.percentComplete)
      .setNumberFormat("0.00");

    if (columns.status) {
      var nextStatus =
        progress >= 100
          ? "Completed"
          : progress > 0
            ? "In Progress"
            : "Not Started";
      activitiesSheet.getRange(i + 1, columns.status).setValue(nextStatus);
    }
    syncProjectProgressFromActivities(rowProjectId, rowProjectName);
    return true;
  }

  return false;
}

function syncAllActivityProgressFromCosts() {
  var costsSheet = getOrCreateSheet(CONFIG.sheetNames.costs);
  ensureSheetHeaders(costsSheet, CONFIG.headers.costs);
  var columns = getCostColumnMap(costsSheet);
  if (!columns.projectId || !columns.progress) return;

  var values = costsSheet.getDataRange().getValues();
  for (var i = 1; i < values.length; i += 1) {
    var row = values[i];
    var rawProgress = row[columns.progress - 1];
    if (
      rawProgress === "" ||
      rawProgress === null ||
      rawProgress === undefined
    ) {
      continue;
    }

    syncActivityProgressFromCost({
      projectId: cleanText(row[columns.projectId - 1]),
      project: columns.project ? cleanText(row[columns.project - 1]) : "",
      activityId: columns.activityId
        ? cleanText(row[columns.activityId - 1])
        : "",
      activity: columns.activity ? cleanText(row[columns.activity - 1]) : "",
      progress: rawProgress,
    });
  }
}

function handleRequest(payload) {
  try {
    const source = payload || {};
    const resource = normalizeResource(source.resource || "dashboard");
    const action = cleanText(source.action).toLowerCase();

    if (action) {
      if (resource === "projects") {
        const projectsSheet = getOrCreateSheet(CONFIG.sheetNames.projects);
        ensureSheetHeaders(projectsSheet, CONFIG.headers.projects);
        return handleProjectMutation(action, source);
      }

      if (resource === "activities") {
        const activitiesSheet = getOrCreateSheet(CONFIG.sheetNames.activities);
        ensureSheetHeaders(activitiesSheet, CONFIG.headers.activities);
        return handleActivityMutation(action, source);
      }

      if (resource === "costs") {
        const costsSheet = getOrCreateSheet(CONFIG.sheetNames.costs);
        ensureSheetHeaders(costsSheet, CONFIG.headers.costs);
        return handleCostMutation(action, source);
      }

      if (resource === "daily_costs") {
        const dailySheet = getOrCreateSheet(CONFIG.sheetNames.dailyCosts);
        ensureSheetHeaders(dailySheet, CONFIG.headers.dailyCosts);
        return handleDailyCostMutation(action, source);
      }

      throw new Error(
        'Only "projects", "activities", "costs", and "daily_costs" are supported for mutations.',
      );
    }

    const projectFilter = {
      id: cleanText(source.projectId || source.project_id || ""),
      name: cleanText(
        source.project || source.projectName || source.project_name || "",
      ),
    };

    const allData = loadDataByResource(resource);
    const filtered = applyProjectFilter(allData, projectFilter);

    const payloadByResource = {
      projects: buildProjectsPayload(filtered.projects),
      activities: buildActivitiesPayload(filtered.activities),
      costs: buildCostsPayload(filtered.costs),
      daily_costs: {
        count: filtered.dailyCosts.length,
        dailyCosts: filtered.dailyCosts,
      },
      dashboard: buildDashboardPayload(filtered),
      reports: buildReportsPayload(filtered),
      all: buildAllPayload(filtered),
    };

    const responsePayload =
      payloadByResource[resource] || payloadByResource.dashboard;

    return jsonResponse({
      ok: true,
      resource: resource,
      filter: projectFilter,
      generatedAt: new Date().toISOString(),
      ...responsePayload,
    });
  } catch (error) {
    return jsonResponse({
      ok: false,
      error:
        error && error.message
          ? error.message
          : "Unexpected error while processing request.",
      generatedAt: new Date().toISOString(),
    });
  }
}

function createEmptyDataBundle() {
  return {
    projects: [],
    activities: [],
    costs: [],
    dailyCosts: [],
    sheets: {},
  };
}

function loadDataByResource(resource) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const bundle = createEmptyDataBundle();

  if (
    resource !== "projects" &&
    resource !== "costs" &&
    resource !== "daily_costs"
  ) {
    syncAllActivityProgressFromCosts();
  }

  if (
    resource === "projects" ||
    resource === "dashboard" ||
    resource === "reports" ||
    resource === "all"
  ) {
    syncAllProjectProgressFromActivities();
  }

  const readResource = function (
    targetKey,
    sheetName,
    expectedHeaders,
    normalizer,
  ) {
    const result = readSheetRows(ss, sheetName, expectedHeaders);
    bundle[targetKey] = result.rows.map(normalizer);
    bundle.sheets[targetKey] = result.meta;
  };

  if (resource === "projects") {
    readResource(
      "projects",
      CONFIG.sheetNames.projects,
      CONFIG.headers.projects,
      normalizeProjectRecord,
    );
    return bundle;
  }

  if (resource === "activities") {
    readResource(
      "activities",
      CONFIG.sheetNames.activities,
      CONFIG.headers.activities,
      normalizeActivityRecord,
    );
    return bundle;
  }

  if (resource === "dashboard") {
    readResource(
      "projects",
      CONFIG.sheetNames.projects,
      CONFIG.headers.projects,
      normalizeProjectRecord,
    );
    readResource(
      "activities",
      CONFIG.sheetNames.activities,
      CONFIG.headers.activities,
      normalizeActivityRecord,
    );
    readResource(
      "costs",
      CONFIG.sheetNames.costs,
      CONFIG.headers.costs,
      normalizeCostRecord,
    );
    readResource(
      "dailyCosts",
      CONFIG.sheetNames.dailyCosts,
      CONFIG.headers.dailyCosts,
      normalizeDailyCostRecord,
    );
    return bundle;
  }

  if (resource === "costs") {
    readResource(
      "costs",
      CONFIG.sheetNames.costs,
      CONFIG.headers.costs,
      normalizeCostRecord,
    );
    return bundle;
  }

  if (resource === "daily_costs") {
    readResource(
      "dailyCosts",
      CONFIG.sheetNames.dailyCosts,
      CONFIG.headers.dailyCosts,
      normalizeDailyCostRecord,
    );
    return bundle;
  }

  if (resource === "reports" || resource === "all") {
    readResource(
      "projects",
      CONFIG.sheetNames.projects,
      CONFIG.headers.projects,
      normalizeProjectRecord,
    );
    readResource(
      "activities",
      CONFIG.sheetNames.activities,
      CONFIG.headers.activities,
      normalizeActivityRecord,
    );
    readResource(
      "costs",
      CONFIG.sheetNames.costs,
      CONFIG.headers.costs,
      normalizeCostRecord,
    );
    readResource(
      "dailyCosts",
      CONFIG.sheetNames.dailyCosts,
      CONFIG.headers.dailyCosts,
      normalizeDailyCostRecord,
    );
    return bundle;
  }

  readResource(
    "projects",
    CONFIG.sheetNames.projects,
    CONFIG.headers.projects,
    normalizeProjectRecord,
  );
  readResource(
    "activities",
    CONFIG.sheetNames.activities,
    CONFIG.headers.activities,
    normalizeActivityRecord,
  );
  readResource(
    "costs",
    CONFIG.sheetNames.costs,
    CONFIG.headers.costs,
    normalizeCostRecord,
  );
  readResource(
    "dailyCosts",
    CONFIG.sheetNames.dailyCosts,
    CONFIG.headers.dailyCosts,
    normalizeDailyCostRecord,
  );
  return bundle;
}

function handleProjectMutation(action, payload) {
  if (action === "create") {
    const project = normalizeIncomingProject(payload.project || payload);
    if (!project.name || !project.id) {
      throw new Error("Project Name and Project ID are required.");
    }

    const sheet = getOrCreateSheet(CONFIG.sheetNames.projects);
    ensureSheetHeaders(sheet, CONFIG.headers.projects);
    const columns = getProjectColumnMap(sheet);
    const lastColumn = Math.max(
      sheet.getLastColumn(),
      CONFIG.headers.projects.length,
      columns.maxColumn,
    );
    const rowValues = new Array(lastColumn).fill("");
    const storedProjectId = cleanText(project.id || project.code);

    if (columns.id) rowValues[columns.id - 1] = storedProjectId;
    if (columns.name) rowValues[columns.name - 1] = project.name;
    if (columns.type) rowValues[columns.type - 1] = project.type;
    if (columns.status) rowValues[columns.status - 1] = project.status;
    if (columns.location) rowValues[columns.location - 1] = project.location;
    if (columns.startDate) rowValues[columns.startDate - 1] = project.startDate;
    if (columns.finishDate)
      rowValues[columns.finishDate - 1] = project.finishDate;
    if (columns.budget) rowValues[columns.budget - 1] = project.budget;
    if (columns.progress) rowValues[columns.progress - 1] = project.progress;
    if (columns.createdAt) rowValues[columns.createdAt - 1] = getSpreadsheetTodayDate();

    sheet
      .getRange(sheet.getLastRow() + 1, 1, 1, lastColumn)
      .setValues([rowValues]);

    return jsonResponse({
      ok: true,
      message: "Project saved successfully.",
      project: {
        ...project,
        id: storedProjectId,
        code: storedProjectId,
      },
      generatedAt: new Date().toISOString(),
    });
  }

  if (action === "update") {
    const project = normalizeIncomingProject(payload.project || payload);
    if (!project.id) {
      throw new Error("Project ID is required for update.");
    }

    const updateResult = updateProjectRow(project);
    return jsonResponse({
      ok: true,
      message: "Project updated successfully.",
      project: updateResult,
      generatedAt: new Date().toISOString(),
    });
  }

  if (action === "delete") {
    const projectId = cleanText(payload.projectId || payload.id);
    if (!projectId) {
      throw new Error("Project ID is required for delete.");
    }

    deleteProjectRow(projectId);
    return jsonResponse({
      ok: true,
      message: "Project deleted successfully.",
      projectId: projectId,
      generatedAt: new Date().toISOString(),
    });
  }

  throw new Error("Unsupported action. Use action=create|update|delete.");
}

function handleActivityMutation(action, payload) {
  if (action === "create") {
    const activity = normalizeIncomingActivity(payload.activity || payload);
    if (
      !activity.name ||
      !activity.id ||
      (!activity.project && !activity.projectId)
    ) {
      throw new Error(
        "Activity ID, Activity Name, and Project (ID or Name) are required.",
      );
    }
    validateActivityForMutation(activity, action);

    const sheet = getOrCreateSheet(CONFIG.sheetNames.activities);
    ensureSheetHeaders(sheet, CONFIG.headers.activities);
    const columns = getActivityColumnMap(sheet);
    assertActivitySheetColumns(columns);
    ensureActivityDoesNotExist(activity, sheet, columns);
    const lastColumn = Math.max(
      sheet.getLastColumn(),
      CONFIG.headers.activities.length,
      columns.maxColumn,
    );
    const rowValues = new Array(lastColumn).fill("");

    if (columns.projectId)
      rowValues[columns.projectId - 1] = activity.projectId;
    if (columns.project) rowValues[columns.project - 1] = activity.project;
    if (columns.id) rowValues[columns.id - 1] = activity.id;
    if (columns.name) rowValues[columns.name - 1] = activity.name;
    if (columns.status) rowValues[columns.status - 1] = activity.status;
    if (columns.plannedStart)
      rowValues[columns.plannedStart - 1] = activity.plannedStart;
    if (columns.plannedFinish)
      rowValues[columns.plannedFinish - 1] = activity.plannedFinish;
    if (columns.duration) rowValues[columns.duration - 1] = activity.duration;
    if (columns.percentComplete)
      rowValues[columns.percentComplete - 1] = activity.percentComplete;
    if (columns.createdAt) rowValues[columns.createdAt - 1] = getSpreadsheetTodayDate();

    sheet
      .getRange(sheet.getLastRow() + 1, 1, 1, lastColumn)
      .setValues([rowValues]);
    syncProjectProgressFromActivities(activity.projectId, activity.project);

    return jsonResponse({
      ok: true,
      message: "Activity saved successfully.",
      activity: activity,
      generatedAt: new Date().toISOString(),
    });
  }

  if (action === "update") {
    const activity = normalizeIncomingActivity(payload.activity || payload);
    if (!activity.id) throw new Error("Activity ID is required for update.");
    validateActivityForMutation(activity, action);

    const updateResult = updateActivityRow(activity);
    syncProjectProgressFromActivities(updateResult.projectId, updateResult.project);
    return jsonResponse({
      ok: true,
      message: "Activity updated successfully.",
      activity: updateResult,
      generatedAt: new Date().toISOString(),
    });
  }

  if (action === "delete") {
    const activity = normalizeIncomingActivity(payload.activity || payload);
    if (!activity.id) throw new Error("Activity ID is required for delete.");

    const deleteResult = deleteActivityRow(
      activity.id,
      activity.projectId,
      activity.project,
    );
    syncProjectProgressFromActivities(
      deleteResult.projectId || activity.projectId,
      deleteResult.project || activity.project,
    );
    return jsonResponse({
      ok: true,
      message: "Activity and related costs deleted successfully.",
      activityId: activity.id,
      deletedCosts: deleteResult.deletedCosts,
      deletedDailyCosts: deleteResult.deletedDailyCosts,
      generatedAt: new Date().toISOString(),
    });
  }

  throw new Error("Unsupported action. Use action=create|update|delete.");
}

function handleCostMutation(action, payload) {
  if (action === "create" || action === "update") {
    const cost = normalizeIncomingCost(payload.cost || payload);
    if (!cost.projectId || !cost.costId) {
      throw new Error("Project ID and Cost ID are required.");
    }
    assertProjectExists(cost.projectId);
    if (cost.activityId) {
      assertActivityExists(cost.projectId, cost.activityId);
    }

    if (action === "create" && costExists(cost.projectId, cost.costId, cost.activityId)) {
      throw new Error(
        "Cost already exists for this project. Use update instead of create.",
      );
    }

    if (action === "update" && !costExists(cost.projectId, cost.costId, cost.activityId)) {
      throw new Error(
        "Cost record not found for update. Create the cost first.",
      );
    }

    upsertCostRow(cost);
    if (!payload.skipDailyCostSync && !payload.summaryOnly) {
      upsertDailyCostRow(buildDailyCostFromCost(cost));
    }
    return jsonResponse({
      ok: true,
      message:
        action === "create"
          ? "Cost saved successfully."
          : "Cost updated successfully.",
      cost: cost,
      generatedAt: new Date().toISOString(),
    });
  }

  throw new Error("Unsupported action for costs. Use action=create|update.");
}

function handleDailyCostMutation(action, payload) {
  if (action === "create" || action === "update") {
    const dailyCost = normalizeIncomingDailyCost(
      payload.dailyCost || payload.daily_cost || payload,
    );
    if (!dailyCost.projectId || !dailyCost.costId || !dailyCost.date) {
      throw new Error("Project ID, Cost ID, and Date are required.");
    }
    assertProjectExists(dailyCost.projectId);
    assertCostExists(dailyCost.projectId, dailyCost.costId, dailyCost.activityId);
    upsertDailyCostRow(dailyCost);
    syncCostActualFromDailyCost(dailyCost.projectId, dailyCost.costId, {
      activityId: dailyCost.activityId,
    });
    return jsonResponse({
      ok: true,
      message: "Daily cost saved successfully.",
      dailyCost: dailyCost,
      generatedAt: new Date().toISOString(),
    });
  }

  if (action === "delete") {
    const dailyCost = normalizeIncomingDailyCost(
      payload.dailyCost || payload.daily_cost || payload,
    );
    if (!dailyCost.projectId || !dailyCost.costId || !dailyCost.date) {
      throw new Error("Project ID, Cost ID, and Date are required for delete.");
    }
    assertProjectExists(dailyCost.projectId);
    assertCostExists(dailyCost.projectId, dailyCost.costId, dailyCost.activityId);
    deleteDailyCostRow(dailyCost);
    syncCostActualFromDailyCost(dailyCost.projectId, dailyCost.costId, {
      activityId: dailyCost.activityId,
    });
    return jsonResponse({
      ok: true,
      message: "Daily cost deleted successfully.",
      generatedAt: new Date().toISOString(),
    });
  }

  throw new Error(
    "Unsupported action for daily costs. Use action=create|update|delete.",
  );
}

function assertProjectExists(projectId) {
  var normalizedProjectId = cleanText(projectId);
  if (!normalizedProjectId) throw new Error("Project ID is required.");
  if (!findProjectSheetRow(normalizedProjectId)) {
    throw new Error(
      "Project not found. Create the project first before adding related records.",
    );
  }
}

function assertActivityExists(projectId, activityId) {
  var normalizedProjectId = cleanText(projectId);
  var normalizedActivityId = cleanText(activityId);
  if (!normalizedProjectId || !normalizedActivityId) {
    throw new Error("Project ID and Activity ID are required.");
  }

  var lookup = findActivitySheetRow(
    normalizedActivityId,
    normalizedProjectId,
    "",
  );
  if (!lookup) {
    throw new Error(
      "Activity not found for the given Project ID and Activity ID. Create the activity first.",
    );
  }
}

function costExists(projectId, costId, activityId) {
  var normalizedProjectId = cleanText(projectId);
  var normalizedCostId = cleanText(costId);
  var normalizedActivityId = cleanText(activityId);
  if (!normalizedProjectId || !normalizedCostId) return false;

  var sheet = getOrCreateSheet(CONFIG.sheetNames.costs);
  ensureSheetHeaders(sheet, CONFIG.headers.costs);
  var columns = getCostColumnMap(sheet);
  if (!columns.projectId || !columns.costId) {
    throw new Error("Costs sheet is missing Project ID or Cost ID columns.");
  }

  var values = sheet.getDataRange().getValues();
  for (var i = 1; i < values.length; i += 1) {
    if (
      cleanText(values[i][columns.projectId - 1]) === normalizedProjectId &&
      cleanText(values[i][columns.costId - 1]) === normalizedCostId &&
      (!normalizedActivityId ||
        !columns.activityId ||
        cleanText(values[i][columns.activityId - 1]) === normalizedActivityId)
    ) {
      return true;
    }
  }

  return false;
}

function assertCostExists(projectId, costId, activityId) {
  var normalizedProjectId = cleanText(projectId);
  var normalizedCostId = cleanText(costId);
  if (!normalizedProjectId || !normalizedCostId) {
    throw new Error("Project ID and Cost ID are required.");
  }

  if (!costExists(normalizedProjectId, normalizedCostId, activityId)) {
    throw new Error(
      "Cost record not found for the given Project ID, Activity ID, and Cost ID. Create the cost first.",
    );
  }
}

function normalizeIncomingDailyCost(input) {
  var source = input || {};
  var hasExplicitValue = function (value) {
    return value !== undefined && value !== null && value !== "";
  };
  var pickFirstDefined = function (candidates) {
    for (var i = 0; i < candidates.length; i += 1) {
      if (hasExplicitValue(candidates[i])) return candidates[i];
    }
    return "";
  };

  var plannedCost = parseNumber(
    pickFirstDefined([
      source.plannedCost,
      source.planned_cost,
      source.plannedValue,
    ]),
  );
  var plannedCostPerDay = parseNumber(
    pickFirstDefined([source.plannedCostPerDay, source.planned_cost_per_day]),
  );
  var rawProgress = pickFirstDefined([
    source.progress,
    source.percentComplete,
    source.percent_complete,
    source["% Complete"],
  ]);
  var hasProgress = hasExplicitValue(rawProgress);
  var progress = hasProgress ? roundTo(parseNumber(rawProgress), 2) : null;
  var explicitEarnedValue = roundTo(
    parseNumber(
      pickFirstDefined([source.earnedValue, source.earned_value, source.ev]),
    ),
    2,
  );
  var computedEarnedValue =
    plannedCostPerDay > 0 && progress >= 0
      ? roundTo(plannedCostPerDay * (progress / 100), 2)
      : 0;

  return {
    projectId: cleanText(source.projectId || source.project_id),
    project: cleanText(
      source.project || source.projectName || source.project_name,
    ),
    costId: cleanText(
      source.costId ||
        source.cost_id ||
        source.id ||
        source.activityId ||
        source.activity_id,
    ),
    activityId: cleanText(
      source.activityId ||
        source.activity_id ||
        source.sourceActivityId ||
        source.activityRefId ||
        source.activity_ref_id,
    ),
    activity: cleanText(source.activity || source.activityName),
    plannedCost: plannedCost,
    plannedCostPerDay: plannedCostPerDay,
    progress: progress,
    date: normalizeDate(source.date),
    actualCost: parseNumber(
      source.actualCost || source.actual_cost || source.amount,
    ),
    earnedValue:
      explicitEarnedValue >= 0 ? explicitEarnedValue : computedEarnedValue,
  };
}

function dailyCostExists(projectId, costId, date) {
  var normalizedProjectId = cleanText(projectId);
  var normalizedCostId = cleanText(costId);
  var normalizedDate = normalizeDate(date);
  if (!normalizedProjectId || !normalizedCostId || !normalizedDate)
    return false;

  var sheet = getOrCreateSheet(CONFIG.sheetNames.dailyCosts);
  ensureSheetHeaders(sheet, CONFIG.headers.dailyCosts);
  var values = sheet.getDataRange().getValues();
  var columns = getDailyCostColumnMap(sheet);

  for (var i = 1; i < values.length; i += 1) {
    if (
      cleanText(values[i][columns.projectId - 1]) === normalizedProjectId &&
      cleanText(values[i][columns.costId - 1]) === normalizedCostId &&
      normalizeDate(values[i][columns.date - 1]) === normalizedDate
    ) {
      return true;
    }
  }

  return false;
}

function upsertDailyCostRow(dailyCost) {
  var sheet = getOrCreateSheet(CONFIG.sheetNames.dailyCosts);
  ensureSheetHeaders(sheet, CONFIG.headers.dailyCosts);
  var values = sheet.getDataRange().getValues();
  var columns = getDailyCostColumnMap(sheet);
  var rowNumber = -1;
  for (var i = 1; i < values.length; i += 1) {
    if (
      cleanText(values[i][columns.projectId - 1]) === dailyCost.projectId &&
      cleanText(values[i][columns.costId - 1]) === dailyCost.costId &&
      (!dailyCost.activityId ||
        !columns.activityId ||
        cleanText(values[i][columns.activityId - 1]) === dailyCost.activityId) &&
      normalizeDate(values[i][columns.date - 1]) === dailyCost.date
    ) {
      rowNumber = i + 1;
      break;
    }
  }
  var rowLength = Math.max(
    sheet.getLastColumn(),
    CONFIG.headers.dailyCosts.length,
    columns.maxColumn,
  );
  var existingRow =
    rowNumber > 1
      ? sheet.getRange(rowNumber, 1, 1, rowLength).getValues()[0]
      : [];
  var row = new Array(rowLength).fill("");
  for (
    var existingIndex = 0;
    existingIndex < existingRow.length;
    existingIndex += 1
  ) {
    row[existingIndex] = existingRow[existingIndex];
  }
  if (columns.projectId) row[columns.projectId - 1] = dailyCost.projectId;
  if (columns.project) row[columns.project - 1] = dailyCost.project;
  if (columns.costId) row[columns.costId - 1] = dailyCost.costId;
  if (columns.activityId) row[columns.activityId - 1] = dailyCost.activityId;
  if (columns.activity) row[columns.activity - 1] = dailyCost.activity;
  if (columns.plannedCost) row[columns.plannedCost - 1] = dailyCost.plannedCost;
  if (columns.plannedCostPerDay)
    row[columns.plannedCostPerDay - 1] = dailyCost.plannedCostPerDay;
  if (columns.progress && dailyCost.progress !== null)
    row[columns.progress - 1] = dailyCost.progress;
  if (columns.date) row[columns.date - 1] = dailyCost.date;
  if (columns.actualCost) row[columns.actualCost - 1] = dailyCost.actualCost;
  if (columns.earnedValue) row[columns.earnedValue - 1] = dailyCost.earnedValue;
  if (columns.createdAt) {
    var createdAtIndex = columns.createdAt - 1;
    var currentCreatedAt = row[createdAtIndex];
    var normalizedCreatedAt = normalizeDate(currentCreatedAt);
    // Guard against legacy column drift where non-date values (for example,
    // Cost IDs like "C1") land in Created At and appear to mutate on refresh.
    if (!normalizedCreatedAt) {
      row[createdAtIndex] = getSpreadsheetTodayDate();
    }
  }
  var targetRow =
    rowNumber > 1 ? rowNumber : Math.max(sheet.getLastRow() + 1, 2);
  sheet.getRange(targetRow, 1, 1, row.length).setValues([row]);
  applyDailyCostRowFormats(sheet, targetRow, columns);
}

function applyDailyCostRowFormats(sheet, rowNumber, columns) {
  if (columns.plannedCost)
    sheet.getRange(rowNumber, columns.plannedCost).setNumberFormat("#,##0.00");
  if (columns.plannedCostPerDay)
    sheet
      .getRange(rowNumber, columns.plannedCostPerDay)
      .setNumberFormat("#,##0.00");
  if (columns.progress)
    sheet.getRange(rowNumber, columns.progress).setNumberFormat("0.00");
  if (columns.actualCost)
    sheet.getRange(rowNumber, columns.actualCost).setNumberFormat("#,##0.00");
  if (columns.earnedValue)
    sheet.getRange(rowNumber, columns.earnedValue).setNumberFormat("#,##0.00");
  if (columns.date)
    sheet.getRange(rowNumber, columns.date).setNumberFormat("yyyy-mm-dd");
}

function deleteDailyCostRow(dailyCost) {
  var sheet = getOrCreateSheet(CONFIG.sheetNames.dailyCosts);
  ensureSheetHeaders(sheet, CONFIG.headers.dailyCosts);
  var values = sheet.getDataRange().getValues();
  var columns = getDailyCostColumnMap(sheet);
  for (var i = values.length - 1; i >= 1; i -= 1) {
    if (
      cleanText(values[i][columns.projectId - 1]) === dailyCost.projectId &&
      cleanText(values[i][columns.costId - 1]) === dailyCost.costId &&
      (!dailyCost.activityId ||
        !columns.activityId ||
        cleanText(values[i][columns.activityId - 1]) === dailyCost.activityId) &&
      normalizeDate(values[i][columns.date - 1]) === dailyCost.date
    )
      sheet.deleteRow(i + 1);
  }
}

function syncCostActualFromDailyCost(projectId, costId, options) {
  var normalizedProjectId = cleanText(projectId);
  var normalizedCostId = cleanText(costId);
  var normalizedActivityId = cleanText(options && options.activityId);
  var syncProgress = !options || options.syncProgress !== false;
  if (!normalizedProjectId || !normalizedCostId) return;

  var dailySheet = getOrCreateSheet(CONFIG.sheetNames.dailyCosts);
  ensureSheetHeaders(dailySheet, CONFIG.headers.dailyCosts);
  var dailyColumns = getDailyCostColumnMap(dailySheet);
  var dailyValues = dailySheet.getDataRange().getValues();
  var totalActualCost = 0;
  var totalProgress = 0;
  var totalEarnedValue = 0;

  for (var i = 1; i < dailyValues.length; i += 1) {
    if (
      cleanText(dailyValues[i][dailyColumns.projectId - 1]) ===
        normalizedProjectId &&
      cleanText(dailyValues[i][dailyColumns.costId - 1]) === normalizedCostId &&
      (!normalizedActivityId ||
        !dailyColumns.activityId ||
        cleanText(dailyValues[i][dailyColumns.activityId - 1]) ===
          normalizedActivityId)
    ) {
      totalActualCost += parseNumber(
        dailyValues[i][dailyColumns.actualCost - 1],
      );
      if (dailyColumns.progress) {
        totalProgress += parseNumber(dailyValues[i][dailyColumns.progress - 1]);
      }
      if (dailyColumns.earnedValue) {
        totalEarnedValue += parseNumber(
          dailyValues[i][dailyColumns.earnedValue - 1],
        );
      }
    }
  }

  var costsSheet = getOrCreateSheet(CONFIG.sheetNames.costs);
  ensureSheetHeaders(costsSheet, CONFIG.headers.costs);
  var costColumns = getCostColumnMap(costsSheet);
  var costValues = costsSheet.getDataRange().getValues();
  var costHeaders = costValues.length
    ? costValues[0].map(function (cell) {
        return normalizeHeader(cell);
      })
    : [];
  var costProgressColumn = costHeaders.indexOf(normalizeHeader("Progress")) + 1;

  for (var rowIndex = 1; rowIndex < costValues.length; rowIndex += 1) {
    if (
      cleanText(costValues[rowIndex][costColumns.projectId - 1]) ===
        normalizedProjectId &&
      cleanText(costValues[rowIndex][costColumns.costId - 1]) ===
        normalizedCostId &&
      (!normalizedActivityId ||
        !costColumns.activityId ||
        cleanText(costValues[rowIndex][costColumns.activityId - 1]) ===
          normalizedActivityId)
    ) {
      if (costColumns.actualCost) {
        var targetRow = rowIndex + 1;
        costsSheet
          .getRange(targetRow, costColumns.actualCost)
          .setValue(totalActualCost);
        costsSheet
          .getRange(targetRow, costColumns.actualCost)
          .setNumberFormat("#,##0.00");
        if (syncProgress && costProgressColumn > 0) {
          costsSheet.getRange(targetRow, costProgressColumn).setValue(
            roundTo(totalProgress, 2),
          );
          costsSheet
            .getRange(targetRow, costProgressColumn)
            .setNumberFormat("0.00");
        }
        if (costColumns.earnedValue) {
          costsSheet
            .getRange(targetRow, costColumns.earnedValue)
            .setValue(roundTo(totalEarnedValue, 2));
          costsSheet
            .getRange(targetRow, costColumns.earnedValue)
            .setNumberFormat("#,##0.00");
        }
        if (syncProgress) {
          syncActivityProgressFromCost({
            projectId: normalizedProjectId,
            project: costColumns.project
              ? cleanText(costValues[rowIndex][costColumns.project - 1])
              : "",
            activityId: costColumns.activityId
              ? cleanText(costValues[rowIndex][costColumns.activityId - 1])
              : "",
            activity: costColumns.activity
              ? cleanText(costValues[rowIndex][costColumns.activity - 1])
              : "",
            progress: totalProgress,
          });
        }
      }
      return;
    }
  }

  // If a matching cost row does not exist yet, create one from the latest daily-cost context
  // so accumulation is still reflected in the Costs sheet.
  var latestDailyRecord = null;
  for (
    var dailyIndex = dailyValues.length - 1;
    dailyIndex >= 1;
    dailyIndex -= 1
  ) {
    if (
      cleanText(dailyValues[dailyIndex][dailyColumns.projectId - 1]) ===
        normalizedProjectId &&
      cleanText(dailyValues[dailyIndex][dailyColumns.costId - 1]) ===
        normalizedCostId &&
      (!normalizedActivityId ||
        !dailyColumns.activityId ||
        cleanText(dailyValues[dailyIndex][dailyColumns.activityId - 1]) ===
          normalizedActivityId)
    ) {
      latestDailyRecord = dailyValues[dailyIndex];
      break;
    }
  }

  if (!latestDailyRecord) return;

  var generatedCost = {
    projectId: normalizedProjectId,
    costId: normalizedCostId,
    activityId:
      costColumns.activityId && dailyColumns.activityId
        ? cleanText(latestDailyRecord[dailyColumns.activityId - 1])
        : "",
    activity:
      costColumns.activity && dailyColumns.activity
        ? cleanText(latestDailyRecord[dailyColumns.activity - 1])
        : "",
    plannedCost:
      costColumns.plannedCost && dailyColumns.plannedCost
        ? parseNumber(latestDailyRecord[dailyColumns.plannedCost - 1])
        : 0,
    plannedCostPerDay:
      costColumns.plannedCostPerDay && dailyColumns.plannedCostPerDay
        ? parseNumber(latestDailyRecord[dailyColumns.plannedCostPerDay - 1])
        : 0,
    actualCost: totalActualCost,
    progress: roundTo(totalProgress, 2),
    earnedValue: roundTo(totalEarnedValue, 2),
    date:
      costColumns.date && dailyColumns.date
        ? normalizeDate(latestDailyRecord[dailyColumns.date - 1])
        : normalizeDate(new Date()),
  };

  upsertCostRow(generatedCost);
}

function computeEarnedValue(cost) {
  var explicitEarnedValue = parseNumber(cost && cost.earnedValue);

  var plannedCost = parseNumber(cost && cost.plannedCost);
  var plannedCostPerDay = parseNumber(cost && cost.plannedCostPerDay);
  if (plannedCost <= 0 && plannedCostPerDay <= 0)
    return explicitEarnedValue > 0 ? explicitEarnedValue : 0;

  var activitiesSheet = getOrCreateSheet(CONFIG.sheetNames.activities);
  ensureSheetHeaders(activitiesSheet, CONFIG.headers.activities);
  var values = activitiesSheet.getDataRange().getValues();
  if (values.length <= 1) return 0;

  var headers = values[0].map(function (cell) {
    return normalizeHeader(cell);
  });
  var idxProjectId = headers.indexOf(normalizeHeader("Project ID"));
  var idxActivityId = headers.indexOf(normalizeHeader("Activity ID"));
  var idxActivity = headers.indexOf(normalizeHeader("Activity"));
  var idxPercent = ["Progress", "% Complete", "Percent Complete"].reduce(
    function (foundIndex, header) {
      if (foundIndex >= 0) return foundIndex;
      return headers.indexOf(normalizeHeader(header));
    },
    -1,
  );
  if (idxPercent < 0) return 0;

  var targetProjectId = cleanText(cost && cost.projectId);
  var targetCostId = cleanText(cost && cost.costId);
  var targetActivityId = cleanText(cost && cost.activityId);
  var targetActivity = cleanText(cost && cost.activity);

  for (var i = 1; i < values.length; i += 1) {
    var row = values[i];
    var rowProjectId = idxProjectId >= 0 ? cleanText(row[idxProjectId]) : "";
    if (targetProjectId && rowProjectId !== targetProjectId) continue;

    var rowActivityId = idxActivityId >= 0 ? cleanText(row[idxActivityId]) : "";
    var rowActivity = idxActivity >= 0 ? cleanText(row[idxActivity]) : "";
    var matches =
      (targetActivityId && rowActivityId === targetActivityId) ||
      (!targetActivityId && targetActivity && rowActivity === targetActivity) ||
      (targetActivityId &&
        !rowActivityId &&
        targetActivity &&
        rowActivity === targetActivity);
    if (!matches) continue;

    var percent = parseNumber(row[idxPercent]);
    var progressFactor = Math.max(0, Math.min(100, percent)) / 100;

    if (progressFactor > 0) {
      if (plannedCost > 0) {
        return plannedCost * progressFactor;
      }

      var recordedDays = countDailyCostRecords(
        targetProjectId,
        targetCostId,
        targetActivityId,
        targetActivity,
      );
      if (plannedCostPerDay > 0 && recordedDays > 0) {
        return plannedCostPerDay * recordedDays * progressFactor;
      }
    }

    return explicitEarnedValue > 0 ? explicitEarnedValue : 0;
  }

  return explicitEarnedValue > 0 ? explicitEarnedValue : 0;
}

function countDailyCostRecords(projectId, costId, activityId, activityName) {
  var dailySheet = getOrCreateSheet(CONFIG.sheetNames.dailyCosts);
  ensureSheetHeaders(dailySheet, CONFIG.headers.dailyCosts);
  var values = dailySheet.getDataRange().getValues();
  if (values.length <= 1) return 0;

  var headers = values[0].map(function (cell) {
    return normalizeHeader(cell);
  });
  var idxProjectId = headers.indexOf(normalizeHeader("Project ID"));
  var idxCostId = headers.indexOf(normalizeHeader("Cost ID"));
  var idxActivityId = headers.indexOf(normalizeHeader("Activity ID"));
  var idxActivity = headers.indexOf(normalizeHeader("Activity"));

  var normalizedActivityId = cleanText(activityId);
  var normalizedActivityName = cleanText(activityName);
  if (!normalizedActivityId && !normalizedActivityName) return 0;

  var count = 0;
  for (var i = 1; i < values.length; i += 1) {
    var row = values[i];
    var rowProjectId = idxProjectId >= 0 ? cleanText(row[idxProjectId]) : "";
    if (projectId && rowProjectId !== projectId) continue;

    var rowCostId = idxCostId >= 0 ? cleanText(row[idxCostId]) : "";
    if (costId && rowCostId !== costId) continue;

    var rowActivityId = idxActivityId >= 0 ? cleanText(row[idxActivityId]) : "";
    var rowActivity = idxActivity >= 0 ? cleanText(row[idxActivity]) : "";
    var matchesActivity =
      (normalizedActivityId && rowActivityId === normalizedActivityId) ||
      (!normalizedActivityId &&
        normalizedActivityName &&
        rowActivity === normalizedActivityName);

    if (matchesActivity) count += 1;
  }

  return count;
}

function normalizeIncomingCost(input) {
  const source = input || {};
  return {
    costId: cleanText(source.costId || source.id),
    projectId: cleanText(source.projectId || source.project_id),
    project: cleanText(
      source.project || source.projectName || source.project_name,
    ),
    activityId: cleanText(
      source.activityId ||
        source.activity_id ||
        source.sourceActivityId ||
        source.activityRefId ||
        source.activity_ref_id,
    ),
    activity: cleanText(source.activity || source.activityName),
    duration: parseNumber(source.duration || source.durationDays),
    progress: parseNumber(source.progress || source.percentComplete || source.percent_complete),
    category:
      cleanText(
        source.category || source.costCategory || source.cost_category,
      ) || "General",
    date: normalizeDate(source.date),
    plannedCost: parseNumber(
      source.plannedCost || source.planned_cost || source.plannedValue,
    ),
    plannedCostPerDay: parseNumber(
      source.plannedCostPerDay || source.planned_cost_per_day,
    ),
    actualCost: parseNumber(source.actualCost || source.actual_cost),
    earnedValue: parseNumber(
      source.earnedValue || source.earned_value || source.ev,
    ),
    notes: cleanText(source.notes || source.note || source.remarks),
  };
}

function buildDailyCostFromCost(cost) {
  const plannedCost = parseNumber(cost.plannedCost);
  const duration = parseNumber(cost.duration);
  const explicitPerDay = parseNumber(cost.plannedCostPerDay);
  const plannedCostPerDay =
    explicitPerDay || (duration > 0 ? plannedCost / duration : 0);
  const projectId = cleanText(cost.projectId);
  const activityId = cleanText(cost.activityId);
  const activityName = cleanText(cost.activity);
  const progress = roundTo(
    findActivityProgress(projectId, activityId, activityName),
    2,
  );
  const activitySchedule = findActivitySchedule(
    projectId,
    activityId,
    activityName,
  );
  const dailyCostDate =
    normalizeDate(activitySchedule.plannedStart) ||
    normalizeDate(cost.date) ||
    normalizeDate(new Date());

  return {
    projectId: projectId,
    project: cleanText(cost.project),
    costId: cleanText(cost.costId),
    activityId: activityId,
    activity: activityName,
    plannedCost: plannedCost,
    plannedCostPerDay: plannedCostPerDay,
    progress: progress >= 0 ? progress : 0,
    date: dailyCostDate,
    actualCost: parseNumber(cost.actualCost),
  };
}

function upsertCostRow(cost) {
  const sheet = getOrCreateSheet(CONFIG.sheetNames.costs);
  ensureSheetHeaders(sheet, CONFIG.headers.costs);
  const values = sheet.getDataRange().getValues();
  const columns = getCostColumnMap(sheet);
  const lastColumn = Math.max(
    sheet.getLastColumn(),
    CONFIG.headers.costs.length,
    columns.maxColumn,
  );

  let rowNumber = -1;
  for (let i = 1; i < values.length; i += 1) {
    const row = values[i];
    const rowCostId = cleanText(row[columns.costId - 1]);
    const rowProjectId = cleanText(row[columns.projectId - 1]);
    const rowActivityId = columns.activityId
      ? cleanText(row[columns.activityId - 1])
      : "";
    if (
      rowCostId === cost.costId &&
      rowProjectId === cost.projectId &&
      (!cost.activityId || !columns.activityId || rowActivityId === cost.activityId)
    ) {
      rowNumber = i + 1;
      break;
    }
  }

  const existingRowValues =
    rowNumber > 1
      ? sheet.getRange(rowNumber, 1, 1, lastColumn).getValues()[0]
      : [];
  const rowValues = new Array(lastColumn).fill("");
  for (
    let existingIndex = 0;
    existingIndex < existingRowValues.length;
    existingIndex += 1
  ) {
    rowValues[existingIndex] = existingRowValues[existingIndex];
  }
  if (columns.costId) rowValues[columns.costId - 1] = cost.costId;
  if (columns.projectId) rowValues[columns.projectId - 1] = cost.projectId;
  if (columns.project) rowValues[columns.project - 1] = cost.project;
  if (columns.activityId) rowValues[columns.activityId - 1] = cost.activityId;
  if (columns.activity) rowValues[columns.activity - 1] = cost.activity;
  if (columns.duration) rowValues[columns.duration - 1] = cost.duration;
  if (columns.progress) rowValues[columns.progress - 1] = cost.progress;
  syncActivityProgressFromCost({
    projectId: columns.projectId
      ? cleanText(rowValues[columns.projectId - 1])
      : cleanText(cost.projectId),
    project: columns.project
      ? cleanText(rowValues[columns.project - 1])
      : cleanText(cost.project),
    activityId: columns.activityId
      ? cleanText(rowValues[columns.activityId - 1])
      : cleanText(cost.activityId),
    activity: columns.activity
      ? cleanText(rowValues[columns.activity - 1])
      : cleanText(cost.activity),
    progress: columns.progress ? rowValues[columns.progress - 1] : cost.progress,
  });
  if (columns.plannedCost)
    rowValues[columns.plannedCost - 1] = cost.plannedCost;
  if (columns.plannedCostPerDay)
    rowValues[columns.plannedCostPerDay - 1] = cost.plannedCostPerDay;
  if (columns.actualCost) rowValues[columns.actualCost - 1] = cost.actualCost;
  const explicitEarnedValue = parseNumber(cost && cost.earnedValue);
  const computedEarnedValue = Number(computeEarnedValue(cost)) || 0;
  const normalizedEarnedValue =
    explicitEarnedValue > 0 ? explicitEarnedValue : computedEarnedValue;
  if (columns.earnedValue)
    rowValues[columns.earnedValue - 1] = Number(normalizedEarnedValue) || 0;
  if (columns.category) rowValues[columns.category - 1] = cost.category;
  if (columns.date) rowValues[columns.date - 1] = cost.date;
  if (columns.notes) rowValues[columns.notes - 1] = cost.notes;
  if (columns.createdAt) {
    const createdAtIndex = columns.createdAt - 1;
    const currentCreatedAt = rowValues[createdAtIndex];
    const normalizedCreatedAt = normalizeDate(currentCreatedAt);
    if (!normalizedCreatedAt) {
      rowValues[createdAtIndex] = getSpreadsheetTodayDate();
    }
  }

  const targetRow =
    rowNumber > 1 ? rowNumber : Math.max(sheet.getLastRow() + 1, 2);
  sheet.getRange(targetRow, 1, 1, lastColumn).setValues([rowValues]);
  applyCostRowFormats(sheet, targetRow, columns);
}

function applyCostRowFormats(sheet, rowNumber, columns) {
  if (columns.plannedCost)
    sheet.getRange(rowNumber, columns.plannedCost).setNumberFormat("#,##0.00");
  if (columns.plannedCostPerDay)
    sheet
      .getRange(rowNumber, columns.plannedCostPerDay)
      .setNumberFormat("#,##0.00");
  if (columns.progress)
    sheet.getRange(rowNumber, columns.progress).setNumberFormat("0.00");
  if (columns.actualCost)
    sheet.getRange(rowNumber, columns.actualCost).setNumberFormat("#,##0.00");
  if (columns.earnedValue)
    sheet.getRange(rowNumber, columns.earnedValue).setNumberFormat("#,##0.00");
  if (columns.date)
    sheet.getRange(rowNumber, columns.date).setNumberFormat("yyyy-mm-dd");
}

function updateProjectRow(project) {
  const lookup = findProjectSheetRow(project.id);
  if (!lookup) {
    throw new Error("Project not found.");
  }

  const rowValues = lookup.sheet
    .getRange(lookup.rowNumber, 1, 1, lookup.lastColumn)
    .getValues()[0];
  rowValues[lookup.columns.id - 1] = project.id;
  rowValues[lookup.columns.name - 1] = project.name;
  rowValues[lookup.columns.type - 1] = project.type;
  rowValues[lookup.columns.status - 1] = project.status;
  rowValues[lookup.columns.location - 1] = project.location;
  rowValues[lookup.columns.startDate - 1] = project.startDate;
  rowValues[lookup.columns.finishDate - 1] = project.finishDate;
  rowValues[lookup.columns.budget - 1] = project.budget;
  if (lookup.columns.progress) {
    const computedProgress = calculateProjectProgressFromActivities(
      project.id,
      project.name,
    );
    rowValues[lookup.columns.progress - 1] =
      computedProgress === null ? project.progress : computedProgress;
  }
  lookup.sheet
    .getRange(lookup.rowNumber, 1, 1, lookup.lastColumn)
    .setValues([rowValues]);
  return project;
}

function deleteProjectRow(projectId) {
  const normalizedProjectId = cleanText(projectId);
  const lookup = findProjectSheetRow(normalizedProjectId);
  if (!lookup) {
    throw new Error("Project not found.");
  }

  var rowValues = lookup.sheet
    .getRange(lookup.rowNumber, 1, 1, lookup.lastColumn)
    .getValues()[0];
  if (lookup.columns.status) rowValues[lookup.columns.status - 1] = "Archived";
  lookup.sheet
    .getRange(lookup.rowNumber, 1, 1, lookup.lastColumn)
    .setValues([rowValues]);
}

function deleteRowsByProjectId(
  sheetName,
  getColumnMap,
  projectColumnKey,
  projectId,
) {
  const normalizedProjectId = cleanText(projectId);
  if (!normalizedProjectId) return;

  const sheet = getOrCreateSheet(sheetName);
  const expectedHeadersBySheet = {
    [CONFIG.sheetNames.activities]: CONFIG.headers.activities,
    [CONFIG.sheetNames.costs]: CONFIG.headers.costs,
    [CONFIG.sheetNames.dailyCosts]: CONFIG.headers.dailyCosts,
  };

  ensureSheetHeaders(sheet, expectedHeadersBySheet[sheetName] || []);
  const columns = getColumnMap(sheet);
  const projectColumn =
    columns && columns[projectColumnKey] ? columns[projectColumnKey] : 0;
  if (!projectColumn) return;

  const values = sheet.getDataRange().getValues();
  for (var rowIdx = values.length - 1; rowIdx >= 1; rowIdx -= 1) {
    if (cleanText(values[rowIdx][projectColumn - 1]) === normalizedProjectId) {
      sheet.deleteRow(rowIdx + 1);
    }
  }
}

function normalizeIncomingActivity(input) {
  const source = input || {};
  const inferredProject = resolveProjectIdentity(
    cleanText(
      source.projectId ||
        source.project_id ||
        source.projectCode ||
        source.project_code,
    ),
    cleanText(source.project || source.projectName || source.project_name),
  );

  const plannedStart = normalizeDate(
    source.plannedStart ||
      source.planned_start ||
      source.startDate ||
      source.start_date,
  );
  const plannedFinish = normalizeDate(
    source.plannedFinish ||
      source.planned_finish ||
      source.finishDate ||
      source.finish_date,
  );
  const duration =
    cleanText(source.duration || source.durationDays || source.duration_days) ||
    calculateDurationDays(plannedStart, plannedFinish);
  const percentComplete = clampPercent(
    source.percentComplete || source.percent_complete || source.progress,
  );

  const rawId = cleanText(
    source.id ||
      source.activityId ||
      source.activity_id ||
      source.sourceActivityId ||
      source.source_activity_id ||
      source.code ||
      source.activityCode ||
      source.activity_code,
  );
  const normalizedId =
    ["-", "--", "n/a", "na", "none", "null", "undefined"].indexOf(
      rawId.toLowerCase(),
    ) >= 0
      ? ""
      : rawId;

  return {
    id: normalizedId || Utilities.getUuid(),
    projectId: inferredProject.id,
    project: inferredProject.name,
    name: cleanText(
      source.name ||
        source.activity ||
        source.activityName ||
        source.activity_name,
    ),
    status: cleanText(source.status) || "Not Started",
    plannedStart: plannedStart,
    plannedFinish: plannedFinish,
    duration: duration,
    percentComplete: percentComplete,
    notes: cleanText(source.notes || source.note || source.remarks),
  };
}

function resolveProjectIdentity(projectIdInput, projectNameInput) {
  var parsedProjectId = splitIdentityLabel(projectIdInput);
  var parsedProjectName = splitIdentityLabel(projectNameInput);
  var normalizedId = cleanText(parsedProjectId.id || projectIdInput);
  var normalizedName = cleanText(
    (parsedProjectName.name && parsedProjectName.name !== parsedProjectName.id
      ? parsedProjectName.name
      : "") ||
      projectNameInput ||
      parsedProjectId.name,
  );

  if (!normalizedId && !normalizedName) {
    return { id: "", name: "" };
  }

  const lookupSheet = getOrCreateSheet(CONFIG.sheetNames.projects);
  ensureSheetHeaders(lookupSheet, CONFIG.headers.projects);
  const columns = getProjectColumnMap(lookupSheet);

  if (!columns.id && !columns.name) {
    return { id: normalizedId, name: normalizedName };
  }

  const values = lookupSheet.getDataRange().getValues();
  for (var rowIdx = 1; rowIdx < values.length; rowIdx += 1) {
    const row = values[rowIdx];
    const rowId = columns.id ? cleanText(row[columns.id - 1]) : "";
    const rowName = columns.name ? cleanText(row[columns.name - 1]) : "";
    if (!rowId && !rowName) continue;

    const matchesId =
      normalizedId &&
      identityCandidatesMatch(
        getIdentityCandidates(rowId),
        getIdentityCandidates(normalizedId),
      );
    const matchesName =
      normalizedName &&
      identityCandidatesMatch(
        getIdentityCandidates(rowName),
        getIdentityCandidates(normalizedName),
      );
    const matchesCombined = identityCandidatesMatch(
      getIdentityCandidates(rowId, rowName),
      getIdentityCandidates(projectIdInput, projectNameInput),
    );
    if (!matchesId && !matchesName && !matchesCombined) continue;

    return {
      id: rowId || normalizedId,
      name: rowName || normalizedName,
    };
  }

  return { id: normalizedId, name: normalizedName };
}

function getActivityColumnMap(sheet) {
  const headers = sheet
    .getRange(
      1,
      1,
      1,
      Math.max(sheet.getLastColumn(), CONFIG.headers.activities.length),
    )
    .getValues()[0]
    .map(function (header) {
      return normalizeHeader(header);
    });

  const indexOfHeader = function (candidates) {
    for (var i = 0; i < candidates.length; i += 1) {
      var candidate = normalizeHeader(candidates[i]);
      var found = headers.indexOf(candidate);
      if (found >= 0) return found + 1;
    }
    return 0;
  };

  return {
    id: indexOfHeader([
      "Activity ID",
      "ActivityID",
      "Source Activity ID",
      "ID",
    ]),
    projectId: indexOfHeader(["Project ID"]),
    project: indexOfHeader(["Project Name", "Project"]),
    name: indexOfHeader(["Activity", "Activity Name", "Name"]),
    plannedStart: indexOfHeader(["Planned Start", "Start Date"]),
    plannedFinish: indexOfHeader(["Planned Finish", "Finish Date"]),
    duration: indexOfHeader(["Duration", "Duration Days"]),
    status: indexOfHeader(["Status"]),
    percentComplete: indexOfHeader([
      "Progress",
      "% Complete",
      "Percent Complete",
    ]),
    createdAt: indexOfHeader(["Created At"]),
    notes: indexOfHeader(["Notes", "Remarks"]),
    maxColumn: headers.length,
  };
}

function findActivitySheetRow(activityId, projectId, projectName) {
  const sheet = getOrCreateSheet(CONFIG.sheetNames.activities);
  ensureSheetHeaders(sheet, CONFIG.headers.activities);

  const values = sheet.getDataRange().getValues();
  if (!values.length) return null;
  const columns = getActivityColumnMap(sheet);
  if (!columns.id) throw new Error("Activity ID column is missing.");

  var rowNumber = 0;
  for (var rowIdx = 1; rowIdx < values.length; rowIdx += 1) {
    var rowId = cleanText(values[rowIdx][columns.id - 1]);
    if (rowId !== activityId) continue;

    var matchesProjectId =
      !projectId ||
      !columns.projectId ||
      cleanText(values[rowIdx][columns.projectId - 1]) === cleanText(projectId);
    var matchesProjectName =
      !projectName ||
      !columns.project ||
      cleanText(values[rowIdx][columns.project - 1]) === cleanText(projectName);
    if (matchesProjectId && matchesProjectName) {
      rowNumber = rowIdx + 1;
      break;
    }
  }

  if (!rowNumber) return null;

  return {
    sheet: sheet,
    rowNumber: rowNumber,
    columns: columns,
    lastColumn: Math.max(
      sheet.getLastColumn(),
      CONFIG.headers.activities.length,
    ),
  };
}

function updateActivityRow(activity) {
  const lookup = findActivitySheetRow(
    activity.id,
    activity.projectId,
    activity.project,
  );
  if (!lookup) throw new Error("Activity not found.");
  assertActivitySheetColumns(lookup.columns);

  const rowValues = lookup.sheet
    .getRange(lookup.rowNumber, 1, 1, lookup.lastColumn)
    .getValues()[0];
  if (lookup.columns.id) rowValues[lookup.columns.id - 1] = activity.id;
  if (lookup.columns.projectId)
    rowValues[lookup.columns.projectId - 1] = activity.projectId;
  if (lookup.columns.project)
    rowValues[lookup.columns.project - 1] = activity.project;
  if (lookup.columns.name) rowValues[lookup.columns.name - 1] = activity.name;
  if (lookup.columns.plannedStart)
    rowValues[lookup.columns.plannedStart - 1] = activity.plannedStart;
  if (lookup.columns.plannedFinish)
    rowValues[lookup.columns.plannedFinish - 1] = activity.plannedFinish;
  if (lookup.columns.duration)
    rowValues[lookup.columns.duration - 1] = activity.duration;
  if (lookup.columns.status)
    rowValues[lookup.columns.status - 1] = activity.status;
  if (lookup.columns.percentComplete)
    rowValues[lookup.columns.percentComplete - 1] = activity.percentComplete;
  if (lookup.columns.createdAt) {
    const createdAtIndex = lookup.columns.createdAt - 1;
    if (!normalizeDate(rowValues[createdAtIndex])) {
      rowValues[createdAtIndex] = getSpreadsheetTodayDate();
    }
  }
  lookup.sheet
    .getRange(lookup.rowNumber, 1, 1, lookup.lastColumn)
    .setValues([rowValues]);
  return activity;
}

function deleteActivityRow(activityId, projectId, projectName) {
  const lookup = findActivitySheetRow(
    cleanText(activityId),
    cleanText(projectId),
    cleanText(projectName),
  );
  if (!lookup) throw new Error("Activity not found.");

  const rowValues = lookup.sheet
    .getRange(lookup.rowNumber, 1, 1, lookup.lastColumn)
    .getValues()[0];
  const activityContext = {
    activityId: cleanText(activityId),
    projectId:
      cleanText(projectId) ||
      (lookup.columns.projectId
        ? cleanText(rowValues[lookup.columns.projectId - 1])
        : ""),
    project:
      cleanText(projectName) ||
      (lookup.columns.project
        ? cleanText(rowValues[lookup.columns.project - 1])
        : ""),
    activity: lookup.columns.name
      ? cleanText(rowValues[lookup.columns.name - 1])
      : "",
  };

  const deleteResult = deleteCostsRelatedToActivity(activityContext);
  lookup.sheet.deleteRow(lookup.rowNumber);
  return {
    deletedCosts: deleteResult.deletedCosts,
    deletedDailyCosts: deleteResult.deletedDailyCosts,
    projectId: activityContext.projectId,
    project: activityContext.project,
  };
}

function deleteCostsRelatedToActivity(activityContext) {
  const costDeleteResult = deleteCostRowsRelatedToActivity(activityContext);
  const deletedDailyCosts = deleteDailyCostRowsRelatedToActivity(
    activityContext,
    costDeleteResult.deletedCostKeys,
  );

  return {
    deletedCosts: costDeleteResult.deletedCosts,
    deletedDailyCosts: deletedDailyCosts,
  };
}

function deleteCostRowsRelatedToActivity(activityContext) {
  const sheet = getOrCreateSheet(CONFIG.sheetNames.costs);
  ensureSheetHeaders(sheet, CONFIG.headers.costs);
  const values = sheet.getDataRange().getValues();
  const columns = getCostColumnMap(sheet);
  const deletedCostKeys = [];
  var deletedCosts = 0;

  for (var rowIdx = values.length - 1; rowIdx >= 1; rowIdx -= 1) {
    if (!isActivityRelatedRow(values[rowIdx], columns, activityContext)) {
      continue;
    }

    deletedCostKeys.push({
      projectId:
        columns.projectId
          ? cleanText(values[rowIdx][columns.projectId - 1])
          : cleanText(activityContext.projectId),
      project:
        columns.project
          ? cleanText(values[rowIdx][columns.project - 1])
          : cleanText(activityContext.project),
      costId:
        columns.costId
          ? cleanText(values[rowIdx][columns.costId - 1])
          : "",
    });
    sheet.deleteRow(rowIdx + 1);
    deletedCosts += 1;
  }

  return {
    deletedCosts: deletedCosts,
    deletedCostKeys: deletedCostKeys,
  };
}

function deleteDailyCostRowsRelatedToActivity(activityContext, costKeys) {
  const sheet = getOrCreateSheet(CONFIG.sheetNames.dailyCosts);
  ensureSheetHeaders(sheet, CONFIG.headers.dailyCosts);
  const values = sheet.getDataRange().getValues();
  const columns = getDailyCostColumnMap(sheet);
  const normalizedCostKeys = (costKeys || [])
    .map(function (costKey) {
      return {
        projectId: cleanText(costKey && costKey.projectId),
        project: cleanText(costKey && costKey.project),
        costId: cleanText(costKey && costKey.costId),
      };
    })
    .filter(function (costKey) {
      return costKey.costId;
    });
  var deletedDailyCosts = 0;

  for (var rowIdx = values.length - 1; rowIdx >= 1; rowIdx -= 1) {
    if (
      !isActivityRelatedRow(values[rowIdx], columns, activityContext) &&
      !isCostKeyRelatedRow(values[rowIdx], columns, normalizedCostKeys)
    ) {
      continue;
    }

    sheet.deleteRow(rowIdx + 1);
    deletedDailyCosts += 1;
  }

  return deletedDailyCosts;
}

function isActivityRelatedRow(row, columns, activityContext) {
  const normalizedProjectId = cleanText(
    activityContext && activityContext.projectId,
  );
  const normalizedProject = cleanText(
    activityContext && activityContext.project,
  );
  const normalizedActivityId = cleanText(
    activityContext && activityContext.activityId,
  );
  const normalizedActivity = cleanText(
    activityContext && activityContext.activity,
  );

  const rowProjectId = columns.projectId
    ? cleanText(row[columns.projectId - 1])
    : "";
  const rowProject = columns.project
    ? cleanText(row[columns.project - 1])
    : "";
  const rowActivityId = columns.activityId
    ? cleanText(row[columns.activityId - 1])
    : "";
  const rowActivity = columns.activity
    ? cleanText(row[columns.activity - 1])
    : "";

  const hasProjectContext = Boolean(normalizedProjectId || normalizedProject);
  const matchesProject =
    !hasProjectContext ||
    (normalizedProjectId && rowProjectId === normalizedProjectId) ||
    (normalizedProject && rowProject === normalizedProject);
  if (!matchesProject) return false;

  if (normalizedActivityId && rowActivityId === normalizedActivityId) {
    return true;
  }

  return Boolean(
    normalizedActivity &&
      rowActivity === normalizedActivity &&
      (hasProjectContext || !normalizedActivityId),
  );
}

function isCostKeyRelatedRow(row, columns, costKeys) {
  if (!columns.costId || !costKeys || !costKeys.length) return false;

  const rowCostId = cleanText(row[columns.costId - 1]);
  const rowProjectId = columns.projectId
    ? cleanText(row[columns.projectId - 1])
    : "";
  const rowProject = columns.project
    ? cleanText(row[columns.project - 1])
    : "";

  for (var i = 0; i < costKeys.length; i += 1) {
    if (rowCostId !== costKeys[i].costId) continue;

    const hasCostProjectContext = Boolean(
      costKeys[i].projectId || costKeys[i].project,
    );
    if (!hasCostProjectContext || (!rowProjectId && !rowProject)) return true;
    if (costKeys[i].projectId && rowProjectId === costKeys[i].projectId) {
      return true;
    }
    if (costKeys[i].project && rowProject === costKeys[i].project) {
      return true;
    }
  }

  return false;
}

function findProjectSheetRow(projectId) {
  const sheet = getOrCreateSheet(CONFIG.sheetNames.projects);
  ensureSheetHeaders(sheet, CONFIG.headers.projects);

  const values = sheet.getDataRange().getValues();
  if (!values.length) return null;

  const headers = values[0].map(function (header) {
    return normalizeHeader(header);
  });

  const indexOfHeader = function (candidates) {
    for (var i = 0; i < candidates.length; i += 1) {
      var normalized = normalizeHeader(candidates[i]);
      var found = headers.indexOf(normalized);
      if (found >= 0) return found + 1;
    }
    return 0;
  };

  const columns = {
    id: indexOfHeader(["Project ID", "ID"]),
    name: indexOfHeader(["Project Name", "Project", "Name"]),
    type: indexOfHeader(["Project Type", "Type"]),
    status: indexOfHeader(["Status"]),
    location: indexOfHeader(["Location", "Site", "Address"]),
    startDate: indexOfHeader(["Start Date", "Planned Start"]),
    finishDate: indexOfHeader(["Finish Date", "End Date", "Target Finish"]),
    budget: indexOfHeader(["Budget", "Planned Value", "Planned Cost"]),
    progress: indexOfHeader(["Progress", "% Complete", "Percent Complete"]),
    createdAt: indexOfHeader(["Created At"]),
  };

  if (!columns.id) {
    throw new Error("Project ID column is missing.");
  }

  var rowNumber = 0;
  for (var rowIdx = 1; rowIdx < values.length; rowIdx += 1) {
    if (cleanText(values[rowIdx][columns.id - 1]) === projectId) {
      rowNumber = rowIdx + 1;
      break;
    }
  }

  if (!rowNumber) return null;

  return {
    sheet: sheet,
    rowNumber: rowNumber,
    columns: columns,
    lastColumn: Math.max(sheet.getLastColumn(), CONFIG.headers.projects.length),
  };
}

function normalizeResource(value) {
  const supported = [
    "dashboard",
    "projects",
    "activities",
    "costs",
    "daily_costs",
    "reports",
    "all",
  ];
  const normalized = cleanText(value).toLowerCase();
  if (
    normalized === "daily-costs" ||
    normalized === "dailycosts" ||
    normalized === "dailycost"
  )
    return "daily_costs";
  return supported.indexOf(normalized) >= 0 ? normalized : "dashboard";
}

function parsePostPayload(e) {
  if (!e) return {};

  const parseFormEncoded = function (raw) {
    const result = {};
    if (!raw) return result;

    raw.split("&").forEach(function (part) {
      if (!part) return;
      const separatorIndex = part.indexOf("=");
      const rawKey = separatorIndex >= 0 ? part.slice(0, separatorIndex) : part;
      const rawValue =
        separatorIndex >= 0 ? part.slice(separatorIndex + 1) : "";
      const key = decodeURIComponent(rawKey.replace(/\+/g, " "));
      const value = decodeURIComponent(rawValue.replace(/\+/g, " "));
      if (key) result[key] = value;
    });

    return result;
  };

  const parameterPayload = e.parameter && e.parameter.payload;
  if (parameterPayload) {
    try {
      return JSON.parse(parameterPayload);
    } catch (error) {
      throw new Error("Invalid payload parameter JSON.");
    }
  }

  if (!e.postData || !e.postData.contents) return {};
  const rawContents = String(e.postData.contents || "");

  if (rawContents.indexOf("payload=") === 0) {
    try {
      const params = parseFormEncoded(rawContents);
      const nestedPayload = params.payload;
      if (nestedPayload) return JSON.parse(nestedPayload);
    } catch (error) {
      throw new Error("Invalid payload parameter JSON.");
    }
  }

  try {
    return JSON.parse(rawContents);
  } catch (error) {
    try {
      const flattened = parseFormEncoded(rawContents);
      if (Object.keys(flattened).length) return flattened;
    } catch (urlError) {
      // Fall through to final error.
    }
    throw new Error("Invalid JSON payload.");
  }
}

function safeParseJson(value) {
  if (!value) return {};
  try {
    return JSON.parse(String(value));
  } catch (error) {
    return {};
  }
}

function normalizeSheetKey(value) {
  return cleanText(value)
    .toLowerCase()
    .replace(/[\s_\-]+/g, "");
}

function getSheetAliases(sheetName) {
  var normalizedName = normalizeSheetKey(sheetName);
  if (normalizedName === normalizeSheetKey(CONFIG.sheetNames.dailyCosts)) {
    return [
      CONFIG.sheetNames.dailyCosts,
      "Daily Cost",
      "Daily Costs",
      "daily_costs",
      "daily-costs",
      "dailycost",
      "dailycosts",
    ];
  }
  if (normalizedName === normalizeSheetKey(CONFIG.sheetNames.costs)) {
    return [CONFIG.sheetNames.costs, "Cost", "costs"];
  }
  return [sheetName];
}

function getOrCreateSheet(sheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const existing = ss.getSheetByName(sheetName);
  if (existing) return existing;

  const aliases = getSheetAliases(sheetName);
  const aliasLookup = {};
  aliases.forEach(function (alias) {
    aliasLookup[normalizeSheetKey(alias)] = true;
  });

  const sheets = ss.getSheets();
  for (var i = 0; i < sheets.length; i += 1) {
    var currentSheet = sheets[i];
    if (aliasLookup[normalizeSheetKey(currentSheet.getName())]) {
      return currentSheet;
    }
  }

  return ss.insertSheet(sheetName);
}

function ensureProjectHeaders(sheet) {
  const expectedHeaders = CONFIG.headers.projects;
  const existingHeaders = sheet
    .getRange(1, 1, 1, expectedHeaders.length)
    .getValues()[0];
  const hasAnyHeader = existingHeaders.some(function (cell) {
    return cleanText(cell) !== "";
  });

  if (!hasAnyHeader) {
    sheet
      .getRange(1, 1, 1, expectedHeaders.length)
      .setValues([expectedHeaders]);
  }
}

function ensureWorkbookStructure() {
  const projectsSheet = getOrCreateSheet(CONFIG.sheetNames.projects);
  const activitiesSheet = getOrCreateSheet(CONFIG.sheetNames.activities);
  const costsSheet = getOrCreateSheet(CONFIG.sheetNames.costs);
  const dailyCostsSheet = getOrCreateSheet(CONFIG.sheetNames.dailyCosts);

  ensureSheetHeaders(projectsSheet, CONFIG.headers.projects);
  ensureSheetHeaders(activitiesSheet, CONFIG.headers.activities);
  ensureSheetHeaders(costsSheet, CONFIG.headers.costs);
  ensureSheetHeaders(dailyCostsSheet, CONFIG.headers.dailyCosts);
}

function isHeaderLikeDailyCostsDataRow(row) {
  if (!row || !row.length) return false;
  const expected = CONFIG.headers.dailyCosts.map(function (header) {
    return normalizeHeader(header);
  });
  const normalizedRow = row.map(function (value) {
    return normalizeHeader(value);
  });
  let matched = 0;
  for (var i = 0; i < expected.length; i += 1) {
    if (!expected[i]) continue;
    if (normalizedRow[i] === expected[i]) matched += 1;
  }
  return matched >= Math.min(expected.length, 6);
}

function migrateLegacyDailyCostsLayoutIfNeeded(sheet) {
  const headers = sheet
    .getRange(1, 1, 1, Math.max(sheet.getLastColumn(), 11))
    .getValues()[0];
  const normalized = headers.map(function (header) {
    return normalizeHeader(header);
  });
  const legacySignatures = [
    [
      "project id",
      "cost id",
      "activity",
      "planned cost",
      "planned cost/day",
      "date",
      "actual cost",
      "created at",
      "actual cost",
      "earned value",
    ],
    [
      "projectid",
      "cost id",
      "activity",
      "planned cost",
      "planned cost/day",
      "date",
      "actual cost",
      "created at",
      "actual cost",
      "earned value",
    ],
  ];
  const signatureMatch = legacySignatures.some(function (signature) {
    return signature.every(function (label, index) {
      return normalized[index] === label;
    });
  });

  const hasLegacyFieldSet =
    normalized.indexOf("project id") >= 0 &&
    normalized.indexOf("cost id") >= 0 &&
    normalized.indexOf("planned cost/day") >= 0 &&
    normalized.indexOf("date") >= 0 &&
    normalized.indexOf("actual cost") >= 0 &&
    normalized.indexOf("earned value") >= 0 &&
    normalized.indexOf("progress") < 0 &&
    normalized.indexOf("activity id") < 0;

  if (!signatureMatch && !hasLegacyFieldSet) return;

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    sheet
      .getRange(1, 1, 1, CONFIG.headers.dailyCosts.length)
      .setValues([CONFIG.headers.dailyCosts]);
    return;
  }

  const sourceWidth = Math.max(sheet.getLastColumn(), 11);
  const values = sheet.getRange(2, 1, lastRow - 1, sourceWidth).getValues();
  const migrated = values
    .map(function (row) {
      const projectId = row[0];
      const costId = row[1];
      const activity = row[2];
      const plannedCost = row[3];
      const plannedCostPerDay = row[4];
      const date = row[5];
      const actualCost = row[6] !== "" ? row[6] : row[8];
      const createdAt = row[7];
      const earnedValue = row[9];

      // Target layout:
      // [Project ID, Project Name, Cost ID, Activity ID, Activity, Progress/Day,
      //  Planned Cost, Planned Cost/Day, Date, Actual Cost/Day, Earned Value/Day, Created At]
      return [
        projectId,
        "",
        costId,
        "",
        activity,
        "",
        plannedCost,
        plannedCostPerDay,
        date,
        actualCost,
        earnedValue,
        createdAt,
      ];
    })
    .filter(function (row) {
      return !isHeaderLikeDailyCostsDataRow(row);
    });

  const targetWidth = CONFIG.headers.dailyCosts.length;
  if (sheet.getMaxColumns() < targetWidth) {
    sheet.insertColumnsAfter(
      sheet.getMaxColumns(),
      targetWidth - sheet.getMaxColumns(),
    );
  }
  sheet.getRange(1, 1, 1, targetWidth).setValues([CONFIG.headers.dailyCosts]);
  sheet.getRange(2, 1, Math.max(lastRow - 1, 1), targetWidth).clearContent();
  sheet.getRange(2, 1, migrated.length, targetWidth).setValues(migrated);
}


function normalizeCostsColumnsIfNeeded(sheet) {
  const targetHeaders = CONFIG.headers.costs;
  const targetWidth = targetHeaders.length;
  const sourceWidth = Math.max(sheet.getLastColumn(), targetWidth);
  const headers = sheet
    .getRange(1, 1, 1, sourceWidth)
    .getValues()[0]
    .map(function (header) {
      return normalizeHeader(header);
    });

  const expectedNormalized = targetHeaders.map(function (header) {
    return normalizeHeader(header);
  });
  const headerAliases = {
    "planned cost": ["planned cost"],
    "actual cost": ["actual cost"],
    "planned cost day": ["planned cost day", "planned cost d"],
  };

  const shouldNormalize =
    headers.indexOf(normalizeHeader("planned cost/day")) >= 0 ||
    headers.indexOf(normalizeHeader("planned cost/d")) >= 0 ||
    expectedNormalized.some(function (expectedHeader, index) {
      const aliases = headerAliases[expectedHeader] || [expectedHeader];
      return aliases.indexOf(headers[index]) < 0;
    });

  if (!shouldNormalize) return;

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    sheet.getRange(1, 1, 1, targetWidth).setValues([targetHeaders]);
    return;
  }

  const values = sheet.getRange(2, 1, lastRow - 1, sourceWidth).getValues();
  const resolvedHeaderIndex = {};
  expectedNormalized.forEach(function (header) {
    const aliases = headerAliases[header] || [header];
    const matchedAlias = aliases.find(function (alias) {
      return headers.lastIndexOf(alias) >= 0;
    });
    resolvedHeaderIndex[header] = matchedAlias
      ? headers.lastIndexOf(matchedAlias)
      : -1;
  });

  const rebuiltRows = values.map(function (row) {
    return expectedNormalized.map(function (header) {
      const idx = resolvedHeaderIndex[header];
      if (idx < 0) return "";
      return row[idx];
    });
  });

  if (sheet.getMaxColumns() < targetWidth) {
    sheet.insertColumnsAfter(
      sheet.getMaxColumns(),
      targetWidth - sheet.getMaxColumns(),
    );
  }
  sheet.getRange(1, 1, 1, targetWidth).setValues([targetHeaders]);
  sheet.getRange(2, 1, Math.max(lastRow - 1, 1), sourceWidth).clearContent();
  sheet.getRange(2, 1, rebuiltRows.length, targetWidth).setValues(rebuiltRows);
}

function normalizeDailyCostsColumnsIfNeeded(sheet) {
  const targetHeaders = CONFIG.headers.dailyCosts;
  const targetWidth = targetHeaders.length;
  const sourceWidth = Math.max(sheet.getLastColumn(), targetWidth);
  const headers = sheet
    .getRange(1, 1, 1, sourceWidth)
    .getValues()[0]
    .map(function (header) {
      return normalizeHeader(header);
    });

  const expectedNormalized = targetHeaders.map(function (header) {
    return normalizeHeader(header);
  });
  const headerAliases = {
    "progress day": ["progress day", "progress"],
    "actual cost day": ["actual cost day", "actual cost"],
    "earned value day": ["earned value day", "earned value", "ev"],
  };
  const duplicateExpectedHeaderExists = expectedNormalized.some(
    function (expectedHeader) {
      return (
        headers.indexOf(expectedHeader) !== headers.lastIndexOf(expectedHeader)
      );
    },
  );
  const headerOrderMismatch = expectedNormalized.some(
    function (expectedHeader, index) {
      return headers[index] !== expectedHeader;
    },
  );

  if (!duplicateExpectedHeaderExists && !headerOrderMismatch) return;

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    sheet.getRange(1, 1, 1, targetWidth).setValues([targetHeaders]);
    return;
  }

  const values = sheet.getRange(2, 1, lastRow - 1, sourceWidth).getValues();
  const resolvedHeaderIndex = {};
  expectedNormalized.forEach(function (header) {
    const aliases = headerAliases[header] || [header];
    const matchedAlias = aliases.find(function (alias) {
      return headers.lastIndexOf(alias) >= 0;
    });
    resolvedHeaderIndex[header] = matchedAlias
      ? headers.lastIndexOf(matchedAlias)
      : -1;
  });

  const rebuiltRows = values
    .map(function (row) {
      return expectedNormalized.map(function (header) {
        const idx = resolvedHeaderIndex[header];
        if (idx < 0) return "";
        return row[idx];
      });
    })
    .filter(function (row) {
      return !isHeaderLikeDailyCostsDataRow(row);
    });

  if (sheet.getMaxColumns() < targetWidth) {
    sheet.insertColumnsAfter(
      sheet.getMaxColumns(),
      targetWidth - sheet.getMaxColumns(),
    );
  }
  sheet.getRange(1, 1, 1, targetWidth).setValues([targetHeaders]);
  sheet.getRange(2, 1, Math.max(lastRow - 1, 1), sourceWidth).clearContent();
  sheet.getRange(2, 1, rebuiltRows.length, targetWidth).setValues(rebuiltRows);
}

function getDefaultProjectProgressForStatus(status) {
  const normalized = cleanText(status).toLowerCase();
  if (normalized === "completed") return 100;
  if (normalized === "in progress") return 55;
  if (normalized === "on hold") return 25;
  return 0;
}

function normalizeProjectProgress(value, status) {
  if (value === undefined || value === null || value === "") {
    return getDefaultProjectProgressForStatus(status);
  }
  return Math.max(0, Math.min(100, parseNumber(value)));
}

function normalizeActivitiesColumnsIfNeeded(sheet) {
  const targetHeaders = CONFIG.headers.activities;
  const targetWidth = targetHeaders.length;
  const sourceWidth = Math.max(sheet.getLastColumn(), targetWidth);
  const headers = sheet
    .getRange(1, 1, 1, sourceWidth)
    .getValues()[0]
    .map(function (header) {
      return normalizeHeader(header);
    });
  const expectedNormalized = targetHeaders.map(function (header) {
    return normalizeHeader(header);
  });
  const headerAliases = {
    "project id": ["project id", "projectid", "project code", "code"],
    "project name": ["project name", "project", "name"],
    "activity id": [
      "activity id",
      "activityid",
      "source activity id",
      "activity code",
      "task id",
      "id",
    ],
    activity: ["activity", "activity name", "name"],
    "planned start": ["planned start", "planned_start", "start date"],
    "planned finish": [
      "planned finish",
      "planned_finish",
      "finish date",
      "end date",
    ],
    duration: ["duration", "duration days"],
    status: ["status"],
    progress: ["progress", "% complete", "percent complete", "completion"],
    "created at": ["created at", "created_at", "timestamp", "date created"],
  };
  const duplicateExpectedHeaderExists = expectedNormalized.some(function (header) {
    return header && headers.indexOf(header) !== headers.lastIndexOf(header);
  });
  const headerOrderMismatch = expectedNormalized.some(function (header, index) {
    return headers[index] !== header;
  });

  if (!headerOrderMismatch && !duplicateExpectedHeaderExists) return;

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    sheet.getRange(1, 1, 1, targetWidth).setValues([targetHeaders]);
    return;
  }

  const resolvedHeaderIndex = {};
  expectedNormalized.forEach(function (header) {
    const aliases = headerAliases[header] || [header];
    const matchedAlias = aliases.find(function (alias) {
      return headers.lastIndexOf(normalizeHeader(alias)) >= 0;
    });
    resolvedHeaderIndex[header] = matchedAlias
      ? headers.lastIndexOf(normalizeHeader(matchedAlias))
      : -1;
  });

  const statusIndex = resolvedHeaderIndex.status;
  const values = sheet.getRange(2, 1, lastRow - 1, sourceWidth).getValues();
  const rebuiltRows = values.map(function (row) {
    return expectedNormalized.map(function (header) {
      const idx = resolvedHeaderIndex[header];
      if (idx >= 0) return row[idx];
      if (header === "progress") {
        const status = statusIndex >= 0 ? row[statusIndex] : "";
        return cleanText(status).toLowerCase() === "completed" ? 100 : 0;
      }
      return "";
    });
  });

  if (sheet.getMaxColumns() < targetWidth) {
    sheet.insertColumnsAfter(
      sheet.getMaxColumns(),
      targetWidth - sheet.getMaxColumns(),
    );
  }
  sheet.getRange(1, 1, 1, targetWidth).setValues([targetHeaders]);
  sheet.getRange(2, 1, Math.max(lastRow - 1, 1), sourceWidth).clearContent();
  sheet.getRange(2, 1, rebuiltRows.length, targetWidth).setValues(rebuiltRows);
}

function normalizeProjectsColumnsIfNeeded(sheet) {
  const targetHeaders = CONFIG.headers.projects;
  const targetWidth = targetHeaders.length;
  const sourceWidth = Math.max(sheet.getLastColumn(), targetWidth);
  const headers = sheet
    .getRange(1, 1, 1, sourceWidth)
    .getValues()[0]
    .map(function (header) {
      return normalizeHeader(header);
    });
  const expectedNormalized = targetHeaders.map(function (header) {
    return normalizeHeader(header);
  });
  const headerOrderMismatch = expectedNormalized.some(function (header, index) {
    return headers[index] !== header;
  });
  const duplicateExpectedHeaderExists = expectedNormalized.some(function (header) {
    return header && headers.indexOf(header) !== headers.lastIndexOf(header);
  });

  if (!headerOrderMismatch && !duplicateExpectedHeaderExists) return;

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    sheet.getRange(1, 1, 1, targetWidth).setValues([targetHeaders]);
    return;
  }

  const headerAliases = {
    "project id": ["project id", "id", "projectid", "project code", "code"],
    "project name": ["project name", "project", "name"],
    "project type": ["project type", "type"],
    status: ["status"],
    location: ["location", "site", "address"],
    "start date": ["start date", "planned start", "planned_start"],
    "finish date": [
      "finish date",
      "target finish",
      "end date",
      "planned finish",
      "planned_finish",
    ],
    budget: ["budget", "planned value", "planned cost"],
    progress: ["progress", "% complete", "percent complete"],
    "created at": ["created at", "created_at", "timestamp", "date created"],
  };

  const resolvedHeaderIndex = {};
  expectedNormalized.forEach(function (header) {
    const aliases = headerAliases[header] || [header];
    const matchedAlias = aliases.find(function (alias) {
      return headers.lastIndexOf(normalizeHeader(alias)) >= 0;
    });
    resolvedHeaderIndex[header] = matchedAlias
      ? headers.lastIndexOf(normalizeHeader(matchedAlias))
      : -1;
  });

  const statusIndex = resolvedHeaderIndex.status;
  const values = sheet.getRange(2, 1, lastRow - 1, sourceWidth).getValues();
  const rebuiltRows = values.map(function (row) {
    return expectedNormalized.map(function (header) {
      const idx = resolvedHeaderIndex[header];
      if (idx >= 0) return row[idx];
      if (header === "progress") {
        return getDefaultProjectProgressForStatus(
          statusIndex >= 0 ? row[statusIndex] : "",
        );
      }
      return "";
    });
  });

  if (sheet.getMaxColumns() < targetWidth) {
    sheet.insertColumnsAfter(
      sheet.getMaxColumns(),
      targetWidth - sheet.getMaxColumns(),
    );
  }
  sheet.getRange(1, 1, 1, targetWidth).setValues([targetHeaders]);
  sheet.getRange(2, 1, Math.max(lastRow - 1, 1), sourceWidth).clearContent();
  sheet.getRange(2, 1, rebuiltRows.length, targetWidth).setValues(rebuiltRows);
}

function ensureSheetHeaders(sheet, expectedHeaders) {
  if (!expectedHeaders || !expectedHeaders.length) return;

  const isDailyCostsSheet =
    cleanText(sheet.getName()) === cleanText(CONFIG.sheetNames.dailyCosts);
  const isDailyCostsHeaderSet =
    expectedHeaders.length === CONFIG.headers.dailyCosts.length &&
    normalizeHeader(expectedHeaders[0]) ===
      normalizeHeader(CONFIG.headers.dailyCosts[0]);
  if (isDailyCostsSheet && isDailyCostsHeaderSet) {
    migrateLegacyDailyCostsLayoutIfNeeded(sheet);
    normalizeDailyCostsColumnsIfNeeded(sheet);
  }

  const isCostsSheet =
    cleanText(sheet.getName()) === cleanText(CONFIG.sheetNames.costs);
  const isCostsHeaderSet =
    expectedHeaders.length === CONFIG.headers.costs.length &&
    normalizeHeader(expectedHeaders[0]) ===
      normalizeHeader(CONFIG.headers.costs[0]);
  if (isCostsSheet && isCostsHeaderSet) {
    normalizeCostsColumnsIfNeeded(sheet);
  }

  const isActivitiesSheet =
    cleanText(sheet.getName()) === cleanText(CONFIG.sheetNames.activities);
  const isActivitiesHeaderSet =
    expectedHeaders.length === CONFIG.headers.activities.length &&
    normalizeHeader(expectedHeaders[0]) ===
      normalizeHeader(CONFIG.headers.activities[0]);
  if (isActivitiesSheet && isActivitiesHeaderSet) {
    normalizeActivitiesColumnsIfNeeded(sheet);
  }

  const normalizeHeaders = function (headers) {
    return headers.map(function (header) {
      return normalizeHeader(header);
    });
  };

  let maxColumns = expectedHeaders.length;
  let lastColumn = Math.max(sheet.getLastColumn(), maxColumns);

  if (sheet.getMaxColumns() < maxColumns) {
    sheet.insertColumnsAfter(
      sheet.getMaxColumns(),
      maxColumns - sheet.getMaxColumns(),
    );
  }

  let firstRow = sheet.getRange(1, 1, 1, lastColumn).getValues()[0];
  const hasAnyHeader = firstRow.some(function (cell) {
    return cleanText(cell) !== "";
  });

  if (!hasAnyHeader) {
    sheet
      .getRange(1, 1, 1, expectedHeaders.length)
      .setValues([expectedHeaders]);
    return;
  }

  const normalizedExpected = normalizeHeaders(expectedHeaders);
  const normalizedExistingFirstRow = normalizeHeaders(firstRow);
  const expectedLookup = {};
  normalizedExpected.forEach(function (header) {
    expectedLookup[header] = true;
  });
  const firstRowHeaderMatchCount = normalizedExistingFirstRow.reduce(function (
    count,
    header,
  ) {
    return count + (header && expectedLookup[header] ? 1 : 0);
  }, 0);
  const firstRowLooksLikeData =
    firstRowHeaderMatchCount <
    Math.max(2, Math.ceil(expectedHeaders.length * 0.35));

  if (firstRowLooksLikeData) {
    sheet.insertRowBefore(1);
    sheet
      .getRange(1, 1, 1, expectedHeaders.length)
      .setValues([expectedHeaders]);
    return;
  }
  const normalizedExisting = normalizedExistingFirstRow;
  const legacyProjectCodeIndex = normalizedExisting.indexOf("project code");
  const expectsProjectCode = normalizedExpected.indexOf("project code") >= 0;
  const expectsDescription = normalizedExpected.indexOf("description") >= 0;

  if (legacyProjectCodeIndex >= 0 && !expectsProjectCode) {
    sheet.deleteColumn(legacyProjectCodeIndex + 1);
    maxColumns = expectedHeaders.length;
    lastColumn = Math.max(sheet.getLastColumn(), maxColumns);
    if (sheet.getMaxColumns() < maxColumns) {
      sheet.insertColumnsAfter(
        sheet.getMaxColumns(),
        maxColumns - sheet.getMaxColumns(),
      );
    }
    firstRow = sheet.getRange(1, 1, 1, lastColumn).getValues()[0];
  }

  var normalizedAfterProjectCodeCleanup = normalizeHeaders(firstRow);
  const legacyDescriptionIndex =
    normalizedAfterProjectCodeCleanup.indexOf("description");

  if (legacyDescriptionIndex >= 0 && !expectsDescription) {
    sheet.deleteColumn(legacyDescriptionIndex + 1);
    maxColumns = expectedHeaders.length;
    lastColumn = Math.max(sheet.getLastColumn(), maxColumns);
    if (sheet.getMaxColumns() < maxColumns) {
      sheet.insertColumnsAfter(
        sheet.getMaxColumns(),
        maxColumns - sheet.getMaxColumns(),
      );
    }
    firstRow = sheet.getRange(1, 1, 1, lastColumn).getValues()[0];
  }

  const normalizedWithoutLegacy = normalizeHeaders(firstRow);
  const duplicateCreatedAtIndexes = [];
  const expectedLastHeader = normalizedExpected[normalizedExpected.length - 1];
  for (
    var idx = expectedHeaders.length;
    idx < normalizedWithoutLegacy.length;
    idx += 1
  ) {
    if (
      normalizedWithoutLegacy[idx] === expectedLastHeader &&
      expectedLastHeader === "created at"
    ) {
      duplicateCreatedAtIndexes.push(idx + 1);
    }
  }
  for (
    var deleteIdx = duplicateCreatedAtIndexes.length - 1;
    deleteIdx >= 0;
    deleteIdx -= 1
  ) {
    sheet.deleteColumn(duplicateCreatedAtIndexes[deleteIdx]);
  }
  if (duplicateCreatedAtIndexes.length) {
    maxColumns = expectedHeaders.length;
    lastColumn = Math.max(sheet.getLastColumn(), maxColumns);
    if (sheet.getMaxColumns() < maxColumns) {
      sheet.insertColumnsAfter(
        sheet.getMaxColumns(),
        maxColumns - sheet.getMaxColumns(),
      );
    }
    firstRow = sheet.getRange(1, 1, 1, lastColumn).getValues()[0];
  }

  const isProjectsSheet =
    cleanText(sheet.getName()) === cleanText(CONFIG.sheetNames.projects);
  const isProjectsHeaderSet =
    expectedHeaders.length === CONFIG.headers.projects.length &&
    normalizeHeader(expectedHeaders[0]) ===
      normalizeHeader(CONFIG.headers.projects[0]);
  if (isProjectsSheet && isProjectsHeaderSet) {
    normalizeProjectsColumnsIfNeeded(sheet);
    maxColumns = expectedHeaders.length;
    lastColumn = Math.max(sheet.getLastColumn(), maxColumns);
    firstRow = sheet.getRange(1, 1, 1, lastColumn).getValues()[0];
  }

  const normalizedAfter = normalizeHeaders(firstRow);
  const needsHeaderSync = expectedHeaders.some(function (header, idx) {
    return normalizedAfter[idx] !== normalizeHeader(header);
  });

  var hasGapsInsideExpectedRange = false;
  for (
    var expectedIndex = 0;
    expectedIndex < expectedHeaders.length;
    expectedIndex += 1
  ) {
    if (cleanText(firstRow[expectedIndex]) === "") {
      hasGapsInsideExpectedRange = true;
      break;
    }
  }

  if (!needsHeaderSync && !hasGapsInsideExpectedRange) return;

  // Enforce canonical headers whenever a mismatch is detected so column
  // mapping remains stable for reads/writes from the web app.
  sheet.getRange(1, 1, 1, expectedHeaders.length).setValues([expectedHeaders]);
}

function normalizeIncomingProject(input) {
  const source = input || {};

  const id = cleanText(
    source.id ||
      source.projectId ||
      source.project_id ||
      source.projectCode ||
      source.project_code,
  );
  const code = cleanText(
    source.code ||
      source.projectCode ||
      source.project_code ||
      source.projectId ||
      source.project_id,
  );
  const name = cleanText(
    source.name || source.project || source.projectName || source.project_name,
  );
  const type =
    cleanText(source.type || source.projectType || source.project_type) ||
    "General";
  let status = cleanText(source.status) || "Not Started";
  let location = cleanText(source.location || source.site || source.address);
  let startDate = normalizeDate(
    source.startDate || source.start_date || source.plannedStart,
  );
  let finishDate = normalizeDate(
    source.finishDate ||
      source.finish_date ||
      source.targetFinish ||
      source.endDate,
  );
  let budget = parseNumber(source.budget);
  const explicitProgress =
    source.progress !== undefined && source.progress !== null
      ? source.progress
      : source.percentComplete !== undefined && source.percentComplete !== null
        ? source.percentComplete
        : source.percent_complete !== undefined && source.percent_complete !== null
          ? source.percent_complete
          : source["% Complete"];
  const descriptionBudget = parseNumber(source.description || source.notes);
  const createdAtBudget = parseNumber(
    source.createdAt ||
      source.created_at ||
      source.timestamp ||
      source.dateCreated,
  );

  const isKnownStatus = function (value) {
    const normalized = cleanText(value).toLowerCase();
    return (
      [
        "not started",
        "in progress",
        "on hold",
        "completed",
        "archived",
      ].indexOf(normalized) >= 0
    );
  };
  const isDateLike = function (value) {
    if (!value) return false;
    const parsed = new Date(value);
    return !Number.isNaN(parsed.getTime());
  };

  const incomingLooksShifted =
    !isKnownStatus(status) &&
    isKnownStatus(location) &&
    !isDateLike(startDate) &&
    isDateLike(finishDate);

  if (incomingLooksShifted) {
    status = location;
    location = startDate;
    startDate = finishDate;
    finishDate = normalizeDate(source.budget);
    budget = descriptionBudget || createdAtBudget || 0;
  }

  return {
    id: id || code || Utilities.getUuid(),
    code: code || id,
    name: name,
    type: type,
    status: status,
    location: location,
    startDate: startDate,
    finishDate: finishDate,
    budget: budget,
    progress: normalizeProjectProgress(explicitProgress, status),
  };
}

function getProjectColumnMap(sheet) {
  const headers = sheet
    .getRange(
      1,
      1,
      1,
      Math.max(sheet.getLastColumn(), CONFIG.headers.projects.length),
    )
    .getValues()[0]
    .map(function (header) {
      return normalizeHeader(header);
    });

  const indexOfHeader = function (candidates) {
    for (var i = 0; i < candidates.length; i += 1) {
      var candidate = normalizeHeader(candidates[i]);
      var found = headers.indexOf(candidate);
      if (found >= 0) return found + 1;
    }
    return 0;
  };

  return {
    id: indexOfHeader(["Project ID", "ID"]),
    code: indexOfHeader(["Project Code", "Code"]),
    name: indexOfHeader(["Project Name", "Project", "Name"]),
    type: indexOfHeader(["Project Type", "Type"]),
    status: indexOfHeader(["Status"]),
    location: indexOfHeader(["Location", "Site", "Address"]),
    startDate: indexOfHeader(["Start Date", "Planned Start"]),
    finishDate: indexOfHeader(["Finish Date", "End Date", "Target Finish"]),
    budget: indexOfHeader(["Budget", "Planned Value", "Planned Cost"]),
    progress: indexOfHeader(["Progress", "% Complete", "Percent Complete"]),
    createdAt: indexOfHeader(["Created At"]),
    maxColumn: headers.length,
  };
}

function getCostColumnMap(sheet) {
  const headers = sheet
    .getRange(
      1,
      1,
      1,
      Math.max(sheet.getLastColumn(), CONFIG.headers.costs.length),
    )
    .getValues()[0]
    .map(function (header) {
      return normalizeHeader(header);
    });

  const indexOfHeader = function (candidates) {
    for (var i = 0; i < candidates.length; i += 1) {
      var candidate = normalizeHeader(candidates[i]);
      var found = headers.indexOf(candidate);
      if (found >= 0) return found + 1;
    }
    return 0;
  };

  return {
    costId: indexOfHeader(["Cost ID", "ID"]),
    projectId: indexOfHeader(["Project ID"]),
    project: indexOfHeader(["Project Name", "Project"]),
    activityId: indexOfHeader([
      "Activity ID",
      "ActivityID",
      "Source Activity ID",
      "Activity Ref ID",
      "Activity Reference ID",
    ]),
    activity: indexOfHeader(["Activity", "Activity Name"]),
    duration: indexOfHeader(["Duration"]),
    category: indexOfHeader(["Cost Category", "Category", "Type"]),
    date: indexOfHeader(["Date", "Cost Date"]),
    plannedCost: indexOfHeader(["Planned Cost", "Planned Value", "Budget"]),
    plannedCostPerDay: indexOfHeader([
      "Planned Cost/Day",
      "Planned Cost Per Day",
    ]),
    progress: indexOfHeader(["Progress", "% Complete", "Percent Complete"]),
    actualCost: indexOfHeader([
      "Actual Cost/Day",
      "Actual Cost",
      "Cost",
      "Amount",
    ]),
    earnedValue: indexOfHeader(["Earned Value/Day", "Earned Value", "EV"]),
    notes: indexOfHeader(["Notes", "Remarks"]),
    createdAt: indexOfHeader(["Created At"]),
    maxColumn: headers.length,
  };
}

function getDailyCostColumnMap(sheet) {
  const headers = sheet
    .getRange(
      1,
      1,
      1,
      Math.max(sheet.getLastColumn(), CONFIG.headers.dailyCosts.length),
    )
    .getValues()[0]
    .map(function (header) {
      return normalizeHeader(header);
    });

  const indexOfHeader = function (candidates) {
    for (var i = 0; i < candidates.length; i += 1) {
      var candidate = normalizeHeader(candidates[i]);
      var found = headers.lastIndexOf(candidate);
      if (found >= 0) return found + 1;
    }
    return 0;
  };

  return {
    projectId: indexOfHeader(["Project ID", "Project"]),
    project: indexOfHeader(["Project Name", "Project"]),
    costId: indexOfHeader(["Cost ID", "ID"]),
    activityId: indexOfHeader([
      "Activity ID",
      "ActivityID",
      "Source Activity ID",
      "Activity Ref ID",
      "Activity Reference ID",
    ]),
    activity: indexOfHeader(["Activity", "Activity Name"]),
    plannedCost: indexOfHeader(["Planned Cost", "Planned Value", "Budget"]),
    progress: indexOfHeader([
      "Progress/Day",
      "Progress",
      "% Complete",
      "Percent Complete",
    ]),
    plannedCostPerDay: indexOfHeader([
      "Planned Cost/Day",
      "Planned Cost per day",
      "Planned Cost Per Day",
    ]),
    date: indexOfHeader(["Date"]),
    actualCost: indexOfHeader([
      "Actual Cost/Day",
      "Actual Cost",
      "Cost",
      "Amount",
    ]),
    earnedValue: indexOfHeader(["Earned Value/Day", "Earned Value", "EV"]),
    createdAt: indexOfHeader(["Created At"]),
    maxColumn: headers.length,
  };
}

function readSheetRows(ss, sheetName, expectedHeaders) {
  const sheet = getOrCreateSheet(sheetName);
  ensureSheetHeaders(sheet, expectedHeaders || []);

  const rawValues = sheet.getDataRange().getValues();

  if (!rawValues.length) {
    return {
      rows: [],
      meta: {
        sheetName: sheetName,
        headerRowIndex: 0,
        rowCount: 0,
      },
    };
  }

  const stringRows = rawValues.map(function (row) {
    return row.map(function (cell) {
      return String(cell || "");
    });
  });
  const headerRowIndex = findHeaderRowIndex(stringRows);
  const headers = stringRows[headerRowIndex].map(function (header) {
    return String(header || "").trim();
  });

  const rows = rawValues.slice(headerRowIndex + 1).reduce(function (acc, row) {
    const rowObj = {};

    headers.forEach(function (header, index) {
      if (!header) return;
      rowObj[header] = row[index];
    });

    const hasValue = Object.keys(rowObj).some(function (key) {
      return String(rowObj[key] || "").trim() !== "";
    });

    if (hasValue) acc.push(rowObj);
    return acc;
  }, []);

  return {
    rows: rows,
    meta: {
      sheetName: sheetName,
      headerRowIndex: headerRowIndex,
      rowCount: rows.length,
      headers: headers,
    },
  };
}

function findHeaderRowIndex(rows) {
  const expectedHeaderAliases = [
    "project",
    "project id",
    "activity",
    "planned value",
    "actual cost",
    "cost",
    "earned value",
    "budget",
  ];

  let bestIndex = 0;
  let bestScore = 0;

  rows.forEach(function (row, index) {
    const normalizedCells = row
      .map(function (cell) {
        return normalizeHeader(cell);
      })
      .filter(Boolean);

    if (!normalizedCells.length) return;

    const score = expectedHeaderAliases.reduce(function (count, alias) {
      const hasAlias = normalizedCells.some(function (cell) {
        return cell.indexOf(alias) >= 0;
      });
      return count + (hasAlias ? 1 : 0);
    }, 0);

    if (score > bestScore) {
      bestScore = score;
      bestIndex = index;
    }
  });

  return bestIndex;
}

function normalizeProjectRecord(row) {
  const projectId = cleanText(getCell(row, ["project id", "id", "projectid"]));
  let status = cleanText(getCell(row, ["status"])) || "Not Started";
  let location = cleanText(getCell(row, ["location", "site", "address"]));
  let startDate = normalizeDate(
    getCell(row, ["start date", "planned start", "planned_start"]),
  );
  let finishDate = normalizeDate(
    getCell(row, [
      "finish date",
      "end date",
      "planned finish",
      "planned_finish",
    ]),
  );
  let budget = parseNumber(
    getCell(row, ["budget", "planned value", "planned cost"]),
  );
  const descriptionBudget = parseNumber(getCell(row, ["description", "notes"]));
  const createdAtBudget = parseNumber(
    getCell(row, ["created at", "created_at", "timestamp", "date created"]),
  );

  const isKnownStatus = function (value) {
    const normalized = cleanText(value).toLowerCase();
    return (
      [
        "not started",
        "in progress",
        "on hold",
        "completed",
        "archived",
      ].indexOf(normalized) >= 0
    );
  };
  const isDateLike = function (value) {
    if (!value) return false;
    const parsed = new Date(value);
    return !Number.isNaN(parsed.getTime());
  };

  const rowLooksShifted =
    !isKnownStatus(status) &&
    isKnownStatus(location) &&
    !isDateLike(startDate) &&
    isDateLike(finishDate);

  if (rowLooksShifted) {
    status = location;
    location = startDate;
    startDate = finishDate;
    finishDate = normalizeDate(
      getCell(row, ["budget", "planned value", "planned cost"]),
    );
    budget = descriptionBudget || createdAtBudget || 0;
  }

  return {
    id: projectId,
    code: projectId,
    name: cleanText(getCell(row, ["project", "project name", "name"])),
    type: cleanText(getCell(row, ["type", "project type"])) || "General",
    status: status,
    location: location,
    startDate: startDate,
    finishDate: finishDate,
    budget: budget,
    progress: normalizeProjectProgress(
      getCell(row, ["progress", "% complete", "percent complete"]),
      status,
    ),
    raw: row,
  };
}

function normalizeActivityRecord(row) {
  const projectId = cleanText(getCell(row, ["project id", "projectid"]));

  return {
    id: cleanText(
      getCell(row, ["activity id", "id", "activity code", "task id"]),
    ),
    projectId: projectId,
    projectCode: projectId,
    project: cleanText(getCell(row, ["project", "project name"])),
    name: cleanText(getCell(row, ["activity", "activity name", "name"])),
    type: cleanText(getCell(row, ["type", "activity type"])) || "General",
    status: cleanText(getCell(row, ["status"])) || "Not Started",
    plannedStart: normalizeDate(
      getCell(row, ["planned start", "start date", "planned_start"]),
    ),
    plannedFinish: normalizeDate(
      getCell(row, ["planned finish", "finish date", "planned_finish"]),
    ),
    duration: cleanText(
      getCell(row, ["duration", "duration days", "duration_days"]),
    ),
    percentComplete: parseNumber(
      getCell(row, ["progress", "% complete", "percent complete"]),
    ),
    createdAt: normalizeDate(
      getCell(row, ["created at", "created_at", "timestamp", "date created"]),
    ),
    plannedValue: parseNumber(
      getCell(row, ["planned value", "planned cost", "budget"]),
    ),
    actualCost: parseNumber(getCell(row, ["actual cost", "ac", "actual"])),
    earnedValue: parseNumber(getCell(row, ["earned value", "ev"])),
    costVariance: parseNumber(getCell(row, ["cost variance", "cv"])),
    raw: row,
  };
}

function normalizeCostRecord(row) {
  const projectId = cleanText(getCell(row, ["project id", "projectid"]));

  return {
    id: cleanText(getCell(row, ["cost id", "id"])),
    costId: cleanText(getCell(row, ["cost id", "id"])),
    projectId: projectId,
    projectCode: projectId,
    project: cleanText(getCell(row, ["project", "project name"])),
    activityId: cleanText(
      getCell(row, [
        "activity id",
        "activityid",
        "activity_id",
        "id activity",
        "source activity id",
        "activity ref id",
        "activity reference id",
      ]),
    ),
    activity: cleanText(getCell(row, ["activity", "activity name", "name"])),
    duration: parseNumber(
      getCell(row, ["duration", "duration days", "duration_days"]),
    ),
    progress: parseNumber(
      getCell(row, ["progress", "progress/day", "% complete", "percent complete"]),
    ),
    category:
      cleanText(getCell(row, ["cost category", "category", "type"])) ||
      "General",
    date: normalizeDate(
      getCell(row, ["date", "cost date", "transaction date"]),
    ),
    plannedCost: parseNumber(
      getCell(row, ["planned cost", "planned value", "budget"]),
    ),
    plannedCostPerDay: parseNumber(
      getCell(row, [
        "planned cost/day",
        "planned cost per day",
        "planned_cost_per_day",
      ]),
    ),
    actualCost: parseNumber(getCell(row, ["actual cost", "cost", "amount"])),
    earnedValue: parseNumber(getCell(row, ["earned value", "ev"])),
    notes: cleanText(getCell(row, ["note", "notes", "remarks"])),
    raw: row,
  };
}

function normalizeDailyCostRecord(row) {
  return {
    projectId: cleanText(
      row["Project ID"] || row["projectId"] || row["project_id"],
    ),
    project: cleanText(
      row["Project Name"] ||
        row["Project"] ||
        row["project"] ||
        row["projectName"],
    ),
    costId: cleanText(row["Cost ID"] || row["costId"] || row["cost_id"]),
    activityId: cleanText(
      row["Activity ID"] ||
        row["activityId"] ||
        row["activity_id"] ||
        row["Source Activity ID"] ||
        row["Activity Ref ID"] ||
        row["Activity Reference ID"],
    ),
    activity: cleanText(
      row["Activity"] || row["activity"] || row["activityName"],
    ),
    plannedCost: parseNumber(
      row["Planned Cost"] ||
        row["plannedCost"] ||
        row["planned_cost"] ||
        row["plannedValue"],
    ),
    progress: parseNumber(
      row["Progress/Day"] ||
        row["Progress"] ||
        row["progress"] ||
        row["progressPerDay"] ||
        row["progress_per_day"] ||
        row["percentComplete"] ||
        row["% Complete"],
    ),
    plannedCostPerDay: parseNumber(
      row["Planned Cost/Day"] ||
        row["plannedCostPerDay"] ||
        row["planned_cost_per_day"],
    ),
    date: normalizeDate(row["Date"] || row["date"]),
    actualCost: parseNumber(
      row["Actual Cost/Day"] ||
        row["Actual Cost"] ||
        row["actualCost"] ||
        row["actual_cost"],
    ),
    earnedValue: parseNumber(
      row["Earned Value/Day"] ||
        row["Earned Value"] ||
        row["earnedValue"] ||
        row["earned_value"] ||
        row["EV"],
    ),
    createdAt: normalizeDate(
      row["Created At"] || row["createdAt"] || row["created_at"],
    ),
  };
}

function applyProjectFilter(data, filter) {
  const hasFilter = filter.id || filter.name;
  if (!hasFilter) return data;

  const projectMatches = function (record) {
    if (!record) return false;
    const id = cleanText(record.projectId || record.id);
    const name = cleanText(record.project || record.name);

    if (filter.id && filter.id === id) return true;
    if (filter.name && filter.name.toLowerCase() === name.toLowerCase())
      return true;
    return false;
  };

  return {
    projects: data.projects.filter(projectMatches),
    activities: data.activities.filter(projectMatches),
    costs: data.costs.filter(projectMatches),
    dailyCosts: data.dailyCosts.filter(projectMatches),
    sheets: data.sheets,
  };
}

function buildProjectsPayload(projects) {
  const typeCounts = projects.reduce(function (acc, row) {
    const key = row.type || "General";
    acc[key] = (acc[key] || 0) + 1;
    return acc;
  }, {});

  const statusCounts = projects.reduce(function (acc, row) {
    const key = row.status || "Not Started";
    acc[key] = (acc[key] || 0) + 1;
    return acc;
  }, {});

  return {
    count: projects.length,
    projects: projects,
    summary: {
      byType: typeCounts,
      byStatus: statusCounts,
      totalBudget: sumBy(projects, "budget"),
    },
  };
}

function buildActivitiesPayload(activities) {
  const statusCounts = activities.reduce(function (acc, row) {
    const key = row.status || "Not Started";
    acc[key] = (acc[key] || 0) + 1;
    return acc;
  }, {});

  return {
    count: activities.length,
    activities: activities,
    summary: {
      byStatus: statusCounts,
      totalPlannedValue: sumBy(activities, "plannedValue"),
      totalActualCost: sumBy(activities, "actualCost"),
      totalEarnedValue: sumBy(activities, "earnedValue"),
      totalCostVariance: sumBy(activities, "costVariance"),
    },
  };
}

function buildCostsPayload(costs) {
  const byCategory = costs.reduce(function (acc, row) {
    const key = row.category || "General";
    acc[key] = (acc[key] || 0) + row.actualCost;
    return acc;
  }, {});

  return {
    count: costs.length,
    costs: costs,
    summary: {
      byCategory: byCategory,
      totalPlannedCost: sumBy(costs, "plannedCost"),
      totalActualCost: sumBy(costs, "actualCost"),
    },
  };
}

function normalizeDashboardIdentityKey(value) {
  return cleanText(value).toLowerCase();
}

function makeDashboardCompositeKey(projectId, activityId, costId) {
  return [projectId, activityId, costId]
    .map(function (value) {
      return normalizeDashboardIdentityKey(value);
    })
    .join("::");
}

function buildDashboardRows(data) {
  const activities = data.activities || [];
  const costs = data.costs || [];
  const dailyCosts = data.dailyCosts || [];
  const activitiesByProjectAndActivityId = {};
  const dailyCostsByCompositeKey = {};
  const dailyCostsByProjectActivityKey = {};

  const addDailyCostTotal = function (map, key, dailyCost) {
    if (!key) return;
    if (!map[key]) {
      map[key] = { actualCost: 0, earnedValue: 0, progress: 0, count: 0 };
    }

    map[key].actualCost += parseNumber(dailyCost.actualCost);
    map[key].earnedValue += parseNumber(dailyCost.earnedValue);
    map[key].progress += parseNumber(dailyCost.progress);
    map[key].count += 1;
  };

  dailyCosts.forEach(function (dailyCost) {
    const projectId = cleanText(dailyCost.projectId);
    const activityId = cleanText(dailyCost.activityId);
    const costId = cleanText(dailyCost.costId || dailyCost.id);
    if (!projectId || (!activityId && !costId)) return;

    addDailyCostTotal(
      dailyCostsByCompositeKey,
      makeDashboardCompositeKey(projectId, activityId, costId),
      dailyCost,
    );
    if (activityId) {
      addDailyCostTotal(
        dailyCostsByProjectActivityKey,
        makeDashboardCompositeKey(projectId, activityId, ""),
        dailyCost,
      );
    }
  });

  const getDailyTotalsForRow = function (projectId, activityId, costId) {
    return (
      dailyCostsByCompositeKey[
        makeDashboardCompositeKey(projectId, activityId, costId)
      ] ||
      (activityId
        ? dailyCostsByProjectActivityKey[
            makeDashboardCompositeKey(projectId, activityId, "")
          ]
        : null) ||
      null
    );
  };

  activities.forEach(function (activity) {
    const projectId = cleanText(activity.projectId);
    const activityId = cleanText(activity.id || activity.activityId);
    if (!projectId || !activityId) return;
    activitiesByProjectAndActivityId[makeDashboardCompositeKey(projectId, activityId, "")] = activity;
  });

  const rows = [];
  const costBackedActivityKeys = {};

  costs.forEach(function (cost) {
    const projectId = cleanText(cost.projectId);
    const activityId = cleanText(cost.activityId);
    const costId = cleanText(cost.costId || cost.id);
    if (!projectId || (!activityId && !costId)) return;

    const activityKey = makeDashboardCompositeKey(projectId, activityId, "");
    const activity = activitiesByProjectAndActivityId[activityKey] || null;
    if (activityId) costBackedActivityKeys[activityKey] = true;

    const percentComplete = parseNumber(
      activity ? activity.percentComplete : cost.progress,
    );
    const plannedCost = parseNumber(
      cost.plannedCost || (activity && activity.plannedValue),
    );
    const dailyTotals = getDailyTotalsForRow(projectId, activityId, costId);
    const actualCost = dailyTotals
      ? parseNumber(dailyTotals.actualCost)
      : parseNumber(cost.actualCost);
    const rowPercentComplete =
      dailyTotals && parseNumber(dailyTotals.progress)
        ? parseNumber(dailyTotals.progress)
        : percentComplete;
    const earnedValue =
      (dailyTotals && parseNumber(dailyTotals.earnedValue)) ||
      parseNumber(cost.earnedValue) ||
      plannedCost * (rowPercentComplete / 100);
    const costVariance = earnedValue - actualCost;

    rows.push({
      projectId: projectId,
      projectCode: projectId,
      project: cleanText(cost.project || (activity && activity.project)),
      activityId: activityId,
      costId: costId,
      id: activityId || costId,
      name: cleanText(
        cost.activity ||
          (activity && activity.name) ||
          (activityId ? "Activity " + activityId : "Cost " + costId),
      ),
      plannedStart: activity ? activity.plannedStart : "",
      plannedFinish: activity ? activity.plannedFinish : "",
      percentComplete: rowPercentComplete,
      plannedValue: plannedCost,
      actualCost: actualCost,
      earnedValue: earnedValue,
      costVariance: costVariance,
    });
  });

  activities.forEach(function (activity) {
    const projectId = cleanText(activity.projectId);
    const activityId = cleanText(activity.id || activity.activityId);
    const activityKey = makeDashboardCompositeKey(projectId, activityId, "");
    if (!projectId || !activityId || costBackedActivityKeys[activityKey]) return;

    const plannedCost = parseNumber(activity.plannedValue);
    const dailyTotals = getDailyTotalsForRow(projectId, activityId, "");
    const actualCost = dailyTotals
      ? parseNumber(dailyTotals.actualCost)
      : parseNumber(activity.actualCost);
    const percentComplete = dailyTotals && parseNumber(dailyTotals.progress)
      ? parseNumber(dailyTotals.progress)
      : parseNumber(activity.percentComplete);
    const earnedValue =
      (dailyTotals && parseNumber(dailyTotals.earnedValue)) ||
      parseNumber(activity.earnedValue) ||
      plannedCost * (percentComplete / 100);

    rows.push({
      projectId: projectId,
      projectCode: projectId,
      project: cleanText(activity.project),
      activityId: activityId,
      costId: "",
      id: activityId,
      name: cleanText(activity.name || "Activity " + activityId),
      plannedStart: activity.plannedStart,
      plannedFinish: activity.plannedFinish,
      percentComplete: percentComplete,
      plannedValue: plannedCost,
      actualCost: actualCost,
      earnedValue: earnedValue,
      costVariance: earnedValue - actualCost,
    });
  });

  return rows;
}

function buildDashboardPayload(data) {
  const dashboardRows = buildDashboardRows(data);
  const projectCount = data.projects.length;
  const activityCount = data.activities.length;
  const costCount = data.costs.length;

  const totalPlanned = sumBy(dashboardRows, "plannedValue");
  const totalActual = sumBy(dashboardRows, "actualCost");
  const totalEarned = sumBy(dashboardRows, "earnedValue");
  const totalCv = sumBy(dashboardRows, "costVariance");

  const budget = totalPlanned;
  const spentPercent = budget ? (totalActual / budget) * 100 : 0;
  const earnedPercent = budget ? (totalEarned / budget) * 100 : 0;

  return {
    kpi: {
      projectCount: projectCount,
      activityCount: activityCount,
      costEntryCount: costCount,
      totalPlannedValue: totalPlanned,
      totalActualCost: totalActual,
      totalEarnedValue: totalEarned,
      totalCostVariance: totalCv,
      spentPercent: roundTo(spentPercent, 2),
      earnedPercent: roundTo(earnedPercent, 2),
      projectStatus: totalCv >= 0 ? "Under Budget" : "Over Budget",
    },
    rows: dashboardRows,
  };
}

function buildReportsPayload(data) {
  const projectsByName = data.projects.reduce(function (acc, project) {
    const key = project.name || project.code || project.id;
    if (!key) return acc;

    acc[key] = {
      project: project,
      activities: [],
      costs: [],
    };
    return acc;
  }, {});

  data.activities.forEach(function (activity) {
    const key = activity.project || activity.projectId;
    if (!key) return;
    if (!projectsByName[key]) {
      projectsByName[key] = { project: null, activities: [], costs: [] };
    }
    projectsByName[key].activities.push(activity);
  });

  data.costs.forEach(function (cost) {
    const key = cost.project || cost.projectId;
    if (!key) return;
    if (!projectsByName[key]) {
      projectsByName[key] = { project: null, activities: [], costs: [] };
    }
    projectsByName[key].costs.push(cost);
  });

  const projectReports = Object.keys(projectsByName).map(function (projectKey) {
    const bundle = projectsByName[projectKey];
    const totalPlanned = sumBy(bundle.activities, "plannedValue");
    const totalActual = sumBy(bundle.activities, "actualCost");
    const totalEv = sumBy(bundle.activities, "earnedValue");
    const totalCv = sumBy(bundle.activities, "costVariance");

    return {
      projectKey: projectKey,
      project: bundle.project,
      summary: {
        activityCount: bundle.activities.length,
        costEntryCount: bundle.costs.length,
        plannedValue: totalPlanned,
        actualCost: totalActual,
        earnedValue: totalEv,
        costVariance: totalCv,
        status: totalCv >= 0 ? "Under Budget" : "Over Budget",
      },
      activities: bundle.activities,
      costs: bundle.costs,
    };
  });

  return {
    count: projectReports.length,
    reports: projectReports,
  };
}

function buildAllPayload(data) {
  return {
    projects: buildProjectsPayload(data.projects),
    activities: buildActivitiesPayload(data.activities),
    costs: buildCostsPayload(data.costs),
    daily_costs: {
      count: data.dailyCosts.length,
      dailyCosts: data.dailyCosts,
    },
    dashboard: buildDashboardPayload(data),
    reports: buildReportsPayload(data),
  };
}

function sumBy(rows, field) {
  return rows.reduce(function (total, row) {
    return total + parseNumber(row[field]);
  }, 0);
}

function parseNumber(value) {
  if (value === null || value === undefined || value === "") return 0;
  if (typeof value === "number") return Number.isFinite(value) ? value : 0;

  const stringValue = String(value).trim();
  // Guard against date-like strings (e.g., 5/5/2026) being misread as numbers (552026).
  if (
    /^\d{1,2}\/\d{1,2}\/\d{2,4}$/.test(stringValue) ||
    /^\d{4}-\d{1,2}-\d{1,2}$/.test(stringValue)
  ) {
    return 0;
  }
  const isAccountingNegative = /^\(.*\)$/.test(stringValue);
  const cleaned = stringValue.replace(/[^0-9.-]/g, "");
  const parsed = Number(cleaned);

  if (!Number.isFinite(parsed)) return 0;
  return isAccountingNegative ? -Math.abs(parsed) : parsed;
}

function roundTo(value, digits) {
  const factor = Math.pow(10, digits || 0);
  return Math.round(parseNumber(value) * factor) / factor;
}

function normalizeDate(value) {
  if (!value) return "";
  if (value instanceof Date && !Number.isNaN(value.getTime())) {
    return Utilities.formatDate(
      value,
      Session.getScriptTimeZone(),
      "yyyy-MM-dd",
    );
  }

  var textValue = cleanText(value);
  var exactIsoDate = textValue.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (exactIsoDate) {
    return exactIsoDate[1] + "-" + exactIsoDate[2] + "-" + exactIsoDate[3];
  }

  var slashDate = textValue.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (slashDate) {
    var month = ("0" + slashDate[1]).slice(-2);
    var day = ("0" + slashDate[2]).slice(-2);
    return slashDate[3] + "-" + month + "-" + day;
  }

  const parsed = new Date(value);
  if (Number.isNaN(parsed.getTime())) return textValue;
  return Utilities.formatDate(
    parsed,
    Session.getScriptTimeZone(),
    "yyyy-MM-dd",
  );
}

function getCell(row, aliases) {
  const keys = Object.keys(row || {});
  const normalizedKeys = keys.map(function (key) {
    return {
      key: key,
      normalized: normalizeHeader(key),
      compact: compactHeader(key),
    };
  });

  for (let i = 0; i < aliases.length; i += 1) {
    const alias = aliases[i];
    const normalizedAlias = normalizeHeader(alias);
    const compactAlias = compactHeader(alias);

    const exact = normalizedKeys.find(function (entry) {
      return (
        entry.normalized === normalizedAlias || entry.compact === compactAlias
      );
    });
    if (exact) return row[exact.key];

    const shouldUsePrefixMatch = normalizedAlias.indexOf(" ") >= 0;
    if (shouldUsePrefixMatch) {
      const prefixed = normalizedKeys.find(function (entry) {
        return entry.normalized.indexOf(normalizedAlias + " ") === 0;
      });
      if (prefixed) return row[prefixed.key];
    }
  }

  return "";
}

function normalizeHeader(value) {
  return String(value || "")
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}

function compactHeader(value) {
  return normalizeHeader(value).replace(/\s+/g, "");
}

function cleanText(value) {
  return String(value || "").trim();
}

function splitIdentityLabel(value) {
  var text = cleanText(value);
  if (!text) return { id: "", name: "" };

  var separatorMatch = text.match(/^(.+?)\s+(?:-|–|—)\s+(.+)$/);
  if (separatorMatch) {
    return {
      id: cleanText(separatorMatch[1]),
      name: cleanText(separatorMatch[2]),
    };
  }

  return { id: text, name: text };
}

function normalizeIdentityForCompare(value) {
  return cleanText(value).toLowerCase();
}

function getIdentityCandidates() {
  var candidates = [];
  for (var i = 0; i < arguments.length; i += 1) {
    var value = cleanText(arguments[i]);
    if (!value) continue;
    var parsed = splitIdentityLabel(value);
    [value, parsed.id, parsed.name].forEach(function (candidate) {
      var normalized = normalizeIdentityForCompare(candidate);
      if (normalized && candidates.indexOf(normalized) < 0) {
        candidates.push(normalized);
      }
    });
  }
  return candidates;
}

function identityCandidatesMatch(rowCandidates, targetCandidates) {
  if (!rowCandidates.length || !targetCandidates.length) return false;
  return rowCandidates.some(function (candidate) {
    return targetCandidates.indexOf(candidate) >= 0;
  });
}

function clampPercent(value) {
  const parsed = parseNumber(value);
  if (!Number.isFinite(parsed)) return 0;
  return Math.max(0, Math.min(100, parsed));
}

function calculateDurationDays(startDate, finishDate) {
  if (!startDate || !finishDate) return "";
  const start = new Date(startDate);
  const finish = new Date(finishDate);
  if (Number.isNaN(start.getTime()) || Number.isNaN(finish.getTime()))
    return "";
  const dayMs = 24 * 60 * 60 * 1000;
  const diff = Math.floor((finish.getTime() - start.getTime()) / dayMs) + 1;
  if (diff <= 0) return "";
  return String(diff);
}

function validateActivityForMutation(activity, action) {
  if (!activity.projectId && !activity.project) {
    throw new Error(
      "Activity is missing project reference. Provide Project ID or Project Name.",
    );
  }

  if (activity.plannedStart && activity.plannedFinish) {
    const start = new Date(activity.plannedStart);
    const finish = new Date(activity.plannedFinish);
    if (
      !Number.isNaN(start.getTime()) &&
      !Number.isNaN(finish.getTime()) &&
      start.getTime() > finish.getTime()
    ) {
      throw new Error("Planned Start cannot be after Planned Finish.");
    }
  }

  if (action === "create" && !activity.id) {
    throw new Error("Activity ID is required for create.");
  }
}

function assertActivitySheetColumns(columns) {
  if (!columns.id)
    throw new Error('Activities sheet is missing "Activity ID" column.');
  if (!columns.projectId)
    throw new Error('Activities sheet is missing "Project ID" column.');
  if (!columns.project)
    throw new Error('Activities sheet is missing "Project Name" column.');
  if (!columns.name)
    throw new Error('Activities sheet is missing "Activity" column.');
}

function ensureActivityDoesNotExist(activity, sheet, columns) {
  const values = sheet.getDataRange().getValues();
  if (values.length <= 1) return;
  const incomingId = cleanText(activity.id).toLowerCase();
  const incomingProjectId = cleanText(activity.projectId).toLowerCase();
  const incomingProjectName = cleanText(activity.project).toLowerCase();

  for (var rowIdx = 1; rowIdx < values.length; rowIdx += 1) {
    const row = values[rowIdx];
    const rowId = cleanText(row[columns.id - 1]).toLowerCase();
    if (!rowId || rowId !== incomingId) continue;

    const rowProjectId = columns.projectId
      ? cleanText(row[columns.projectId - 1]).toLowerCase()
      : "";
    const rowProjectName = columns.project
      ? cleanText(row[columns.project - 1]).toLowerCase()
      : "";
    const projectIdMatch =
      !incomingProjectId || incomingProjectId === rowProjectId;
    const projectNameMatch =
      !incomingProjectName || incomingProjectName === rowProjectName;

    if (projectIdMatch && projectNameMatch) {
      throw new Error(
        "Activity already exists for this project. Use update instead of create.",
      );
    }
  }
}

function jsonResponse(payload) {
  return ContentService.createTextOutput(JSON.stringify(payload)).setMimeType(
    ContentService.MimeType.JSON,
  );
}
