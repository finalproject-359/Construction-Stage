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
    projects: 'Projects',
    activities: 'Activities',
    costs: 'Costs',
    dailyCosts: 'DailyCosts',
  },
  headers: {
    projects: [
      'Project ID',
      'Project Name',
      'Project Type',
      'Status',
      'Location',
      'Start Date',
      'Finish Date',
      'Budget',
      'Created At',
    ],
    activities: [
      'Project ID',
      'Project Name',
      'Activity ID',
      'Activity',
      'Planned Start',
      'Planned Finish',
      'Duration',
      'Status',
      '% Complete',
      'Notes',
    ],
    costs: [
      'Project ID',
      'Project Name',
      'Cost ID',
      'Activity ID',
      'Activity',
      'Duration',
      'Planned Cost',
      'Planned Cost/Day',
      'Actual Cost',
    ],
    dailyCosts: [
      'Project ID',
      'Cost ID',
      'Activity ID',
      'Date',
      'Actual Cost',
      'Created At',
    ],
  },
};

function doGet(e) {
  const params = (e && e.parameter) || {};
  const payloadParam = params.payload;
  let payload = payloadParam ? safeParseJson(payloadParam) : {};

  if (!payload || typeof payload !== 'object') payload = {};
  payload.resource = payload.resource || params.resource || params.view || params.type;
  payload.action = payload.action || params.action;
  payload.projectId = payload.projectId || params.projectId || params.project_id;

  return handleRequest(payload);
}

function doPost(e) {
  return handleRequest(parsePostPayload(e));
}

function handleRequest(payload) {
  try {
    const source = payload || {};
    const resource = normalizeResource(source.resource || 'dashboard');
    const action = cleanText(source.action).toLowerCase();

    if (action) {
      if (resource === 'projects') {
        const projectsSheet = getOrCreateSheet(CONFIG.sheetNames.projects);
        ensureSheetHeaders(projectsSheet, CONFIG.headers.projects);
        return handleProjectMutation(action, source);
      }

      if (resource === 'activities') {
        const activitiesSheet = getOrCreateSheet(CONFIG.sheetNames.activities);
        ensureSheetHeaders(activitiesSheet, CONFIG.headers.activities);
        return handleActivityMutation(action, source);
      }

      if (resource === 'costs') {
        const costsSheet = getOrCreateSheet(CONFIG.sheetNames.costs);
        ensureSheetHeaders(costsSheet, CONFIG.headers.costs);
        return handleCostMutation(action, source);
      }

      if (resource === 'daily_costs') {
        const dailySheet = getOrCreateSheet(CONFIG.sheetNames.dailyCosts);
        ensureSheetHeaders(dailySheet, CONFIG.headers.dailyCosts);
        return handleDailyCostMutation(action, source);
      }

      throw new Error('Only "projects", "activities", "costs", and "daily_costs" are supported for mutations.');
    }

    const projectFilter = {
      id: cleanText(source.projectId || source.project_id || ''),
      name: cleanText(source.project || source.projectName || source.project_name || ''),
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

    const responsePayload = payloadByResource[resource] || payloadByResource.dashboard;

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
      error: error && error.message ? error.message : 'Unexpected error while processing request.',
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

  const readResource = function(targetKey, sheetName, expectedHeaders, normalizer) {
    const result = readSheetRows(ss, sheetName, expectedHeaders);
    bundle[targetKey] = result.rows.map(normalizer);
    bundle.sheets[targetKey] = result.meta;
  };

  if (resource === 'projects') {
    readResource('projects', CONFIG.sheetNames.projects, CONFIG.headers.projects, normalizeProjectRecord);
    return bundle;
  }

  if (resource === 'activities') {
    readResource('activities', CONFIG.sheetNames.activities, CONFIG.headers.activities, normalizeActivityRecord);
    return bundle;
  }

  if (resource === 'dashboard') {
    readResource('projects', CONFIG.sheetNames.projects, CONFIG.headers.projects, normalizeProjectRecord);
    readResource('activities', CONFIG.sheetNames.activities, CONFIG.headers.activities, normalizeActivityRecord);
    readResource('costs', CONFIG.sheetNames.costs, CONFIG.headers.costs, normalizeCostRecord);
    readResource('dailyCosts', CONFIG.sheetNames.dailyCosts, CONFIG.headers.dailyCosts, normalizeDailyCostRecord);
    return bundle;
  }

  if (resource === 'costs') {
    readResource('costs', CONFIG.sheetNames.costs, CONFIG.headers.costs, normalizeCostRecord);
    return bundle;
  }

  if (resource === 'daily_costs') {
    readResource('dailyCosts', CONFIG.sheetNames.dailyCosts, CONFIG.headers.dailyCosts, normalizeDailyCostRecord);
    return bundle;
  }

  if (resource === 'reports' || resource === 'all') {
    readResource('projects', CONFIG.sheetNames.projects, CONFIG.headers.projects, normalizeProjectRecord);
    readResource('activities', CONFIG.sheetNames.activities, CONFIG.headers.activities, normalizeActivityRecord);
    readResource('costs', CONFIG.sheetNames.costs, CONFIG.headers.costs, normalizeCostRecord);
    readResource('dailyCosts', CONFIG.sheetNames.dailyCosts, CONFIG.headers.dailyCosts, normalizeDailyCostRecord);
    return bundle;
  }

  readResource('projects', CONFIG.sheetNames.projects, CONFIG.headers.projects, normalizeProjectRecord);
  readResource('activities', CONFIG.sheetNames.activities, CONFIG.headers.activities, normalizeActivityRecord);
  readResource('costs', CONFIG.sheetNames.costs, CONFIG.headers.costs, normalizeCostRecord);
  readResource('dailyCosts', CONFIG.sheetNames.dailyCosts, CONFIG.headers.dailyCosts, normalizeDailyCostRecord);
  return bundle;
}

function handleProjectMutation(action, payload) {
  if (action === 'create') {
    const project = normalizeIncomingProject(payload.project || payload);
    if (!project.name || !project.id) {
      throw new Error('Project Name and Project ID are required.');
    }

    const sheet = getOrCreateSheet(CONFIG.sheetNames.projects);
    ensureSheetHeaders(sheet, CONFIG.headers.projects);
    const columns = getProjectColumnMap(sheet);
    const lastColumn = Math.max(sheet.getLastColumn(), CONFIG.headers.projects.length, columns.maxColumn);
    const rowValues = new Array(lastColumn).fill('');
    const storedProjectId = cleanText(project.id || project.code);

    if (columns.id) rowValues[columns.id - 1] = storedProjectId;
    if (columns.name) rowValues[columns.name - 1] = project.name;
    if (columns.type) rowValues[columns.type - 1] = project.type;
    if (columns.status) rowValues[columns.status - 1] = project.status;
    if (columns.location) rowValues[columns.location - 1] = project.location;
    if (columns.startDate) rowValues[columns.startDate - 1] = project.startDate;
    if (columns.finishDate) rowValues[columns.finishDate - 1] = project.finishDate;
    if (columns.budget) rowValues[columns.budget - 1] = project.budget;
    if (columns.createdAt) rowValues[columns.createdAt - 1] = new Date();

    sheet.getRange(sheet.getLastRow() + 1, 1, 1, lastColumn).setValues([rowValues]);

    return jsonResponse({
      ok: true,
      message: 'Project saved successfully.',
      project: {
        ...project,
        id: storedProjectId,
        code: storedProjectId,
      },
      generatedAt: new Date().toISOString(),
    });
  }

  if (action === 'update') {
    const project = normalizeIncomingProject(payload.project || payload);
    if (!project.id) {
      throw new Error('Project ID is required for update.');
    }

    const updateResult = updateProjectRow(project);
    return jsonResponse({
      ok: true,
      message: 'Project updated successfully.',
      project: updateResult,
      generatedAt: new Date().toISOString(),
    });
  }

  if (action === 'delete') {
    const projectId = cleanText(payload.projectId || payload.id);
    if (!projectId) {
      throw new Error('Project ID is required for delete.');
    }

    deleteProjectRow(projectId);
    return jsonResponse({
      ok: true,
      message: 'Project deleted successfully.',
      projectId: projectId,
      generatedAt: new Date().toISOString(),
    });
  }

  throw new Error('Unsupported action. Use action=create|update|delete.');
}

function handleActivityMutation(action, payload) {
  if (action === 'create') {
    const activity = normalizeIncomingActivity(payload.activity || payload);
    if (!activity.name || !activity.id || (!activity.project && !activity.projectId)) {
      throw new Error('Activity ID, Activity Name, and Project (ID or Name) are required.');
    }
    validateActivityForMutation(activity, action);

    const sheet = getOrCreateSheet(CONFIG.sheetNames.activities);
    ensureSheetHeaders(sheet, CONFIG.headers.activities);
    const columns = getActivityColumnMap(sheet);
    assertActivitySheetColumns(columns);
    ensureActivityDoesNotExist(activity, sheet, columns);
    const lastColumn = Math.max(sheet.getLastColumn(), CONFIG.headers.activities.length, columns.maxColumn);
    const rowValues = new Array(lastColumn).fill('');

    if (columns.projectId) rowValues[columns.projectId - 1] = activity.projectId;
    if (columns.project) rowValues[columns.project - 1] = activity.project;
    if (columns.id) rowValues[columns.id - 1] = activity.id;
    if (columns.name) rowValues[columns.name - 1] = activity.name;
    if (columns.status) rowValues[columns.status - 1] = activity.status;
    if (columns.plannedStart) rowValues[columns.plannedStart - 1] = activity.plannedStart;
    if (columns.plannedFinish) rowValues[columns.plannedFinish - 1] = activity.plannedFinish;
    if (columns.duration) rowValues[columns.duration - 1] = activity.duration;
    if (columns.percentComplete) rowValues[columns.percentComplete - 1] = activity.percentComplete;
    if (columns.notes) rowValues[columns.notes - 1] = activity.notes;

    sheet.getRange(sheet.getLastRow() + 1, 1, 1, lastColumn).setValues([rowValues]);

    return jsonResponse({
      ok: true,
      message: 'Activity saved successfully.',
      activity: activity,
      generatedAt: new Date().toISOString(),
    });
  }

  if (action === 'update') {
    const activity = normalizeIncomingActivity(payload.activity || payload);
    if (!activity.id) throw new Error('Activity ID is required for update.');
    validateActivityForMutation(activity, action);

    const updateResult = updateActivityRow(activity);
    return jsonResponse({
      ok: true,
      message: 'Activity updated successfully.',
      activity: updateResult,
      generatedAt: new Date().toISOString(),
    });
  }

  if (action === 'delete') {
    const activity = normalizeIncomingActivity(payload.activity || payload);
    if (!activity.id) throw new Error('Activity ID is required for delete.');

    deleteActivityRow(activity.id, activity.projectId, activity.project);
    return jsonResponse({
      ok: true,
      message: 'Activity deleted successfully.',
      activityId: activity.id,
      generatedAt: new Date().toISOString(),
    });
  }

  throw new Error('Unsupported action. Use action=create|update|delete.');
}

function handleCostMutation(action, payload) {
  if (action === 'create' || action === 'update') {
    const cost = normalizeIncomingCost(payload.cost || payload);
    if (!cost.projectId || !cost.costId) {
      throw new Error('Project ID and Cost ID are required.');
    }

    upsertCostRow(cost);
    return jsonResponse({
      ok: true,
      message: action === 'create' ? 'Cost saved successfully.' : 'Cost updated successfully.',
      cost: cost,
      generatedAt: new Date().toISOString(),
    });
  }

  throw new Error('Unsupported action for costs. Use action=create|update.');
}


function handleDailyCostMutation(action, payload) {
  if (action === 'create' || action === 'update') {
    const dailyCost = normalizeIncomingDailyCost(payload.dailyCost || payload.daily_cost || payload);
    if (!dailyCost.projectId || !dailyCost.costId || !dailyCost.activityId || !dailyCost.date) {
      throw new Error('Project ID, Cost ID, Activity ID, and Date are required.');
    }
    upsertDailyCostRow(dailyCost);
    return jsonResponse({ ok: true, message: 'Daily cost saved successfully.', dailyCost: dailyCost, generatedAt: new Date().toISOString() });
  }

  if (action === 'delete') {
    const dailyCost = normalizeIncomingDailyCost(payload.dailyCost || payload.daily_cost || payload);
    if (!dailyCost.projectId || !dailyCost.activityId || !dailyCost.date) {
      throw new Error('Project ID, Activity ID, and Date are required for delete.');
    }
    deleteDailyCostRow(dailyCost);
    return jsonResponse({ ok: true, message: 'Daily cost deleted successfully.', generatedAt: new Date().toISOString() });
  }

  throw new Error('Unsupported action for daily costs. Use action=create|update|delete.');
}

function normalizeIncomingDailyCost(input) {
  var source = input || {};
  return {
    projectId: cleanText(source.projectId || source.project_id),
    costId: cleanText(source.costId || source.cost_id || source.id),
    activityId: cleanText(source.activityId || source.activity_id),
    date: normalizeDate(source.date),
    actualCost: parseNumber(source.actualCost || source.actual_cost || source.amount),
  };
}

function upsertDailyCostRow(dailyCost) {
  var sheet = getOrCreateSheet(CONFIG.sheetNames.dailyCosts);
  ensureSheetHeaders(sheet, CONFIG.headers.dailyCosts);
  var values = sheet.getDataRange().getValues();
  var headers = sheet.getRange(1, 1, 1, Math.max(sheet.getLastColumn(), CONFIG.headers.dailyCosts.length)).getValues()[0];
  var idx = {
    projectId: headers.indexOf('Project ID'),
    costId: headers.indexOf('Cost ID'),
    activityId: headers.indexOf('Activity ID'),
    date: headers.indexOf('Date'),
    actualCost: headers.indexOf('Actual Cost'),
    createdAt: headers.indexOf('Created At'),
  };
  var rowNumber = -1;
  for (var i = 1; i < values.length; i += 1) {
    if (cleanText(values[i][idx.projectId]) === dailyCost.projectId
      && cleanText(values[i][idx.activityId]) === dailyCost.activityId
      && normalizeDate(values[i][idx.date]) === dailyCost.date) { rowNumber = i + 1; break; }
  }
  var row = new Array(Math.max(sheet.getLastColumn(), CONFIG.headers.dailyCosts.length)).fill('');
  row[idx.projectId] = dailyCost.projectId; row[idx.costId] = dailyCost.costId; row[idx.activityId] = dailyCost.activityId; row[idx.date] = dailyCost.date; row[idx.actualCost] = dailyCost.actualCost; row[idx.createdAt] = new Date();
  if (rowNumber > 0) sheet.getRange(rowNumber, 1, 1, row.length).setValues([row]);
  else sheet.getRange(sheet.getLastRow() + 1, 1, 1, row.length).setValues([row]);
}

function deleteDailyCostRow(dailyCost) {
  var sheet = getOrCreateSheet(CONFIG.sheetNames.dailyCosts);
  ensureSheetHeaders(sheet, CONFIG.headers.dailyCosts);
  var values = sheet.getDataRange().getValues();
  var headers = values[0] || [];
  var p = headers.indexOf('Project ID'); var a = headers.indexOf('Activity ID'); var d = headers.indexOf('Date');
  for (var i = values.length - 1; i >= 1; i -= 1) {
    if (cleanText(values[i][p]) === dailyCost.projectId && cleanText(values[i][a]) === dailyCost.activityId && normalizeDate(values[i][d]) === dailyCost.date) sheet.deleteRow(i + 1);
  }
}

function normalizeIncomingCost(input) {
  const source = input || {};
  return {
    costId: cleanText(source.costId || source.id),
    projectId: cleanText(source.projectId || source.project_id),
    project: cleanText(source.project || source.projectName || source.project_name),
    activityId: cleanText(source.activityId || source.activity_id || source.sourceActivityId),
    activity: cleanText(source.activity || source.activityName),
    duration: parseNumber(source.duration || source.durationDays),
    category: cleanText(source.category || source.costCategory || source.cost_category) || 'General',
    date: normalizeDate(source.date),
    plannedCost: parseNumber(source.plannedCost || source.planned_cost || source.plannedValue),
    plannedCostPerDay: parseNumber(source.plannedCostPerDay || source.planned_cost_per_day),
    actualCost: parseNumber(source.actualCost || source.actual_cost),
    notes: cleanText(source.notes || source.note || source.remarks),
  };
}

function upsertCostRow(cost) {
  const sheet = getOrCreateSheet(CONFIG.sheetNames.costs);
  ensureSheetHeaders(sheet, CONFIG.headers.costs);
  const values = sheet.getDataRange().getValues();
  const columns = getCostColumnMap(sheet);
  const lastColumn = Math.max(sheet.getLastColumn(), CONFIG.headers.costs.length, columns.maxColumn);

  let rowNumber = -1;
  for (let i = 1; i < values.length; i += 1) {
    const row = values[i];
    const rowCostId = cleanText(row[columns.costId - 1]);
    const rowProjectId = cleanText(row[columns.projectId - 1]);
    const rowDate = columns.date ? normalizeDate(row[columns.date - 1]) : '';
    const isSameDate = columns.date ? rowDate === cost.date : true;
    if (rowCostId === cost.costId && rowProjectId === cost.projectId && isSameDate) {
      rowNumber = i + 1;
      break;
    }
  }

  const rowValues = new Array(lastColumn).fill('');
  if (columns.costId) rowValues[columns.costId - 1] = cost.costId;
  if (columns.projectId) rowValues[columns.projectId - 1] = cost.projectId;
  if (columns.project) rowValues[columns.project - 1] = cost.project;
  if (columns.activityId) rowValues[columns.activityId - 1] = cost.activityId;
  if (columns.activity) rowValues[columns.activity - 1] = cost.activity;
  if (columns.duration) rowValues[columns.duration - 1] = cost.duration;
  if (columns.plannedCost) rowValues[columns.plannedCost - 1] = cost.plannedCost;
  if (columns.plannedCostPerDay) rowValues[columns.plannedCostPerDay - 1] = cost.plannedCostPerDay;
  if (columns.actualCost) rowValues[columns.actualCost - 1] = cost.actualCost;
  if (columns.category) rowValues[columns.category - 1] = cost.category;
  if (columns.date) rowValues[columns.date - 1] = cost.date;
  if (columns.notes) rowValues[columns.notes - 1] = cost.notes;
  if (columns.createdAt) rowValues[columns.createdAt - 1] = new Date();

  if (rowNumber > 0) {
    sheet.getRange(rowNumber, 1, 1, lastColumn).setValues([rowValues]);
    return;
  }
  sheet.getRange(sheet.getLastRow() + 1, 1, 1, lastColumn).setValues([rowValues]);
}

function updateProjectRow(project) {
  const lookup = findProjectSheetRow(project.id);
  if (!lookup) {
    throw new Error('Project not found.');
  }

  const rowValues = lookup.sheet.getRange(lookup.rowNumber, 1, 1, lookup.lastColumn).getValues()[0];
  rowValues[lookup.columns.id - 1] = project.id;
  rowValues[lookup.columns.name - 1] = project.name;
  rowValues[lookup.columns.type - 1] = project.type;
  rowValues[lookup.columns.status - 1] = project.status;
  rowValues[lookup.columns.location - 1] = project.location;
  rowValues[lookup.columns.startDate - 1] = project.startDate;
  rowValues[lookup.columns.finishDate - 1] = project.finishDate;
  rowValues[lookup.columns.budget - 1] = project.budget;
  lookup.sheet.getRange(lookup.rowNumber, 1, 1, lookup.lastColumn).setValues([rowValues]);
  return project;
}

function deleteProjectRow(projectId) {
  const lookup = findProjectSheetRow(projectId);
  if (!lookup) {
    throw new Error('Project not found.');
  }

  lookup.sheet.deleteRow(lookup.rowNumber);
}

function normalizeIncomingActivity(input) {
  const source = input || {};
  const inferredProject = resolveProjectIdentity(
    cleanText(source.projectId || source.project_id || source.projectCode || source.project_code),
    cleanText(source.project || source.projectName || source.project_name)
  );

  const plannedStart = normalizeDate(source.plannedStart || source.planned_start || source.startDate || source.start_date);
  const plannedFinish = normalizeDate(source.plannedFinish || source.planned_finish || source.finishDate || source.finish_date);
  const duration = cleanText(source.duration || source.durationDays || source.duration_days) || calculateDurationDays(plannedStart, plannedFinish);
  const percentComplete = clampPercent(source.percentComplete || source.percent_complete || source.progress);

  const rawId = cleanText(source.id || source.activityId || source.activity_id || source.code || source.activityCode || source.activity_code);
  const normalizedId = ['-', '--', 'n/a', 'na', 'none', 'null', 'undefined'].indexOf(rawId.toLowerCase()) >= 0 ? '' : rawId;

  return {
    id: normalizedId || Utilities.getUuid(),
    projectId: inferredProject.id,
    project: inferredProject.name,
    name: cleanText(source.name || source.activity || source.activityName || source.activity_name),
    status: cleanText(source.status) || 'Not Started',
    plannedStart: plannedStart,
    plannedFinish: plannedFinish,
    duration: duration,
    percentComplete: percentComplete,
    notes: cleanText(source.notes || source.note || source.remarks),
  };
}

function resolveProjectIdentity(projectIdInput, projectNameInput) {
  var normalizedId = cleanText(projectIdInput);
  var normalizedName = cleanText(projectNameInput);

  if (!normalizedId && !normalizedName) {
    return { id: '', name: '' };
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
    const rowId = columns.id ? cleanText(row[columns.id - 1]) : '';
    const rowName = columns.name ? cleanText(row[columns.name - 1]) : '';
    if (!rowId && !rowName) continue;

    const matchesId = normalizedId && rowId.toLowerCase() === normalizedId.toLowerCase();
    const matchesName = normalizedName && rowName.toLowerCase() === normalizedName.toLowerCase();
    if (!matchesId && !matchesName) continue;

    return {
      id: rowId || normalizedId,
      name: rowName || normalizedName,
    };
  }

  return { id: normalizedId, name: normalizedName };
}

function getActivityColumnMap(sheet) {
  const headers = sheet.getRange(1, 1, 1, Math.max(sheet.getLastColumn(), CONFIG.headers.activities.length)).getValues()[0]
    .map(function(header) { return normalizeHeader(header); });

  const indexOfHeader = function(candidates) {
    for (var i = 0; i < candidates.length; i += 1) {
      var candidate = normalizeHeader(candidates[i]);
      var found = headers.indexOf(candidate);
      if (found >= 0) return found + 1;
    }
    return 0;
  };

  return {
    id: indexOfHeader(['Activity ID', 'ID']),
    projectId: indexOfHeader(['Project ID']),
    project: indexOfHeader(['Project Name', 'Project']),
    name: indexOfHeader(['Activity', 'Activity Name', 'Name']),
    plannedStart: indexOfHeader(['Planned Start', 'Start Date']),
    plannedFinish: indexOfHeader(['Planned Finish', 'Finish Date']),
    duration: indexOfHeader(['Duration', 'Duration Days']),
    status: indexOfHeader(['Status']),
    percentComplete: indexOfHeader(['% Complete', 'Percent Complete', 'Progress']),
    notes: indexOfHeader(['Notes', 'Remarks']),
    maxColumn: headers.length,
  };
}

function findActivitySheetRow(activityId, projectId, projectName) {
  const sheet = getOrCreateSheet(CONFIG.sheetNames.activities);
  ensureSheetHeaders(sheet, CONFIG.headers.activities);

  const values = sheet.getDataRange().getValues();
  if (!values.length) return null;
  const columns = getActivityColumnMap(sheet);
  if (!columns.id) throw new Error('Activity ID column is missing.');

  var rowNumber = 0;
  for (var rowIdx = 1; rowIdx < values.length; rowIdx += 1) {
    var rowId = cleanText(values[rowIdx][columns.id - 1]);
    if (rowId !== activityId) continue;

    var matchesProjectId = !projectId || !columns.projectId || cleanText(values[rowIdx][columns.projectId - 1]) === cleanText(projectId);
    var matchesProjectName = !projectName || !columns.project || cleanText(values[rowIdx][columns.project - 1]) === cleanText(projectName);
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
    lastColumn: Math.max(sheet.getLastColumn(), CONFIG.headers.activities.length),
  };
}

function updateActivityRow(activity) {
  const lookup = findActivitySheetRow(activity.id, activity.projectId, activity.project);
  if (!lookup) throw new Error('Activity not found.');
  assertActivitySheetColumns(lookup.columns);

  const rowValues = lookup.sheet.getRange(lookup.rowNumber, 1, 1, lookup.lastColumn).getValues()[0];
  if (lookup.columns.id) rowValues[lookup.columns.id - 1] = activity.id;
  if (lookup.columns.projectId) rowValues[lookup.columns.projectId - 1] = activity.projectId;
  if (lookup.columns.project) rowValues[lookup.columns.project - 1] = activity.project;
  if (lookup.columns.name) rowValues[lookup.columns.name - 1] = activity.name;
  if (lookup.columns.plannedStart) rowValues[lookup.columns.plannedStart - 1] = activity.plannedStart;
  if (lookup.columns.plannedFinish) rowValues[lookup.columns.plannedFinish - 1] = activity.plannedFinish;
  if (lookup.columns.duration) rowValues[lookup.columns.duration - 1] = activity.duration;
  if (lookup.columns.status) rowValues[lookup.columns.status - 1] = activity.status;
  if (lookup.columns.percentComplete) rowValues[lookup.columns.percentComplete - 1] = activity.percentComplete;
  if (lookup.columns.notes) rowValues[lookup.columns.notes - 1] = activity.notes;
  lookup.sheet.getRange(lookup.rowNumber, 1, 1, lookup.lastColumn).setValues([rowValues]);
  return activity;
}

function deleteActivityRow(activityId, projectId, projectName) {
  const lookup = findActivitySheetRow(cleanText(activityId), cleanText(projectId), cleanText(projectName));
  if (!lookup) throw new Error('Activity not found.');
  lookup.sheet.deleteRow(lookup.rowNumber);
}

function findProjectSheetRow(projectId) {
  const sheet = getOrCreateSheet(CONFIG.sheetNames.projects);
  ensureSheetHeaders(sheet, CONFIG.headers.projects);

  const values = sheet.getDataRange().getValues();
  if (!values.length) return null;

  const headers = values[0].map(function(header) {
    return normalizeHeader(header);
  });

  const indexOfHeader = function(candidates) {
    for (var i = 0; i < candidates.length; i += 1) {
      var normalized = normalizeHeader(candidates[i]);
      var found = headers.indexOf(normalized);
      if (found >= 0) return found + 1;
    }
    return 0;
  };

  const columns = {
    id: indexOfHeader(['Project ID', 'ID']),
    name: indexOfHeader(['Project Name', 'Project', 'Name']),
    type: indexOfHeader(['Project Type', 'Type']),
    status: indexOfHeader(['Status']),
    location: indexOfHeader(['Location', 'Site', 'Address']),
    startDate: indexOfHeader(['Start Date', 'Planned Start']),
    finishDate: indexOfHeader(['Finish Date', 'End Date', 'Target Finish']),
    budget: indexOfHeader(['Budget', 'Planned Value', 'Planned Cost']),
  };

  if (!columns.id) {
    throw new Error('Project ID column is missing.');
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
  const supported = ['dashboard', 'projects', 'activities', 'costs', 'daily_costs', 'reports', 'all'];
  const normalized = cleanText(value).toLowerCase();
  if (normalized === 'daily-costs' || normalized === 'dailycosts' || normalized === 'dailycost') return 'daily_costs';
  return supported.indexOf(normalized) >= 0 ? normalized : 'dashboard';
}

function parsePostPayload(e) {
  if (!e) return {};

  const parseFormEncoded = function(raw) {
    const result = {};
    if (!raw) return result;

    raw.split('&').forEach(function(part) {
      if (!part) return;
      const separatorIndex = part.indexOf('=');
      const rawKey = separatorIndex >= 0 ? part.slice(0, separatorIndex) : part;
      const rawValue = separatorIndex >= 0 ? part.slice(separatorIndex + 1) : '';
      const key = decodeURIComponent(rawKey.replace(/\+/g, ' '));
      const value = decodeURIComponent(rawValue.replace(/\+/g, ' '));
      if (key) result[key] = value;
    });

    return result;
  };

  const parameterPayload = e.parameter && e.parameter.payload;
  if (parameterPayload) {
    try {
      return JSON.parse(parameterPayload);
    } catch (error) {
      throw new Error('Invalid payload parameter JSON.');
    }
  }

  if (!e.postData || !e.postData.contents) return {};
  const rawContents = String(e.postData.contents || '');

  if (rawContents.indexOf('payload=') === 0) {
    try {
      const params = parseFormEncoded(rawContents);
      const nestedPayload = params.payload;
      if (nestedPayload) return JSON.parse(nestedPayload);
    } catch (error) {
      throw new Error('Invalid payload parameter JSON.');
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
    throw new Error('Invalid JSON payload.');
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

function getOrCreateSheet(sheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const existing = ss.getSheetByName(sheetName);
  if (existing) return existing;
  return ss.insertSheet(sheetName);
}

function ensureProjectHeaders(sheet) {
  const expectedHeaders = CONFIG.headers.projects;
  const existingHeaders = sheet.getRange(1, 1, 1, expectedHeaders.length).getValues()[0];
  const hasAnyHeader = existingHeaders.some(function(cell) {
    return cleanText(cell) !== '';
  });

  if (!hasAnyHeader) {
    sheet.getRange(1, 1, 1, expectedHeaders.length).setValues([expectedHeaders]);
  }
}

function ensureWorkbookStructure() {
  const projectsSheet = getOrCreateSheet(CONFIG.sheetNames.projects);
  const activitiesSheet = getOrCreateSheet(CONFIG.sheetNames.activities);
  const costsSheet = getOrCreateSheet(CONFIG.sheetNames.costs);

  ensureSheetHeaders(projectsSheet, CONFIG.headers.projects);
  ensureSheetHeaders(activitiesSheet, CONFIG.headers.activities);
  ensureSheetHeaders(costsSheet, CONFIG.headers.costs);
}

function ensureSheetHeaders(sheet, expectedHeaders) {
  if (!expectedHeaders || !expectedHeaders.length) return;

  const normalizeHeaders = function(headers) {
    return headers.map(function(header) {
      return normalizeHeader(header);
    });
  };

  let maxColumns = expectedHeaders.length;
  let lastColumn = Math.max(sheet.getLastColumn(), maxColumns);

  if (sheet.getMaxColumns() < maxColumns) {
    sheet.insertColumnsAfter(sheet.getMaxColumns(), maxColumns - sheet.getMaxColumns());
  }

  let firstRow = sheet.getRange(1, 1, 1, lastColumn).getValues()[0];
  const hasAnyHeader = firstRow.some(function(cell) {
    return cleanText(cell) !== '';
  });

  if (!hasAnyHeader) {
    sheet.getRange(1, 1, 1, expectedHeaders.length).setValues([expectedHeaders]);
    return;
  }

  const normalizedExpected = normalizeHeaders(expectedHeaders);
  const normalizedExistingFirstRow = normalizeHeaders(firstRow);
  const expectedLookup = {};
  normalizedExpected.forEach(function(header) {
    expectedLookup[header] = true;
  });
  const firstRowHeaderMatchCount = normalizedExistingFirstRow.reduce(function(count, header) {
    return count + (header && expectedLookup[header] ? 1 : 0);
  }, 0);
  const firstRowLooksLikeData = firstRowHeaderMatchCount < Math.max(2, Math.ceil(expectedHeaders.length * 0.35));

  if (firstRowLooksLikeData) {
    sheet.insertRowBefore(1);
    sheet.getRange(1, 1, 1, expectedHeaders.length).setValues([expectedHeaders]);
    return;
  }
  const normalizedExisting = normalizedExistingFirstRow;
  const legacyProjectCodeIndex = normalizedExisting.indexOf('project code');
  const expectsProjectCode = normalizedExpected.indexOf('project code') >= 0;
  const legacyDescriptionIndex = normalizedExisting.indexOf('description');
  const expectsDescription = normalizedExpected.indexOf('description') >= 0;

  if (legacyProjectCodeIndex >= 0 && !expectsProjectCode) {
    sheet.deleteColumn(legacyProjectCodeIndex + 1);
    maxColumns = expectedHeaders.length;
    lastColumn = Math.max(sheet.getLastColumn(), maxColumns);
    if (sheet.getMaxColumns() < maxColumns) {
      sheet.insertColumnsAfter(sheet.getMaxColumns(), maxColumns - sheet.getMaxColumns());
    }
    firstRow = sheet.getRange(1, 1, 1, lastColumn).getValues()[0];
  }

  if (legacyDescriptionIndex >= 0 && !expectsDescription) {
    sheet.deleteColumn(legacyDescriptionIndex + 1);
    maxColumns = expectedHeaders.length;
    lastColumn = Math.max(sheet.getLastColumn(), maxColumns);
    if (sheet.getMaxColumns() < maxColumns) {
      sheet.insertColumnsAfter(sheet.getMaxColumns(), maxColumns - sheet.getMaxColumns());
    }
    firstRow = sheet.getRange(1, 1, 1, lastColumn).getValues()[0];
  }

  const normalizedWithoutLegacy = normalizeHeaders(firstRow);
  const duplicateCreatedAtIndexes = [];
  const expectedLastHeader = normalizedExpected[normalizedExpected.length - 1];
  for (var idx = expectedHeaders.length; idx < normalizedWithoutLegacy.length; idx += 1) {
    if (normalizedWithoutLegacy[idx] === expectedLastHeader && expectedLastHeader === 'created at') {
      duplicateCreatedAtIndexes.push(idx + 1);
    }
  }
  for (var deleteIdx = duplicateCreatedAtIndexes.length - 1; deleteIdx >= 0; deleteIdx -= 1) {
    sheet.deleteColumn(duplicateCreatedAtIndexes[deleteIdx]);
  }
  if (duplicateCreatedAtIndexes.length) {
    maxColumns = expectedHeaders.length;
    lastColumn = Math.max(sheet.getLastColumn(), maxColumns);
    if (sheet.getMaxColumns() < maxColumns) {
      sheet.insertColumnsAfter(sheet.getMaxColumns(), maxColumns - sheet.getMaxColumns());
    }
    firstRow = sheet.getRange(1, 1, 1, lastColumn).getValues()[0];
  }

  const normalizedAfter = normalizeHeaders(firstRow);
  const needsHeaderSync = expectedHeaders.some(function(header, idx) {
    return normalizedAfter[idx] !== normalizeHeader(header);
  });

  if (needsHeaderSync) {
    sheet.getRange(1, 1, 1, expectedHeaders.length).setValues([expectedHeaders]);
  }
}

function normalizeIncomingProject(input) {
  const source = input || {};

  const id = cleanText(source.id || source.projectId || source.project_id || source.projectCode || source.project_code);
  const code = cleanText(source.code || source.projectCode || source.project_code || source.projectId || source.project_id);
  const name = cleanText(source.name || source.project || source.projectName || source.project_name);
  const type = cleanText(source.type || source.projectType || source.project_type) || 'General';
  let status = cleanText(source.status) || 'Not Started';
  let location = cleanText(source.location || source.site || source.address);
  let startDate = normalizeDate(source.startDate || source.start_date || source.plannedStart);
  let finishDate = normalizeDate(source.finishDate || source.finish_date || source.targetFinish || source.endDate);
  let budget = parseNumber(source.budget);
  const descriptionBudget = parseNumber(source.description || source.notes);
  const createdAtBudget = parseNumber(source.createdAt || source.created_at || source.timestamp || source.dateCreated);

  const isKnownStatus = function(value) {
    const normalized = cleanText(value).toLowerCase();
    return ['not started', 'in progress', 'on hold', 'completed', 'archived'].indexOf(normalized) >= 0;
  };
  const isDateLike = function(value) {
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
  };
}

function getProjectColumnMap(sheet) {
  const headers = sheet.getRange(1, 1, 1, Math.max(sheet.getLastColumn(), CONFIG.headers.projects.length)).getValues()[0]
    .map(function(header) { return normalizeHeader(header); });

  const indexOfHeader = function(candidates) {
    for (var i = 0; i < candidates.length; i += 1) {
      var candidate = normalizeHeader(candidates[i]);
      var found = headers.indexOf(candidate);
      if (found >= 0) return found + 1;
    }
    return 0;
  };

  return {
    id: indexOfHeader(['Project ID', 'ID']),
    code: indexOfHeader(['Project Code', 'Code']),
    name: indexOfHeader(['Project Name', 'Project', 'Name']),
    type: indexOfHeader(['Project Type', 'Type']),
    status: indexOfHeader(['Status']),
    location: indexOfHeader(['Location', 'Site', 'Address']),
    startDate: indexOfHeader(['Start Date', 'Planned Start']),
    finishDate: indexOfHeader(['Finish Date', 'End Date', 'Target Finish']),
    budget: indexOfHeader(['Budget', 'Planned Value', 'Planned Cost']),
    createdAt: indexOfHeader(['Created At']),
    maxColumn: headers.length,
  };
}

function getCostColumnMap(sheet) {
  const headers = sheet.getRange(1, 1, 1, Math.max(sheet.getLastColumn(), CONFIG.headers.costs.length)).getValues()[0]
    .map(function(header) { return normalizeHeader(header); });

  const indexOfHeader = function(candidates) {
    for (var i = 0; i < candidates.length; i += 1) {
      var candidate = normalizeHeader(candidates[i]);
      var found = headers.indexOf(candidate);
      if (found >= 0) return found + 1;
    }
    return 0;
  };

  return {
    costId: indexOfHeader(['Cost ID', 'ID']),
    projectId: indexOfHeader(['Project ID']),
    project: indexOfHeader(['Project Name', 'Project']),
    activityId: indexOfHeader(['Activity ID', 'ActivityID', 'Source Activity ID']),
    activity: indexOfHeader(['Activity', 'Activity Name']),
    duration: indexOfHeader(['Duration']),
    category: indexOfHeader(['Cost Category', 'Category', 'Type']),
    date: indexOfHeader(['Date', 'Cost Date']),
    plannedCost: indexOfHeader(['Planned Cost', 'Planned Value', 'Budget']),
    plannedCostPerDay: indexOfHeader(['Planned Cost/Day', 'Planned Cost Per Day']),
    actualCost: indexOfHeader(['Actual Cost', 'Cost', 'Amount']),
    notes: indexOfHeader(['Notes', 'Remarks']),
    createdAt: indexOfHeader(['Created At']),
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

  const stringRows = rawValues.map(function(row) {
    return row.map(function(cell) {
      return String(cell || '');
    });
  });
  const headerRowIndex = findHeaderRowIndex(stringRows);
  const headers = stringRows[headerRowIndex].map(function(header) {
    return String(header || '').trim();
  });

  const rows = rawValues.slice(headerRowIndex + 1).reduce(function(acc, row) {
    const rowObj = {};

    headers.forEach(function(header, index) {
      if (!header) return;
      rowObj[header] = row[index];
    });

    const hasValue = Object.keys(rowObj).some(function(key) {
      return String(rowObj[key] || '').trim() !== '';
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
    'project',
    'project id',
    'activity',
    'planned value',
    'actual cost',
    'cost',
    'earned value',
    'budget',
  ];

  let bestIndex = 0;
  let bestScore = 0;

  rows.forEach(function(row, index) {
    const normalizedCells = row.map(function(cell) {
      return normalizeHeader(cell);
    }).filter(Boolean);

    if (!normalizedCells.length) return;

    const score = expectedHeaderAliases.reduce(function(count, alias) {
      const hasAlias = normalizedCells.some(function(cell) {
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
  const projectId = cleanText(getCell(row, ['project id', 'id', 'projectid']));
  let status = cleanText(getCell(row, ['status'])) || 'Not Started';
  let location = cleanText(getCell(row, ['location', 'site', 'address']));
  let startDate = normalizeDate(getCell(row, ['start date', 'planned start', 'planned_start']));
  let finishDate = normalizeDate(getCell(row, ['finish date', 'end date', 'planned finish', 'planned_finish']));
  let budget = parseNumber(getCell(row, ['budget', 'planned value', 'planned cost']));
  const descriptionBudget = parseNumber(getCell(row, ['description', 'notes']));
  const createdAtBudget = parseNumber(getCell(row, ['created at', 'created_at', 'timestamp', 'date created']));

  const isKnownStatus = function(value) {
    const normalized = cleanText(value).toLowerCase();
    return ['not started', 'in progress', 'on hold', 'completed', 'archived'].indexOf(normalized) >= 0;
  };
  const isDateLike = function(value) {
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
    finishDate = normalizeDate(getCell(row, ['budget', 'planned value', 'planned cost']));
    budget = descriptionBudget || createdAtBudget || 0;
  }

  return {
    id: projectId,
    code: projectId,
    name: cleanText(getCell(row, ['project', 'project name', 'name'])),
    type: cleanText(getCell(row, ['type', 'project type'])) || 'General',
    status: status,
    location: location,
    startDate: startDate,
    finishDate: finishDate,
    budget: budget,
    raw: row,
  };
}

function normalizeActivityRecord(row) {
  const projectId = cleanText(getCell(row, ['project id', 'projectid']));

  return {
    id: cleanText(getCell(row, ['activity id', 'id', 'activity code', 'task id'])),
    projectId: projectId,
    projectCode: projectId,
    project: cleanText(getCell(row, ['project', 'project name'])),
    name: cleanText(getCell(row, ['activity', 'activity name', 'name'])),
    type: cleanText(getCell(row, ['type', 'activity type'])) || 'General',
    status: cleanText(getCell(row, ['status'])) || 'Not Started',
    plannedStart: normalizeDate(getCell(row, ['planned start', 'start date', 'planned_start'])),
    plannedFinish: normalizeDate(getCell(row, ['planned finish', 'finish date', 'planned_finish'])),
    duration: cleanText(getCell(row, ['duration', 'duration days', 'duration_days'])),
    percentComplete: parseNumber(getCell(row, ['% complete', 'percent complete', 'progress'])),
    createdAt: normalizeDate(getCell(row, ['created at', 'created_at', 'timestamp', 'date created'])),
    plannedValue: parseNumber(getCell(row, ['planned value', 'planned cost', 'budget'])),
    actualCost: parseNumber(getCell(row, ['actual cost', 'ac', 'actual'])),
    earnedValue: parseNumber(getCell(row, ['earned value', 'ev'])),
    costVariance: parseNumber(getCell(row, ['cost variance', 'cv'])),
    raw: row,
  };
}

function normalizeCostRecord(row) {
  const projectId = cleanText(getCell(row, ['project id', 'projectid']));

  return {
    id: cleanText(getCell(row, ['cost id', 'id'])),
    costId: cleanText(getCell(row, ['cost id', 'id'])),
    projectId: projectId,
    projectCode: projectId,
    project: cleanText(getCell(row, ['project', 'project name'])),
    activityId: cleanText(getCell(row, ['activity id', 'activityid', 'activity_id', 'id activity', 'source activity id'])),
    activity: cleanText(getCell(row, ['activity', 'activity name', 'name'])),
    duration: parseNumber(getCell(row, ['duration', 'duration days', 'duration_days'])),
    category: cleanText(getCell(row, ['cost category', 'category', 'type'])) || 'General',
    date: normalizeDate(getCell(row, ['date', 'cost date', 'transaction date'])),
    plannedCost: parseNumber(getCell(row, ['planned cost', 'planned value', 'budget'])),
    plannedCostPerDay: parseNumber(getCell(row, ['planned cost/day', 'planned cost per day', 'planned_cost_per_day'])),
    actualCost: parseNumber(getCell(row, ['actual cost', 'cost', 'amount'])),
    notes: cleanText(getCell(row, ['note', 'notes', 'remarks'])),
    raw: row,
  };
}

function normalizeDailyCostRecord(row) {
  return {
    projectId: cleanText(row['Project ID'] || row['projectId'] || row['project_id']),
    costId: cleanText(row['Cost ID'] || row['costId'] || row['cost_id']),
    activityId: cleanText(row['Activity ID'] || row['activityId'] || row['activity_id']),
    date: normalizeDate(row['Date'] || row['date']),
    actualCost: parseNumber(row['Actual Cost'] || row['actualCost'] || row['actual_cost']),
    createdAt: normalizeDate(row['Created At'] || row['createdAt'] || row['created_at']),
  };
}

function applyProjectFilter(data, filter) {
  const hasFilter = filter.id || filter.name;
  if (!hasFilter) return data;

  const projectMatches = function(record) {
    if (!record) return false;
    const id = cleanText(record.projectId || record.id);
    const name = cleanText(record.project || record.name);

    if (filter.id && filter.id === id) return true;
    if (filter.name && filter.name.toLowerCase() === name.toLowerCase()) return true;
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
  const typeCounts = projects.reduce(function(acc, row) {
    const key = row.type || 'General';
    acc[key] = (acc[key] || 0) + 1;
    return acc;
  }, {});

  const statusCounts = projects.reduce(function(acc, row) {
    const key = row.status || 'Not Started';
    acc[key] = (acc[key] || 0) + 1;
    return acc;
  }, {});

  return {
    count: projects.length,
    projects: projects,
    summary: {
      byType: typeCounts,
      byStatus: statusCounts,
      totalBudget: sumBy(projects, 'budget'),
    },
  };
}

function buildActivitiesPayload(activities) {
  const statusCounts = activities.reduce(function(acc, row) {
    const key = row.status || 'Not Started';
    acc[key] = (acc[key] || 0) + 1;
    return acc;
  }, {});

  return {
    count: activities.length,
    activities: activities,
    summary: {
      byStatus: statusCounts,
      totalPlannedValue: sumBy(activities, 'plannedValue'),
      totalActualCost: sumBy(activities, 'actualCost'),
      totalEarnedValue: sumBy(activities, 'earnedValue'),
      totalCostVariance: sumBy(activities, 'costVariance'),
    },
  };
}

function buildCostsPayload(costs) {
  const byCategory = costs.reduce(function(acc, row) {
    const key = row.category || 'General';
    acc[key] = (acc[key] || 0) + row.actualCost;
    return acc;
  }, {});

  return {
    count: costs.length,
    costs: costs,
    summary: {
      byCategory: byCategory,
      totalPlannedCost: sumBy(costs, 'plannedCost'),
      totalActualCost: sumBy(costs, 'actualCost'),
    },
  };
}

function buildDashboardPayload(data) {
  const projectCount = data.projects.length;
  const activityCount = data.activities.length;
  const costCount = data.costs.length;

  const totalPlanned = sumBy(data.activities, 'plannedValue');
  const totalActual = sumBy(data.activities, 'actualCost');
  const totalEarned = sumBy(data.activities, 'earnedValue');
  const totalCv = sumBy(data.activities, 'costVariance');

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
      projectStatus: totalCv >= 0 ? 'Under Budget' : 'Over Budget',
    },
    rows: data.activities,
  };
}

function buildReportsPayload(data) {
  const projectsByName = data.projects.reduce(function(acc, project) {
    const key = project.name || project.code || project.id;
    if (!key) return acc;

    acc[key] = {
      project: project,
      activities: [],
      costs: [],
    };
    return acc;
  }, {});

  data.activities.forEach(function(activity) {
    const key = activity.project || activity.projectId;
    if (!key) return;
    if (!projectsByName[key]) {
      projectsByName[key] = { project: null, activities: [], costs: [] };
    }
    projectsByName[key].activities.push(activity);
  });

  data.costs.forEach(function(cost) {
    const key = cost.project || cost.projectId;
    if (!key) return;
    if (!projectsByName[key]) {
      projectsByName[key] = { project: null, activities: [], costs: [] };
    }
    projectsByName[key].costs.push(cost);
  });

  const projectReports = Object.keys(projectsByName).map(function(projectKey) {
    const bundle = projectsByName[projectKey];
    const totalPlanned = sumBy(bundle.activities, 'plannedValue');
    const totalActual = sumBy(bundle.activities, 'actualCost');
    const totalEv = sumBy(bundle.activities, 'earnedValue');
    const totalCv = sumBy(bundle.activities, 'costVariance');

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
        status: totalCv >= 0 ? 'Under Budget' : 'Over Budget',
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
    dashboard: buildDashboardPayload(data),
    reports: buildReportsPayload(data),
  };
}

function sumBy(rows, field) {
  return rows.reduce(function(total, row) {
    return total + parseNumber(row[field]);
  }, 0);
}

function parseNumber(value) {
  if (value === null || value === undefined || value === '') return 0;
  if (typeof value === 'number') return Number.isFinite(value) ? value : 0;

  const stringValue = String(value).trim();
  const isAccountingNegative = /^\(.*\)$/.test(stringValue);
  const cleaned = stringValue.replace(/[^0-9.-]/g, '');
  const parsed = Number(cleaned);

  if (!Number.isFinite(parsed)) return 0;
  return isAccountingNegative ? -Math.abs(parsed) : parsed;
}

function roundTo(value, digits) {
  const factor = Math.pow(10, digits || 0);
  return Math.round(parseNumber(value) * factor) / factor;
}

function normalizeDate(value) {
  if (!value) return '';
  if (value instanceof Date && !Number.isNaN(value.getTime())) {
    return Utilities.formatDate(value, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }

  const parsed = new Date(value);
  if (Number.isNaN(parsed.getTime())) return cleanText(value);
  return Utilities.formatDate(parsed, Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

function getCell(row, aliases) {
  const keys = Object.keys(row || {});
  const normalizedKeys = keys.map(function(key) {
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

    const exact = normalizedKeys.find(function(entry) {
      return entry.normalized === normalizedAlias || entry.compact === compactAlias;
    });
    if (exact) return row[exact.key];

    const shouldUsePrefixMatch = normalizedAlias.indexOf(' ') >= 0;
    if (shouldUsePrefixMatch) {
      const prefixed = normalizedKeys.find(function(entry) {
        return entry.normalized.indexOf(normalizedAlias + ' ') === 0;
      });
      if (prefixed) return row[prefixed.key];
    }
  }

  return '';
}

function normalizeHeader(value) {
  return String(value || '')
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

function compactHeader(value) {
  return normalizeHeader(value).replace(/\s+/g, '');
}

function cleanText(value) {
  return String(value || '').trim();
}

function clampPercent(value) {
  const parsed = parseNumber(value);
  if (!Number.isFinite(parsed)) return 0;
  return Math.max(0, Math.min(100, parsed));
}

function calculateDurationDays(startDate, finishDate) {
  if (!startDate || !finishDate) return '';
  const start = new Date(startDate);
  const finish = new Date(finishDate);
  if (Number.isNaN(start.getTime()) || Number.isNaN(finish.getTime())) return '';
  const dayMs = 24 * 60 * 60 * 1000;
  const diff = Math.floor((finish.getTime() - start.getTime()) / dayMs) + 1;
  if (diff <= 0) return '';
  return String(diff);
}

function validateActivityForMutation(activity, action) {
  if (!activity.projectId && !activity.project) {
    throw new Error('Activity is missing project reference. Provide Project ID or Project Name.');
  }

  if (activity.plannedStart && activity.plannedFinish) {
    const start = new Date(activity.plannedStart);
    const finish = new Date(activity.plannedFinish);
    if (!Number.isNaN(start.getTime()) && !Number.isNaN(finish.getTime()) && start.getTime() > finish.getTime()) {
      throw new Error('Planned Start cannot be after Planned Finish.');
    }
  }

  if (action === 'create' && !activity.id) {
    throw new Error('Activity ID is required for create.');
  }
}

function assertActivitySheetColumns(columns) {
  if (!columns.id) throw new Error('Activities sheet is missing "Activity ID" column.');
  if (!columns.projectId) throw new Error('Activities sheet is missing "Project ID" column.');
  if (!columns.project) throw new Error('Activities sheet is missing "Project Name" column.');
  if (!columns.name) throw new Error('Activities sheet is missing "Activity" column.');
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

    const rowProjectId = columns.projectId ? cleanText(row[columns.projectId - 1]).toLowerCase() : '';
    const rowProjectName = columns.project ? cleanText(row[columns.project - 1]).toLowerCase() : '';
    const projectIdMatch = !incomingProjectId || incomingProjectId === rowProjectId;
    const projectNameMatch = !incomingProjectName || incomingProjectName === rowProjectName;

    if (projectIdMatch && projectNameMatch) {
      throw new Error('Activity already exists for this project. Use update instead of create.');
    }
  }
}

function jsonResponse(payload) {
  return ContentService.createTextOutput(JSON.stringify(payload)).setMimeType(ContentService.MimeType.JSON);
}
