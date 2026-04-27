/**
 * Google Apps Script Web App endpoint for Construction Stage data.
 *
 * Supported query params:
 * - resource: dashboard | projects | activities | costs | reports | all
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
      'Activity ID',
      'Project ID',
      'Project',
      'Activity',
      'Type',
      'Status',
      'Planned Start',
      'Planned Finish',
      '% Complete',
      'Planned Value',
      'Actual Cost',
      'Earned Value',
      'Cost Variance',
      'Notes',
      'Created At',
    ],
    costs: [
      'Cost ID',
      'Project ID',
      'Project',
      'Cost Category',
      'Date',
      'Planned Cost',
      'Actual Cost',
      'Notes',
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

    ensureWorkbookStructure();

    if (action) {
      if (resource !== 'projects') {
        throw new Error('Only "projects" is supported for mutations.');
      }
      return handleProjectMutation(action, source);
    }

    const projectFilter = {
      id: cleanText(source.projectId || source.project_id || ''),
      name: cleanText(source.project || source.projectName || source.project_name || ''),
    };

    const allData = loadAllData();
    const filtered = applyProjectFilter(allData, projectFilter);

    const payloadByResource = {
      projects: buildProjectsPayload(filtered.projects),
      activities: buildActivitiesPayload(filtered.activities),
      costs: buildCostsPayload(filtered.costs),
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

function handleProjectMutation(action, payload) {
  if (action === 'create') {
    const project = normalizeIncomingProject(payload.project || payload);
    if (!project.name || !project.id) {
      throw new Error('Project Name and Project ID are required.');
    }

    const sheet = getOrCreateSheet(CONFIG.sheetNames.projects);
    ensureProjectHeaders(sheet);
    const storedProjectId = cleanText(project.id);

    sheet.appendRow([
      storedProjectId,
      project.name,
      project.type,
      project.status,
      project.location,
      project.startDate,
      project.finishDate,
      project.budget,
      new Date(),
    ]);

    return jsonResponse({
      ok: true,
      message: 'Project saved successfully.',
      project: {
        ...project,
        id: storedProjectId,
        code: cleanText(project.code || storedProjectId),
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

function findProjectSheetRow(projectId) {
  const sheet = getOrCreateSheet(CONFIG.sheetNames.projects);
  ensureProjectHeaders(sheet);

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
  const supported = ['dashboard', 'projects', 'activities', 'costs', 'reports', 'all'];
  const normalized = cleanText(value).toLowerCase();
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
  const normalizedExisting = normalizeHeaders(firstRow);
  const legacyDeleteIndexes = [];
  const legacyProjectCodeIndex = normalizedExisting.indexOf('project code');
  const expectsProjectCode = normalizedExpected.indexOf('project code') >= 0;
  if (legacyProjectCodeIndex >= 0 && !expectsProjectCode) {
    legacyDeleteIndexes.push(legacyProjectCodeIndex + 1);
  }

  const legacyDescriptionIndex = normalizedExisting.indexOf('description');
  const expectsDescription = normalizedExpected.indexOf('description') >= 0;
  if (legacyDescriptionIndex >= 0 && !expectsDescription) {
    legacyDeleteIndexes.push(legacyDescriptionIndex + 1);
  }

  legacyDeleteIndexes.sort(function(a, b) {
    return b - a;
  });
  for (var legacyIdx = 0; legacyIdx < legacyDeleteIndexes.length; legacyIdx += 1) {
    sheet.deleteColumn(legacyDeleteIndexes[legacyIdx]);
  }

  if (legacyDeleteIndexes.length) {
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

  const id = cleanText(source.id || source.projectId || source.project_id);
  const code = cleanText(source.code || source.projectCode || source.project_code);
  const name = cleanText(source.name || source.project || source.projectName || source.project_name);
  const type = cleanText(source.type || source.projectType || source.project_type) || 'General';
  const status = cleanText(source.status) || 'Not Started';
  const location = cleanText(source.location || source.site || source.address);
  const startDate = normalizeDate(source.startDate || source.start_date || source.plannedStart);
  const finishDate = normalizeDate(source.finishDate || source.finish_date || source.targetFinish || source.endDate);
  const budget = parseNumber(source.budget);
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

function loadAllData() {
  ensureWorkbookStructure();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const projects = readSheetRows(ss, CONFIG.sheetNames.projects);
  const activities = readSheetRows(ss, CONFIG.sheetNames.activities);
  const costs = readSheetRows(ss, CONFIG.sheetNames.costs);

  return {
    projects: projects.rows.map(normalizeProjectRecord),
    activities: activities.rows.map(normalizeActivityRecord),
    costs: costs.rows.map(normalizeCostRecord),
    sheets: {
      projects: projects.meta,
      activities: activities.meta,
      costs: costs.meta,
    },
  };
}

function readSheetRows(ss, sheetName) {
  const sheet = getOrCreateSheet(sheetName);
  const expectedHeadersBySheet = {};
  expectedHeadersBySheet[CONFIG.sheetNames.projects] = CONFIG.headers.projects;
  expectedHeadersBySheet[CONFIG.sheetNames.activities] = CONFIG.headers.activities;
  expectedHeadersBySheet[CONFIG.sheetNames.costs] = CONFIG.headers.costs;
  ensureSheetHeaders(sheet, expectedHeadersBySheet[sheetName] || []);

  const displayValues = sheet.getDataRange().getDisplayValues();
  const rawValues = sheet.getDataRange().getValues();

  if (!displayValues.length) {
    return {
      rows: [],
      meta: {
        sheetName: sheetName,
        headerRowIndex: 0,
        rowCount: 0,
      },
    };
  }

  const headerRowIndex = findHeaderRowIndex(displayValues);
  const headers = displayValues[headerRowIndex].map(function(header) {
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

  return {
    id: projectId,
    code: projectId,
    name: cleanText(getCell(row, ['project', 'project name', 'name'])),
    type: cleanText(getCell(row, ['type', 'project type'])) || 'General',
    status: cleanText(getCell(row, ['status'])) || 'Not Started',
    location: cleanText(getCell(row, ['location', 'site', 'address'])),
    startDate: normalizeDate(getCell(row, ['start date', 'planned start', 'planned_start'])),
    finishDate: normalizeDate(getCell(row, ['finish date', 'end date', 'planned finish', 'planned_finish'])),
    budget: parseNumber(getCell(row, ['budget', 'planned value', 'planned cost'])),
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
    percentComplete: parseNumber(getCell(row, ['% complete', 'percent complete', 'progress'])),
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
    projectId: projectId,
    projectCode: projectId,
    project: cleanText(getCell(row, ['project', 'project name'])),
    category: cleanText(getCell(row, ['cost category', 'category', 'type'])) || 'General',
    date: normalizeDate(getCell(row, ['date', 'cost date', 'transaction date'])),
    plannedCost: parseNumber(getCell(row, ['planned cost', 'planned value', 'budget'])),
    actualCost: parseNumber(getCell(row, ['actual cost', 'cost', 'amount'])),
    notes: cleanText(getCell(row, ['note', 'notes', 'remarks'])),
    raw: row,
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

function jsonResponse(payload) {
  return ContentService.createTextOutput(JSON.stringify(payload)).setMimeType(ContentService.MimeType.JSON);
}
