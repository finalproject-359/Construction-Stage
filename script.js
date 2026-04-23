const fileInput = document.getElementById("fileInput");
const totalPlannedEl = document.getElementById("totalPlanned");
const totalActualEl = document.getElementById("totalActual");
const totalCvEl = document.getElementById("totalCv");
const varianceCardEl = document.getElementById("varianceCard");
const projectStatusEl = document.getElementById("projectStatus");
const statusMetaEl = document.getElementById("statusMeta");
const statusCardEl = document.getElementById("statusCard");
const messageEl = document.getElementById("message");
const tableBodyEl = document.getElementById("activityTableBody");

let cvChart;
let planActualChart;

const asCurrency = (value) =>
  new Intl.NumberFormat("en-US", {
    style: "currency",
    currency: "USD",
    minimumFractionDigits: 2,
    maximumFractionDigits: 2,
  }).format(Number.isFinite(value) ? value : 0);

const parseNumber = (value) => {
  if (value === null || value === undefined || value === "") return 0;
  if (typeof value === "number") return Number.isFinite(value) ? value : 0;
  const cleaned = String(value).replace(/[^0-9.-]/g, "");
  const parsed = Number(cleaned);
  return Number.isFinite(parsed) ? parsed : 0;
};

const normalizeText = (value, fallback = "Not provided") => {
  if (value === null || value === undefined || value === "") return fallback;
  const str = String(value).trim();
  return str || fallback;
};

const getValueByAliases = (row, aliases) => {
  for (const key of Object.keys(row)) {
    const normalized = key.toLowerCase().replace(/\s+/g, " ").trim();
    if (aliases.includes(normalized)) {
      return row[key];
    }
  }
  return "";
};

const extractRows = (rawRows) =>
  rawRows
    .map((row) => ({
      activityId: normalizeText(getValueByAliases(row, ["activity id", "activity", "activityid", "id"]), "Unspecified"),
      plannedCost: parseNumber(getValueByAliases(row, ["planned cost", "planned", "plannedcost", "planned cost (pv)", "pv"])),
      actualCost: parseNumber(getValueByAliases(row, ["actual cost", "actual", "actualcost", "actual cost (ac)", "ac"])),
      earnedValue: parseNumber(getValueByAliases(row, ["earned value", "earned value (ev)", "ev"])),
      costVariance: parseNumber(getValueByAliases(row, ["cost variance", "cost variance (cv)", "cv"])),
      budgetStatus: normalizeText(getValueByAliases(row, ["budget status", "status"])),
    }))
    .filter((row) => row.activityId !== "Unspecified" || row.plannedCost || row.actualCost || row.earnedValue || row.costVariance);

const renderKpis = (rows) => {
  const totals = rows.reduce(
    (acc, row) => {
      acc.totalPlanned += row.plannedCost;
      acc.totalActual += row.actualCost;
      acc.totalCv += row.costVariance;
      return acc;
    },
    { totalPlanned: 0, totalActual: 0, totalCv: 0 }
  );

  totalPlannedEl.textContent = asCurrency(totals.totalPlanned);
  totalActualEl.textContent = asCurrency(totals.totalActual);
  totalCvEl.textContent = asCurrency(totals.totalCv);

  varianceCardEl.classList.remove("status-under", "status-over");
  statusCardEl.classList.remove("status-under", "status-over");

  if (totals.totalCv > 0) {
    projectStatusEl.textContent = "Under Budget";
    statusMetaEl.textContent = "Positive total CV";
    statusCardEl.classList.add("status-under");
    varianceCardEl.classList.add("status-under");
  } else if (totals.totalCv < 0) {
    projectStatusEl.textContent = "Over Budget";
    statusMetaEl.textContent = "Negative total CV";
    statusCardEl.classList.add("status-over");
    varianceCardEl.classList.add("status-over");
  } else {
    projectStatusEl.textContent = "On Budget";
    statusMetaEl.textContent = "Total CV equals 0";
  }
};

const statusBadgeClass = (row) => {
  const value = row.budgetStatus.toLowerCase();
  if (value.includes("under") || row.costVariance > 0) return "good";
  if (value.includes("over") || row.costVariance < 0) return "bad";
  return row.costVariance >= 0 ? "good" : "bad";
};

const renderTable = (rows) => {
  if (!rows.length) {
    tableBodyEl.innerHTML = '<tr><td class="placeholder" colspan="6">No valid activity rows found in the Dashboard sheet.</td></tr>';
    return;
  }

  tableBodyEl.innerHTML = rows
    .map((row) => `
      <tr>
        <td>${row.activityId}</td>
        <td>${asCurrency(row.plannedCost)}</td>
        <td>${asCurrency(row.actualCost)}</td>
        <td>${asCurrency(row.earnedValue)}</td>
        <td class="${row.costVariance >= 0 ? "cv-good" : "cv-bad"}">${asCurrency(row.costVariance)}</td>
        <td><span class="badge ${statusBadgeClass(row)}">${row.budgetStatus}</span></td>
      </tr>
    `)
    .join("");
};

const clearChart = (chartRef) => {
  if (chartRef) chartRef.destroy();
};

const renderCharts = (rows) => {
  const labels = rows.map((r) => r.activityId);
  const cvValues = rows.map((r) => r.costVariance);

  clearChart(cvChart);
  clearChart(planActualChart);

  cvChart = new Chart(document.getElementById("cvChart"), {
    type: "bar",
    data: {
      labels,
      datasets: [
        {
          data: cvValues,
          label: "Cost Variance",
          backgroundColor: cvValues.map((v) => (v >= 0 ? "rgba(22, 163, 74, 0.8)" : "rgba(220, 38, 38, 0.8)")),
          borderColor: cvValues.map((v) => (v >= 0 ? "rgba(22, 163, 74, 1)" : "rgba(220, 38, 38, 1)")),
          borderWidth: 1,
          borderRadius: 4,
        },
      ],
    },
    options: {
      maintainAspectRatio: false,
      plugins: { legend: { display: false } },
      scales: {
        y: {
          ticks: { callback: (value) => asCurrency(Number(value)) },
        },
      },
    },
  });

  planActualChart = new Chart(document.getElementById("planActualChart"), {
    type: "line",
    data: {
      labels,
      datasets: [
        {
          label: "Planned Cost",
          data: rows.map((r) => r.plannedCost),
          borderColor: "#3d5afe",
          backgroundColor: "rgba(61, 90, 254, 0.2)",
          tension: 0.3,
          fill: false,
          pointRadius: 3,
        },
        {
          label: "Actual Cost",
          data: rows.map((r) => r.actualCost),
          borderColor: "#16a34a",
          backgroundColor: "rgba(22, 163, 74, 0.2)",
          tension: 0.3,
          fill: false,
          pointRadius: 3,
        },
      ],
    },
    options: {
      maintainAspectRatio: false,
      plugins: { legend: { position: "top" } },
      scales: {
        y: {
          ticks: { callback: (value) => asCurrency(Number(value)) },
        },
      },
    },
  });
};

const setMessage = (text, isError = false) => {
  messageEl.textContent = text;
  messageEl.style.color = isError ? "#dc2626" : "#64748b";
};

const processWorkbook = (arrayBuffer) => {
  const workbook = XLSX.read(arrayBuffer, { type: "array" });
  const dashboardSheet = workbook.Sheets["Dashboard"];

  if (!dashboardSheet) {
    setMessage('Could not find a sheet named "Dashboard" in this file.', true);
    return;
  }

  const rawRows = XLSX.utils.sheet_to_json(dashboardSheet, { defval: "" });
  const rows = extractRows(rawRows);

  renderKpis(rows);
  renderTable(rows);
  renderCharts(rows);

  setMessage(`Loaded ${rows.length} activity rows from Dashboard.`);
};

fileInput.addEventListener("change", async (event) => {
  const file = event.target.files?.[0];
  if (!file) return;

  try {
    const fileBuffer = await file.arrayBuffer();
    processWorkbook(fileBuffer);
  } catch (error) {
    setMessage(`Failed to read the file: ${error.message}`, true);
  }
});
