const fileInput = document.getElementById("fileInput");
const totalPlannedEl = document.getElementById("totalPlanned");
const totalActualEl = document.getElementById("totalActual");
const totalCvEl = document.getElementById("totalCv");
const projectStatusEl = document.getElementById("projectStatus");
const statusCardEl = document.getElementById("statusCard");
const messageEl = document.getElementById("message");
const tableBodyEl = document.getElementById("activityTableBody");

let cvChart;
let planActualChart;

const formatCurrency = (value) =>
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

const normalize = (value, fallback = "N/A") => {
  if (value === null || value === undefined || value === "") return fallback;
  return String(value).trim() || fallback;
};

const getCell = (row, aliases) => {
  for (const key of Object.keys(row)) {
    const normalizedKey = key.toLowerCase().replace(/\s+/g, " ").trim();
    if (aliases.includes(normalizedKey)) {
      return row[key];
    }
  }
  return null;
};

const extractDashboardRows = (rawRows) =>
  rawRows
    .map((row) => ({
      activityId: normalize(
        getCell(row, ["activity id", "activity", "id", "activityid"]),
        "Unspecified"
      ),
      plannedCost: parseNumber(
        getCell(row, ["planned cost", "planned", "plannedcost", "budgeted cost"])
      ),
      actualCost: parseNumber(getCell(row, ["actual cost", "actual", "actualcost"])),
      ev: parseNumber(getCell(row, ["earned value (ev)", "earned value", "ev"])),
      cv: parseNumber(getCell(row, ["cost variance (cv)", "cost variance", "cv"])),
      budgetStatus: normalize(
        getCell(row, ["budget status", "status"]),
        "Not provided"
      ),
    }))
    .filter((row) => row.activityId !== "Unspecified" || row.plannedCost || row.actualCost || row.cv || row.ev);

const renderKpis = (rows) => {
  const totals = rows.reduce(
    (acc, row) => {
      acc.planned += row.plannedCost;
      acc.actual += row.actualCost;
      acc.cv += row.cv;
      return acc;
    },
    { planned: 0, actual: 0, cv: 0 }
  );

  totalPlannedEl.textContent = formatCurrency(totals.planned);
  totalActualEl.textContent = formatCurrency(totals.actual);
  totalCvEl.textContent = formatCurrency(totals.cv);

  statusCardEl.classList.remove("status-under", "status-over");

  if (totals.cv > 0) {
    projectStatusEl.textContent = "Under Budget";
    statusCardEl.classList.add("status-under");
  } else if (totals.cv < 0) {
    projectStatusEl.textContent = "Over Budget";
    statusCardEl.classList.add("status-over");
  } else {
    projectStatusEl.textContent = "On Budget";
  }
};

const renderTable = (rows) => {
  if (!rows.length) {
    tableBodyEl.innerHTML = '<tr><td colspan="6" class="placeholder">No valid rows found in Dashboard sheet.</td></tr>';
    return;
  }

  tableBodyEl.innerHTML = rows
    .map(
      (row) => `
      <tr>
        <td>${row.activityId}</td>
        <td>${formatCurrency(row.plannedCost)}</td>
        <td>${formatCurrency(row.actualCost)}</td>
        <td>${formatCurrency(row.ev)}</td>
        <td>${formatCurrency(row.cv)}</td>
        <td>${row.budgetStatus}</td>
      </tr>
    `
    )
    .join("");
};

const destroyChart = (chart) => {
  if (chart) chart.destroy();
};

const renderCharts = (rows) => {
  const labels = rows.map((row) => row.activityId);
  const cvValues = rows.map((row) => row.cv);

  destroyChart(cvChart);
  destroyChart(planActualChart);

  cvChart = new Chart(document.getElementById("cvChart"), {
    type: "bar",
    data: {
      labels,
      datasets: [
        {
          label: "Cost Variance (CV)",
          data: cvValues,
          backgroundColor: cvValues.map((value) => (value >= 0 ? "rgba(5, 150, 105, 0.75)" : "rgba(220, 38, 38, 0.75)")),
          borderColor: cvValues.map((value) => (value >= 0 ? "rgba(5, 150, 105, 1)" : "rgba(220, 38, 38, 1)")),
          borderWidth: 1,
          borderRadius: 4,
        },
      ],
    },
    options: {
      maintainAspectRatio: false,
      plugins: {
        legend: { display: false },
      },
      scales: {
        y: {
          ticks: {
            callback: (value) => formatCurrency(Number(value)),
          },
        },
      },
    },
  });

  planActualChart = new Chart(document.getElementById("planActualChart"), {
    type: "bar",
    data: {
      labels,
      datasets: [
        {
          label: "Planned Cost",
          data: rows.map((row) => row.plannedCost),
          backgroundColor: "rgba(37, 99, 235, 0.75)",
          borderColor: "rgba(37, 99, 235, 1)",
          borderWidth: 1,
          borderRadius: 4,
        },
        {
          label: "Actual Cost",
          data: rows.map((row) => row.actualCost),
          backgroundColor: "rgba(245, 158, 11, 0.75)",
          borderColor: "rgba(245, 158, 11, 1)",
          borderWidth: 1,
          borderRadius: 4,
        },
      ],
    },
    options: {
      maintainAspectRatio: false,
      plugins: {
        legend: {
          position: "top",
        },
      },
      scales: {
        y: {
          ticks: {
            callback: (value) => formatCurrency(Number(value)),
          },
        },
      },
    },
  });
};

const showMessage = (text, isError = false) => {
  messageEl.textContent = text;
  messageEl.style.color = isError ? "#dc2626" : "#6b7280";
};

const processWorkbook = (arrayBuffer) => {
  const workbook = XLSX.read(arrayBuffer, { type: "array" });
  const sheet = workbook.Sheets["Dashboard"];

  if (!sheet) {
    showMessage('Sheet "Dashboard" not found. Please upload the correct file.', true);
    return;
  }

  const rawRows = XLSX.utils.sheet_to_json(sheet, { defval: "" });
  const rows = extractDashboardRows(rawRows);

  renderKpis(rows);
  renderTable(rows);
  renderCharts(rows);

  showMessage(`Loaded ${rows.length} activity row(s) from Dashboard sheet.`);
};

fileInput.addEventListener("change", async (event) => {
  const file = event.target.files?.[0];
  if (!file) return;

  try {
    const buffer = await file.arrayBuffer();
    processWorkbook(buffer);
  } catch (error) {
    showMessage(`Error reading file: ${error.message}`, true);
  }
});
