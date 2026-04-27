const projectModal = document.getElementById("projectModal");
const projectForm = document.getElementById("projectForm");
const projectModalClose = document.getElementById("projectModalClose");
const projectModalBackdrop = document.getElementById("projectModalBackdrop");
const projectFormCancel = document.getElementById("projectFormCancel");
const openAddProjectModalBtn = document.getElementById("openAddProjectModalBtn");
const openAddProjectModalEmptyBtn = document.getElementById("openAddProjectModalEmptyBtn");

if (!projectModal || !projectForm) {
  throw new Error("Projects page is missing modal form elements.");
}

const budgetInput = projectForm.elements.namedItem("budget");
const pesoBudgetFormatter = new Intl.NumberFormat("en-PH", {
  style: "currency",
  currency: "PHP",
  minimumFractionDigits: 2,
  maximumFractionDigits: 2,
});

const toNumericBudgetValue = (value) => {
  const sanitized = value.replace(/[^\d.]/g, "");
  const [integerPart, ...fractionParts] = sanitized.split(".");
  const normalized = fractionParts.length
    ? `${integerPart}.${fractionParts.join("")}`
    : integerPart;

  return normalized.replace(/^0+(\d)/, "$1");
};

const formatBudgetAsPeso = (value) => {
  const normalized = toNumericBudgetValue(value);

  if (!normalized) {
    return "";
  }

  const numericValue = Number.parseFloat(normalized);
  if (!Number.isFinite(numericValue)) {
    return "";
  }

  return pesoBudgetFormatter.format(numericValue);
};

if (budgetInput instanceof HTMLInputElement) {
  budgetInput.addEventListener("focus", () => {
    budgetInput.value = toNumericBudgetValue(budgetInput.value);
  });

  budgetInput.addEventListener("input", () => {
    budgetInput.value = toNumericBudgetValue(budgetInput.value);

    if (budgetInput.value) {
      budgetInput.setCustomValidity("");
      return;
    }

    budgetInput.setCustomValidity("Please enter a budget amount.");
  });

  budgetInput.addEventListener("blur", () => {
    if (!budgetInput.value) {
      return;
    }

    const formattedValue = formatBudgetAsPeso(budgetInput.value);

    if (!formattedValue) {
      budgetInput.setCustomValidity("Please enter a valid budget amount.");
      return;
    }

    budgetInput.setCustomValidity("");
    budgetInput.value = formattedValue;
  });
}

const openProjectModal = () => {
  projectModal.classList.remove("hidden");
  document.body.style.overflow = "hidden";
};

const closeProjectModal = () => {
  projectModal.classList.add("hidden");
  document.body.style.overflow = "";
};

openAddProjectModalBtn?.addEventListener("click", openProjectModal);
openAddProjectModalEmptyBtn?.addEventListener("click", openProjectModal);
projectModalClose?.addEventListener("click", closeProjectModal);
projectFormCancel?.addEventListener("click", closeProjectModal);
projectModalBackdrop?.addEventListener("click", closeProjectModal);

projectModal.addEventListener("click", (event) => {
  if (event.target === projectModal) {
    closeProjectModal();
  }
});

document.addEventListener("keydown", (event) => {
  if (event.key === "Escape" && !projectModal.classList.contains("hidden")) {
    closeProjectModal();
  }
});

projectForm.addEventListener("submit", (event) => {
  event.preventDefault();
  closeProjectModal();
  projectForm.reset();
});
