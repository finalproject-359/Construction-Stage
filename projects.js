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
