(() => {
  const CONTAINER_ID = "appToastContainer";

  const ensureContainer = () => {
    let container = document.getElementById(CONTAINER_ID);
    if (container) return container;
    container = document.createElement("div");
    container.id = CONTAINER_ID;
    container.className = "app-toast-container";
    container.setAttribute("aria-live", "polite");
    container.setAttribute("aria-atomic", "false");
    document.body.appendChild(container);
    return container;
  };

  window.notify = (message, type = "info", timeout = 2800) => {
    if (!message) return;
    const container = ensureContainer();
    const toast = document.createElement("div");
    toast.className = `app-toast app-toast-${type}`;
    toast.setAttribute("role", "status");
    toast.innerHTML = `<span>${String(message)}</span><button aria-label="Dismiss notification">×</button>`;

    const close = () => {
      toast.classList.add("is-hiding");
      toast.addEventListener("transitionend", () => toast.remove(), { once: true });
    };

    toast.querySelector("button")?.addEventListener("click", close);
    container.appendChild(toast);

    window.setTimeout(close, Math.max(1400, timeout));
  };

  const nativeAlert = window.alert.bind(window);
  window.alert = (message) => {
    if (typeof window.notify === "function") {
      const text = String(message || "");
      const type = /success|saved|updated|deleted|done/i.test(text) ? "success" : /unable|error|failed|missing|invalid/i.test(text) ? "error" : "info";
      window.notify(text, type, 3600);
      return;
    }
    nativeAlert(message);
  };
})();
