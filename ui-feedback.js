(() => {
  const CONTAINER_ID = "appToastContainer";
  const DIALOG_ID = "appFeedbackDialog";
  const ICONS = {
    success: "✓",
    error: "!",
    warning: "!",
    info: "i",
  };

  const escapeHtml = (value) => String(value ?? "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#39;");

  const getToneFromMessage = (message) => {
    const text = String(message || "");
    if (/success|saved|updated|deleted|restored|done|synced/i.test(text)) return "success";
    if (/unable|error|failed|missing|invalid|cannot|strict mode/i.test(text)) return "error";
    if (/archive|delete|warning|overwrite|no changes|already/i.test(text)) return "warning";
    return "info";
  };

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

  window.notify = (message, type = "info", timeout = 3200) => {
    if (!message) return null;
    const tone = ICONS[type] ? type : "info";
    const container = ensureContainer();
    const toast = document.createElement("div");
    toast.className = `app-toast app-toast-${tone}`;
    toast.setAttribute("role", tone === "error" ? "alert" : "status");
    toast.innerHTML = `
      <span class="app-toast-icon" aria-hidden="true">${ICONS[tone]}</span>
      <span class="app-toast-message">${escapeHtml(message)}</span>
      <button type="button" class="app-toast-dismiss" aria-label="Dismiss notification">×</button>
      <span class="app-toast-progress" aria-hidden="true"></span>`;

    let closed = false;
    const close = () => {
      if (closed) return;
      closed = true;
      toast.classList.add("is-hiding");
      toast.addEventListener("transitionend", () => toast.remove(), { once: true });
      window.setTimeout(() => toast.remove(), 260);
    };

    toast.querySelector("button")?.addEventListener("click", close);
    container.appendChild(toast);
    window.setTimeout(close, Math.max(1800, timeout));
    return { close, element: toast };
  };

  window.showAppDialog = ({
    title = "Notification",
    message = "",
    type = "info",
    confirmText = "OK",
    cancelText = "Cancel",
    showCancel = false,
  } = {}) => new Promise((resolve) => {
    const existing = document.getElementById(DIALOG_ID);
    existing?.remove();

    const tone = ICONS[type] ? type : getToneFromMessage(`${title} ${message}`);
    const dialog = document.createElement("div");
    dialog.id = DIALOG_ID;
    dialog.className = `app-feedback-dialog app-feedback-${tone}`;
    dialog.innerHTML = `
      <div class="app-feedback-backdrop" data-feedback-cancel></div>
      <section class="app-feedback-card" role="dialog" aria-modal="true" aria-labelledby="appFeedbackTitle" aria-describedby="appFeedbackMessage">
        <div class="app-feedback-icon" aria-hidden="true">${ICONS[tone]}</div>
        <div class="app-feedback-content">
          <p class="app-feedback-eyebrow">${showCancel ? "Please confirm" : "CosTrack notification"}</p>
          <h2 id="appFeedbackTitle">${escapeHtml(title)}</h2>
          <p id="appFeedbackMessage">${escapeHtml(message)}</p>
        </div>
        <div class="app-feedback-actions">
          ${showCancel ? `<button type="button" class="ghost-btn app-feedback-cancel" data-feedback-cancel>${escapeHtml(cancelText)}</button>` : ""}
          <button type="button" class="primary-btn app-feedback-confirm" data-feedback-confirm>${escapeHtml(confirmText)}</button>
        </div>
      </section>`;

    let settled = false;
    const settle = (value) => {
      if (settled) return;
      settled = true;
      dialog.classList.add("is-hiding");
      document.removeEventListener("keydown", onKeydown);
      window.setTimeout(() => dialog.remove(), 180);
      resolve(value);
    };
    const onKeydown = (event) => {
      if (event.key === "Escape") settle(false);
      if (event.key === "Enter") settle(true);
    };

    dialog.querySelectorAll("[data-feedback-cancel]").forEach((button) => button.addEventListener("click", () => settle(false)));
    dialog.querySelector("[data-feedback-confirm]")?.addEventListener("click", () => settle(true));
    document.addEventListener("keydown", onKeydown);
    document.body.appendChild(dialog);
    window.setTimeout(() => dialog.querySelector("[data-feedback-confirm]")?.focus(), 30);
  });

  window.confirmAction = (message, options = {}) => {
    const normalizedOptions = typeof options === "string" ? { title: options } : options;
    return window.showAppDialog({
      title: normalizedOptions.title || "Confirm action",
      message,
      type: normalizedOptions.type || "warning",
      confirmText: normalizedOptions.confirmText || "Confirm",
      cancelText: normalizedOptions.cancelText || "Cancel",
      showCancel: true,
    });
  };

  const nativeAlert = window.alert.bind(window);
  window.alert = (message) => {
    if (typeof window.notify === "function") {
      const text = String(message || "");
      window.notify(text, getToneFromMessage(text), 4200);
      return;
    }
    nativeAlert(message);
  };
})();
