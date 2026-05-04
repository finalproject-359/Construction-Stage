(() => {
  const MIN_VISIBLE_MS = 450;
  const pageStart = performance.now();
  const isCostManagementPage = document.body.classList.contains("page-cost-management");

  const ensureLoaderStyles = () => {
    if (document.getElementById("page-loader-inline-styles")) return;

    const style = document.createElement("style");
    style.id = "page-loader-inline-styles";
    style.textContent = `
      .page-loader-overlay {
        position: absolute;
        inset: 0;
        z-index: 80;
        background: rgba(15, 23, 42, 0.18);
        backdrop-filter: blur(3px);
        display: grid;
        place-items: center;
        transition: opacity .24s ease, visibility .24s ease;
      }

      .page-loader-overlay.is-hiding {
        opacity: 0;
        visibility: hidden;
      }

      .page-loader-card {
        width: min(300px, calc(100% - 2rem));
        background: #fff;
        border: 1px solid #e5eaf3;
        border-radius: 20px;
        padding: 26px 22px;
        text-align: center;
        box-shadow: 0 18px 44px rgba(15, 23, 42, 0.14);
      }

      .page-loader-spinner {
        width: 62px;
        height: 62px;
        border-radius: 50%;
        border: 6px solid #d9dfed;
        border-top-color: #3b6ff8;
        margin: 0 auto 16px;
        animation: page-loader-spin 0.95s linear infinite;
      }

      .page-loader-card h2 { margin: 0; font-size: 1.5rem; font-weight: 700; color: #0f2547; }
      .page-loader-card p { margin: 8px 0 0; font-size: 1rem; color: #5e6f8d; }
      @keyframes page-loader-spin { to { transform: rotate(360deg); } }
    `;

    document.head.appendChild(style);
  };

  const overlay = document.createElement("div");
  overlay.className = "page-loader-overlay";
  overlay.setAttribute("role", "status");
  overlay.setAttribute("aria-live", "polite");
  overlay.setAttribute("aria-busy", "true");
  overlay.innerHTML = `
    <div class="page-loader-card" aria-label="Loading content">
      <div class="page-loader-spinner" aria-hidden="true"></div>
      <h2>Fetching data...</h2>
      <p>Please wait while we load the data.</p>
    </div>
  `;

  const removeLoader = () => {
    if (!overlay.isConnected) return;
    const elapsed = performance.now() - pageStart;
    const wait = Math.max(0, MIN_VISIBLE_MS - elapsed);
    window.setTimeout(() => {
      overlay.classList.add("is-hiding");
      overlay.addEventListener("transitionend", () => overlay.remove(), { once: true });
    }, wait);
  };

  const mountLoader = () => {
    ensureLoaderStyles();
    const host = document.querySelector(".main-content") || document.body;
    if (getComputedStyle(host).position === "static") host.style.position = "relative";
    host.appendChild(overlay);
  };

  const waitForReadySignal = () => {
    if (!isCostManagementPage) {
      if (document.readyState === "complete") removeLoader();
      else window.addEventListener("load", removeLoader, { once: true });
      return;
    }

    const onLoaded = () => removeLoader();
    window.addEventListener("cost-management:data-loaded", onLoaded, { once: true });
  };

  if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", () => {
      mountLoader();
      waitForReadySignal();
    }, { once: true });
  } else {
    mountLoader();
    waitForReadySignal();
  }
})();
