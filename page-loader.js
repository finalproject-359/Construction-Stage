(() => {
  const MIN_VISIBLE_MS = 550;
  const pageStart = performance.now();

  const overlay = document.createElement('div');
  overlay.className = 'page-loader-overlay';
  overlay.setAttribute('role', 'status');
  overlay.setAttribute('aria-live', 'polite');
  overlay.setAttribute('aria-busy', 'true');
  overlay.innerHTML = `
    <div class="page-loader-card">
      <div class="page-loader-spinner" aria-hidden="true"></div>
      <h2>Loading page...</h2>
      <p>Please wait while we fetch the latest data.</p>
    </div>
  `;

  const removeLoader = () => {
    const elapsed = performance.now() - pageStart;
    const wait = Math.max(0, MIN_VISIBLE_MS - elapsed);

    window.setTimeout(() => {
      overlay.classList.add('is-hiding');
      overlay.addEventListener('transitionend', () => overlay.remove(), { once: true });
    }, wait);
  };

  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', () => document.body.appendChild(overlay), { once: true });
  } else {
    document.body.appendChild(overlay);
  }

  if (document.readyState === 'complete') {
    removeLoader();
  } else {
    window.addEventListener('load', removeLoader, { once: true });
  }
})();
