(() => {
  const fromEl = document.getElementById('reportDateFrom');
  const toEl = document.getElementById('reportDateTo');
  const periodLabelEl = document.getElementById('periodLabel');

  const format = (value) => {
    if (!value) return '—';
    const date = new Date(`${value}T00:00:00`);
    return Number.isNaN(date.getTime()) ? '—' : date.toLocaleDateString();
  };

  const syncPeriodLabel = () => {
    if (!periodLabelEl) return;
    periodLabelEl.textContent = `${format(fromEl?.value)} to ${format(toEl?.value)}`;
  };

  fromEl?.addEventListener('change', syncPeriodLabel);
  toEl?.addEventListener('change', syncPeriodLabel);
  syncPeriodLabel();
})();
