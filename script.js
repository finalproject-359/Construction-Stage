async function loadData() {
  const res = await fetch('data/activities.csv');
  const text = await res.text();

  const rows = text.split('\n').slice(1);
  let totalBudget = 0;

  rows.forEach(row => {
    const cols = row.split(',');
    totalBudget += Number(cols[2]);
  });

  document.getElementById("budget").innerText = totalBudget;
}

loadData();
