
const DATA = JSON.parse(document.getElementById('dashboard-data').textContent);
const SECTION_ORDER = ["Data Check", "Infra Check", "Process Check", "Aspirant Interaction"];
const SECTION_COLORS = {
  "Data Check": "#1F3864",
  "Infra Check": "#2E7D32",
  "Process Check": "#C77700",
  "Aspirant Interaction": "#7B1FA2"
};

const managerSel = document.getElementById('managerSel');
const centreSel = document.getElementById('centreSel');

// Build manager -> centres map
const managerCentres = {};
DATA.visits.forEach(v => {
  if (!managerCentres[v.manager]) managerCentres[v.manager] = new Set();
  managerCentres[v.manager].add(v.centre);
});
const allManagers = Object.keys(managerCentres).sort();
const allCentres = Array.from(new Set(DATA.visits.map(v => v.centre))).sort();

function populateManagerSelect() {
  managerSel.innerHTML = '<option value="__ALL__">All Managers (' + allManagers.length + ')</option>' +
    allManagers.map(m => `<option value="${m}">${m} (${managerCentres[m].size})</option>`).join('');
}
function populateCentreSelect(filterManager) {
  let centres;
  if (filterManager && filterManager !== '__ALL__') {
    centres = Array.from(managerCentres[filterManager]).sort();
  } else {
    centres = allCentres;
  }
  const current = centreSel.value;
  centreSel.innerHTML = '<option value="__ALL__">All Centres (' + centres.length + ')</option>' +
    centres.map(c => `<option value="${c}">${c}</option>`).join('');
  if (centres.includes(current)) centreSel.value = current;
}

populateManagerSelect();
populateCentreSelect('__ALL__');

let sectionBarChart, sectionStackChart;

function getFilteredVisits() {
  const m = managerSel.value;
  const c = centreSel.value;
  return DATA.visits.filter(v => (m === '__ALL__' || v.manager === m) && (c === '__ALL__' || v.centre === c));
}

function fmtPct(n) { return isFinite(n) ? n.toFixed(1) + '%' : '—'; }

function render() {
  const visits = getFilteredVisits();
  const scopeParts = [];
  if (managerSel.value !== '__ALL__') scopeParts.push(managerSel.value);
  if (centreSel.value !== '__ALL__') scopeParts.push(centreSel.value);
  document.getElementById('scopeBadge').textContent = scopeParts.length ? 'Viewing: ' + scopeParts.join(' → ') : 'Viewing: All managers, all centres';

  // Section aggregates
  const sectionAgg = {};
  SECTION_ORDER.forEach(s => sectionAgg[s] = { aw: 0, ni: 0, na: 0 });
  // Item aggregates keyed by section|label
  const itemAgg = {};

  visits.forEach(v => {
    v.items.forEach(it => {
      const sec = sectionAgg[it.section];
      if (!sec) return;
      if (it.status === 'All is Well') sec.aw++;
      else if (it.status === 'Need Improvement') sec.ni++;
      else sec.na++;

      const key = it.section + '|||' + it.label;
      if (!itemAgg[key]) itemAgg[key] = { section: it.section, label: it.label, aw: 0, ni: 0, na: 0 };
      if (it.status === 'All is Well') itemAgg[key].aw++;
      else if (it.status === 'Need Improvement') itemAgg[key].ni++;
      else itemAgg[key].na++;
    });
  });

  // Cards
  document.getElementById('cardVisits').textContent = visits.length;
  const distinctCentres = new Set(visits.map(v => v.centre)).size;
  const distinctManagers = new Set(visits.map(v => v.manager)).size;
  document.getElementById('cardVisitsSub').textContent = distinctCentres + ' centre(s), ' + distinctManagers + ' manager(s)';

  let totalAw = 0, totalNi = 0;
  SECTION_ORDER.forEach(s => { totalAw += sectionAgg[s].aw; totalNi += sectionAgg[s].ni; });
  const overallPct = totalAw + totalNi > 0 ? (100 * totalNi / (totalAw + totalNi)) : NaN;
  document.getElementById('cardNI').textContent = fmtPct(overallPct);

  let weakestSection = null, weakestPct = -1;
  SECTION_ORDER.forEach(s => {
    const app = sectionAgg[s].aw + sectionAgg[s].ni;
    const pct = app > 0 ? 100 * sectionAgg[s].ni / app : -1;
    if (pct > weakestPct) { weakestPct = pct; weakestSection = s; }
  });
  document.getElementById('cardWeakSection').textContent = weakestSection || '—';
  document.getElementById('cardWeakSectionSub').textContent = weakestPct >= 0 ? fmtPct(weakestPct) + ' Need Improvement' : 'no applicable data';

  const itemList = Object.values(itemAgg).map(it => {
    const app = it.aw + it.ni;
    return { ...it, app, pct: app > 0 ? 100 * it.ni / app : -1 };
  }).filter(it => it.ni > 0).sort((a, b) => b.ni - a.ni || b.pct - a.pct);

  if (itemList.length) {
    document.getElementById('cardTopItem').textContent = itemList[0].label.length > 70 ? itemList[0].label.slice(0, 68) + '…' : itemList[0].label;
    document.getElementById('cardTopItemSub').textContent = itemList[0].ni + ' of ' + itemList[0].app + ' visits (' + fmtPct(itemList[0].pct) + ') — ' + itemList[0].section;
  } else {
    document.getElementById('cardTopItem').textContent = 'None';
    document.getElementById('cardTopItemSub').textContent = 'No Need Improvement flags in this view';
  }

  // Charts
  const barLabels = SECTION_ORDER;
  const barData = SECTION_ORDER.map(s => {
    const app = sectionAgg[s].aw + sectionAgg[s].ni;
    return app > 0 ? +(100 * sectionAgg[s].ni / app).toFixed(1) : 0;
  });
  const barColors = SECTION_ORDER.map(s => SECTION_COLORS[s]);

  if (sectionBarChart) sectionBarChart.destroy();
  sectionBarChart = new Chart(document.getElementById('sectionBarChart'), {
    type: 'bar',
    data: { labels: barLabels, datasets: [{ label: '% Need Improvement', data: barData, backgroundColor: barColors, borderRadius: 6 }] },
    options: {
      indexAxis: 'y',
      responsive: true, maintainAspectRatio: false,
      plugins: { legend: { display: false }, title: { display: true, text: '% Need Improvement by Section', font: { size: 13 } } },
      scales: { x: { beginAtZero: true, max: 100, ticks: { callback: v => v + '%' } } }
    }
  });

  if (sectionStackChart) sectionStackChart.destroy();
  sectionStackChart = new Chart(document.getElementById('sectionStackChart'), {
    type: 'bar',
    data: {
      labels: SECTION_ORDER,
      datasets: [
        { label: 'All is Well', data: SECTION_ORDER.map(s => sectionAgg[s].aw), backgroundColor: '#2E7D32', stack: 's' },
        { label: 'Need Improvement', data: SECTION_ORDER.map(s => sectionAgg[s].ni), backgroundColor: '#C77700', stack: 's' },
        { label: 'Not Applicable', data: SECTION_ORDER.map(s => sectionAgg[s].na), backgroundColor: '#B9BDC3', stack: 's' },
      ]
    },
    options: {
      responsive: true, maintainAspectRatio: false,
      plugins: { legend: { display: false }, title: { display: true, text: 'Item Counts by Section (stacked)', font: { size: 13 } } },
      scales: { x: { stacked: true }, y: { stacked: true, beginAtZero: true } }
    }
  });

  // Spotlight table
  const tbody = document.getElementById('spotlightBody');
  const emptyNote = document.getElementById('spotlightEmpty');
  const top = itemList.slice(0, 15);
  if (!top.length) {
    tbody.innerHTML = '';
    emptyNote.style.display = 'block';
  } else {
    emptyNote.style.display = 'none';
    tbody.innerHTML = top.map(it => `
      <tr>
        <td><span class="sec-tag" style="background:${SECTION_COLORS[it.section]}">${it.section}</span></td>
        <td>${it.label}</td>
        <td class="ni-pct">${it.ni} / ${it.app}</td>
        <td>
          <div class="bar-cell"><div class="bar-fill" style="width:${Math.max(it.pct,0)}%"></div></div>
          <span style="font-size:11.5px; color:var(--text-muted);">${fmtPct(it.pct)}</span>
        </td>
      </tr>
    `).join('');
  }

  // Reflections
  const container = document.getElementById('reflectionsContainer');
  if (!visits.length) {
    container.innerHTML = '<div class="empty-note">No visits match this selection.</div>';
  } else {
    container.innerHTML = visits
      .slice()
      .sort((a, b) => a.centre.localeCompare(b.centre))
      .map(v => `
        <div class="refl-card">
          <div class="meta"><span class="centre">${v.centre}</span> · ${v.program || ''} · ${v.manager} · ${v.date}</div>
          ${qaBlock('Key bottleneck resolved', v.q_bottleneck)}
          ${qaBlock('Action to improve placement %', v.q_placement)}
          ${qaBlock('Action to help batch start on time', v.q_batch_ontime)}
        </div>
      `).join('');
  }
}

function qaBlock(q, a) {
  const text = (a && a.trim()) ? escapeHtml(a) : '<span style="color:var(--text-muted); font-style:italic;">Not answered</span>';
  return `<div class="qa"><div class="q">${q}</div><div class="a">${text}</div></div>`;
}
function escapeHtml(s) {
  return s.replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;');
}

managerSel.addEventListener('change', () => {
  populateCentreSelect(managerSel.value);
  centreSel.value = '__ALL__';
  render();
});
centreSel.addEventListener('change', render);
document.getElementById('resetBtn').addEventListener('click', () => {
  managerSel.value = '__ALL__';
  populateCentreSelect('__ALL__');
  centreSel.value = '__ALL__';
  render();
});

render();
