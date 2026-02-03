/*
  Nomogram Calculator
  - Uses Excel file (nomogram_points_tables_12_36_60.xlsx) if available and the XLSX library is loaded.
  - Falls back to embedded data for offline use.
*/

const EMBEDDED_DATA = {
  "variables": [
    {
      "key": "Bloodloss",
      "label": "Bloodloss",
      "options": [
        {
          "label": "≤150",
          "value": "≤150",
          "points": 8
        },
        {
          "label": ">150",
          "value": ">150",
          "points": 26
        }
      ],
      "minPoints": 8,
      "maxPoints": 26
    },
    {
      "key": "CA125",
      "label": "CA125",
      "options": [
        {
          "label": "≤14.05",
          "value": "≤14.05",
          "points": 4
        },
        {
          "label": ">14.05",
          "value": ">14.05",
          "points": 26
        }
      ],
      "minPoints": 4,
      "maxPoints": 26
    },
    {
      "key": "PNII",
      "label": "PNII",
      "options": [
        {
          "label": "Low",
          "value": "Low",
          "points": 0
        },
        {
          "label": "High",
          "value": "High",
          "points": 26
        }
      ],
      "minPoints": 0,
      "maxPoints": 26
    },
    {
      "key": "Lymph_metastases",
      "label": "Lymph metastases",
      "options": [
        {
          "label": "≤3",
          "value": "≤3",
          "points": 0
        },
        {
          "label": ">3",
          "value": ">3",
          "points": 26
        }
      ],
      "minPoints": 0,
      "maxPoints": 26
    },
    {
      "key": "pTNM",
      "label": "pTNM",
      "options": [
        {
          "label": "Stage I",
          "value": "Stage I",
          "points": 26
        },
        {
          "label": "Stage II",
          "value": "Stage II",
          "points": 78
        },
        {
          "label": "Stage III",
          "value": "Stage III",
          "points": 100
        }
      ],
      "minPoints": 26,
      "maxPoints": 100
    },
    {
      "key": "Age",
      "label": "Age",
      "options": [
        {
          "label": "≤66",
          "value": "≤66",
          "points": 2
        },
        {
          "label": ">66",
          "value": ">66",
          "points": 26
        }
      ],
      "minPoints": 2,
      "maxPoints": 26
    }
  ],
  "survivalTable": [
    {
      "points": 40,
      "s12": 0.9899,
      "s36": 0.9668,
      "s60": 0.9521
    },
    {
      "points": 60,
      "s12": 0.9844,
      "s36": 0.9491,
      "s60": 0.927
    },
    {
      "points": 80,
      "s12": 0.976,
      "s36": 0.9225,
      "s60": 0.8894
    },
    {
      "points": 100,
      "s12": 0.9631,
      "s36": 0.8828,
      "s60": 0.8344
    },
    {
      "points": 120,
      "s12": 0.9436,
      "s36": 0.8248,
      "s60": 0.7559
    },
    {
      "points": 140,
      "s12": 0.9142,
      "s36": 0.7426,
      "s60": 0.649
    },
    {
      "points": 160,
      "s12": 0.8705,
      "s36": 0.6313,
      "s60": 0.5126
    },
    {
      "points": 180,
      "s12": 0.8071,
      "s36": 0.4912,
      "s60": 0.3561
    },
    {
      "points": 200,
      "s12": 0.718,
      "s36": 0.3334,
      "s60": 0.2028
    },
    {
      "points": 220,
      "s12": 0.5994,
      "s36": 0.1831,
      "s60": 0.08494
    },
    {
      "points": 240,
      "s12": 0.4534,
      "s36": 0.07257,
      "s60": 0.02214
    },
    {
      "points": 260,
      "s12": 0.2945,
      "s36": 0.01736,
      "s60": 0.002771
    }
  ],
  "totalPointsRange": {
    "minByVariables": 40,
    "maxByVariables": 230,
    "minByTable": 40,
    "maxByTable": 260
  }
};

function prettyLabel(key) {
  return String(key).replace(/_/g, ' ');
}

function clamp(v, min, max) {
  return Math.min(max, Math.max(min, v));
}

function interpolateSurvival(table, totalPoints, field) {
  // Piecewise linear interpolation over the survival table.
  // If totalPoints is outside the table range, it will be clamped to the nearest boundary.
  if (!Array.isArray(table) || table.length === 0) return NaN;

  const sorted = [...table].sort((a, b) => Number(a.points) - Number(b.points));
  const x = clamp(Number(totalPoints), Number(sorted[0].points), Number(sorted[sorted.length - 1].points));

  // Exact match
  for (const row of sorted) {
    if (Number(row.points) === x) return Number(row[field]);
  }

  // Find segment
  for (let i = 0; i < sorted.length - 1; i++) {
    const x0 = Number(sorted[i].points);
    const x1 = Number(sorted[i + 1].points);
    if (x >= x0 && x <= x1) {
      const y0 = Number(sorted[i][field]);
      const y1 = Number(sorted[i + 1][field]);
      if (!Number.isFinite(y0) || !Number.isFinite(y1) || x1 === x0) return NaN;
      const t = (x - x0) / (x1 - x0);
      return y0 + t * (y1 - y0);
    }
  }

  return Number(sorted[sorted.length - 1][field]);
}

function formatPercent(p) {
  if (!Number.isFinite(p)) return '—';
  return (p * 100).toFixed(1) + '%';
}

function sheetToRows(ws) {
  // rows: Array<Array<any>>
  return XLSX.utils.sheet_to_json(ws, { header: 1, blankrows: false, defval: null });
}

function parseWorkbookToData(workbook) {
  const sheetNames = workbook.SheetNames || [];

  const variables = [];

  for (const sheetName of sheetNames) {
    if (sheetName === 'TotalPoints_AllTimes') continue;

    const ws = workbook.Sheets[sheetName];
    if (!ws) continue;

    const rows = sheetToRows(ws);
    if (!rows || rows.length < 2) continue;

    const key = rows[0]?.[0];
    if (!key) continue;

    const options = [];
    for (let i = 1; i < rows.length; i++) {
      const label = rows[i]?.[0];
      const pts = rows[i]?.[1];
      if (label == null || pts == null) continue;
      const pointsNum = Number(pts);
      if (!Number.isFinite(pointsNum)) continue;
      const labelStr = String(label).trim();
      options.push({ label: labelStr, value: labelStr, points: pointsNum });
    }

    if (options.length === 0) continue;

    const minPoints = Math.min(...options.map(o => o.points));
    const maxPoints = Math.max(...options.map(o => o.points));

    variables.push({
      key: String(key).trim(),
      label: prettyLabel(String(key).trim()),
      options,
      minPoints,
      maxPoints,
    });
  }

  // Survival table
  const survivalSheet = workbook.Sheets['TotalPoints_AllTimes'];
  if (!survivalSheet) {
    throw new Error('Missing sheet: TotalPoints_AllTimes');
  }

  const sRows = sheetToRows(survivalSheet);
  if (!sRows || sRows.length < 2) {
    throw new Error('Empty sheet: TotalPoints_AllTimes');
  }

  // Expect headers: Total Points | Survival_12 | Survival_36 | Survival_60
  const survivalTable = [];
  for (let i = 1; i < sRows.length; i++) {
    const p = sRows[i]?.[0];
    const s12 = sRows[i]?.[1];
    const s36 = sRows[i]?.[2];
    const s60 = sRows[i]?.[3];
    if (p == null || s12 == null || s36 == null || s60 == null) continue;
    const row = {
      points: Number(p),
      s12: Number(s12),
      s36: Number(s36),
      s60: Number(s60),
    };
    if ([row.points, row.s12, row.s36, row.s60].every(Number.isFinite)) {
      survivalTable.push(row);
    }
  }

  const minByVariables = variables.reduce((acc, v) => acc + Number(v.minPoints), 0);
  const maxByVariables = variables.reduce((acc, v) => acc + Number(v.maxPoints), 0);

  const sortedTable = [...survivalTable].sort((a, b) => a.points - b.points);
  const minByTable = sortedTable.length ? sortedTable[0].points : minByVariables;
  const maxByTable = sortedTable.length ? sortedTable[sortedTable.length - 1].points : maxByVariables;

  return {
    variables,
    survivalTable,
    totalPointsRange: {
      minByVariables,
      maxByVariables,
      minByTable,
      maxByTable,
    }
  };
}

async function loadNomogramData() {
  // If XLSX isn't available (offline/no CDN), use embedded.
  if (typeof XLSX === 'undefined') {
    return EMBEDDED_DATA;
  }

  try {
    const res = await fetch('./nomogram_points_tables_12_36_60.xlsx', { cache: 'no-store' });
    if (!res.ok) throw new Error('Failed to load Excel file');
    const buf = await res.arrayBuffer();
    const workbook = XLSX.read(buf, { type: 'array' });
    return parseWorkbookToData(workbook);
  } catch (e) {
    console.warn('[Nomogram] Excel load failed, using embedded data.', e);
    return EMBEDDED_DATA;
  }
}

function createEl(tag, className, text) {
  const el = document.createElement(tag);
  if (className) el.className = className;
  if (text != null) el.textContent = text;
  return el;
}

function renderHeader(container) {
  const head = createEl('div', 'header');

  const title = createEl('div', 'title');
  title.innerHTML = `
    <h1>Nomogram Calculator</h1>
    <p class="subtitle">Select each variable to calculate total points and the predicted survival probability at <b>12</b>, <b>36</b>, and <b>60</b> months.</p>
  `;

  head.appendChild(title);

  const axis = createEl('div', 'axis');
  const axisLabel = createEl('div', 'axis-label', 'Points');
  const axisLine = createEl('div', 'axis-line');

  // Ticks 0..100 step 10
  for (let t = 0; t <= 100; t += 10) {
    const tick = createEl('div', 'axis-tick');
    tick.style.left = t + '%';
    tick.textContent = String(t);
    axisLine.appendChild(tick);
  }

  axis.appendChild(axisLabel);
  axis.appendChild(axisLine);

  container.appendChild(head);
  container.appendChild(axis);

  // Optional: reference image (if present in the same folder)
  const details = createEl('details', 'ref');
  const summary = createEl('summary', 'ref-summary', 'Show reference nomogram');
  const imgWrap = createEl('div', 'ref-body');
  imgWrap.innerHTML = `
    <img src="./nomogram_reference.png" alt="Reference nomogram" loading="lazy" />
  `;
  details.appendChild(summary);
  details.appendChild(imgWrap);
  container.appendChild(details);
}

function renderVariableRow(container, variable, state, onChange) {
  const field = createEl('div', 'field');

  const left = createEl('div', 'field-left');
  const label = createEl('div', 'field-label', variable.label);
  const sub = createEl('div', 'field-sub');
  sub.innerHTML = `Selected points: <span class="mono" id="pts_${variable.key}">—</span>`;
  left.appendChild(label);
  left.appendChild(sub);

  const right = createEl('div', 'field-right');

  // Scale / markers
  const scale = createEl('div', 'scale');
  const line = createEl('div', 'line');
  scale.appendChild(line);

  // Markers
  variable.options.forEach((opt, idx) => {
    const pos = clamp(Number(opt.points), 0, 100);

    const btn = createEl('button', 'mark');
    btn.type = 'button';
    btn.style.left = pos + '%';
    btn.dataset.varKey = variable.key;
    btn.dataset.value = opt.value;
    btn.setAttribute('aria-label', `${variable.label}: ${opt.label} (${opt.points} points)`);

    if (state[variable.key] == null && idx === 0) {
      state[variable.key] = opt.value;
    }

    if (state[variable.key] === opt.value) {
      btn.classList.add('active');
    }

    btn.addEventListener('click', () => {
      state[variable.key] = opt.value;
      onChange();
    });

    const optLabel = createEl('div', 'mark-label', opt.label);
    optLabel.style.left = pos + '%';

    scale.appendChild(btn);
    scale.appendChild(optLabel);
  });

  // Select fallback
  const select = createEl('select', 'select');
  select.name = variable.key;
  variable.options.forEach((opt) => {
    const optionEl = document.createElement('option');
    optionEl.value = opt.value;
    optionEl.textContent = `${opt.label} (${opt.points} pts)`;
    select.appendChild(optionEl);
  });
  select.value = state[variable.key];
  select.addEventListener('change', (e) => {
    state[variable.key] = e.target.value;
    onChange();
  });

  right.appendChild(scale);
  right.appendChild(select);

  field.appendChild(left);
  field.appendChild(right);

  container.appendChild(field);
}

function renderTotalBlock(container, data, state) {
  const total = createEl('div', 'total');

  const head = createEl('div', 'total-head');
  const ttl = createEl('div', 'total-label', 'Total points');
  const val = createEl('div', 'total-value');
  val.innerHTML = `<span class="mono" id="totalPts">—</span>`;
  head.appendChild(ttl);
  head.appendChild(val);

  const scale = createEl('div', 'total-scale');
  const line = createEl('div', 'line');
  scale.appendChild(line);

  // Range ticks based on the survival table
  const minP = Number(data.totalPointsRange?.minByTable ?? 0);
  const maxP = Number(data.totalPointsRange?.maxByTable ?? 100);

  const tickCount = 6;
  for (let i = 0; i <= tickCount; i++) {
    const p = minP + (i * (maxP - minP)) / tickCount;
    const t = createEl('div', 'total-tick', String(Math.round(p)));
    t.style.left = (i * 100) / tickCount + '%';
    scale.appendChild(t);
  }

  const marker = createEl('div', 'total-marker');
  marker.id = 'totalMarker';
  marker.style.left = '0%';
  scale.appendChild(marker);

  total.appendChild(head);
  total.appendChild(scale);

  container.appendChild(total);
}

function getSelectedOption(variable, value) {
  return variable.options.find(o => o.value === value) || variable.options[0];
}

function updateUI(data, state) {
  // Update active markers and selects
  data.variables.forEach(v => {
    const selectedValue = state[v.key];

    // Update marker classes
    document.querySelectorAll(`.mark[data-var-key="${CSS.escape(v.key)}"]`).forEach(btn => {
      const isActive = btn.dataset.value === selectedValue;
      btn.classList.toggle('active', isActive);
    });

    // Update select value
    const sel = document.querySelector(`select[name="${CSS.escape(v.key)}"]`);
    if (sel) sel.value = selectedValue;

    // Update per-variable points
    const opt = getSelectedOption(v, selectedValue);
    const span = document.getElementById(`pts_${v.key}`);
    if (span) span.textContent = String(opt.points);
  });

  // Total points
  let totalPoints = 0;
  data.variables.forEach(v => {
    const opt = getSelectedOption(v, state[v.key]);
    totalPoints += Number(opt.points) || 0;
  });

  const totalPtsEl = document.getElementById('totalPts');
  if (totalPtsEl) totalPtsEl.textContent = String(totalPoints);

  // Total marker position
  const minP = Number(data.totalPointsRange?.minByTable ?? 0);
  const maxP = Number(data.totalPointsRange?.maxByTable ?? 100);
  const t = (clamp(totalPoints, minP, maxP) - minP) / (maxP - minP || 1);
  const marker = document.getElementById('totalMarker');
  if (marker) marker.style.left = (t * 100) + '%';

  // Survival
  const s12 = interpolateSurvival(data.survivalTable, totalPoints, 's12');
  const s36 = interpolateSurvival(data.survivalTable, totalPoints, 's36');
  const s60 = interpolateSurvival(data.survivalTable, totalPoints, 's60');

  const s1 = document.getElementById('s1');
  const s3 = document.getElementById('s3');
  const s5 = document.getElementById('s5');

  if (s1) s1.textContent = formatPercent(s12);
  if (s3) s3.textContent = formatPercent(s36);
  if (s5) s5.textContent = formatPercent(s60);

  // Optional debug
  const dbg = document.getElementById('debug');
  if (dbg) {
    dbg.textContent = `Total=${totalPoints} | s12=${(Number.isFinite(s12) ? s12.toFixed(4) : 'NaN')} s36=${(Number.isFinite(s36) ? s36.toFixed(4) : 'NaN')} s60=${(Number.isFinite(s60) ? s60.toFixed(4) : 'NaN')}`;
  }
}

function attachReset(container, data, state, onChange) {
  const bar = createEl('div', 'actions');
  const btn = createEl('button', 'btn', 'Reset');
  btn.type = 'button';
  btn.addEventListener('click', () => {
    data.variables.forEach(v => {
      state[v.key] = v.options[0].value;
    });
    onChange();
  });

  const hint = createEl('div', 'actions-hint');
  hint.textContent = 'Tip: You can click the markers on the point scale, or use the dropdown.';

  bar.appendChild(btn);
  bar.appendChild(hint);
  container.appendChild(bar);
}

async function main() {
  const inputs = document.getElementById('inputs');
  if (!inputs) return;

  const data = await loadNomogramData();

  // State: selected option value per variable
  const state = {};

  // Render
  inputs.innerHTML = '';
  renderHeader(inputs);

  const onChange = () => updateUI(data, state);

  data.variables.forEach(v => {
    renderVariableRow(inputs, v, state, onChange);
  });

  renderTotalBlock(inputs, data, state);
  attachReset(inputs, data, state, onChange);

  // Initial update
  onChange();
}

document.addEventListener('DOMContentLoaded', main);
