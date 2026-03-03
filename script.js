// Offline-ready Nomogram Calculator
// Data embedded from nomogram_points_tables_12_36_60.xlsx
const CONFIG = {
  "variables": [
    {
      "key": "Age",
      "label": "Age",
      "type": "select",
      "options": [
        {
          "label": "≤66",
          "value": "≤66",
          "points": 11.0
        },
        {
          "label": ">66",
          "value": ">66",
          "points": 35.0
        }
      ]
    },
    {
      "key": "pT",
      "label": "pT",
      "type": "select",
      "options": [
        {
          "label": "T1",
          "value": "T1",
          "points": 35.0
        },
        {
          "label": "T2",
          "value": "T2",
          "points": 46.0
        },
        {
          "label": "T4",
          "value": "T4",
          "points": 88.0
        },
        {
          "label": "T3",
          "value": "T3",
          "points": 100.0
        }
      ]
    },
    {
      "key": "pN",
      "label": "pN",
      "type": "select",
      "options": [
        {
          "label": "N0",
          "value": "N0",
          "points": 35.0
        },
        {
          "label": "N1",
          "value": "N1",
          "points": 57.0
        },
        {
          "label": "N2",
          "value": "N2",
          "points": 83.0
        },
        {
          "label": "N3",
          "value": "N3",
          "points": 99.0
        }
      ]
    },
    {
      "key": "PNII",
      "label": "PNII",
      "type": "select",
      "options": [
        {
          "label": "Low",
          "value": "Low",
          "points": 0.0
        },
        {
          "label": "High",
          "value": "High",
          "points": 35.0
        }
      ]
    },
    {
      "key": "CA125",
      "label": "CA125",
      "type": "select",
      "options": [
        {
          "label": "Low",
          "value": "Low",
          "points": 4.0
        },
        {
          "label": "High",
          "value": "High",
          "points": 35.0
        }
      ]
    }
  ],
  "totalPointsCurve": {
    "points": [
      50.0,
      100.0,
      150.0,
      200.0,
      250.0,
      300.0,
      350.0
    ],
    "s12": [
      0.9934,
      0.9824,
      0.954,
      0.8823,
      0.717,
      0.413,
      0.0953
    ],
    "s36": [
      0.9781,
      0.9429,
      0.8553,
      0.6601,
      0.3315,
      0.05313,
      0.0004091
    ],
    "s60": [
      0.9681,
      0.9176,
      0.7956,
      0.5445,
      0.1987,
      0.01364,
      1.102e-05
    ]
  }
};

const elInputs = document.getElementById('inputs');
const elS1 = document.getElementById('s1');
const elS3 = document.getElementById('s3');
const elS5 = document.getElementById('s5');

function clamp(x, a, b) {
  return Math.max(a, Math.min(b, x));
}

function lerp(x0, y0, x1, y1, x) {
  if (x1 === x0) return y0;
  return y0 + (y1 - y0) * ((x - x0) / (x1 - x0));
}

function interp(xs, ys, x) {
  const n = xs.length;
  if (n === 0) return NaN;
  if (x <= xs[0]) return ys[0];
  if (x >= xs[n-1]) return ys[n-1];
  for (let i = 0; i < n - 1; i++) {
    const x0 = xs[i], x1 = xs[i+1];
    if (x >= x0 && x <= x1) {
      return lerp(x0, ys[i], x1, ys[i+1], x);
    }
  }
  return ys[n-1];
}

function fmtProb(p) {
  if (!isFinite(p)) return '—';
  const v = clamp(p, 0, 1) * 100;
  return v.toFixed(1) + '%';
}

function buildUI() {
  elInputs.innerHTML = `
    <h1>Nomogram Calculator</h1>
    <p class="sub">Select each variable level to calculate total points and the predicted survival probability at 12, 36, and 60 months.</p>
    <div id="fields"></div>
    <hr class="sep" />
    <div class="total">
      <div class="tlabel">Total points</div>
      <div class="tvalue" id="totalPts">—</div>
    </div>
  `;

  const fields = elInputs.querySelector('#fields');

  for (const v of CONFIG.variables) {
    const row = document.createElement('div');
    row.className = 'field';
    row.innerHTML = `
      <div class="label">${v.label}</div>
      <div>
        <select data-key="${v.key}"></select>
      </div>
      <div class="badge" id="pts_${v.key}">—</div>
    `;
    fields.appendChild(row);

    const sel = row.querySelector('select');
    v.options.forEach((opt, idx) => {
      const o = document.createElement('option');
      o.value = opt.value;
      o.textContent = opt.label;
      sel.appendChild(o);
      if (idx === 0) sel.value = opt.value;
    });
    sel.addEventListener('change', compute);
  }

  compute();
}

function getSelectedPoints(varKey, levelValue) {
  const v = CONFIG.variables.find(x => x.key === varKey);
  if (!v) return 0;
  const opt = v.options.find(o => String(o.value) === String(levelValue));
  return opt ? Number(opt.points) : 0;
}

function compute() {
  let total = 0;

  for (const v of CONFIG.variables) {
    const sel = document.querySelector(`select[data-key="${v.key}"]`);
    const val = sel ? sel.value : null;
    const pts = getSelectedPoints(v.key, val);
    total += pts;

    const badge = document.getElementById(`pts_${v.key}`);
    if (badge) badge.textContent = isFinite(pts) ? pts.toFixed(0) : '—';
  }

  document.getElementById('totalPts').textContent = isFinite(total) ? total.toFixed(0) : '—';

  const xs = CONFIG.totalPointsCurve.points;
  const s12 = interp(xs, CONFIG.totalPointsCurve.s12, total);
  const s36 = interp(xs, CONFIG.totalPointsCurve.s36, total);
  const s60 = interp(xs, CONFIG.totalPointsCurve.s60, total);

  elS1.textContent = fmtProb(s12);
  elS3.textContent = fmtProb(s36);
  elS5.textContent = fmtProb(s60);
}

buildUI();
