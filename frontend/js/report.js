// ── Report tab ────────────────────────────────────────────────────────────────

const reportState = { files: [] };

const reportDropzone      = document.getElementById('reportDropzone');
const reportFileInput     = document.getElementById('reportFileInput');
const reportFileList      = document.getElementById('reportFileList');
const btnGenReport        = document.getElementById('btnGenReport');
const btnClearReport      = document.getElementById('btnClearReport');
const reportError         = document.getElementById('reportError');
const reportProgressCard  = document.getElementById('reportProgressCard');
const reportLog           = document.getElementById('reportLog');
const reportDownloads     = document.getElementById('reportDownloads');
const reportDownloadLinks = document.getElementById('reportDownloadLinks');

function detectFileType(name) {
  const n = name.toLowerCase();
  if (n === 'metadata.xlsx' || n === 'metadata.xls') return { label: 'Metadata', cls: '' };
  if (/^\d{8}\s*-/.test(name)) return { label: 'AI Result', cls: 'warn' };
  if (/-\s*tiab\.\w+$/i.test(name)) return { label: 'Human TIAB', cls: 'ok' };
  if (/-\s*fulltext\.\w+$/i.test(name)) return { label: 'Fulltext', cls: 'ok' };
  if (/-\s*listfinal\.\w+$/i.test(name)) return { label: 'Listfinal', cls: 'ok' };
  return { label: 'File', cls: '' };
}

function renderReportFileList() {
  if (!reportState.files.length) { reportFileList.classList.add('hidden'); return; }
  reportFileList.classList.remove('hidden');
  reportFileList.innerHTML = reportState.files.map((f, i) => {
    const type = detectFileType(f.name);
    const size = (f.size / 1024).toFixed(1) + ' KB';
    return `<div class="report-file-item">
      <span class="report-file-name">${f.name}</span>
      <span class="badge ${type.cls}">${type.label}</span>
      <span class="report-file-size">${size}</span>
      <button class="icon-btn remove-report-file" data-idx="${i}" title="Remove">×</button>
    </div>`;
  }).join('');
  reportFileList.querySelectorAll('.remove-report-file').forEach(btn => {
    btn.addEventListener('click', () => {
      reportState.files.splice(Number(btn.dataset.idx), 1);
      renderReportFileList();
      refreshProjectClassification();
      btnGenReport.disabled = !reportState.files.length;
    });
  });
  btnGenReport.disabled = false;
}

function addReportFiles(fileList) {
  for (const f of fileList) {
    if (!reportState.files.find(x => x.name === f.name)) reportState.files.push(f);
  }
  renderReportFileList();
  refreshProjectClassification();
}

// ── Project classification (used by Paper Figures & Tables) ────────────────

const AI_FILENAME_RE  = /^(\d{8})\s*-\s*(.+?)\s*-\s*(\d)[ºo°]\s*teste\s*-\s*(.+)\.(?:xlsx|xls|csv)$/i;
const HUMAN_FILENAME_RE = /^(.+?)\s*-\s*(?:TIAB|Fulltext|Listfinal)\.(?:xlsx|xls|csv)$/i;
const _CLASS_LS_KEY = "reportProjectClasses";

/** Extract normalized project keys from the current file list. */
function detectProjectsFromFiles() {
  const set = new Map();  // project_norm -> display name (last seen)
  for (const f of reportState.files) {
    const name = f.name;
    let display = null;
    const mAi = name.match(AI_FILENAME_RE);
    if (mAi) display = mAi[4].trim();
    const mHu = name.match(HUMAN_FILENAME_RE);
    if (!display && mHu) display = mHu[1].trim();
    if (display) {
      const norm = display.trim().toLowerCase();
      set.set(norm, display);
    }
  }
  return Array.from(set.entries()).map(([norm, disp]) => ({ norm, display: disp }));
}

/** Guess the default category from a project name (falls back to 'official'). */
function guessDefaultCategory(projNorm) {
  const s = String(projNorm).toLowerCase();
  if (/mino/.test(s))     return "pilot";
  if (/modified|mod$/.test(s)) return "sensitivity";
  return "official";
}

/** Render or refresh the classification UI based on currently uploaded files. */
function refreshProjectClassification() {
  const section = document.getElementById("projectClassSection");
  const list = document.getElementById("projectClassList");
  if (!section || !list) return;

  const detected = detectProjectsFromFiles();
  if (!detected.length) {
    section.classList.add("hidden");
    list.innerHTML = "";
    return;
  }

  // Load previously saved choices (persisted across reloads)
  let saved = {};
  try { saved = JSON.parse(localStorage.getItem(_CLASS_LS_KEY) || "{}"); } catch {}

  section.classList.remove("hidden");
  list.innerHTML = detected.map(({ norm, display }) => {
    const current = saved[norm] || guessDefaultCategory(norm);
    const opt = (val, label) => `
      <label>
        <input type="radio" name="pclass_${norm}" value="${val}" ${current === val ? "checked" : ""} />
        <span>${label}</span>
      </label>`;
    return `<div class="project-class-row" data-project="${norm}">
      <span class="project-class-name" title="${display}">${display}
        <span class="muted" style="font-size:11px; margin-left:6px;">(${norm})</span>
      </span>
      <div class="project-class-radios">
        ${opt("pilot", "Pilot")}
        ${opt("official", "Official")}
        ${opt("sensitivity", "Sensitivity")}
      </div>
    </div>`;
  }).join("");

  // Persist choices whenever the user changes them
  list.querySelectorAll('input[type="radio"]').forEach(radio => {
    radio.addEventListener("change", () => {
      const row = radio.closest(".project-class-row");
      if (!row) return;
      const pn = row.dataset.project;
      let store = {};
      try { store = JSON.parse(localStorage.getItem(_CLASS_LS_KEY) || "{}"); } catch {}
      store[pn] = radio.value;
      localStorage.setItem(_CLASS_LS_KEY, JSON.stringify(store));
    });
  });
}

/** Read the current classification from the UI (project_norm -> category). */
function readProjectClassification() {
  const out = {};
  document.querySelectorAll("#projectClassList .project-class-row").forEach(row => {
    const pn = row.dataset.project;
    const checked = row.querySelector('input[type="radio"]:checked');
    if (pn && checked) out[pn] = checked.value;
  });
  return out;
}

reportDropzone.addEventListener('dragover', e => { e.preventDefault(); reportDropzone.classList.add('hover'); });
reportDropzone.addEventListener('dragleave', () => reportDropzone.classList.remove('hover'));
reportDropzone.addEventListener('drop', e => { e.preventDefault(); reportDropzone.classList.remove('hover'); if (e.dataTransfer.files.length) addReportFiles(e.dataTransfer.files); });
reportFileInput.addEventListener('change', e => { if (e.target.files.length) { addReportFiles(e.target.files); e.target.value = ''; } });

btnClearReport.addEventListener('click', () => {
  reportState.files = [];
  renderReportFileList();
  refreshProjectClassification();
  btnGenReport.disabled = true;
  reportProgressCard.classList.add('hidden');
  reportDownloads.classList.add('hidden');
  reportError.classList.add('hidden');
  document.getElementById('reportResultsCard').classList.add('hidden');
  _currentReportJobId = null;
});

// ── Chart gallery ──────────────────────────────────────────────────────────────

// Global tooltip element for chart info buttons
const _chartGlobalTip = document.createElement('div');
_chartGlobalTip.className = 'chart-global-tip hidden';
document.body.appendChild(_chartGlobalTip);

function _setupChartInfoBtn(btn, text) {
  btn.addEventListener('mouseenter', () => {
    _chartGlobalTip.textContent = text;
    _chartGlobalTip.style.visibility = 'hidden';
    _chartGlobalTip.style.top = '0px';
    _chartGlobalTip.style.left = '0px';
    _chartGlobalTip.classList.remove('hidden');
    const r = btn.getBoundingClientRect();
    const h = _chartGlobalTip.offsetHeight;
    const w = 280;
    let left = r.left + r.width / 2 - w / 2;
    left = Math.max(8, Math.min(left, window.innerWidth - w - 8));
    let top = r.top - h - 8;
    if (top < 8) top = r.bottom + 8;
    _chartGlobalTip.style.left = `${left}px`;
    _chartGlobalTip.style.top = `${top}px`;
    _chartGlobalTip.style.visibility = '';
  });
  btn.addEventListener('mouseleave', () => _chartGlobalTip.classList.add('hidden'));
}

const CHART_LABELS = {
  'sensitivity_per_model_per_project':   'Sensitivity by Model per Project',
  'listfinal_capture_heatmap':           'Listfinal Capture Rate Heatmap',
  'test_retest_kappa':                   'Test-Retest Kappa with 95% CI',
  'model_comparison_radar':              'Model Comparison Radar',
  'cost_vs_sensitivity_bubble':          'Cost vs Sensitivity',
  'workload_reduction':                  'Workload Reduction',
  'efficiency_frontier_grid':            'Efficiency Frontier (Individual Runs)',
  'efficiency_frontier_by_project_avg':  'Efficiency Frontier (by Project)',
  'efficiency_frontier_averaged':        'Efficiency Frontier (Overall)',
  'efficiency_score_by_project':         'Efficiency Score by Project',
  'efficiency_score_aggregated':         'Efficiency Score Aggregated',
  'sensitivity_specificity_dual_gold':   'Sensitivity & Specificity (Dual Gold Standard)',
  'aggregated_performance_metrics':      'Aggregated Performance Metrics',
  'f1_score_vs_cost':                    'F1 Score vs Cost',
  'sensitivity_specificity_tradeoff':    'Sensitivity vs Specificity Trade-off',
  'model_ranking_heatmap':               'Model Ranking Heatmap',
};

const CHART_TOOLTIPS = {
  'sensitivity_per_model_per_project':   'Shows the sensitivity (recall) of each AI model per project. Sensitivity = proportion of human-included articles that the AI also kept. Higher is better; 1.0 means no relevant article was missed.',
  'listfinal_capture_heatmap':           'Heatmap of Listfinal capture rate by model and project. Each cell shows what percentage of the final included articles (gold standard) the AI retained. Values near 100% indicate the AI missed very few relevant articles.',
  'test_retest_kappa':                   "Cohen's Kappa measuring consistency between two runs of the same AI model (test-retest reliability). Values ≥ 0.81 = almost perfect, ≥ 0.61 = substantial. Error bars show 95% confidence intervals.",
  'model_comparison_radar':              'Radar chart comparing models across key dimensions: Sensitivity (based on Listfinal capture), Specificity, F1 Score, Fulltext Capture, and Test-Retest Kappa. Each axis ranges 0–1. A larger filled area means more balanced performance.',
  'cost_vs_sensitivity_bubble':          'Cost (USD) vs. Sensitivity per model. Bubble size = F1 Score. Ideal models appear upper-left: high sensitivity at low cost. The dashed line marks 95% sensitivity.',
  'workload_reduction':                  'Compares human review time vs. AI time for the same task. Speed Factor = Human Hours ÷ AI Hours. A factor of 10× means the AI finished in 1/10th the time.',
  'efficiency_frontier_grid':            'Each point = one AI run. X-axis: share of articles flagged as positive by AI (lower = more selective). Y-axis: Listfinal capture rate. Points in the upper-left corner are most efficient.',
  'efficiency_frontier_by_project_avg':  'Same as Efficiency Frontier but averaged per project, reducing noise from individual runs.',
  'efficiency_frontier_averaged':        'Overall efficiency frontier averaged across all projects and runs. Gives a high-level view of each model\'s screening efficiency.',
  'efficiency_score_by_project':         'Efficiency score broken down by project. Efficiency = Capture Rate × (1 − AI Positive Rate). Higher = captures more relevant articles while excluding more irrelevant ones.',
  'efficiency_score_aggregated':         'Aggregated efficiency score per model, averaged across all projects. Error bars show standard deviation. Useful for ranking models by overall efficiency.',
  'sensitivity_specificity_dual_gold':   'Compares sensitivity and specificity using two gold standards: TIAB (human title/abstract decisions) and Listfinal (final included articles). Reveals how model performance differs at each review phase.',
  'aggregated_performance_metrics':      'Summary of all key metrics per model: Sensitivity, Specificity, F1 Score, Accuracy, Listfinal Capture, and Test-Retest Kappa. Allows fast cross-model comparison across all dimensions.',
  'f1_score_vs_cost':                    'F1 Score vs. total cost (USD) per model. F1 balances precision and recall. Models in the upper-left offer the best quality-per-dollar ratio.',
  'sensitivity_specificity_tradeoff':    'Trade-off between sensitivity and specificity for each model and project. High sensitivity often costs lower specificity (more false positives). This reveals each model\'s operating point.',
  'model_ranking_heatmap':               'Heatmap ranking all models across multiple metrics. Color intensity = performance level. The Overall Score column aggregates all metrics into a single ranking.',
};

function chartLabel(filename) {
  const key = filename.replace(/\.png$/, '');
  return CHART_LABELS[key] || key.replace(/_/g, ' ').replace(/\b\w/g, c => c.toUpperCase());
}

function renderChartsGallery(jobId, chartFiles) {
  const grid = document.getElementById('chartsGrid');
  if (!chartFiles.length) { grid.innerHTML = '<p class="muted" style="padding:0 0 8px">No charts were generated.</p>'; return; }
  grid.innerHTML = chartFiles.map(fname => {
    const label = chartLabel(fname);
    const src = `/api/report/chart/${jobId}/${encodeURIComponent(fname)}`;
    const key = fname.replace(/\.png$/, '');
    const hasTip = !!CHART_TOOLTIPS[key];
    return `<div class="chart-card" data-src="${src}" data-label="${label}">
      <img src="${src}" alt="${label}" loading="lazy" />
      <div class="chart-footer">
        <p class="chart-label">${label}</p>
        ${hasTip ? `<button class="chart-info-btn" data-key="${key}" aria-label="Chart information" tabindex="0">ℹ</button>` : ''}
      </div>
    </div>`;
  }).join('');
  grid.querySelectorAll('.chart-card').forEach(card => {
    card.addEventListener('click', e => {
      if (e.target.closest('.chart-info-btn')) return;
      openLightbox(card.dataset.src, card.dataset.label);
    });
  });
  grid.querySelectorAll('.chart-info-btn').forEach(btn => {
    const text = CHART_TOOLTIPS[btn.dataset.key] || '';
    if (text) _setupChartInfoBtn(btn, text);
  });
}

// ── Lightbox ──────────────────────────────────────────────────────────────────

const lightbox        = document.getElementById('chartLightbox');
const lightboxImg     = document.getElementById('lightboxImg');
const lightboxCaption = document.getElementById('lightboxCaption');

function openLightbox(src, caption) {
  lightboxImg.src = src;
  lightboxCaption.textContent = caption;
  lightbox.classList.remove('hidden');
  document.body.style.overflow = 'hidden';
}

function closeLightbox() {
  lightbox.classList.add('hidden');
  lightboxImg.src = '';
  document.body.style.overflow = '';
}

document.querySelector('.lightbox-close').addEventListener('click', closeLightbox);
document.querySelector('.lightbox-backdrop').addEventListener('click', closeLightbox);
document.addEventListener('keydown', e => { if (e.key === 'Escape') closeLightbox(); });

// ── Table viewer ──────────────────────────────────────────────────────────────

let _currentReportJobId = null;

async function loadSheetList(jobId) {
  const sel = document.getElementById('sheetSelector');
  sel.innerHTML = '<option>Loading…</option>';
  try {
    const r = await fetch(`/api/report/tabledata/${jobId}`);
    if (!r.ok) { sel.innerHTML = '<option>Unavailable</option>'; return; }
    const data = await r.json();
    if (!data.sheets || !data.sheets.length) { sel.innerHTML = '<option>No sheets found</option>'; return; }
    sel.innerHTML = data.sheets.map(s => `<option value="${s}">${s}</option>`).join('');
    sel.addEventListener('change', () => loadSheetData(jobId, sel.value));
    loadSheetData(jobId, data.sheets[0]);
  } catch { sel.innerHTML = '<option>Error loading sheets</option>'; }
}

async function loadSheetData(jobId, sheet) {
  const wrap = document.getElementById('sheetTableWrap');
  wrap.innerHTML = '<p class="muted" style="padding:12px">Loading…</p>';
  try {
    const r = await fetch(`/api/report/tabledata/${jobId}/${encodeURIComponent(sheet)}`);
    if (!r.ok) { wrap.innerHTML = '<p class="muted" style="padding:12px">Failed to load sheet.</p>'; return; }
    const data = await r.json();
    if (!data.headers.length) { wrap.innerHTML = '<p class="muted" style="padding:12px">Sheet is empty.</p>'; return; }
    const esc = s => String(s ?? '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
    wrap.innerHTML = `<table>
      <thead><tr>${data.headers.map(h => `<th>${esc(h)}</th>`).join('')}</tr></thead>
      <tbody>${data.rows.map(row => `<tr>${row.map(cell => `<td>${esc(cell)}</td>`).join('')}</tr>`).join('')}</tbody>
    </table>`;
  } catch { wrap.innerHTML = '<p class="muted" style="padding:12px">Error loading data.</p>'; }
}

// ── Report generation ─────────────────────────────────────────────────────────

btnGenReport.addEventListener('click', async () => {
  if (!reportState.files.length) return;
  reportError.classList.add('hidden');
  reportProgressCard.classList.remove('hidden');
  reportLog.textContent = '';
  reportDownloads.classList.add('hidden');
  btnGenReport.disabled = true;

  const form = new FormData();
  reportState.files.forEach(f => form.append('files', f, f.name));

  // Include the user-provided project classification (pilot/official/sensitivity)
  const classification = readProjectClassification();
  if (Object.keys(classification).length) {
    form.append('project_classes', JSON.stringify(classification));
  }

  let jobId;
  try {
    const resp = await fetch('/api/report/start', { method: 'POST', body: form });
    if (!resp.ok) {
      const txt = await resp.text().catch(() => '');
      reportError.textContent = `Failed to start report: ${resp.status}${txt ? ' — ' + txt : ''}`;
      reportError.classList.remove('hidden');
      btnGenReport.disabled = false;
      return;
    }
    const data = await resp.json();
    jobId = data.job_id;
  } catch (e) {
    reportError.textContent = String(e);
    reportError.classList.remove('hidden');
    btnGenReport.disabled = false;
    return;
  }

  const es = new EventSource(`/api/report/stream/${jobId}`);
  es.onmessage = ev => {
    try {
      const msg = JSON.parse(ev.data);
      if (msg.type === 'log') {
        reportLog.textContent += msg.line + '\n';
        reportLog.scrollTop = reportLog.scrollHeight;
      } else if (msg.type === 'done') {
        es.close();
        btnGenReport.disabled = false;
        if (msg.status === 'done') {
          const dlFiles = (msg.files || []).filter(f => !f.startsWith('charts'));
          if (dlFiles.length) {
            reportDownloadLinks.innerHTML = dlFiles.map(fname => {
              const icon = fname.endsWith('.docx')
                ? '<svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M14 2H6a2 2 0 00-2 2v16a2 2 0 002 2h12a2 2 0 002-2V8z"/><polyline points="14 2 14 8 20 8"/></svg>'
                : '<svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M14 2H6a2 2 0 00-2 2v16a2 2 0 002 2h12a2 2 0 002-2V8z"/><polyline points="14 2 14 8 20 8"/><line x1="16" y1="13" x2="8" y2="13"/><line x1="16" y1="17" x2="8" y2="17"/></svg>';
              return `<a class="btn btn-secondary" href="/api/report/download/${jobId}/${encodeURIComponent(fname)}" download="${fname}">${icon} ${fname}</a>`;
            }).join('');
            reportDownloads.classList.remove('hidden');
          }
          _currentReportJobId = jobId;
          const resultsCard = document.getElementById('reportResultsCard');
          resultsCard.classList.remove('hidden');
          renderChartsGallery(jobId, msg.charts || []);
          loadSheetList(jobId);
          const badge = document.getElementById('chartsCountBadge');
          if (badge && msg.charts && msg.charts.length) { badge.textContent = `${msg.charts.length} charts`; badge.classList.remove('hidden'); }
        } else if (msg.status === 'error') {
          reportError.textContent = 'Report generation failed. See log above for details.';
          reportError.classList.remove('hidden');
        }
      }
    } catch {}
  };
  es.onerror = () => {
    es.close();
    btnGenReport.disabled = false;
    reportError.textContent = 'Lost connection to server during report generation.';
    reportError.classList.remove('hidden');
  };
});
