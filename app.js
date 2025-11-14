// ===== State =====
const state = {
  workbook: null,
  filename: null,
  selectedSheet: null,
  rows: [],
};

// ===== Restore + persist preferences =====
document.addEventListener("DOMContentLoaded", () => {
  const savedModel = localStorage.getItem("modelSelect");
  const savedStudy = localStorage.getItem("studySynopsis");
  const savedInc = localStorage.getItem("inclusionCriteria");
  const savedExc = localStorage.getItem("exclusionCriteria");
  const savedKey = localStorage.getItem("apiKey");

  if (savedModel) document.getElementById("modelSelect").value = savedModel;
  if (savedStudy) {
    const el = document.getElementById("studySynopsis"); if (el) el.value = savedStudy;
  }
  if (savedInc) {
    const el = document.getElementById("inclusionCriteria"); if (el) el.value = savedInc;
  }
  if (savedExc) {
    const el = document.getElementById("exclusionCriteria"); if (el) el.value = savedExc;
  }
  if (savedKey) { const el=document.getElementById("apiKey"); if(el) el.value = savedKey; }

  // initialize parameter groups visibility
  const ms = document.getElementById("modelSelect");
  if (ms) {
    ms.addEventListener("change", toggleParamGroups);
    toggleParamGroups();
  }

  // Restore saved parameter values
  const paramIds = [
    "verbosity","reasoningEffort",
    "temperature",
  ];
  paramIds.forEach(id => {
    const v = localStorage.getItem(id);
    const el = document.getElementById(id);
    if (el && v != null) el.value = v;
  });
  // Attach tooltips for parameters
  addParamTooltips();

  initCriteriaLists();
});

const debounce = (fn, ms = 300) => { let t; return (...a) => { clearTimeout(t); t = setTimeout(() => fn(...a), ms); }; };
const savePref = debounce((id, value) => localStorage.setItem(id, value), 200);
[
  "modelSelect",
  "studySynopsis",
  "inclusionCriteria",
  "exclusionCriteria",
  "apiKey",
  "verbosity","reasoningEffort",
  "temperature",
].forEach((id) => {
  const el = document.getElementById(id);
  if (!el) return;
  const evt = el.tagName === "SELECT" ? "change" : "input";
  el.addEventListener(evt, () => savePref(id, el.value));
});

// ===== Upload & read =====
const fileInput = document.getElementById("fileInput");
const dropzone = document.getElementById("dropzone");
const fileMeta = document.getElementById("fileMeta");
const sheetPickerWrap = document.getElementById("sheetPickerWrap");
const sheetSelect = document.getElementById("sheetSelect");
const columnsStatus = document.getElementById("columnsStatus");
const previewWrap = document.getElementById("previewWrap");
const previewTable = document.getElementById("previewTable");
const errorBanner = document.getElementById("errorBanner");
let lastPayload = null; // stores last built payload for restart

function showError(msg) { if (!errorBanner) { alert(msg); return; } errorBanner.textContent = msg; errorBanner.classList.remove("hidden"); }
function clearError() { if (!errorBanner) return; errorBanner.textContent = ""; errorBanner.classList.add("hidden"); }

dropzone.addEventListener("dragover", (e) => { e.preventDefault(); dropzone.classList.add("hover"); });
dropzone.addEventListener("dragleave", () => dropzone.classList.remove("hover"));
dropzone.addEventListener("drop", (e) => { e.preventDefault(); dropzone.classList.remove("hover"); const f = e.dataTransfer.files?.[0]; if (f) handleFile(f); });
fileInput.addEventListener("change", (e) => { const f = e.target.files?.[0]; if (f) handleFile(f); });

async function ensureXLSX() {
  if (typeof XLSX !== "undefined") return true;
  // Try to load from CDN dynamically (fallback if head script failed)
  const trySrc = (src) => new Promise((resolve, reject) => { const s=document.createElement('script'); s.src=src; s.defer=true; s.onload=resolve; s.onerror=reject; document.head.appendChild(s); });
  try {
    await trySrc('https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js');
    if (typeof XLSX !== "undefined") return true;
  } catch {}
  try {
    await trySrc('https://unpkg.com/xlsx@0.18.5/dist/xlsx.full.min.js');
    if (typeof XLSX !== "undefined") return true;
  } catch {}
  return false;
}

async function handleFile(file) {
  state.filename = file.name;
  fileMeta.textContent = `${file.name} • ${(file.size / 1024 / 1024).toFixed(2)} MB`;
  const reader = new FileReader();
  reader.onload = (evt) => {
    try {
      const data = evt.target.result;
      const lower = file.name.toLowerCase();
      if (lower.endsWith(".csv")) {
        const text = decodeTextWithFallback(data);
        const rows = csvToArray(text);
        if (!rows.length) showError("No rows could be parsed from the CSV. Check delimiter, encoding and header row.");
        state.workbook = null; state.rows = rows; sheetPickerWrap.classList.add("hidden"); validateAndPreview();
      } else if (lower.endsWith(".xlsx") || lower.endsWith(".xls")) {
        if (typeof XLSX === "undefined") {
          // Attempt dynamic load
          ensureXLSX().then(ok => {
            if (!ok) { showError("Excel library not loaded. Ensure internet or load the page via http://localhost:8000."); return; }
            // Re-run parsing after loading
            try {
              const wb = XLSX.read(data, { type: "array" });
              if (!wb.SheetNames?.length) { showError("No sheet found in this Excel file."); return; }
              state.workbook = wb; sheetSelect.innerHTML = "";
              wb.SheetNames.forEach((s, i) => { const opt = document.createElement("option"); opt.value = s; opt.textContent = s; if (i===0) opt.selected = true; sheetSelect.appendChild(opt); });
              sheetPickerWrap.classList.remove("hidden"); state.selectedSheet = wb.SheetNames[0]; loadSheet(state.selectedSheet);
            } catch (e) { showError(`Failed to parse Excel: ${e?.message || e}`); }
          });
          return;
        }
        const wb = XLSX.read(data, { type: "array" });
        if (!wb.SheetNames?.length) { showError("No sheet found in this Excel file."); return; }
        state.workbook = wb; sheetSelect.innerHTML = "";
        wb.SheetNames.forEach((s, i) => { const opt = document.createElement("option"); opt.value = s; opt.textContent = s; if (i===0) opt.selected = true; sheetSelect.appendChild(opt); });
        sheetPickerWrap.classList.remove("hidden"); state.selectedSheet = wb.SheetNames[0]; loadSheet(state.selectedSheet);
      } else {
        showError("Unsupported file type. Upload .csv, .xlsx, or .xls.");
      }
    } catch (e) { showError(`Failed to read file: ${e?.message || e}`); }
  };
  reader.readAsArrayBuffer(file);
}

sheetSelect.addEventListener("change", () => { state.selectedSheet = sheetSelect.value; loadSheet(state.selectedSheet); });

function loadSheet(name) {
  const ws = state.workbook.Sheets[name];
  const json = XLSX.utils.sheet_to_json(ws, { defval: "" });
  state.rows = json; validateAndPreview();
}

// ===== Criteria dynamic lists =====
function initCriteriaLists(){
  const incContainer = document.getElementById('inclusionList');
  const excContainer = document.getElementById('exclusionList');
  const addInc = document.getElementById('addInc');
  const addExc = document.getElementById('addExc');
  if (!incContainer || !excContainer) return;
  // seed from saved newline-separated values for compatibility
  const savedInc = localStorage.getItem('inclusionCriteria') || '';
  const savedExc = localStorage.getItem('exclusionCriteria') || '';
  const incArr = savedInc.split(/\r?\n/).map(s=>s.trim()).filter(Boolean);
  const excArr = savedExc.split(/\r?\n/).map(s=>s.trim()).filter(Boolean);
  if (!incArr.length) incArr.push('');
  if (!excArr.length) excArr.push('');
  incArr.forEach(v=>addCriteriaInput(incContainer, v, 'inc'));
  excArr.forEach(v=>addCriteriaInput(excContainer, v, 'exc'));
  if (addInc) addInc.onclick = () => addCriteriaInput(incContainer, '', 'inc');
  if (addExc) addExc.onclick = () => addCriteriaInput(excContainer, '', 'exc');
}
function addCriteriaInput(container, value, kind){
  const row = document.createElement('div');
  row.className = 'criteria-row';
  const input = document.createElement('input');
  input.type = 'text';
  input.placeholder = kind === 'inc' ? 'e.g., population: rats or mice (preclinical)' : 'e.g., case reports, reviews, editorials';
  input.value = value || '';
  input.addEventListener('input', persistCriteria);
  input.addEventListener('keydown', (e) => {
    if (e.key === 'Enter' && !e.shiftKey) {
      e.preventDefault();
      // add a new row of the same kind
      addCriteriaInput(container, '', kind);
    }
  });
  const remove = document.createElement('button');
  remove.type = 'button';
  remove.className = 'icon-btn remove-criterion';
  remove.title = 'Remove this criterion';
  remove.setAttribute('aria-label', 'Remove');
  remove.textContent = 'x';
  remove.addEventListener('click', () => {
    const total = container.querySelectorAll('input').length;
    if (total > 1) {
      container.removeChild(row);
    } else {
      input.value = '';
    }
    persistCriteria();
  });
  row.appendChild(input);
  row.appendChild(remove);
  container.appendChild(row);
  // focus the new input for quick expansion UX
  setTimeout(() => input.focus(), 0);
}
function persistCriteria(){
  const incVals = Array.from(document.querySelectorAll('#inclusionList input')).map(el=>el.value.trim()).filter(Boolean);
  const excVals = Array.from(document.querySelectorAll('#exclusionList input')).map(el=>el.value.trim()).filter(Boolean);
  localStorage.setItem('inclusionCriteria', incVals.join('\n'));
  localStorage.setItem('exclusionCriteria', excVals.join('\n'));
}

// ===== Model parameter helpers =====
function isReasoningModel(model){ return /^gpt-5/.test(model || ""); }
function isChatModel(model){ return /^(gpt-4\.1|gpt-4o|gpt-3\.5)/.test(model || ""); }
function toggleParamGroups(){
  const model = (document.getElementById("modelSelect")?.value || "");
  const r = document.getElementById("paramsReasoning");
  const c = document.getElementById("paramsChat");
  if (!r || !c) return;
  if (isReasoningModel(model)) { r.classList.remove("hidden"); c.classList.add("hidden"); }
  else if (isChatModel(model)) { c.classList.remove("hidden"); r.classList.add("hidden"); }
  else { r.classList.add("hidden"); c.classList.add("hidden"); }
}

function validateParamsOrWarn(){
  const model = (document.getElementById("modelSelect")?.value || "");
  const warn = document.getElementById("paramsWarning");
  if (warn) { warn.classList.add("hidden"); warn.textContent = ""; }
  const num = (id) => {
    const el=document.getElementById(id); if(!el) return undefined; const v=el.value?.trim(); if(v===""||v==null) return undefined; const n=Number(v); return Number.isFinite(n)?n:NaN;
  };
  if (isReasoningModel(model)){
    // no numeric validation needed for reasoning controls
  } else if (isChatModel(model)){
    const t = num("temperature");
    if (t!==undefined && (t<0 || t>2)) { if(warn){ warn.textContent = "Temperature must be between 0 and 2."; warn.classList.remove("hidden"); } return false; }
  }
  return true;
}

function addParamTooltips(){
  const tips = {
    verbosity: 'Controls the level of detail in the response.\nlow: concise • medium: balanced (recommended) • high: very detailed.',
    reasoningEffort: 'Controls how much the model “thinks” before answering.\nminimal: fastest • low/medium: good balance (recommended) • high: most thorough but slower.',
    maxTokensReasoning: 'Upper bound on tokens generated for the answer. Use a positive integer. Typical: 256–1024.',
    temperature: 'Controls randomness: 0 = deterministic, 2 = very creative. Recommended: 0.2–0.8.',
    topP: 'Nucleus sampling: consider tokens whose cumulative probability ≤ top_p. Recommended: 0.8–1.0. Prefer tuning either temperature or top_p, not both.',
    presencePenalty: 'Encourages introducing new topics (higher = more encouragement). Recommended: 0–0.5.',
    frequencyPenalty: 'Discourages repetition of the same tokens (higher = less repetition). Recommended: 0–0.5.',
    maxTokensChat: 'Upper bound on tokens generated for the answer. Use a positive integer. Typical: 256–1024.',
  };
  Object.entries(tips).forEach(([id, text]) => {
    const el = document.getElementById(id);
    if (!el) return;
    const label = el.closest('label.field');
    if (!label) return;
    const titleSpan = label.querySelector('span');
    if (!titleSpan) return;
    if (titleSpan.querySelector('.help')) return; // avoid duplicates
    const help = document.createElement('span');
    help.className = 'help';
    help.tabIndex = 0;
    const icon = document.createElement('span'); icon.className = 'icon'; icon.textContent = '?';
    const tip = document.createElement('span'); tip.className = 'tip'; tip.textContent = text;
    help.appendChild(icon); help.appendChild(tip);
    titleSpan.appendChild(help);
  });
}

// ===== CSV helpers =====
function csvToArray(text) {
  const lines = text.split(/\r?\n/).filter((l) => l.length > 0);
  if (!lines.length) return [];
  const delim = detectDelimiter(lines[0]);
  if (lines[0].charCodeAt(0) === 0xfeff) lines[0] = lines[0].slice(1);
  const header = splitCsvLine(lines[0], delim).map((s) => s.trim());
  const out = [];
  for (let i = 1; i < lines.length; i++) {
    const cols = splitCsvLine(lines[i], delim);
    const obj = {}; header.forEach((h, idx) => obj[h] = cols[idx] ?? "");
    out.push(obj);
  }
  return out;
}
function detectDelimiter(headerLine) { const commas = (headerLine.match(/,/g) || []).length; const semis = (headerLine.match(/;/g) || []).length; return semis > commas ? ";" : ","; }
function splitCsvLine(line, delim) {
  const out = []; let cur = ""; let inQ = false;
  for (let i=0;i<line.length;i++){ const c=line[i]; if(c==='"'){ if(inQ && line[i+1]==='"'){ cur+='"'; i++; } else { inQ=!inQ; } continue; } if(c===delim && !inQ){ out.push(cur); cur=""; continue;} cur+=c; }
  out.push(cur); return out.map((s)=>s.trim());
}
function decodeTextWithFallback(buffer){ const dec=(enc)=>new TextDecoder(enc,{fatal:false}).decode(buffer); let txt=dec("utf-8"); const bad=(txt.match(/\uFFFD/g)||[]).length; if(bad>0 && bad/Math.max(1,txt.length)>0.001){ try{ txt=dec("windows-1252"); }catch{} } return txt; }

// ===== Validation + preview =====
function validateAndPreview(){
  const cols = inferColumns(state.rows);
  const hasTitle = cols.includes("title"); const hasAbstract = cols.includes("abstract");
  let statusHtml = "";
  if (hasTitle && hasAbstract) statusHtml = `<span class="badge ok">OK</span> Detected columns: <b>title</b>, <b>abstract</b>.`;
  else if (!state.rows.length) statusHtml = `No rows could be read from this sheet/file.`;
  else { const miss=[]; if(!hasTitle) miss.push("title"); if(!hasAbstract) miss.push("abstract"); statusHtml = `<span class="badge bad">Missing</span> Missing columns: <b>${miss.join(", ")}</b>.`; }
  columnsStatus.innerHTML = statusHtml;

  const head = previewTable.querySelector("thead"); const body = previewTable.querySelector("tbody"); head.innerHTML = ""; body.innerHTML = "";
  if (!state.rows.length) { previewWrap.classList.add("hidden"); return; }
  const colsToShow = cols.length ? cols : Object.keys(state.rows[0] ?? {});
  const trh=document.createElement("tr"); colsToShow.forEach(c=>{ const th=document.createElement("th"); th.textContent=c; trh.appendChild(th); }); head.appendChild(trh);
  const ell=(s,n=200)=>{ if(s==null) return ""; const str=String(s); return str.length>n?str.slice(0,n-1)+"…":str; };
  state.rows.slice(0,5).forEach(r=>{ const tr=document.createElement("tr"); colsToShow.forEach(c=>{ const td=document.createElement("td"); const full=r[c]??""; td.textContent=ell(full); if(String(full).length>200) td.title=String(full); tr.appendChild(td); }); body.appendChild(tr); });
  previewWrap.classList.remove("hidden");
  document.getElementById("btnSend").disabled = !(hasTitle && hasAbstract);
  if (state.filename) {
    const sizePart = (fileMeta.textContent.match(/•\s[\d\.]+\sMB/) || [""])[0];
    fileMeta.textContent = sizePart ? `${state.filename} ${sizePart} • ${state.rows.length.toLocaleString()} rows` : `${state.filename} • ${state.rows.length.toLocaleString()} rows`;
  }
}

function inferColumns(rows){ if(!rows.length) return []; const rawCols=Object.keys(rows[0]); const map=rawCols.reduce((a,c)=>{a[c.toLowerCase().trim()]=c; return a;},{}); const variants={ title:["title","título","titulo"], abstract:["abstract","resumo","summary"]}; const normalized=[]; const chosen={}; for(const k in variants){ for(const v of variants[k]){ if(map[v]){ chosen[k]=map[v]; break; } } } if(chosen.title||chosen.abstract){ for(const row of rows){ const r2={...row}; if(chosen.title) r2.title=row[chosen.title]; if(chosen.abstract) r2.abstract=row[chosen.abstract]; normalized.push(r2);} state.rows=normalized; return Object.keys(normalized[0]??{});} return rawCols; }

// ===== Payload =====
function buildPayload(){
  const model = (document.getElementById("modelSelect")?.value || "");
  const synopsis = (document.getElementById("studySynopsis")?.value || "");
  const incl = (document.getElementById("inclusionCriteria")?.value || "");
  const excl = (document.getElementById("exclusionCriteria")?.value || "");
  // Gather criteria from dynamic lists if present
  const incListEls = document.querySelectorAll('#inclusionList input');
  const excListEls = document.querySelectorAll('#exclusionList input');
  const inclusionArr = incListEls.length ? Array.from(incListEls).map(el=>el.value.trim()).filter(Boolean) : splitLines(incl);
  const exclusionArr = excListEls.length ? Array.from(excListEls).map(el=>el.value.trim()).filter(Boolean) : splitLines(excl);
  const api_key = (document.getElementById("apiKey").value || "").trim();
  if (!model) throw new Error("Select a model.");
  if (!state.rows.length) throw new Error("Load a spreadsheet.");
  if (!api_key) throw new Error("Enter your OpenAI API key.");
  if (!validateParamsOrWarn()) throw new Error("Please fix parameters.");
  const first = state.rows[0] ?? {};
  if (!("title" in first) || !("abstract" in first)) throw new Error("The spreadsheet must have 'title' and 'abstract' columns (or common variants).");
  // build params by family
  const params = {};
  if (isReasoningModel(model)){
    const v = document.getElementById("verbosity")?.value?.trim();
    const re = document.getElementById("reasoningEffort")?.value?.trim();
    if (v) params.verbosity = v;
    if (re) params.reasoning_effort = re;
  } else if (isChatModel(model)){
    const t = document.getElementById("temperature")?.value?.trim();
    if (t) params.temperature = Number(t);
  }
  return {
    model,
    api_key,
    study_synopsis: synopsis,
    inclusion_criteria: inclusionArr,
    exclusion_criteria: exclusionArr,
    params,
    sheet: state.selectedSheet || state.filename || "",
    filename: state.filename || "",
    sample_preview: state.rows.slice(0,5).map(r=>({ title: r.title ?? "", abstract: r.abstract ?? "" })),
    normalized_columns: true,
    records: state.rows.map((r,i)=>({ id: r.id ?? r.ID ?? r.Id ?? i+1, title: r.title ?? "", abstract: r.abstract ?? "" })),
  };
}
function splitLines(txt){ return txt.split(/\r?\n/).map(s=>s.replace(/^[-•\s]+/, "").trim()).filter(Boolean); }

// ===== Actions =====
const btnGenerate = document.getElementById("btnGenerate");
const btnSend = document.getElementById("btnSend");
const payloadCard = document.getElementById("payloadCard");
const payloadOut = document.getElementById("payloadOut");
const btnCopy = document.getElementById("btnCopy");
const btnClear = document.getElementById("btnClear");

btnGenerate.addEventListener("click", () => { try { const payload=buildPayload(); payloadOut.textContent = JSON.stringify(payload,null,2); payloadCard.classList.remove("hidden"); } catch(e){ showError(e.message); } });
btnCopy.addEventListener("click", async () => { try { await navigator.clipboard.writeText(payloadOut.textContent); btnCopy.textContent="Copied!"; setTimeout(()=>btnCopy.textContent="Copy",1200);} catch{ const ta=document.createElement("textarea"); ta.value=payloadOut.textContent; document.body.appendChild(ta); ta.select(); document.execCommand("copy"); document.body.removeChild(ta);} });

btnSend.addEventListener("click", async () => {
  try {
    clearError();
    const payload = buildPayload();
    lastPayload = payload;
    const prev = btnSend.textContent; btnSend.textContent = "Sending..."; btnSend.disabled = true;
    const controller = new AbortController(); const timeout = setTimeout(()=>controller.abort(), 30000);
    const resp = await fetch(`/api/start`, { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify(payload), signal: controller.signal }).finally(()=>clearTimeout(timeout));
    if (!resp.ok) {
      const txt = await resp.text().catch(()=>"");
      let friendly = `Send failed (${resp.status}).`;
      if (resp.status === 501) friendly += " The server does not support POST (likely a static server or file://). Run 'uvicorn backend:app --port 8000' and open http://localhost:8000.";
      else if (resp.status === 404) friendly += " Endpoint not found. Access via http://localhost:8000 and ensure /api/start exists.";
      else if (resp.status === 405) friendly += " Method not allowed. You're likely hitting a static server, not FastAPI.";
      showError(friendly + (txt ? `\nDetails: ${txt}` : ""));
      return;
    }
    const data = await resp.json(); const jobId = data.job_id; showProgress(jobId);
  } catch (e) { showError(e.message || String(e)); }
  finally { btnSend.textContent = "Send to backend"; btnSend.disabled = false; }
});

// Clear spreadsheet handler
if (btnClear){
  btnClear.addEventListener("click", () => {
    // Reset state
    state.workbook = null;
    state.filename = null;
    state.selectedSheet = null;
    state.rows = [];
    // Reset UI
    const fi = document.getElementById("fileInput"); if (fi) fi.value = "";
    fileMeta.textContent = "";
    sheetPickerWrap.classList.add("hidden");
    columnsStatus.textContent = "Waiting for file...";
    const head = previewTable.querySelector("thead"); const body = previewTable.querySelector("tbody"); if (head) head.innerHTML = ""; if (body) body.innerHTML = "";
    previewWrap.classList.add("hidden");
    payloadCard.classList.add("hidden");
    document.getElementById("btnSend").disabled = true;
    lastPayload = null;
  });
}

// ===== Progress (SSE) + cancel =====
function showProgress(jobId){
  const card = document.getElementById("progressCard"); const bar = document.getElementById("progressBar"); const label = document.getElementById("progressLabel"); const link = document.getElementById("downloadLink"); const linkX = document.getElementById("downloadLinkXlsx"); const btnCancel = document.getElementById("btnCancel");
  const btnRestart = document.getElementById("btnRestart");
  const liveLog = document.getElementById("liveLog");
  const btnToggleLog = document.getElementById("btnToggleLog");
  card.classList.remove("hidden"); link.classList.add("hidden"); if(linkX) linkX.classList.add("hidden"); bar.style.width = "0%"; label.textContent = "Starting...";
  // Reset log and keep hidden by default to avoid UI pollution
  if (liveLog) { liveLog.textContent = ""; liveLog.classList.add("hidden"); }
  if (btnToggleLog) {
    btnToggleLog.textContent = "Show live log";
    btnToggleLog.onclick = () => {
      if (!liveLog) return;
      const hidden = liveLog.classList.contains("hidden");
      if (hidden) {
        liveLog.classList.remove("hidden");
        btnToggleLog.textContent = "Hide live log";
      } else {
        liveLog.classList.add("hidden");
        btnToggleLog.textContent = "Show live log";
      }
    };
  }

  // polling state for partial results
  let pollHandle = null;
  let delivered = 0;
  function startPolling(){
    if (pollHandle) return;
    pollHandle = setInterval(async () => {
      try {
        const resp = await fetch(`/api/partial/${jobId}?since=${delivered}`);
        if (!resp.ok) return;
        const data = await resp.json();
        const items = data.items || [];
        if (items.length && liveLog){
          const lines = items.map(it => JSON.stringify({ id: it.id, decision: it.decision, rationale: it.rationale }));
          liveLog.textContent = (liveLog.textContent ? liveLog.textContent + "\n" : "") + lines.join("\n");
          if (liveLog.textContent.length > 100000) liveLog.textContent = liveLog.textContent.slice(-80000);
          liveLog.scrollTop = liveLog.scrollHeight;
        }
        delivered = data.next || delivered;
        if (data.status && (data.status === 'done' || data.status === 'error' || data.status === 'cancelled')){
          clearInterval(pollHandle); pollHandle = null;
        }
      } catch {}
    }, 1200);
  }
  function stopPolling(){ if (pollHandle) { clearInterval(pollHandle); pollHandle = null; } }
  // start polling by default in background; user can toggle visibility separately
  startPolling();
  const es = new EventSource(`/api/progress/${jobId}`);
  es.onmessage = (ev) => {
    try {
      const data = JSON.parse(ev.data);
      const processed = data.processed || 0; const total = data.total || 0; const pct = total ? Math.floor((processed/total)*100) : 0;
      bar.style.width = `${pct}%`; label.textContent = `Processed ${processed} of ${total} (${pct}%)`;
      // SSE still updates the progress; detailed rows come via polling
      if (data.status === "done") { es.close(); label.textContent = `Completed: ${processed} of ${total} (100%)`; link.href = `/api/result/${jobId}`; link.classList.remove("hidden"); if(linkX){ linkX.href = `/api/result/${jobId}?format=xlsx`; linkX.classList.remove("hidden"); } if(btnCancel) btnCancel.disabled = true; if(btnRestart) btnRestart.classList.remove("hidden"); }
      if (data.status === "cancelled") { es.close(); label.textContent = `Cancelled at ${processed} of ${total}`; link.classList.add("hidden"); if(linkX) linkX.classList.add("hidden"); if(btnCancel) btnCancel.disabled = true; if(btnRestart) btnRestart.classList.remove("hidden"); }
      if (data.status === "error") { es.close(); showError("Backend reported an error (check server logs)."); if(btnCancel) btnCancel.disabled = true; }
    } catch {}
  };
  es.onerror = () => { showError("Lost connection to backend while streaming progress. Is the server running?"); };
  if (btnCancel) {
    btnCancel.disabled = false; btnCancel.onclick = async () => { try { btnCancel.disabled = true; label.textContent = `Cancelling…`; await fetch(`/api/cancel/${jobId}`, { method: "POST" }); } catch { btnCancel.disabled = false; } };
  }
  if (btnRestart){
    btnRestart.onclick = async () => {
      try {
        clearError();
        // Build a fresh payload so the user can change model/params
        // while reusing the same spreadsheet currently loaded in state
        const newPayload = buildPayload();
        btnRestart.disabled = true;
        label.textContent = "Restarting...";
        const resp = await fetch(`/api/start`, { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify(newPayload) });
        if (!resp.ok) { const txt = await resp.text().catch(()=>""); showError(`Restart failed: ${resp.status} ${txt}`); btnRestart.disabled = false; return; }
        const data = await resp.json(); const newId = data.job_id; btnRestart.disabled = false; btnRestart.classList.add("hidden"); stopPolling(); showProgress(newId);
      } catch (e) { showError(e.message || String(e)); btnRestart.disabled = false; }
    };
  }
}

