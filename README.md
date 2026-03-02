<h1>
  <img src="scripts/logo.avif" alt="AI-Tools Logo" height="52" align="left" style="margin-right:14px">
  AI-Tools — Systematic Review AI Screening &amp; Diagnostic Evaluation
</h1>

<p>
  <strong>Author:</strong> Francisco Marques (<a href="https://github.com/francisco-marques-silva">francisco-marques-silva</a>)<br>
  <strong>Contact:</strong> <a href="mailto:francisco@pesquisaemdados.com.br">francisco@pesquisaemdados.com.br</a>
</p>

&nbsp;

A platform for **automated screening** of scientific articles via OpenAI models and **diagnostic evaluation** comparing AI decisions against human reviewers in systematic reviews.

---

## Table of Contents

1. [Project Structure](#project-structure)
2. [Quick Start](#quick-start)
3. [Web Application (AI Screening)](#part-1--web-application-ai-screening)
4. [Unified Multi-Project Report](#part-2--unified-multi-project-report)
5. [External Validation Guide](#part-3--external-validation-guide)
6. [Kappa Interpretation](#kappa-interpretation-landis--koch-1977)

---

## Project Structure

```
AI-tools/
├── scripts/                    ← Web application files
│   ├── backend.py              ← FastAPI backend (AI screening)
│   ├── index.html              ← Web application frontend
│   ├── app.js                  ← Frontend logic (JS)
│   ├── style.css               ← Frontend styles
│   └── logo.avif               ← Application logo
├── report/                     ← Unified report (multi-project)
│   └── relatorio_unificado.py  ← Generates consolidated Word report
├── input/                      ← Input files (not versioned)
│   ├── YYYYMMDD - model - Xº teste - project.xlsx  ← AI results
│   ├── Project - TIAB.xlsx       ← Human decision (TIAB)
│   ├── Project - Fulltext.xlsx   ← Articles selected for full-text reading
│   ├── Project - Listfinal.xlsx  ← Final included articles
│   └── metadata.xlsx             ← Execution metadata (cost, tokens, etc.)
├── output/                     ← Generated reports (not versioned)
│   └── relatorio_unificado_*.docx
├── .gitignore
└── README.md
```

---

## Quick Start

> This guide assumes **no prior programming knowledge**. Follow every step in order.

---

### Step 1 — Install Python

1. Go to **https://www.python.org/downloads/**
2. Click the yellow **"Download Python 3.x.x"** button (any version ≥ 3.10)
3. Run the downloaded installer
4. **Important:** On the first screen of the installer, check **"Add Python to PATH"** before clicking Install Now
5. After installation finishes, close the installer

**Verify it worked:** Open a terminal (see Step 2) and type `python --version`. You should see something like `Python 3.12.0`. If you see an error, reinstall Python and check that box.

---

### Step 2 — Open a Terminal

**Windows:** Press `Windows + R`, type `powershell`, press Enter.

---

### Step 3 — Download the Project

If you have Git installed:
```powershell
git clone https://github.com/BRISAcollab/AI-tools.git
cd AI-tools
```

Otherwise, on GitHub click **Code → Download ZIP**, extract it, then navigate to the folder in the terminal:
```powershell
cd "C:\path\to\AI-tools"
```

---

### Step 4 — Create a Virtual Environment

Run once to create an isolated Python environment for this project:

```powershell
python -m venv .venv
```

A `.venv` folder will appear — this is normal.

---

### Step 5 — Install Dependencies

```powershell
.\.venv\Scripts\pip.exe install -r requirements.txt
```

This downloads all required libraries. It may take 1–2 minutes. Wait until the prompt returns.

---

### Step 6 — Run the Web Application

```powershell
.\.venv\Scripts\python.exe -m uvicorn scripts.backend:app --reload --port 8000
```

You should see `Application startup complete.` — then open **http://localhost:8000** in your browser.

> To stop the server: press `Ctrl + C` in the terminal.

---

### Step 7 — Generate a Report (optional)

After placing your input files in `input/` (see [External Validation Guide](#part-3--external-validation-guide)):

```powershell
$env:PYTHONIOENCODING="utf-8"
.\.venv\Scripts\python.exe report\relatorio_unificado.py
```

The report is saved to `output\relatorio_unificado_YYYYMMDD_HHMMSS.docx`.

> Run `$env:PYTHONIOENCODING="utf-8"` once per terminal session before generating reports — it ensures accented characters display correctly on Windows.

---

### Every Time You Return

You do **not** need to re-run Steps 1–5 again. Just open a terminal in the project folder:

```powershell
# Web app
.\.venv\Scripts\python.exe -m uvicorn scripts.backend:app --reload --port 8000

# Report
$env:PYTHONIOENCODING="utf-8"
.\.venv\Scripts\python.exe report\relatorio_unificado.py
```

---

# Part 1 — Web Application (AI Screening)

## What It Does

The web application allows batch submission of scientific articles (title + abstract) for automated screening via OpenAI models. The backend processes each article and returns a screening decision (`include`, `exclude`, or `maybe`) with rationale.

## Frontend

Modern, responsive UI with:
- **Model selector** — GPT-5 family (reasoning models), GPT-4.1, GPT-4o, and legacy
- **API key** input with localStorage persistence
- **Parameters** — Reasoning (verbosity, effort) for GPT-5; Temperature for chat models
- **Study context** — PICO synopsis + dynamic inclusion/exclusion criteria lists
- **Advanced Backend Settings** — Collapsible panel with:
  - **Tier selector** (Free, Tier 1–5) with auto-fill presets for all concurrency/retry fields
  - Concurrent Workers, Max/Min Concurrency, Record Max Retries, AIMD Increase After, Max API Retries, Base Backoff
  - Explanatory descriptions for each field; values persist in localStorage
- **Spreadsheet upload** — Drag-and-drop for `.xlsx`, `.xls`, `.csv` with auto-detection of `title` and `abstract` columns
- **Real-time progress** — SSE streaming with live log, concurrency display, cancel/restart
- **Export** — Download results as CSV or XLSX

## Backend (`backend.py`)

### Technology

- **FastAPI** with CORS
- **OpenAI API** communication (GPT-4o, GPT-5 families)
- In-memory job management with thread pools
- Real-time progress via **Server-Sent Events (SSE)**
- **Adaptive concurrency** (AIMD algorithm) — automatically adjusts parallelism based on rate-limit responses
- Export results as **CSV** and **XLSX**

### API Endpoints

| Method | Route | Description |
|--------|-------|-------------|
| `POST` | `/api/start` | Start a screening job (returns `job_id`) |
| `POST` | `/api/cancel/{job_id}` | Cancel a running job |
| `GET` | `/api/status/{job_id}` | Job status (running/done/error/cancelled) |
| `GET` | `/api/progress/{job_id}` | SSE stream of real-time progress |
| `GET` | `/api/partial/{job_id}?since=N` | Partial results (paginated) |
| `GET` | `/api/errors/{job_id}` | Errors encountered during processing |
| `GET` | `/api/result/{job_id}?format=csv\|xlsx` | Final result as CSV or XLSX |
| `GET` | `/api/health` | Server health check |

### Environment Variables

| Variable | Default | Description |
|----------|---------|-------------|
| `CONCURRENT_WORKERS` | `20` | Starting concurrent workers (AIMD initial) |
| `CONCURRENT_MAX` | `40` | Maximum concurrent workers |
| `CONCURRENT_MIN` | `2` | Minimum concurrent workers |
| `RECORD_MAX_RETRIES` | `3` | Per-record retry attempts on error |
| `AIUP_AFTER` | `5` | Consecutive successes before adding +1 worker |
| `OPENAI_MAX_RETRIES` | `5` | Max retries per API call (5xx errors) |
| `OPENAI_BASE_BACKOFF` | `1.0` | Base exponential backoff (seconds) |

All of the above can also be overridden **per-job** via the frontend Advanced Backend Settings panel. Values sent by the frontend take precedence over environment variables.

---

# Part 2 — Unified Multi-Project Report

## What It Does

Generates a **single Word document** consolidating all analyses across **all projects and models** found in the `input/` folder. The script automatically detects all files by naming convention.

## File Naming Conventions

### AI Results

```
YYYYMMDD - model - Xº teste - project.xlsx
```

| Field | Example | Description |
|-------|---------|-------------|
| `YYYYMMDD` | `20260227` | Date/code of the spreadsheet |
| `model` | `gpt-5-mini` | Model used |
| `Xº teste` | `2º teste` | Test number (for test-retest) |
| `project` | `zebra` | Project name |

**Required columns:** `title`, `screening_decision`

### Human Reference Spreadsheets

| File | Example | Description |
|------|---------|-------------|
| TIAB | `zebra - TIAB.xlsx` | Human decision at title/abstract phase |
| Fulltext | `zebra - Fulltext.xlsx` | Articles selected for full-text reading |
| Listfinal | `zebra - Listfinal.xlsx` | Final included articles (gold standard) |

**Required columns:** `title`, `decision` (+ `abstract` when available)

### Metadata File (`metadata.xlsx`)

| Column | Type | Description |
|--------|------|-------------|
| `project` | text | Project name (must match file naming, e.g., `mino`, `zebra`, `NMDA`) |
| `code` | text | Date/code matching the AI result filename |
| `model` | text | Model name (must match AI result filename) |
| `parameter` | text | Parameter configuration description |
| `version` | text | Test number: `1º teste` or `2º teste` (used for test-retest pairing) |
| `time_ia` | timedelta | AI execution time (HH:MM:SS format) |
| `time_human` | timedelta | TIAB human screening time: 1 min/record × 2 reviewers (in researcher-hours) |
| `tokens input` | numeric | Input tokens consumed |
| `tokens_output` | numeric | Output tokens generated |
| `cost_input` | numeric | Input token cost (USD) |
| `cost_output` | numeric | Output token cost (USD) |
| `cost_total` | numeric | Total cost (USD) |

## Report Sections (12 + Appendices)

The report begins with a **Methodological Notes & Report Guide** section that explains each subsequent section and defines key methodological concepts (binarization, gold standard hierarchy, Cohen's Kappa, workload reduction, absolute efficiency, and cost-effectiveness).

### 1. Data Validation
- Inventory of all detected files (includes Listfinal article counts)
- Verification of AI files ↔ metadata correspondence
- Alerts for missing data

### 2. Metadata and Costs
- Complete execution metadata table (model, parameters, time, tokens, cost)
- Cost summary by project and by model

### 3. Diagnostic Analysis (AI vs Human TIAB)
For each **project × model × test**:
- Comparative table (sensitivity, specificity, PPV, NPV, accuracy, F1, Kappa)
- 2×2 confusion matrices (TP, FP, FN, TN)
- Visual highlight: sensitivity ≥ 95% (green), < 80% (red)

### 4. Fulltext Verification (Capture Rate)
For each **project × model × test**:
- Proportion of fulltext-selected articles that AI would have retained
- List of missed articles (details in appendix)

### 5. Listfinal Verification (Definitive Gold Standard)
For each **project × model × test**:
- Capture rate over the **final** included articles (post full-text reading)
- This is the definitive measure: proportion of truly relevant articles the AI would have retained
- Summary table with capture rate and miss rate

### 6. Test-Retest (Reproducibility)
For each **project × model**:
- Exact and binarized agreement
- Kappa with 95% CI
- Confusion matrices (1st test × 2nd test)

### 7. False Negatives
- Articles included by humans but excluded by AI, per model and test

### 8. False Positives
- Articles excluded by humans but included by AI
- FP rate over human-excluded articles

### 9. General Comparative Table
- Consolidated view: all metrics in one table (sensitivity, specificity, F1, Kappa, fulltext capture, **Listfinal capture**, test-retest Kappa, cost)

### 10. Cost-Effectiveness
- Cost (USD) vs. mean sensitivity per model
- Cost per sensitivity point

### 11. Workload Reduction Analysis
- Per-execution table comparing human time vs. AI time
- Time saved, reduction percentage, and speed factor
- Per-project summary aggregating all executions

### 12. Absolute Efficiency Analysis
- TIAB volume vs. AI positive rate vs. Listfinal capture
- Shows how much the AI reduces the screening workload while retaining relevant articles
- Efficiency score = Listfinal Capture Rate × (1 − AI Positive Rate)

### Appendices
- **Appendix A — False Positives TIAB**: title, abstract, and which models flagged as FP
- **Appendix B — Missed Fulltext Articles**: per-model matrix + detail with abstract (includes per-article × per-model miss/capture breakdown)

### Methodological Notes

Now placed at the **beginning** of the report (before Section 1) as a combined "Methodological Notes & Report Guide":
- Section-by-section guide explaining what each section contains
- Binarization: `include`/`maybe` → positive, `exclude` → negative
- Gold standard hierarchy: Listfinal > Fulltext > TIAB
- Workload reduction: based on `time_human` and `time_ia` columns
- Absolute efficiency: measures both selectivity and capture simultaneously
- Cohen's Kappa interpretation (Landis & Koch, 1977)
- Cost-effectiveness methodology

---

# Part 3 — External Validation Guide

This section explains how to **set up your own validation study** using this platform with your own systematic review data.

## Overview

The AI-Tools platform compares AI screening decisions against human reviewer decisions across three reference levels:

| Level | File | Purpose |
|-------|------|---------|
| **TIAB** | `Project - TIAB.xlsx` | Title/abstract screening decisions (primary comparison) |
| **Fulltext** | `Project - Fulltext.xlsx` | Articles selected for full-text evaluation |
| **Listfinal** | `Project - Listfinal.xlsx` | Final included articles after full-text reading (gold standard) |

## Step-by-Step Setup

### Step 1: Prepare Your TIAB Spreadsheet

This is the **most important** file — it contains the human reviewer's screening decisions for every article.

**File name:** `YourProject - TIAB.xlsx`

**Required columns:**
| Column | Description |
|--------|-------------|
| `title` | Article title — **must match exactly** with the AI result spreadsheet |
| `abstract` | Article abstract (used for false-negative/positive analysis) |
| `decision` | Human screening decision: `include`, `exclude`, or `maybe` |

**Important:** The `title` column is the **join key** between human and AI spreadsheets. Titles must be character-for-character identical. If you exported TIABs from a reference manager (e.g., EndNote, Rayyan, Covidence), use the **exact same export** as the source for AI screening.

### Step 2: Prepare Your Fulltext Spreadsheet

Contains articles that passed TIAB screening and were selected for full-text evaluation.

**File name:** `YourProject - Fulltext.xlsx`

**Required columns:** `title`, `decision`

All articles here should have `decision = include` (they were selected for full-text reading). The report checks whether the AI would have retained these articles during TIAB screening.

### Step 3: Prepare Your Listfinal Spreadsheet

Contains the **definitive** set of included articles after full-text reading. This is the gold standard.

**File name:** `YourProject - Listfinal.xlsx`

**Required columns:** `title`, `decision`

Only articles that survived full-text evaluation should appear here. The report measures how many of these the AI would have retained — this is the most clinically relevant metric.

### Step 4: Run AI Screening

1. Start the web app: `.venv\Scripts\python.exe -m uvicorn scripts.backend:app --reload --port 8000` (from repo root)
2. Open http://localhost:8000
3. Select a model (e.g., GPT-5.2)
4. Enter your API key
5. Write your PICO synopsis and inclusion/exclusion criteria
6. Upload your TIAB spreadsheet (the one with `title` + `abstract` columns)
7. Click **Send to Backend**
8. Wait for completion, then download the result

**Naming the output:** Rename the downloaded file to match the convention:
```
YYYYMMDD - model - 1º teste - YourProject.xlsx
```

For test-retest analysis, run the same screening again and name it `2º teste`.

### Step 5: Create Metadata Spreadsheet

Create `metadata.xlsx` in the `input/` folder with one row per AI execution:

| project | code | model | parameter | version | time_ia | time_human | tokens input | tokens_output | cost_input | cost_output | cost_total |
|---------|------|-------|-----------|---------|---------|------------|--------------|---------------|------------|-------------|------------|
| mino | 20260227 | gpt-5-mini | reasoning=medium | 1º teste | 0:04:32 | 2:00:00 | 152340 | 28450 | 0.023 | 0.114 | 0.137 |

**Column details:**

- **`project`**: Must match the project name in your file names (case-sensitive).
- **`code`**: The YYYYMMDD prefix of your AI result file.
- **`model`**: Must match the model name in your AI result file.
- **`parameter`**: Free-text description of AI parameters used.
- **`version`**: Test number — `1º teste` or `2º teste`. This identifies replication runs for the test-retest analysis.
- **`time_ia`**: How long the AI took (format: `H:MM:SS` or `HH:MM:SS`).
- **`time_human`**: Human time for TIAB screening, calculated as **1 minute per record × 2 reviewers**, expressed in researcher-hours (format: `H:MM:SS`). Used in the **Workload Reduction** analysis (Section 11).
- **`tokens input` / `tokens_output`**: Token counts from the execution, available in the **OpenAI dashboard** (Usage section).
- **`cost_input` / `cost_output` / `cost_total`**: Costs in USD (available from OpenAI usage dashboard).

### Step 6: Place All Files in `input/`

Your `input/` folder should look like:

```
input/
├── 20260227 - gpt-5-mini - 1º teste - mino.xlsx    ← AI result
├── 20260227 - gpt-5-mini - 2º teste - mino.xlsx    ← AI result (retest)
├── 20260228 - gpt-4o - 1º teste - mino.xlsx        ← AI result (different model)
├── mino - TIAB.xlsx                                  ← Human TIAB decisions
├── mino - Fulltext.xlsx                              ← Fulltext selections
├── mino - Listfinal.xlsx                             ← Final included articles
└── metadata.xlsx                                     ← Execution metadata
```

You can have **multiple projects** in the same folder. Just use different project names:

```
input/
├── ... mino files ...
├── ... zebra files ...
├── ... NMDA files ...
└── metadata.xlsx        ← Contains rows for ALL projects
```

### Step 7: Generate the Report

```bash
# Windows
$env:PYTHONIOENCODING="utf-8"
python report/relatorio_unificado.py

# Linux / macOS
PYTHONIOENCODING=utf-8 python report/relatorio_unificado.py
```

The report is saved to `output/relatorio_unificado_YYYYMMDD_HHMMSS.docx`.

## Title Consistency Checklist

Title matching is critical for accurate results. Common pitfalls:

| Problem | Solution |
|---------|----------|
| Different capitalization | Matching is case-insensitive (handled automatically) |
| Extra whitespace | Whitespace is normalized (handled automatically) |
| Different special characters (e.g., `–` vs. `-`) | Export both human and AI data from the same source file |
| Truncated titles | Ensure full titles in both spreadsheets |
| HTML entities (e.g., `&amp;`) | Clean titles before use, or ensure consistent encoding |

**Best practice:** Use the **exact same TIAB spreadsheet** as the source for both human review and AI screening. This guarantees title consistency.

## Reusing TIAB Data

The TIAB spreadsheet serves dual purposes:

1. **Source for AI screening** — Upload it to the web app (the app uses `title` + `abstract` columns)
2. **Human reference** — The `decision` column contains the human reviewer's judgment

This means you only need **one** TIAB spreadsheet per project. The AI result file will have a `screening_decision` column added by the AI, while the original TIAB keeps the human `decision`.

## Interpreting the Report

### Key Metrics

| Metric | What It Tells You | Good Value |
|--------|-------------------|------------|
| **Sensitivity** | % of human-included articles the AI also included | ≥ 95% |
| **Specificity** | % of human-excluded articles the AI also excluded | ≥ 50% |
| **Listfinal Capture** | % of final included articles the AI retained | ≥ 95% |
| **Test-Retest Kappa** | Reproducibility of AI decisions across runs | ≥ 0.80 |
| **Workload Reduction** | Time saved using AI vs. human screening | Higher is better |
| **Efficiency Score** | Combined selectivity × capture (Section 12) | Higher is better |

### Decision Flow

```
TIAB Screening → Sensitivity/Specificity (Section 3)
       ↓
Fulltext Capture → Section 4
       ↓
Listfinal Capture → Section 5 (most important!)
       ↓
Cost-Effectiveness → Section 10
Workload Reduction → Section 11
Absolute Efficiency → Section 12
```

---

## Kappa Interpretation (Landis & Koch, 1977)

| Kappa | Agreement |
|-------|-----------|
| < 0 | Poor |
| 0.00–0.20 | Slight |
| 0.21–0.40 | Fair |
| 0.41–0.60 | Moderate |
| 0.61–0.80 | Substantial |
| 0.81–1.00 | Almost Perfect |

---

## Troubleshooting

| Issue | Solution |
|-------|----------|
| `python` is not recognized | Python not added to PATH — reinstall Python and check "Add Python to PATH" on the first installer screen |
| `.venv\Scripts\pip.exe` not found | Virtual environment not created — run `python -m venv .venv` from the project root folder |
| `UnicodeEncodeError` on Windows | Run `$env:PYTHONIOENCODING="utf-8"` in the same terminal before the report command |
| `No module named 'docx'` | Run `.\.venv\Scripts\pip.exe install python-docx` |
| `No module named 'fastapi'` | Run `.\.venv\Scripts\pip.exe install -r requirements.txt` |
| `Could not import module "backend"` | You must run the uvicorn command from the **project root folder** (`AI-tools/`), not from inside `scripts/` |
| Report finds no files | Check that input files follow the exact naming conventions described in [File Naming Conventions](#file-naming-conventions) |
| Metadata mismatch warnings | Ensure `project`, `code`, `model` in `metadata.xlsx` match the AI result filenames exactly (case-sensitive) |
| XLSX library not loaded (web app) | Access the app via http://localhost:8000, not by opening `index.html` directly as a file |

---

## License

