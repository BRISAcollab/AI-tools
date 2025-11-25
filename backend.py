import os
import uuid
import json
import threading
import time
import io
import random
from typing import Dict, Any, List, Optional

from fastapi import FastAPI, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import StreamingResponse, JSONResponse, Response
from fastapi.staticfiles import StaticFiles
from dotenv import load_dotenv
from pydantic import BaseModel, Field
import requests


load_dotenv()

app = FastAPI(title="Triage Backend", version="0.1.0")
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)


class StartPayload(BaseModel):
    model: str = Field(default="gpt-5")
    api_key: str | None = None
    study_synopsis: str
    inclusion_criteria: List[str] = []
    exclusion_criteria: List[str] = []
    filename: str = ""
    sheet: str = ""
    records: List[Dict[str, Any]]
    temperature: float | None = None
    params: Optional[Dict[str, Any]] = None


JOBS: Dict[str, Dict[str, Any]] = {}

# Rate limiting knobs (seconds)
RATE_LIMIT_MIN_INTERVAL = float(os.getenv("RATE_LIMIT_MIN_INTERVAL", "0.6"))  # min spacing between calls
MAX_RETRIES = int(os.getenv("OPENAI_MAX_RETRIES", "5"))
BASE_BACKOFF = float(os.getenv("OPENAI_BASE_BACKOFF", "1.0"))


def build_prompt(synopsis: str, inc: List[str], exc: List[str], title: str, abstract: str) -> str:
    inc_lines = "\n".join(f"- {i}" for i in inc) if inc else "- (none provided)"
    exc_lines = "\n".join(f"- {e}" for e in exc) if exc else "- (none provided)"
    parts: List[str] = [
                "You are a knowledgeable AI assistant tasked with high-sensitivity title and abstract screening of a research article for a systematic review. Follow a step-by-step evaluation focusing on not missing any potentially relevant study.",
        "",
        f"Synopsis/PICO: {synopsis.strip()}",
        "",
        "Inclusion Criteria:",
        inc_lines,
        "",
        "Exclusion Criteria:",
        exc_lines,
        "",
        f"Study Title: {title or ''}",
        f"Study Abstract: {abstract or ''}",
        "",
        "Instructions:",
        "1. Identify PICO elements and type record from the title and abstract: determine the studied population (animals/population), intervention/exposure end type of record (review, systematic review, original research article, case report).",
        "2. Check each inclusion criterion against the information: for each inclusion criterion, assess whether the abstract suggests the study fulfills it. (Treat unspecified details as uncertain rather than negative.)",
        "3. Check each exclusion criterion: assess if any exclusion criterion is clearly met by the study.",
        "4. Perform the above reasoning internally – do not output these steps.",
        "",
        "Decision logic (high recall focus):",
        "- If all inclusion criteria are met met and no exclusion criteria apply, Include the study.",
        "- If any inclusion criterion is clearly unmet or any exclusion criterion is definitely met, decide Exclude",
        "- If there is any uncertainty (e.g. some PICO elements are unclear from the abstract) and no clear exclusion, mark as Maybe rather than risk wrongful exclusion.",
        "",
        "When in doubt, err on the side of inclusion (include or maybe).",
        "",
        "Output (JSON only): Return a single JSON object with keys:",
        "- decision: \"include\" | \"exclude\" | \"maybe\"",
        "- rationale: brief reason (<=12 words)",
        "- inclusion_evaluation: array of { \"criterion\": string, \"status\": \"met\"|\"unclear\"|\"unmet\" }",
        "- exclusion_evaluation: array of { \"criterion\": string, \"status\": \"met\"|\"unclear\"|\"unmet\" }",
        "No other text should be produced outside the JSON.",
        "",
        "Example format:",
        "{",
        "  \"decision\": \"maybe\",",
        "  \"rationale\": \"Population matches, but intervention details unclear from abstract\",",
        "  \"inclusion_evaluation\": [ { \"criterion\": \"population: adults with T2D\", \"status\": \"met\" } ],",
        "  \"exclusion_evaluation\": [ { \"criterion\": \"non-human study\", \"status\": \"unmet\" } ]",
        "}",
        "",
        "Now, based on the above criteria and the article's title/abstract, output the JSON decision.",
    ]
    return "\n".join(parts)


def call_openai_chat(model: str, prompt: str, api_key: str, params: Optional[Dict[str, Any]] = None) -> Dict[str, Any]:
    is_reasoning = model.startswith("gpt-5")
    headers = {"Authorization": f"Bearer {api_key}", "Content-Type": "application/json"}
    # Build request target/body depending on model family
    if is_reasoning:
        url = "https://api.openai.com/v1/responses"
        base_body: Dict[str, Any] = {"model": model, "input": prompt}
        if params:
            if params.get("reasoning_effort"):
                base_body["reasoning"] = {"effort": params["reasoning_effort"]}
            if params.get("verbosity"):
                base_body["text"] = {"verbosity": params["verbosity"]}
    else:
        url = "https://api.openai.com/v1/chat/completions"
        base_body: Dict[str, Any] = {
            "model": model,
            "messages": [
                {"role": "system", "content": "You return only strict JSON and nothing else."},
                {"role": "user", "content": prompt},
            ],
        }
        if params and "temperature" in params:
            base_body["temperature"] = float(params["temperature"])

    # Perform request with retries and exponential backoff (handles 429/5xx)
    r = None
    for attempt in range(1, MAX_RETRIES + 1):
        r = requests.post(url, headers=headers, json=base_body, timeout=60)
        if r.status_code == 200:
            break
        # Respect Retry-After when present
        if r.status_code in (429, 500, 502, 503, 504):
            ra = r.headers.get("Retry-After")
            if ra:
                try:
                    sleep_s = float(ra)
                except Exception:
                    sleep_s = BASE_BACKOFF * (2 ** (attempt - 1))
            else:
                jitter = random.uniform(0, 0.25)
                sleep_s = BASE_BACKOFF * (2 ** (attempt - 1)) + jitter
            time.sleep(min(sleep_s, 20.0))
            continue
        # Other errors: do not retry further
        break
    if r is None or r.status_code != 200:
        raise RuntimeError(f"OpenAI error {r.status_code if r else 'no_response'}: {r.text[:200] if r else ''}")
    data = r.json()
    # Extract assistant text depending on API used
    content: Optional[str] = None
    def deep_extract_text(obj) -> Optional[str]:
        try:
            from collections import deque
            q = deque([obj])
            while q:
                node = q.popleft()
                if isinstance(node, dict):
                    # direct output_text
                    ot = node.get("output_text")
                    if isinstance(ot, str) and ot.strip():
                        return ot.strip()
                    # Responses content blocks
                    txt = node.get("text")
                    if isinstance(txt, dict):
                        val = txt.get("value")
                        if isinstance(val, str) and val.strip():
                            return val.strip()
                    if isinstance(txt, str) and txt.strip():
                        return txt.strip()
                    cont = node.get("content")
                    if isinstance(cont, str) and cont.strip():
                        return cont.strip()
                    for v in node.values():
                        q.append(v)
                elif isinstance(node, list):
                    q.extend(node)
        except Exception:
            return None
        return None
    if is_reasoning:
        # Responses API. Prefer 'output_text'; otherwise dig into 'output' items.
        resp_obj = data.get("response") or data
        ot = resp_obj.get("output_text") or data.get("output_text")
        if isinstance(ot, str) and ot.strip():
            content = ot.strip()
        if not content:
            try:
                for item in (resp_obj.get("output") or data.get("output") or []):
                    for part in (item.get("content") or []):
                        if isinstance(part, dict):
                            # Newer schema: part = { type: 'output_text', text: { value: '...' } }
                            txt = None
                            if isinstance(part.get("text"), dict):
                                txt = part.get("text", {}).get("value")
                            elif isinstance(part.get("text"), str):
                                txt = part.get("text")
                            if isinstance(txt, str) and txt.strip():
                                content = txt.strip()
                                break
                    if content:
                        break
            except Exception:
                content = None
    else:
        # Chat Completions
        try:
            content = data["choices"][0]["message"]["content"].strip()
        except Exception:
            # Fallback for rare structures (e.g., text field)
            try:
                content = data["choices"][0].get("text", "").strip()
            except Exception:
                content = None
    if not content:
        # last-resort: deep search
        content = deep_extract_text(data)
    if not content:
        raise RuntimeError(f"Invalid OpenAI response structure (keys={list(data.keys())[:10]})")
    # Strip code fences if present
    if content.startswith("```"):
        content = content.strip("`")
        if "\n" in content:
            content = content.split("\n", 1)[1].strip()
    try:
        parsed = json.loads(content)
    except Exception:
        lower = content.lower()
        decision = "maybe"
        if "include" in lower and "exclude" not in lower:
            decision = "include"
        elif "exclude" in lower and "include" not in lower:
            decision = "exclude"
        rationale = content.split("\n")[0][:80]
        return {"decision": decision, "rationale": rationale}
    decision = str(parsed.get("decision", "")).strip().lower()
    rationale = str(parsed.get("rationale", "")).strip()
    # Optional per-criterion evaluations
    def coerce_eval(arr: Any) -> List[Dict[str, str]]:
        out: List[Dict[str, str]] = []
        if isinstance(arr, str):
            try:
                arr = json.loads(arr)
            except Exception:
                arr = []
        if isinstance(arr, list):
            for it in arr:
                if isinstance(it, dict):
                    crit = str(it.get("criterion", "")).strip()
                    status = str(it.get("status", "")).strip().lower()
                    if status not in {"met", "unclear", "unmet"}:
                        status = "unclear"
                    if crit:
                        out.append({"criterion": crit, "status": status})
                elif isinstance(it, (list, tuple)) and len(it) >= 2:
                    crit = str(it[0]).strip()
                    status = str(it[1]).strip().lower()
                    if status not in {"met", "unclear", "unmet"}:
                        status = "unclear"
                    if crit:
                        out.append({"criterion": crit, "status": status})
        return out
    inc_eval = coerce_eval(
        parsed.get("inclusion_evaluation")
        or parsed.get("inclusionEvaluations")
        or parsed.get("inclusion")
        or []
    )
    exc_eval = coerce_eval(
        parsed.get("exclusion_evaluation")
        or parsed.get("exclusionEvaluations")
        or parsed.get("exclusion")
        or []
    )
    if decision not in {"include", "exclude", "maybe"}:
        decision = "maybe"
    if len(rationale.split()) > 12:
        rationale = " ".join(rationale.split()[:12])
    if not rationale:
        rationale = "insufficient information"
    return {
        "decision": decision,
        "rationale": rationale,
        "inclusion_evaluation": inc_eval,
        "exclusion_evaluation": exc_eval,
    }


def worker(job_id: str):
    job = JOBS[job_id]
    api_key = job.get("api_key") or os.getenv("OPENAI_API_KEY")
    if not api_key:
        job["status"] = "error"
        job["error"] = "API key missing (provide api_key in request or set OPENAI_API_KEY env var)"
        return
    records = job["records"]
    total = len(records)
    job["total"] = total
    job["processed"] = 0
    results = []
    model = job["model"]
    synopsis = job["study_synopsis"]
    inc = job["inclusion_criteria"]
    exc = job["exclusion_criteria"]
    params = job.get("params")

    last_call_ts = 0.0
    for idx, rec in enumerate(records, start=1):
        if job.get("status") == "cancelled":
            break
        title = rec.get("title") or ""
        abstract = rec.get("abstract") or ""
        rid = rec.get("id", idx)
        try:
            # Pace requests to avoid rate limiting
            now = time.time()
            delta = now - last_call_ts
            if delta < RATE_LIMIT_MIN_INTERVAL:
                time.sleep(RATE_LIMIT_MIN_INTERVAL - delta)
            # Build prompt (single canonical version)
            prompt = build_prompt(synopsis, inc, exc, title, abstract)
            out = call_openai_chat(model, prompt, api_key, params=params)
            # Ensure per-criterion arrays exist even if the model omitted them
            inc_eval = out.get("inclusion_evaluation") or []
            exc_eval = out.get("exclusion_evaluation") or []
            if not inc_eval and inc:
                inc_eval = [{"criterion": c, "status": "unclear"} for c in inc]
            if not exc_eval and exc:
                exc_eval = [{"criterion": c, "status": "unclear"} for c in exc]
            out["inclusion_evaluation"] = inc_eval
            out["exclusion_evaluation"] = exc_eval
            last_call_ts = time.time()
        except Exception as e:
            out = {"decision": "maybe", "rationale": f"error: {str(e)[:70]}"}
        results.append({
            "id": rid,
            "title": title,
            "abstract": abstract,
            "screening_decision": out["decision"],
            "screening_reason": out["rationale"],
            "inclusion_evaluation": out.get("inclusion_evaluation", []),
            "exclusion_evaluation": out.get("exclusion_evaluation", []),
        })
        job["processed"] = idx
        job["results"] = results
        # small additional jitter
        time.sleep(random.uniform(0.0, 0.05))

    job["status"] = "done"


@app.post("/api/start")
def start_job(payload: StartPayload):
    if not payload.records:
        raise HTTPException(status_code=400, detail="empty records")
    job_id = str(uuid.uuid4())
    JOBS[job_id] = {
        "status": "running",
        "model": payload.model or "gpt-5",
        "api_key": payload.api_key,
        "study_synopsis": payload.study_synopsis or "",
        "inclusion_criteria": payload.inclusion_criteria or [],
        "exclusion_criteria": payload.exclusion_criteria or [],
        "filename": payload.filename,
        "sheet": payload.sheet,
        "records": payload.records,
        "params": payload.params or {},
        "processed": 0,
        "total": len(payload.records),
        "results": [],
    }
    th = threading.Thread(target=worker, args=(job_id,), daemon=True)
    th.start()
    return {"job_id": job_id}


@app.post("/api/cancel/{job_id}")
def cancel(job_id: str):
    job = JOBS.get(job_id)
    if not job:
        raise HTTPException(status_code=404, detail="job not found")
    job["status"] = "cancelled"
    return {"status": "cancelled"}

@app.get("/api/status/{job_id}")
def status(job_id: str):
    job = JOBS.get(job_id)
    if not job:
        raise HTTPException(status_code=404, detail="job nÃ£o encontrado")
    return {
        "status": job.get("status"),
        "processed": job.get("processed", 0),
        "total": job.get("total", 0),
        "filename": job.get("filename"),
    }


@app.get("/api/progress/{job_id}")
def progress(job_id: str):
    job = JOBS.get(job_id)
    if not job:
        raise HTTPException(status_code=404, detail="job nÃ£o encontrado")

    def event_stream():
        last = -1
        while True:
            j = JOBS.get(job_id)
            if not j:
                yield f"data: {json.dumps({'status': 'error', 'detail': 'job missing'})}\n\n"
                break
            processed = j.get("processed", 0)
            total = j.get("total", 0)
            status = j.get("status")
            if processed != last or status != "running":
                payload = {"status": status, "processed": processed, "total": total}
                # attach last result summary if present
                try:
                    results = j.get("results") or []
                    if results:
                        lr = results[-1]
                        payload["last"] = {
                            "id": lr.get("id"),
                            "decision": lr.get("screening_decision"),
                            "rationale": lr.get("screening_reason"),
                        }
                except Exception:
                    pass
                if status == "error" and j.get("error"):
                    payload["error"] = str(j.get("error"))
                yield f"data: {json.dumps(payload)}\n\n"
                last = processed
            if status in {"done", "error", "cancelled"}:
                break
            time.sleep(0.5)

    return StreamingResponse(event_stream(), media_type="text/event-stream")


@app.get("/api/partial/{job_id}")
def partial(job_id: str, since: int = 0, limit: int = 200):
    job = JOBS.get(job_id)
    if not job:
        raise HTTPException(status_code=404, detail="job not found")
    results = job.get("results") or []
    total = job.get("total", 0)
    processed = job.get("processed", 0)
    status = job.get("status")
    start = max(0, int(since))
    end = min(len(results), start + max(1, int(limit)))
    items = []
    for idx in range(start, end):
        r = results[idx]
        items.append({
            "index": idx + 1,
            "id": r.get("id"),
            "decision": r.get("screening_decision"),
            "rationale": r.get("screening_reason"),
        })
    payload = {
        "status": status,
        "processed": processed,
        "total": total,
        "items": items,
        "next": end,
    }
    if status == "error" and job.get("error"):
        payload["error"] = str(job.get("error"))
    return JSONResponse(payload)

@app.get("/api/result/{job_id}")
def result(job_id: str, format: str = "csv"):
    job = JOBS.get(job_id)
    if not job:
        raise HTTPException(status_code=404, detail="job nÃ£o encontrado")
    if job.get("status") != "done":
        raise HTTPException(status_code=400, detail="job nÃ£o finalizado")
    rows = job.get("results", [])
    if not rows:
        return JSONResponse({"rows": []})
    fieldnames = [
        "id",
        "title",
        "abstract",
        "screening_decision",
        "screening_reason",
        "inclusion_evaluation",
        "exclusion_evaluation",
    ]

    if (format or "").lower() == "xlsx":
        try:
            from openpyxl import Workbook
        except Exception:
            raise HTTPException(status_code=500, detail="openpyxl not installed on server. Install with 'pip install openpyxl'.")
        wb = Workbook()
        ws = wb.active
        ws.title = "triage"
        ws.append(fieldnames)
        for r in rows:
            def _cell(val):
                if isinstance(val, (list, dict)):
                    try:
                        return json.dumps(val, ensure_ascii=False)
                    except Exception:
                        return str(val)
                return val
            ws.append([_cell(r.get(k, "")) for k in fieldnames])
        bio = io.BytesIO()
        wb.save(bio)
        data = bio.getvalue()
        return Response(
            content=data,
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            headers={
                "Content-Disposition": f"attachment; filename=triage_{job_id}.xlsx"
            },
        )
    else:
        # CSV in memory
        import csv
        output = io.StringIO()
        writer = csv.DictWriter(output, fieldnames=fieldnames)
        writer.writeheader()
        for r in rows:
            row_out = {}
            for k in fieldnames:
                v = r.get(k, "")
                if isinstance(v, (list, dict)):
                    try:
                        v = json.dumps(v, ensure_ascii=False)
                    except Exception:
                        v = str(v)
                row_out[k] = v
            writer.writerow(row_out)
        data = output.getvalue().encode("utf-8")
        return Response(
            content=data,
            media_type="text/csv; charset=utf-8",
            headers={
                "Content-Disposition": f"attachment; filename=triage_{job_id}.csv"
            },
        )


# Serve static frontend from current directory at root
app.mount("/", StaticFiles(directory=".", html=True), name="static")

# Run with: uvicorn backend:app --reload --port 8000


@app.get("/api/health")
def health():
    return {"status": "ok"}
