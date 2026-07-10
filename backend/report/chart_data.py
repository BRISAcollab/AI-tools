"""
chart_data.py — Collect analysis results and export data_grafics_*.xlsx.

This module builds one DataFrame per chart from the dict returned by
run_all_analyses() and writes them as sheets in a single XLSX file
that the graphic scripts (graphic.py / graphic.R) consume.
"""

import datetime
from pathlib import Path

import numpy as np
import pandas as pd

from .utils import normalise_model_name, compute_f1_lf, compute_metrics_vs_lf
from .constants import project_category


# =====================================================================
#  INDIVIDUAL SHEET BUILDERS
# =====================================================================

def _build_sensitivity_per_model(projects, all_results):
    """Sheet: sensitivity_per_model  —  Project, Model, Test, Sensitivity."""
    diag = all_results.get("diagnostic", {})
    rows = []
    for pn in sorted(projects):
        proj = projects[pn]
        for mn in sorted(proj["models"]):
            model_name = proj["models"][mn]["name"]
            for tn, r in sorted(diag.get(pn, {}).get(mn, {}).items()):
                if r is None:
                    continue
                rows.append({
                    "Project": proj["name"],
                    "Model": model_name,
                    "Test": tn,
                    "Sensitivity": r["metrics"]["Sensitivity"],
                })
    return pd.DataFrame(rows) if rows else None


def _build_lf_capture_heatmap(projects, all_results):
    """Sheet: lf_capture_heatmap  —  Project, Model, Capture_Rate_pct (avg across tests)."""
    lf = all_results.get("listfinal", {})
    rows = []
    for pn in sorted(projects):
        proj = projects[pn]
        for mn in sorted(proj["models"]):
            model_name = proj["models"][mn]["name"]
            rates = [
                lf[pn][mn][tn]["capture_rate"]
                for tn in lf.get(pn, {}).get(mn, {})
                if lf[pn][mn][tn].get("capture_rate") is not None
                and not np.isnan(lf[pn][mn][tn]["capture_rate"])
            ]
            if rates:
                rows.append({
                    "Project": proj["name"],
                    "Model": model_name,
                    "Capture_Rate_pct": np.mean(rates) * 100,
                })
    return pd.DataFrame(rows) if rows else None


def _build_test_retest_kappa(projects, all_results):
    """Sheet: test_retest_kappa  —  Label, Kappa, CI_lo, CI_hi."""
    tr = all_results.get("test_retest", {})
    rows = []
    for pn in sorted(tr):
        proj = projects[pn]
        for mn in sorted(tr[pn]):
            r = tr[pn][mn]
            model_name = proj["models"][mn]["name"]
            rows.append({
                "Label": f"{model_name}\n{proj['name']}",
                "Project": proj["name"],
                "Model": model_name,
                "Kappa": r["kappa"],
                "CI_lo": r["kappa_ci_lo"],
                "CI_hi": r["kappa_ci_hi"],
            })
    return pd.DataFrame(rows) if rows else None


def _build_model_comparison_radar(projects, all_results):
    """Sheet: model_comparison_radar  —  Model + metric columns (0-1 scale).

    Sensitivity is based on Listfinal capture rate (not TIAB diagnostic),
    making this chart reflect final-inclusion quality rather than TIAB agreement.
    """
    lf = all_results.get("listfinal", {})
    tr = all_results.get("test_retest", {})
    ft = all_results.get("fulltext", {})
    diag = all_results.get("diagnostic", {})

    model_agg = {}
    for pn in sorted(projects):
        proj = projects[pn]
        for mn in sorted(proj["models"]):
            mname = proj["models"][mn]["name"]
            if mname not in model_agg:
                model_agg[mname] = {
                    "Sens_Listfinal": [], "Specificity": [], "F1": [],
                    "FT_Capture": [], "Test_Retest_Kappa": [],
                }

            # Sensitivity from Listfinal capture rate (primary quality measure)
            for tn in lf.get(pn, {}).get(mn, {}):
                v = lf[pn][mn][tn]["capture_rate"]
                if v is not None and not np.isnan(v):
                    model_agg[mname]["Sens_Listfinal"].append(v)

            # Specificity from diagnostic; F1 against Listfinal gold standard
            for tn, r in diag.get(pn, {}).get(mn, {}).items():
                if r is None:
                    continue
                v = r["metrics"]["Specificity"]
                if not np.isnan(v):
                    model_agg[mname]["Specificity"].append(v)
                lr = lf.get(pn, {}).get(mn, {}).get(tn)
                f1_lf = compute_f1_lf(r, lr)
                if not np.isnan(f1_lf):
                    model_agg[mname]["F1"].append(f1_lf)

            # Fulltext capture rate
            for tn in ft.get(pn, {}).get(mn, {}):
                v = ft[pn][mn][tn]["capture_rate"]
                if v is not None and not np.isnan(v):
                    model_agg[mname]["FT_Capture"].append(v)

            # Test-retest kappa
            tr_r = tr.get(pn, {}).get(mn)
            if tr_r and not np.isnan(tr_r["kappa"]):
                model_agg[mname]["Test_Retest_Kappa"].append(tr_r["kappa"])

    rows = []
    for mname, agg in model_agg.items():
        row = {"Model": mname}
        for metric, vals in agg.items():
            row[metric] = np.mean(vals) if vals else float("nan")
        rows.append(row)
    return pd.DataFrame(rows) if rows else None


def _build_cost_vs_sensitivity(projects, all_results, metadados):
    """Sheet: cost_vs_sensitivity  —  Model, Avg_Sensitivity_pct, Avg_Cost_USD, Avg_F1_LF.

    F1 is computed against the Listfinal gold standard (not TIAB diagnostic).
    """
    if metadados is None:
        return None
    if "code" not in metadados.columns:
        return None
    diag = all_results.get("diagnostic", {})
    lf = all_results.get("listfinal", {})
    model_data = {}  # model_name -> {sens: [], cost: [], f1: []}
    for pn in sorted(projects):
        proj = projects[pn]
        for mn in sorted(proj["models"]):
            mname = proj["models"][mn]["name"]
            if mname not in model_data:
                model_data[mname] = {"sens": [], "cost": [], "f1": []}
            for tn, r in diag.get(pn, {}).get(mn, {}).items():
                if r is None:
                    continue
                s = r["metrics"]["Sensitivity"]
                if not np.isnan(s):
                    model_data[mname]["sens"].append(s)
                lr = lf.get(pn, {}).get(mn, {}).get(tn)
                f1_lf = compute_f1_lf(r, lr)
                if not np.isnan(f1_lf):
                    model_data[mname]["f1"].append(f1_lf)
                code = proj["models"][mn]["tests"][tn]["code"]
                meta_m = metadados[metadados["code"].astype(str) == str(code)]
                if not meta_m.empty and pd.notna(meta_m.iloc[0].get("cost_total")):
                    model_data[mname]["cost"].append(meta_m.iloc[0]["cost_total"])

    rows = []
    for mname, d in model_data.items():
        if d["sens"] and d["cost"]:
            rows.append({
                "Model": mname,
                "Avg_Sensitivity_pct": np.mean(d["sens"]) * 100,
                "Avg_Cost_USD": np.mean(d["cost"]),
                "Avg_F1_LF": np.mean(d["f1"]) if d["f1"] else float("nan"),
            })
    return pd.DataFrame(rows) if rows else None


def _build_workload_reduction(projects, all_results, metadados):
    """Sheet: workload_reduction  —  Model, Human_Hours, AI_Hours, Speed_Factor."""
    if metadados is None:
        return None
    if "time_human" not in metadados.columns or "time_ia" not in metadados.columns:
        return None

    def _parse_td(val):
        if pd.isna(val):
            return float("nan")
        if isinstance(val, pd.Timedelta):
            return val.total_seconds() / 3600.0
        try:
            return pd.to_timedelta(str(val).strip()).total_seconds() / 3600.0
        except Exception:
            return float("nan")

    meta = metadados.copy()
    meta["_h_human"] = meta["time_human"].apply(_parse_td)
    meta["_h_ia"] = meta["time_ia"].apply(_parse_td)

    model_data = {}
    for _, row in meta.iterrows():
        mname = str(row.get("model", ""))
        if mname not in model_data:
            model_data[mname] = {"human": [], "ai": []}
        h_h = row["_h_human"]
        h_a = row["_h_ia"]
        if not np.isnan(h_h):
            model_data[mname]["human"].append(h_h)
        if not np.isnan(h_a):
            model_data[mname]["ai"].append(h_a)

    rows = []
    for mname, d in model_data.items():
        if d["human"] and d["ai"]:
            avg_h = np.mean(d["human"])
            avg_a = np.mean(d["ai"])
            factor = avg_h / avg_a if avg_a > 0 else float("nan")
            rows.append({
                "Model": mname,
                "Human_Hours": avg_h,
                "AI_Hours": avg_a,
                "Speed_Factor": factor,
            })
    return pd.DataFrame(rows) if rows else None


def _build_eff_frontier_runs(projects, all_results):
    """Sheet: eff_frontier_runs  —  Project, Model, Test, AI_Positive_Rate_pct,
    LF_Capture_pct, Specificity_LF_pct, F1_LF, Efficiency_Score.

    Specificity_LF_pct is the model's specificity against the Listfinal gold
    standard — used as the X axis of the efficiency frontier charts.
    """
    diag = all_results.get("diagnostic", {})
    lf = all_results.get("listfinal", {})
    rows = []
    for pn in sorted(projects):
        proj = projects[pn]
        for mn in sorted(proj["models"]):
            mname = proj["models"][mn]["name"]
            for tn in sorted(proj["models"][mn]["tests"]):
                d = diag.get(pn, {}).get(mn, {}).get(tn)
                l = lf.get(pn, {}).get(mn, {}).get(tn)
                if not d:
                    continue
                n_total = d["n_paired"]
                ai_pos = d["tp"] + d["fp"]
                ai_pos_rate = ai_pos / n_total * 100 if n_total > 0 else float("nan")
                lf_cap = l["capture_rate"] * 100 if l and not np.isnan(l["capture_rate"]) else float("nan")
                eff = (l["capture_rate"] * (1 - ai_pos / n_total)
                       if l and not np.isnan(l["capture_rate"]) and n_total > 0
                       else float("nan"))
                mlf = compute_metrics_vs_lf(n_total, ai_pos, l)
                rows.append({
                    "Project": proj["name"],
                    "Model": mname,
                    "Test": tn,
                    "AI_Positive_Rate_pct": ai_pos_rate,
                    "LF_Capture_pct": lf_cap,
                    "Specificity_LF_pct": mlf["spec_lf"] * 100 if not np.isnan(mlf["spec_lf"]) else float("nan"),
                    "F1_LF": mlf["f1_lf"],
                    "Efficiency_Score": eff,
                })
    return pd.DataFrame(rows) if rows else None


def _build_human_vs_lf(projects, all_results):
    """Sheet: human_vs_lf  —  Project, Sens_LF_pct, Spec_LF_pct, F1_LF,
    Inclusion_Rate_pct, N_universe.

    Human reviewer's TIAB decisions vs Listfinal gold standard. One row per project.
    """
    hu_lf = all_results.get("human_listfinal", {})
    rows = []
    for pn in sorted(projects):
        h = hu_lf.get(pn)
        if not h:
            continue
        proj = projects[pn]
        n_uni = h.get("n_universe")
        n_pos = h.get("n_human_pos")
        m = compute_metrics_vs_lf(n_uni, n_pos, h)
        rows.append({
            "Project": proj["name"],
            "Sens_LF_pct": m["sens_lf"] * 100 if not np.isnan(m["sens_lf"]) else float("nan"),
            "Spec_LF_pct": m["spec_lf"] * 100 if not np.isnan(m["spec_lf"]) else float("nan"),
            "F1_LF": m["f1_lf"],
            "Inclusion_Rate_pct": m["inclusion_rate"] * 100 if not np.isnan(m["inclusion_rate"]) else float("nan"),
            "N_universe": n_uni,
            "N_positives": n_pos,
            "LF_total": h.get("n_listfinal"),
            "LF_not_in_TIAB": h.get("n_not_found"),
        })
    return pd.DataFrame(rows) if rows else None


def _build_eff_score_by_project(projects, all_results):
    """Sheet: eff_score_by_project  —  same structure as eff_frontier_runs
    (used for per-project efficiency bar chart)."""
    return _build_eff_frontier_runs(projects, all_results)


def _build_eff_score_aggregated(projects, all_results):
    """Sheet: eff_score_aggregated  —  Model, Mean_Efficiency_Score, SD."""
    df_runs = _build_eff_frontier_runs(projects, all_results)
    if df_runs is None or df_runs.empty:
        return None
    agg = df_runs.groupby("Model")["Efficiency_Score"].agg(["mean", "std"]).reset_index()
    agg.columns = ["Model", "Mean_Efficiency_Score", "SD"]
    agg["SD"] = agg["SD"].fillna(0)
    return agg.sort_values("Mean_Efficiency_Score", ascending=False)


def _build_sens_spec_dual_gold(projects, all_results):
    """Sheet: sens_spec_dual_gold  —  Model, Sens_TIAB_pct, Sens_LF_pct,
    Spec_TIAB_pct, Spec_LF_pct (averages across projects/tests)."""
    diag = all_results.get("diagnostic", {})
    lf = all_results.get("listfinal", {})
    model_data = {}
    for pn in sorted(projects):
        proj = projects[pn]
        for mn in sorted(proj["models"]):
            mname = proj["models"][mn]["name"]
            if mname not in model_data:
                model_data[mname] = {"sens_tiab": [], "spec_tiab": [],
                                     "sens_lf": [], "spec_lf": []}
            for tn, r in diag.get(pn, {}).get(mn, {}).items():
                if r is None:
                    continue
                s = r["metrics"]["Sensitivity"]
                sp = r["metrics"]["Specificity"]
                if not np.isnan(s):
                    model_data[mname]["sens_tiab"].append(s * 100)
                if not np.isnan(sp):
                    model_data[mname]["spec_tiab"].append(sp * 100)

            # For listfinal gold standard, sensitivity = capture rate,
            # specificity is estimated from diagnostic data (AI excludes / total excludes)
            for tn in lf.get(pn, {}).get(mn, {}):
                lr = lf[pn][mn][tn]
                cr = lr["capture_rate"]
                if not np.isnan(cr):
                    model_data[mname]["sens_lf"].append(cr * 100)
                # Specificity vs LF: proportion of non-LF articles the AI correctly excluded
                d = diag.get(pn, {}).get(mn, {}).get(tn)
                if d:
                    spec_val = d["metrics"]["Specificity"]
                    if not np.isnan(spec_val):
                        model_data[mname]["spec_lf"].append(spec_val * 100)

    rows = []
    for mname, d in model_data.items():
        rows.append({
            "Model": mname,
            "Sens_TIAB_pct": np.mean(d["sens_tiab"]) if d["sens_tiab"] else float("nan"),
            "Sens_LF_pct": np.mean(d["sens_lf"]) if d["sens_lf"] else float("nan"),
            "Spec_TIAB_pct": np.mean(d["spec_tiab"]) if d["spec_tiab"] else float("nan"),
            "Spec_LF_pct": np.mean(d["spec_lf"]) if d["spec_lf"] else float("nan"),
        })
    return pd.DataFrame(rows) if rows else None


def _build_aggregated_performance(projects, all_results):
    """Sheet: aggregated_performance  —  Model + metric_mean / metric_sd pairs.

    F1 is computed against the Listfinal gold standard (not TIAB diagnostic).
    """
    diag = all_results.get("diagnostic", {})
    lf = all_results.get("listfinal", {})
    tr = all_results.get("test_retest", {})
    diag_metrics = ["Sensitivity", "Specificity", "Accuracy"]
    model_vals = {}
    for pn in sorted(projects):
        proj = projects[pn]
        for mn in sorted(proj["models"]):
            mname = proj["models"][mn]["name"]
            if mname not in model_vals:
                model_vals[mname] = {m: [] for m in diag_metrics}
                model_vals[mname]["F1_LF"] = []
                model_vals[mname]["LF_Capture"] = []
                model_vals[mname]["Kappa_TR"] = []
            for tn, r in diag.get(pn, {}).get(mn, {}).items():
                if r is None:
                    continue
                for m in diag_metrics:
                    v = r["metrics"][m]
                    if not np.isnan(v):
                        model_vals[mname][m].append(v)
                lr = lf.get(pn, {}).get(mn, {}).get(tn)
                f1_lf = compute_f1_lf(r, lr)
                if not np.isnan(f1_lf):
                    model_vals[mname]["F1_LF"].append(f1_lf)
            for tn in lf.get(pn, {}).get(mn, {}):
                v = lf[pn][mn][tn]["capture_rate"]
                if v is not None and not np.isnan(v):
                    model_vals[mname]["LF_Capture"].append(v)
            tr_r = tr.get(pn, {}).get(mn)
            if tr_r and not np.isnan(tr_r["kappa"]):
                model_vals[mname]["Kappa_TR"].append(tr_r["kappa"])

    rows = []
    all_metric_keys = diag_metrics + ["F1_LF", "LF_Capture", "Kappa_TR"]
    for mname, vals in model_vals.items():
        row = {"Model": mname}
        for mk in all_metric_keys:
            col_name = mk.replace(" ", "_")
            v = vals[mk]
            row[f"{col_name}_mean"] = np.mean(v) if v else float("nan")
            row[f"{col_name}_sd"] = np.std(v) if len(v) > 1 else 0.0
        rows.append(row)
    return pd.DataFrame(rows) if rows else None


def _build_f1_vs_cost(projects, all_results, metadados):
    """Sheet: f1_vs_cost  —  Model, Avg_Cost_USD, Avg_F1_LF.

    F1 is computed against the Listfinal gold standard.
    """
    if metadados is None:
        return None
    if "code" not in metadados.columns:
        return None
    lf = all_results.get("listfinal", {})
    diag = all_results.get("diagnostic", {})
    model_data = {}
    for pn in sorted(projects):
        proj = projects[pn]
        for mn in sorted(proj["models"]):
            mname = proj["models"][mn]["name"]
            if mname not in model_data:
                model_data[mname] = {"cost": [], "f1_lf": []}
            for tn in proj["models"][mn]["tests"]:
                code = proj["models"][mn]["tests"][tn]["code"]
                meta_m = metadados[metadados["code"].astype(str) == str(code)]
                if not meta_m.empty and pd.notna(meta_m.iloc[0].get("cost_total")):
                    model_data[mname]["cost"].append(meta_m.iloc[0]["cost_total"])
                r = diag.get(pn, {}).get(mn, {}).get(tn)
                lr = lf.get(pn, {}).get(mn, {}).get(tn)
                f1_lf = compute_f1_lf(r, lr)
                if not np.isnan(f1_lf):
                    model_data[mname]["f1_lf"].append(f1_lf)
    rows = []
    for mname, d in model_data.items():
        if d["cost"] and d["f1_lf"]:
            rows.append({
                "Model": mname,
                "Avg_Cost_USD": np.mean(d["cost"]),
                "Avg_F1_LF": np.mean(d["f1_lf"]),
            })
    return pd.DataFrame(rows) if rows else None


def _build_sens_spec_tradeoff(projects, all_results):
    """Sheet: sens_spec_tradeoff  —  Model, Project, Test, Gold_Standard,
    Sensitivity_pct, Specificity_pct."""
    diag = all_results.get("diagnostic", {})
    lf = all_results.get("listfinal", {})
    rows = []
    for pn in sorted(projects):
        proj = projects[pn]
        for mn in sorted(proj["models"]):
            mname = proj["models"][mn]["name"]
            for tn, r in diag.get(pn, {}).get(mn, {}).items():
                if r is None:
                    continue
                rows.append({
                    "Model": mname,
                    "Project": proj["name"],
                    "Test": tn,
                    "Gold_Standard": "TIAB",
                    "Sensitivity_pct": r["metrics"]["Sensitivity"] * 100,
                    "Specificity_pct": r["metrics"]["Specificity"] * 100,
                })
                lr = lf.get(pn, {}).get(mn, {}).get(tn)
                if lr:
                    rows.append({
                        "Model": mname,
                        "Project": proj["name"],
                        "Test": tn,
                        "Gold_Standard": "Listfinal",
                        "Sensitivity_pct": lr["capture_rate"] * 100 if not np.isnan(lr["capture_rate"]) else float("nan"),
                        "Specificity_pct": r["metrics"]["Specificity"] * 100,
                    })
    return pd.DataFrame(rows) if rows else None


def _build_model_ranking_heatmap(projects, all_results, metadados):
    """Sheet: model_ranking_heatmap  —  Model + metric columns + Overall_Score.

    Ranks models by averaging normalised metrics (higher = better).
    F1 is computed against the Listfinal gold standard.
    """
    if metadados is not None and "code" not in metadados.columns:
        metadados = None
    diag = all_results.get("diagnostic", {})
    lf = all_results.get("listfinal", {})
    tr = all_results.get("test_retest", {})

    model_avgs = {}
    for pn in sorted(projects):
        proj = projects[pn]
        for mn in sorted(proj["models"]):
            mname = proj["models"][mn]["name"]
            if mname not in model_avgs:
                model_avgs[mname] = {
                    "Sensitivity": [], "Specificity": [], "F1_LF": [],
                    "LF_Capture": [], "Kappa_TR": [], "Cost_USD": [],
                }
            for tn, r in diag.get(pn, {}).get(mn, {}).items():
                if r is None:
                    continue
                for mk, dk in [("Sensitivity", "Sensitivity"), ("Specificity", "Specificity")]:
                    v = r["metrics"][dk]
                    if not np.isnan(v):
                        model_avgs[mname][mk].append(v * 100)
                lr = lf.get(pn, {}).get(mn, {}).get(tn)
                f1_lf = compute_f1_lf(r, lr)
                if not np.isnan(f1_lf):
                    model_avgs[mname]["F1_LF"].append(f1_lf * 100)
                code = proj["models"][mn]["tests"][tn]["code"]
                if metadados is not None:
                    mc = metadados[metadados["code"].astype(str) == str(code)]
                    if not mc.empty and pd.notna(mc.iloc[0].get("cost_total")):
                        model_avgs[mname]["Cost_USD"].append(mc.iloc[0]["cost_total"])
            for tn in lf.get(pn, {}).get(mn, {}):
                v = lf[pn][mn][tn]["capture_rate"]
                if v is not None and not np.isnan(v):
                    model_avgs[mname]["LF_Capture"].append(v * 100)
            tr_r = tr.get(pn, {}).get(mn)
            if tr_r and not np.isnan(tr_r["kappa"]):
                model_avgs[mname]["Kappa_TR"].append(tr_r["kappa"] * 100)

    metric_keys = ["Sensitivity", "Specificity", "F1_LF", "LF_Capture", "Kappa_TR"]
    rows = []
    for mname, avgs in model_avgs.items():
        row = {"Model": mname}
        vals_for_score = []
        for mk in metric_keys:
            v = np.mean(avgs[mk]) if avgs[mk] else float("nan")
            row[mk] = v
            if not np.isnan(v):
                vals_for_score.append(v / 100.0)  # normalise to 0-1 for score
        row["Overall_Score"] = np.mean(vals_for_score) if vals_for_score else float("nan")
        rows.append(row)

    df = pd.DataFrame(rows)
    if not df.empty:
        df = df.sort_values("Overall_Score", ascending=False).reset_index(drop=True)
    return df if not df.empty else None


def _build_fulltext_hours_saved(projects, all_results):
    """Sheet: fulltext_hours_saved  —  hours of full-text reading the AI
    saves by including fewer articles at TIAB than the human screener.

    Time model: 2 reviewers × 5 min per article = 10 reviewer-minutes per article
    (= 1/6 hour). Hours saved = (Human_TIAB_Included − AI_TIAB_Included) × 10 / 60.
    Negative values mean the AI was more inclusive than the human (extra workload).
    """
    diag = all_results.get("diagnostic", {})
    rows = []
    for pn in sorted(projects):
        proj = projects[pn]
        for mn in sorted(proj["models"]):
            mname = proj["models"][mn]["name"]
            for tn in sorted(proj["models"][mn]["tests"]):
                d = diag.get(pn, {}).get(mn, {}).get(tn)
                if not d:
                    continue
                ai_pos = d["tp"] + d["fp"]
                hu_pos = d["tp"] + d["fn"]
                articles_saved = hu_pos - ai_pos
                hours_saved = articles_saved * 2 * 5 / 60.0
                rows.append({
                    "Project": proj["name"],
                    "Model": mname,
                    "Test": tn,
                    "Human_TIAB_Included": hu_pos,
                    "AI_TIAB_Included": ai_pos,
                    "Articles_Saved": articles_saved,
                    "Hours_Saved_Fulltext": hours_saved,
                })
    return pd.DataFrame(rows) if rows else None


# =====================================================================
#  PAPER-ONLY SHEETS (used by Figures 2 & 3 in the paper section)
# =====================================================================
#
# Both sheets restrict to `official` projects (see constants.PROJECT_CATEGORIES).
# One row per model, averaged across the 3 official projects × 2 tests, with
# standard deviation for error bars.

def _build_paper_sens_spec_lf(projects, all_results, project_classes=None):
    """Sheet: paper_sens_spec_lf — Model, Sens_LF_pct_mean/sd, Spec_LF_pct_mean/sd."""
    diag = all_results.get("diagnostic", {})
    lf = all_results.get("listfinal", {})
    model_agg = {}
    for pn in sorted(projects):
        if project_category(pn, overrides=project_classes) != "official":
            continue
        proj = projects[pn]
        for mn in sorted(proj["models"]):
            mname = proj["models"][mn]["name"]
            entry = model_agg.setdefault(mname, {"sens": [], "spec": []})
            for tn, r in diag.get(pn, {}).get(mn, {}).items():
                if r is None:
                    continue
                lf_res = lf.get(pn, {}).get(mn, {}).get(tn)
                mlf = compute_metrics_vs_lf(r["n_paired"], r["tp"] + r["fp"], lf_res)
                if not np.isnan(mlf["sens_lf"]): entry["sens"].append(mlf["sens_lf"])
                if not np.isnan(mlf["spec_lf"]): entry["spec"].append(mlf["spec_lf"])

    rows = []
    for mname, d in model_agg.items():
        if not d["sens"] and not d["spec"]:
            continue
        rows.append({
            "Model": mname,
            "N_points": len(d["sens"]),
            "Sens_LF_pct_mean": (np.mean(d["sens"]) * 100) if d["sens"] else float("nan"),
            "Sens_LF_pct_sd":   (np.std(d["sens"])  * 100) if len(d["sens"]) > 1 else 0.0,
            "Spec_LF_pct_mean": (np.mean(d["spec"]) * 100) if d["spec"] else float("nan"),
            "Spec_LF_pct_sd":   (np.std(d["spec"])  * 100) if len(d["spec"]) > 1 else 0.0,
        })
    return pd.DataFrame(rows) if rows else None


def _build_paper_f1_vs_cost_official(projects, all_results, metadados, project_classes=None):
    """Sheet: paper_f1_vs_cost_official — Model, Avg_Cost_USD, Avg_F1_LF (± SD).

    Same signature as f1_vs_cost but restricted to official projects only.
    """
    if metadados is None or "code" not in metadados.columns:
        return None
    diag = all_results.get("diagnostic", {})
    lf = all_results.get("listfinal", {})
    model_data = {}
    for pn in sorted(projects):
        if project_category(pn, overrides=project_classes) != "official":
            continue
        proj = projects[pn]
        for mn in sorted(proj["models"]):
            mname = proj["models"][mn]["name"]
            entry = model_data.setdefault(mname, {"cost": [], "f1_lf": []})
            for tn in proj["models"][mn]["tests"]:
                code = proj["models"][mn]["tests"][tn]["code"]
                meta_m = metadados[metadados["code"].astype(str) == str(code)]
                if not meta_m.empty and pd.notna(meta_m.iloc[0].get("cost_total")):
                    entry["cost"].append(meta_m.iloc[0]["cost_total"])
                r = diag.get(pn, {}).get(mn, {}).get(tn)
                lr = lf.get(pn, {}).get(mn, {}).get(tn)
                f1_lf = compute_f1_lf(r, lr)
                if not np.isnan(f1_lf):
                    entry["f1_lf"].append(f1_lf)
    rows = []
    for mname, d in model_data.items():
        if not (d["cost"] and d["f1_lf"]):
            continue
        rows.append({
            "Model": mname,
            "Avg_Cost_USD":   np.mean(d["cost"]),
            "Avg_Cost_USD_sd": (np.std(d["cost"]) if len(d["cost"]) > 1 else 0.0),
            "Avg_F1_LF":      np.mean(d["f1_lf"]),
            "Avg_F1_LF_sd":   (np.std(d["f1_lf"]) if len(d["f1_lf"]) > 1 else 0.0),
        })
    return pd.DataFrame(rows) if rows else None


# =====================================================================
#  PUBLIC API
# =====================================================================

def export_chart_data(projects, all_results, metadados, output_dir: Path,
                       project_classes: dict | None = None) -> Path:
    """Build all chart DataFrames and write them to data_grafics_<timestamp>.xlsx.

    Returns the path to the generated XLSX.
    """
    builders = [
        ("sensitivity_per_model",    lambda: _build_sensitivity_per_model(projects, all_results)),
        ("lf_capture_heatmap",       lambda: _build_lf_capture_heatmap(projects, all_results)),
        ("test_retest_kappa",        lambda: _build_test_retest_kappa(projects, all_results)),
        ("model_comparison_radar",   lambda: _build_model_comparison_radar(projects, all_results)),
        ("cost_vs_sensitivity",      lambda: _build_cost_vs_sensitivity(projects, all_results, metadados)),
        ("workload_reduction",       lambda: _build_workload_reduction(projects, all_results, metadados)),
        ("eff_frontier_runs",        lambda: _build_eff_frontier_runs(projects, all_results)),
        ("eff_score_by_project",     lambda: _build_eff_score_by_project(projects, all_results)),
        ("eff_score_aggregated",     lambda: _build_eff_score_aggregated(projects, all_results)),
        ("sens_spec_dual_gold",      lambda: _build_sens_spec_dual_gold(projects, all_results)),
        ("aggregated_performance",   lambda: _build_aggregated_performance(projects, all_results)),
        ("f1_vs_cost",               lambda: _build_f1_vs_cost(projects, all_results, metadados)),
        ("sens_spec_tradeoff",       lambda: _build_sens_spec_tradeoff(projects, all_results)),
        ("model_ranking_heatmap",    lambda: _build_model_ranking_heatmap(projects, all_results, metadados)),
        ("fulltext_hours_saved",     lambda: _build_fulltext_hours_saved(projects, all_results)),
        ("human_vs_lf",              lambda: _build_human_vs_lf(projects, all_results)),
        # Paper-only sheets (official projects only) — used by Figures 2 & 3.
        # Honour the classification supplied by the web UI (project_classes).
        ("paper_sens_spec_lf",           lambda: _build_paper_sens_spec_lf(projects, all_results, project_classes=project_classes)),
        ("paper_f1_vs_cost_official",    lambda: _build_paper_f1_vs_cost_official(projects, all_results, metadados, project_classes=project_classes)),
    ]

    ts = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    xlsx_path = output_dir / f"data_grafics_{ts}.xlsx"
    output_dir.mkdir(parents=True, exist_ok=True)

    with pd.ExcelWriter(str(xlsx_path), engine="openpyxl") as writer:
        written = 0
        for sheet_name, builder_fn in builders:
            try:
                df = builder_fn()
            except Exception as exc:
                print(f"    ⚠ chart_data: error building '{sheet_name}': {exc}")
                continue
            if df is not None and not df.empty:
                df.to_excel(writer, sheet_name=sheet_name, index=False)
                written += 1
            else:
                print(f"    ⚠ chart_data: no data for '{sheet_name}' — skipped")

    print(f"  ✓ Chart data exported ({written} sheets): {xlsx_path.name}")
    return xlsx_path
