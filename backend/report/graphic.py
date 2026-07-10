#!/usr/bin/env python3
"""
graphic.py — Standalone Chart Generator (procedural, no functions)
===================================================================
Reads the data_grafics_*.xlsx file produced by 001_report.py
and recreates all 16 publication-ready charts as standalone PNGs.

Usage:
    python report/graphic.py                             # auto-finds latest XLSX in output/
    python report/graphic.py path/to/data_grafics.xlsx   # specify XLSX explicitly

Output:
    output/figures_custom/*.png   (16 charts at 300 DPI)

Every chart section below is self-contained.  Edit colors, labels,
sizes, fonts, or layout freely for your publication needs.
"""

import argparse
import sys
from pathlib import Path

import numpy as np
import pandas as pd

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib.lines import Line2D
from matplotlib.patches import FancyBboxPatch

# ──────────────────────────────────────────────────────────────────────
#  CONFIGURATION — edit these to customize all charts at once
# ──────────────────────────────────────────────────────────────────────

MODEL_COLORS = {
    "gpt_4o":     "#4472C4",
    "gpt_5_2":    "#ED7D31",
    "gpt_5_mini": "#A5A5A5",
    "gpt_5_nano": "#FFC000",
}

DEFAULT_COLORS = [
    "#5B9BD5", "#ED7D31", "#A5A5A5", "#FFC000", "#70AD47",
    "#9B59B6", "#E74C3C", "#1ABC9C", "#34495E", "#F39C12",
]

DPI = 300
FIG_FORMAT = "png"
FACECOLOR = "white"


# ──────────────────────────────────────────────────────────────────────
#  LOCATE XLSX
# ──────────────────────────────────────────────────────────────────────

parser = argparse.ArgumentParser(description="Generate charts from data_grafics XLSX")
parser.add_argument("xlsx", nargs="?", default=None,
                    help="Path to data_grafics_*.xlsx (auto-detects latest if omitted)")
parser.add_argument("-o", "--output", default=None,
                    help="Output directory for PNGs (default: output/figures_custom)")
_args = parser.parse_args()

if _args.xlsx:
    xlsx_path = Path(_args.xlsx)
else:
    _base = Path(__file__).resolve().parent.parent / "output"
    _candidates = sorted(_base.glob("data_grafics_*.xlsx"), reverse=True)
    if not _candidates:
        print("ERROR: No data_grafics_*.xlsx found in output/. "
              "Run relatorio_unificado.py first.")
        sys.exit(1)
    xlsx_path = _candidates[0]

if not xlsx_path.exists():
    print(f"ERROR: File not found: {xlsx_path}")
    sys.exit(1)

out_dir = Path(_args.output) if _args.output else xlsx_path.parent / "figures_custom"
out_dir.mkdir(parents=True, exist_ok=True)

print(f"  Reading: {xlsx_path.name}")
sheets = pd.read_excel(str(xlsx_path), sheet_name=None, engine="openpyxl")
print(f"  Sheets found: {len(sheets)} — {', '.join(sheets.keys())}")
print(f"  Output: {out_dir}\n")

generated = 0


# ──────────────────────────────────────────────────────────────────────
#  CHART 1 — Sensitivity by Model per Project (grouped bar)
# ──────────────────────────────────────────────────────────────────────

if "sensitivity_per_model" in sheets:
    df = sheets["sensitivity_per_model"]
    if not df.empty:
        projects = df["Project"].unique()
        models = df["Model"].unique()
        n_proj = len(projects)

        fig, axes = plt.subplots(1, n_proj, figsize=(4.5 * n_proj, 4.5),
                                 sharey=True, squeeze=False)
        fig.suptitle("Sensitivity by Model per Project (vs Human TIAB)",
                     fontsize=12, fontweight="bold", y=1.02)

        for ax_idx, proj in enumerate(projects):
            ax = axes[0][ax_idx]
            means, indiv, colors, names = [], [], [], []
            for mi, model in enumerate(models):
                sub = df[(df["Project"] == proj) & (df["Model"] == model)]
                vals = sub["Sensitivity"].dropna().values
                means.append(np.mean(vals) * 100 if len(vals) else 0)
                indiv.append(vals * 100 if len(vals) else np.array([]))
                key = model.strip().lower().replace(" ", "_").replace("-", "_")
                colors.append(MODEL_COLORS.get(key, DEFAULT_COLORS[mi % len(DEFAULT_COLORS)]))
                names.append(model)

            x = np.arange(len(models))
            bars = ax.bar(x, means, width=0.65, color=colors, edgecolor="#333",
                          linewidth=0.5, alpha=0.85)
            for bi, vals in enumerate(indiv):
                if len(vals):
                    jitter = np.linspace(-0.1, 0.1, len(vals)) if len(vals) > 1 else [0]
                    for ji, v in zip(jitter, vals):
                        ax.scatter(bi + ji, v, color="#222", s=20, zorder=5,
                                   edgecolors="white", linewidth=0.4)
            ax.axhline(y=95, color="#27AE60", linestyle="--", alpha=0.6, linewidth=1)
            ax.axhline(y=80, color="#E74C3C", linestyle=":", alpha=0.4, linewidth=0.8)
            ax.set_title(proj, fontsize=10, fontweight="bold", pad=8)
            if ax_idx == 0:
                ax.set_ylabel("Sensitivity (%)", fontsize=9)
            ax.spines["top"].set_visible(False)
            ax.spines["right"].set_visible(False)
            ax.spines["left"].set_linewidth(0.6)
            ax.spines["bottom"].set_linewidth(0.6)
            ax.set_ylim(0, 108)
            ax.set_xticks(x)
            ax.set_xticklabels(names, rotation=25, ha="right", fontsize=7)
            for bar, val in zip(bars, means):
                if val > 0:
                    ax.text(bar.get_x() + bar.get_width() / 2, bar.get_height() + 1,
                            f"{val:.0f}%", ha="center", va="bottom",
                            fontsize=7, fontweight="bold")

        plt.tight_layout(rect=[0, 0, 1, 0.96])
        fig.savefig(str(out_dir / f"sensitivity_per_model_per_project.{FIG_FORMAT}"),
                    dpi=DPI, bbox_inches="tight", facecolor=FACECOLOR, edgecolor="none")
        plt.close(fig)
        print(f"    ✓ sensitivity_per_model_per_project.{FIG_FORMAT}")
        generated += 1
else:
    print("    ⚠ Sheet 'sensitivity_per_model' not found — skipping")


# ──────────────────────────────────────────────────────────────────────
#  CHART 2 — Listfinal Capture Rate Heatmap
# ──────────────────────────────────────────────────────────────────────

if "lf_capture_heatmap" in sheets:
    df = sheets["lf_capture_heatmap"]
    if not df.empty:
        models = df["Model"].unique()
        projects = df["Project"].unique()
        matrix = np.full((len(models), len(projects)), np.nan)
        for mi, m in enumerate(models):
            for pi, p in enumerate(projects):
                sub = df[(df["Model"] == m) & (df["Project"] == p)]
                if not sub.empty and pd.notna(sub.iloc[0]["Capture_Rate_pct"]):
                    matrix[mi, pi] = sub.iloc[0]["Capture_Rate_pct"]

        fig, ax = plt.subplots(figsize=(3 + len(projects) * 1.5,
                                        1.2 + len(models) * 0.8))
        masked = np.ma.array(matrix, mask=np.isnan(matrix))
        im = ax.imshow(masked, cmap="RdYlGn", vmin=60, vmax=100, aspect="auto")
        ax.set_xticks(np.arange(len(projects)))
        ax.set_xticklabels(projects, fontsize=9, fontweight="bold")
        ax.set_yticks(np.arange(len(models)))
        ax.set_yticklabels(models, fontsize=9)
        ax.tick_params(top=True, bottom=False, labeltop=True, labelbottom=False)

        for mi in range(len(models)):
            for pi in range(len(projects)):
                val = matrix[mi, pi]
                if not np.isnan(val):
                    color = "white" if val < 80 else "black"
                    ax.text(pi, mi, f"{val:.1f}%", ha="center", va="center",
                            fontsize=9, fontweight="bold", color=color)
                else:
                    ax.text(pi, mi, "—", ha="center", va="center",
                            fontsize=9, color="gray")

        cbar = fig.colorbar(im, ax=ax, shrink=0.8, pad=0.04)
        cbar.set_label("Capture Rate (%)", fontsize=8)
        cbar.ax.tick_params(labelsize=7)
        ax.set_title("Listfinal Capture Rate by Model and Project\n"
                     "(Average across Runs)",
                     fontsize=11, fontweight="bold", pad=12)
        plt.tight_layout()
        fig.savefig(str(out_dir / f"listfinal_capture_heatmap.{FIG_FORMAT}"),
                    dpi=DPI, bbox_inches="tight", facecolor=FACECOLOR, edgecolor="none")
        plt.close(fig)
        print(f"    ✓ listfinal_capture_heatmap.{FIG_FORMAT}")
        generated += 1
else:
    print("    ⚠ Sheet 'lf_capture_heatmap' not found — skipping")


# ──────────────────────────────────────────────────────────────────────
#  CHART 3 — Test-Retest Kappa with 95 % CI
# ──────────────────────────────────────────────────────────────────────

if "test_retest_kappa" in sheets:
    df = sheets["test_retest_kappa"]
    if not df.empty:
        labels = df["Label"].values
        kappas = df["Kappa"].values
        ci_los = df["CI_lo"].values
        ci_his = df["CI_hi"].values

        colors = []
        for i, lbl in enumerate(labels):
            model_name = lbl.split("\n")[0] if "\n" in str(lbl) else str(lbl)
            key = model_name.strip().lower().replace(" ", "_").replace("-", "_")
            colors.append(MODEL_COLORS.get(key, DEFAULT_COLORS[i % len(DEFAULT_COLORS)]))

        fig, ax = plt.subplots(figsize=(10, 5))
        x = np.arange(len(labels))
        yerr_lo = [max(0, k - lo) for k, lo in zip(kappas, ci_los)]
        yerr_hi = [max(0, hi - k) for k, hi in zip(kappas, ci_his)]

        bars = ax.bar(x, kappas, width=0.65, color=colors, edgecolor="#333",
                      linewidth=0.5, alpha=0.85,
                      yerr=[yerr_lo, yerr_hi], capsize=4,
                      error_kw={"linewidth": 1.2, "color": "#444", "capthick": 1.2})

        ax.axhline(y=0.81, color="#27AE60", linestyle="--", alpha=0.6, linewidth=1,
                   label="Almost Perfect (0.81)")
        ax.axhline(y=0.61, color="#F39C12", linestyle="--", alpha=0.5, linewidth=0.8,
                   label="Substantial (0.61)")

        for i, (bar, val, lo, hi) in enumerate(zip(bars, kappas, ci_los, ci_his)):
            ax.text(bar.get_x() + bar.get_width() / 2, min(hi, 1.0) + 0.015,
                    f"{val:.3f}\n[{lo:.2f}, {hi:.2f}]",
                    ha="center", va="bottom", fontsize=6,
                    fontweight="bold", linespacing=1.1)

        ax.set_title("Test-Retest Reproducibility (Cohen's Kappa with 95% CI)",
                     fontsize=10, fontweight="bold", pad=8)
        ax.set_ylabel("Kappa", fontsize=9)
        ax.spines["top"].set_visible(False)
        ax.spines["right"].set_visible(False)
        ax.spines["left"].set_linewidth(0.6)
        ax.spines["bottom"].set_linewidth(0.6)
        ax.set_ylim(0, 1.18)
        ax.set_xticks(x)
        ax.set_xticklabels(labels, fontsize=7, rotation=0)
        ax.legend(loc="lower right", fontsize=7, framealpha=0.8)
        plt.tight_layout()
        fig.savefig(str(out_dir / f"test_retest_kappa.{FIG_FORMAT}"),
                    dpi=DPI, bbox_inches="tight", facecolor=FACECOLOR, edgecolor="none")
        plt.close(fig)
        print(f"    ✓ test_retest_kappa.{FIG_FORMAT}")
        generated += 1
else:
    print("    ⚠ Sheet 'test_retest_kappa' not found — skipping")


# ──────────────────────────────────────────────────────────────────────
#  CHART 4 — Model Comparison Radar
# ──────────────────────────────────────────────────────────────────────

if "model_comparison_radar" in sheets:
    df = sheets["model_comparison_radar"]
    if not df.empty:
        categories = [c for c in df.columns if c != "Model"]
        N_cat = len(categories)
        angles = np.linspace(0, 2 * np.pi, N_cat, endpoint=False).tolist()
        angles += angles[:1]

        fig, ax = plt.subplots(figsize=(7, 7), subplot_kw=dict(polar=True))
        ax.set_theta_offset(np.pi / 2)
        ax.set_theta_direction(-1)
        ax.set_xticks(angles[:-1])
        ax.set_xticklabels(categories, fontsize=9, fontweight="bold")
        ax.set_ylim(0, 1.05)
        ax.set_yticks([0.2, 0.4, 0.6, 0.8, 1.0])
        ax.set_yticklabels(["20%", "40%", "60%", "80%", "100%"],
                           fontsize=7, alpha=0.6)

        for mi, (_, row) in enumerate(df.iterrows()):
            values = [row[c] if pd.notna(row[c]) else 0 for c in categories]
            values += values[:1]
            key = row["Model"].strip().lower().replace(" ", "_").replace("-", "_")
            color = MODEL_COLORS.get(key, DEFAULT_COLORS[mi % len(DEFAULT_COLORS)])
            ax.plot(angles, values, "o-", linewidth=2, label=row["Model"],
                    color=color, markersize=5)
            ax.fill(angles, values, alpha=0.08, color=color)

        ax.legend(loc="upper right", bbox_to_anchor=(1.35, 1.1),
                  fontsize=8, framealpha=0.9)
        ax.set_title("Model Comparison — Key Metrics Radar\n"
                     "(Average across All Projects & Runs)",
                     fontsize=11, fontweight="bold", pad=30)
        plt.tight_layout()
        fig.savefig(str(out_dir / f"model_comparison_radar.{FIG_FORMAT}"),
                    dpi=DPI, bbox_inches="tight", facecolor=FACECOLOR, edgecolor="none")
        plt.close(fig)
        print(f"    ✓ model_comparison_radar.{FIG_FORMAT}")
        generated += 1
else:
    print("    ⚠ Sheet 'model_comparison_radar' not found — skipping")


# ──────────────────────────────────────────────────────────────────────
#  CHART 5 — Cost vs Sensitivity (Bubble)
# ──────────────────────────────────────────────────────────────────────

if "cost_vs_sensitivity" in sheets:
    df = sheets["cost_vs_sensitivity"]
    if not df.empty:
        fig, ax = plt.subplots(figsize=(9, 5.5))
        for mi, (_, row) in enumerate(df.iterrows()):
            avg_f1 = row["Avg_F1_LF"] if pd.notna(row.get("Avg_F1_LF")) else 0.5
            size = max(avg_f1 * 600, 60)
            key = row["Model"].strip().lower().replace(" ", "_").replace("-", "_")
            color = MODEL_COLORS.get(key, DEFAULT_COLORS[mi % len(DEFAULT_COLORS)])
            ax.scatter(row["Avg_Cost_USD"], row["Avg_Sensitivity_pct"], s=size,
                       color=color, alpha=0.75, edgecolors="#333", linewidth=1.2,
                       zorder=5)
            ax.annotate(row["Model"],
                        (row["Avg_Cost_USD"], row["Avg_Sensitivity_pct"]),
                        textcoords="offset points", xytext=(10, 8), fontsize=8,
                        fontweight="bold",
                        arrowprops=dict(arrowstyle="-", color="gray", lw=0.5))

        ax.axhline(y=95, color="#27AE60", linestyle="--", alpha=0.5, linewidth=1,
                   label="95% Sensitivity")
        ax.set_title("Cost vs Sensitivity (Bubble Size = F1 vs Listfinal)",
                     fontsize=10, fontweight="bold", pad=8)
        ax.set_ylabel("Average Sensitivity (%)", fontsize=9)
        ax.set_xlabel("Average Cost (USD)", fontsize=9)
        ax.spines["top"].set_visible(False)
        ax.spines["right"].set_visible(False)
        ax.spines["left"].set_linewidth(0.6)
        ax.spines["bottom"].set_linewidth(0.6)
        ax.set_ylim(0, 108)
        ax.legend(loc="lower right", fontsize=8, framealpha=0.8)
        plt.tight_layout()
        fig.savefig(str(out_dir / f"cost_vs_sensitivity_bubble.{FIG_FORMAT}"),
                    dpi=DPI, bbox_inches="tight", facecolor=FACECOLOR, edgecolor="none")
        plt.close(fig)
        print(f"    ✓ cost_vs_sensitivity_bubble.{FIG_FORMAT}")
        generated += 1
else:
    print("    ⚠ Sheet 'cost_vs_sensitivity' not found — skipping")


# ──────────────────────────────────────────────────────────────────────
#  CHART 6 — Workload Reduction
# ──────────────────────────────────────────────────────────────────────

if "workload_reduction" in sheets:
    df = sheets["workload_reduction"]
    if not df.empty:
        model_names = df["Model"].values
        human_means = df["Human_Hours"].values
        ai_means = df["AI_Hours"].values
        speed_factors = df["Speed_Factor"].values
        bar_colors = []
        for i, m in enumerate(model_names):
            key = m.strip().lower().replace(" ", "_").replace("-", "_")
            bar_colors.append(MODEL_COLORS.get(key, DEFAULT_COLORS[i % len(DEFAULT_COLORS)]))

        fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(12, 4.5),
                                        gridspec_kw={"width_ratios": [2, 1]})
        y_pos = np.arange(len(model_names))
        bar_h = 0.35

        ax1.barh(y_pos + bar_h / 2, human_means, bar_h, label="Human Time",
                 color="#E74C3C", alpha=0.85, edgecolor="#333", linewidth=0.5)
        ax1.barh(y_pos - bar_h / 2, ai_means, bar_h, label="AI Time",
                 color="#2ECC71", alpha=0.85, edgecolor="#333", linewidth=0.5)
        for i, (hv, av) in enumerate(zip(human_means, ai_means)):
            hr_h, mn_h = int(hv), int((hv - int(hv)) * 60)
            hr_a, mn_a = int(av), int((av - int(av)) * 60)
            lbl_h = f"{hr_h}h {mn_h:02d}m" if hr_h > 0 else f"{mn_h}m"
            lbl_a = f"{hr_a}h {mn_a:02d}m" if hr_a > 0 else f"{mn_a}m"
            ax1.text(hv + 0.3, i + bar_h / 2, lbl_h, va="center", fontsize=7)
            ax1.text(av + 0.3, i - bar_h / 2, lbl_a, va="center", fontsize=7)
        ax1.set_yticks(y_pos)
        ax1.set_yticklabels(model_names, fontsize=8)
        ax1.set_xlabel("Time (hours)", fontsize=9)
        ax1.set_title("Average Screening Time: Human vs AI",
                      fontsize=11, fontweight="bold")
        ax1.legend(loc="lower right", fontsize=8, framealpha=0.8)
        ax1.invert_yaxis()
        ax1.spines["top"].set_visible(False)
        ax1.spines["right"].set_visible(False)

        ax2.barh(y_pos, speed_factors, 0.5, color=bar_colors,
                 edgecolor="#333", linewidth=0.5, alpha=0.85)
        for i, sf in enumerate(speed_factors):
            ax2.text(sf + 0.5, i, f"{sf:.0f}×", va="center", fontsize=9,
                     fontweight="bold")
        ax2.set_yticks(y_pos)
        ax2.set_yticklabels(model_names, fontsize=8)
        ax2.set_xlabel("Speed Factor (×)", fontsize=9)
        ax2.set_title("Speed Factor (Human ÷ AI)", fontsize=11, fontweight="bold")
        ax2.invert_yaxis()
        ax2.spines["top"].set_visible(False)
        ax2.spines["right"].set_visible(False)

        plt.tight_layout()
        fig.savefig(str(out_dir / f"workload_reduction.{FIG_FORMAT}"),
                    dpi=DPI, bbox_inches="tight", facecolor=FACECOLOR, edgecolor="none")
        plt.close(fig)
        print(f"    ✓ workload_reduction.{FIG_FORMAT}")
        generated += 1
else:
    print("    ⚠ Sheet 'workload_reduction' not found — skipping")


# ──────────────────────────────────────────────────────────────────────
#  CHART 7 — Efficiency Frontier Grid (Individual Runs)
# ──────────────────────────────────────────────────────────────────────

if "eff_frontier_runs" in sheets:
    df = sheets["eff_frontier_runs"]
    if not df.empty:
        projects = df["Project"].unique()
        n_panels = len(projects)
        ncols = min(n_panels, 3)
        nrows = (n_panels + ncols - 1) // ncols

        # Human-vs-LF baseline (one row per project) if available
        df_human = sheets.get("human_vs_lf") if "human_vs_lf" in sheets else None

        fig, axes = plt.subplots(nrows, ncols,
                                 figsize=(5.5 * ncols, 5 * nrows), squeeze=False)
        fig.suptitle("Efficiency Frontier by Project (Individual Runs)",
                     fontsize=13, fontweight="bold", y=1.01)

        models_seen = {}
        for _, row in df.iterrows():
            if row["Model"] not in models_seen:
                key = row["Model"].strip().lower().replace(" ", "_").replace("-", "_")
                models_seen[row["Model"]] = MODEL_COLORS.get(
                    key, DEFAULT_COLORS[len(models_seen) % len(DEFAULT_COLORS)])

        for idx, proj in enumerate(projects):
            r, c = divmod(idx, ncols)
            ax = axes[r][c]
            ax.set_facecolor("#FAFAFA")
            # Ideal zone — top-right (high spec + high sens)
            ax.axhspan(95, 106, xmin=0.5, color="#D5F5E3", alpha=0.2, zorder=0)
            sub = df[df["Project"] == proj]
            for _, row in sub.iterrows():
                marker = ("o" if row["Test"] == 1
                          else ("s" if row["Test"] == 2 else "^"))
                ax.scatter(row["Specificity_LF_pct"], row["LF_Capture_pct"],
                           s=120, color=models_seen[row["Model"]],
                           marker=marker, edgecolors="white", linewidth=1.2,
                           zorder=5, alpha=0.9)
            # Human baseline overlay (star)
            if df_human is not None and not df_human.empty:
                hsub = df_human[df_human["Project"] == proj]
                for _, hrow in hsub.iterrows():
                    ax.scatter(hrow["Spec_LF_pct"], hrow["Sens_LF_pct"],
                               s=220, color="#000", marker="*",
                               edgecolors="white", linewidth=1.5,
                               zorder=6, label="Human baseline")
            ax.axhline(y=95, color="#27AE60", linestyle="--", alpha=0.5,
                       linewidth=1)
            ax.set_title(proj, fontsize=10, fontweight="bold", pad=8)
            if c == 0:
                ax.set_ylabel("Sensitivity vs Listfinal (%)", fontsize=9)
            ax.spines["top"].set_visible(False)
            ax.spines["right"].set_visible(False)
            ax.spines["left"].set_linewidth(0.6)
            ax.spines["bottom"].set_linewidth(0.6)
            ax.set_ylim(0, 105)
            if r == nrows - 1:
                ax.set_xlabel("Specificity vs Listfinal (%)", fontsize=8)
            ax.set_xlim(0, 105)

        for idx in range(n_panels, nrows * ncols):
            r2, c2 = divmod(idx, ncols)
            axes[r2][c2].set_visible(False)

        handles = [
            Line2D([0], [0], marker="o", color=col, markersize=8,
                   linestyle="None", markeredgecolor="white",
                   markeredgewidth=0.6, label=m)
            for m, col in models_seen.items()
        ]
        handles.append(Line2D([0], [0], marker="o", color="gray",
                              markersize=6, linestyle="None", label="Test 1"))
        handles.append(Line2D([0], [0], marker="s", color="gray",
                              markersize=6, linestyle="None", label="Test 2"))
        if df_human is not None and not df_human.empty:
            handles.append(Line2D([0], [0], marker="*", color="#000",
                                   markersize=12, linestyle="None",
                                   markeredgecolor="white", markeredgewidth=1,
                                   label="Human baseline"))
        fig.legend(handles=handles, loc="lower center", ncol=len(handles),
                   fontsize=8, framealpha=0.9, bbox_to_anchor=(0.5, -0.02))
        plt.tight_layout(rect=[0, 0.04, 1, 0.97])
        fig.savefig(str(out_dir / f"efficiency_frontier_grid.{FIG_FORMAT}"),
                    dpi=DPI, bbox_inches="tight", facecolor=FACECOLOR, edgecolor="none")
        plt.close(fig)
        print(f"    ✓ efficiency_frontier_grid.{FIG_FORMAT}")
        generated += 1
else:
    print("    ⚠ Sheet 'eff_frontier_runs' not found — skipping charts 7-9")


# ──────────────────────────────────────────────────────────────────────
#  CHART 8 — Efficiency Frontier Average (per Project)
# ──────────────────────────────────────────────────────────────────────

if "eff_frontier_runs" in sheets:
    df = sheets["eff_frontier_runs"]
    if not df.empty:
        projects = df["Project"].unique()
        df_human = sheets.get("human_vs_lf") if "human_vs_lf" in sheets else None
        models_seen = {}
        for _, row in df.iterrows():
            if row["Model"] not in models_seen:
                key = row["Model"].strip().lower().replace(" ", "_").replace("-", "_")
                models_seen[row["Model"]] = MODEL_COLORS.get(
                    key, DEFAULT_COLORS[len(models_seen) % len(DEFAULT_COLORS)])

        fig, axes = plt.subplots(1, len(projects),
                                 figsize=(5.5 * len(projects), 5.5), squeeze=False)
        fig.suptitle("Efficiency Frontier by Project (Mean of Runs per Model)",
                     fontsize=13, fontweight="bold", y=1.01)

        for idx, proj in enumerate(projects):
            ax = axes[0][idx]
            ax.set_facecolor("#FAFAFA")
            ax.axhspan(95, 106, xmin=0.5, color="#D5F5E3", alpha=0.2, zorder=0)
            sub = df[df["Project"] == proj]
            for model, color in models_seen.items():
                ms = sub[sub["Model"] == model]
                if ms.empty:
                    continue
                mx = ms["Specificity_LF_pct"].mean()
                my = ms["LF_Capture_pct"].mean()
                sx = ms["Specificity_LF_pct"].std() if len(ms) > 1 else 0
                sy = ms["LF_Capture_pct"].std() if len(ms) > 1 else 0
                ax.errorbar(mx, my, xerr=sx, yerr=sy, fmt="none",
                            ecolor=color, elinewidth=1.2, capsize=4,
                            capthick=1, alpha=0.5, zorder=4)
                ax.scatter(mx, my, s=200, color=color, edgecolors="white",
                           linewidth=2, zorder=6)
                ax.scatter(ms["Specificity_LF_pct"], ms["LF_Capture_pct"],
                           s=25, color=color, alpha=0.35, edgecolors="none",
                           zorder=3)
            # Human baseline overlay
            if df_human is not None and not df_human.empty:
                hsub = df_human[df_human["Project"] == proj]
                for _, hrow in hsub.iterrows():
                    ax.scatter(hrow["Spec_LF_pct"], hrow["Sens_LF_pct"],
                               s=260, color="#000", marker="*",
                               edgecolors="white", linewidth=1.8, zorder=7)
            ax.axhline(y=95, color="#27AE60", linestyle="--", alpha=0.5,
                       linewidth=1)
            ax.set_title(proj, fontsize=10, fontweight="bold", pad=8)
            if idx == 0:
                ax.set_ylabel("Sensitivity vs Listfinal (%)", fontsize=9)
            ax.spines["top"].set_visible(False)
            ax.spines["right"].set_visible(False)
            ax.spines["left"].set_linewidth(0.6)
            ax.spines["bottom"].set_linewidth(0.6)
            ax.set_ylim(0, 105)
            ax.set_xlabel("Specificity vs Listfinal (%)", fontsize=8)
            ax.set_xlim(0, 105)

        handles = []
        for model, color in models_seen.items():
            ms_all = df[df["Model"] == model]
            eff = ms_all["Efficiency_Score"].mean() if not ms_all.empty else 0
            handles.append(Line2D([0], [0], marker="o", color=color,
                                  markersize=8, linestyle="None",
                                  markeredgecolor="white", markeredgewidth=0.6,
                                  label=f"{model} (Eff: {eff:.3f})"))
        if df_human is not None and not df_human.empty:
            handles.append(Line2D([0], [0], marker="*", color="#000",
                                   markersize=12, linestyle="None",
                                   markeredgecolor="white", markeredgewidth=1,
                                   label="Human baseline"))
        fig.legend(handles=handles, loc="lower center", ncol=len(handles),
                   fontsize=8, framealpha=0.9, bbox_to_anchor=(0.5, -0.04))
        plt.tight_layout(rect=[0, 0.05, 1, 0.97])
        fig.savefig(str(out_dir / f"efficiency_frontier_by_project_avg.{FIG_FORMAT}"),
                    dpi=DPI, bbox_inches="tight", facecolor=FACECOLOR, edgecolor="none")
        plt.close(fig)
        print(f"    ✓ efficiency_frontier_by_project_avg.{FIG_FORMAT}")
        generated += 1


# ──────────────────────────────────────────────────────────────────────
#  CHART 9 — Efficiency Frontier Overall (all projects)
# ──────────────────────────────────────────────────────────────────────

if "eff_frontier_runs" in sheets:
    df = sheets["eff_frontier_runs"]
    if not df.empty:
        df_human = sheets.get("human_vs_lf") if "human_vs_lf" in sheets else None
        models_seen = {}
        for _, row in df.iterrows():
            if row["Model"] not in models_seen:
                key = row["Model"].strip().lower().replace(" ", "_").replace("-", "_")
                models_seen[row["Model"]] = MODEL_COLORS.get(
                    key, DEFAULT_COLORS[len(models_seen) % len(DEFAULT_COLORS)])

        fig, ax = plt.subplots(figsize=(9, 6))
        ax.set_facecolor("#FAFAFA")
        # Ideal zone — top-right (high specificity + high sensitivity)
        ax.axhspan(95, 106, xmin=0.5, xmax=1, color="#D5F5E3", alpha=0.25,
                   zorder=0)
        # Bottom-left = poor zone
        ax.axhspan(0, 50, xmin=0, xmax=0.5, color="#FADBD8", alpha=0.18,
                   zorder=0)

        legend_items = []
        for model, color in models_seen.items():
            ms = df[df["Model"] == model]
            if ms.empty:
                continue
            mean_x = ms["Specificity_LF_pct"].mean()
            mean_y = ms["LF_Capture_pct"].mean()
            std_x = ms["Specificity_LF_pct"].std() if len(ms) > 1 else 0
            std_y = ms["LF_Capture_pct"].std() if len(ms) > 1 else 0
            mean_score = ms["Efficiency_Score"].mean()
            ax.errorbar(mean_x, mean_y, xerr=std_x, yerr=std_y, fmt="none",
                        ecolor=color, elinewidth=1.5, capsize=5, capthick=1.2,
                        alpha=0.5, zorder=4)
            ax.scatter(mean_x, mean_y, s=280, color=color, edgecolors="white",
                       linewidth=2.5, zorder=6, marker="o")
            legend_items.append(
                Line2D([0], [0], marker="o", color=color, markersize=10,
                       linestyle="None", markeredgecolor="white",
                       markeredgewidth=1,
                       label=f"{model} (Eff: {mean_score:.3f})"))

        # Human baseline — pooled across projects
        if df_human is not None and not df_human.empty:
            hx = df_human["Spec_LF_pct"].mean()
            hy = df_human["Sens_LF_pct"].mean()
            ax.scatter(hx, hy, s=320, color="#000", marker="*",
                       edgecolors="white", linewidth=2, zorder=7)
            legend_items.append(
                Line2D([0], [0], marker="*", color="#000", markersize=14,
                       linestyle="None", markeredgecolor="white",
                       markeredgewidth=1.2, label="Human baseline (mean)"))

        ax.axhline(y=95, color="#27AE60", linestyle="--", alpha=0.6,
                   linewidth=1.2)
        ax.annotate("IDEAL ZONE", xy=(96, 102), fontsize=9, fontweight="bold",
                    color="#27AE60", alpha=0.5, ha="right")
        ax.set_title("Efficiency Frontier: Overall Mean per Model (± 1 SD)",
                     fontsize=10, fontweight="bold", pad=8)
        ax.set_ylabel("Sensitivity vs Listfinal (%)", fontsize=9)
        ax.set_xlabel("Specificity vs Listfinal (%) — higher = better at excluding non-relevant",
                      fontsize=9)
        ax.spines["top"].set_visible(False)
        ax.spines["right"].set_visible(False)
        ax.spines["left"].set_linewidth(0.6)
        ax.spines["bottom"].set_linewidth(0.6)
        ax.set_ylim(0, 105)
        ax.set_xlim(0, 105)
        legend_items.append(Line2D([0], [0], color="#27AE60", linestyle="--",
                                   label="95% Sensitivity target"))
        ax.legend(handles=legend_items, loc="lower left", fontsize=8,
                  framealpha=0.9)
        plt.tight_layout()
        fig.savefig(str(out_dir / f"efficiency_frontier_averaged.{FIG_FORMAT}"),
                    dpi=DPI, bbox_inches="tight", facecolor=FACECOLOR, edgecolor="none")
        plt.close(fig)
        print(f"    ✓ efficiency_frontier_averaged.{FIG_FORMAT}")
        generated += 1


# ──────────────────────────────────────────────────────────────────────
#  CHART 10 — Efficiency Score per Model per Project
# ──────────────────────────────────────────────────────────────────────

if "eff_score_by_project" in sheets:
    df = sheets["eff_score_by_project"]
    if not df.empty:
        projects = df["Project"].unique()
        n_projs = len(projects)
        fig, axes = plt.subplots(1, n_projs, figsize=(5 * n_projs, 5),
                                 sharey=True, squeeze=False)
        fig.suptitle("Efficiency Score by Model per Project (Individual Runs)",
                     fontsize=12, fontweight="bold", y=1.02)

        for ax_i, proj in enumerate(projects):
            ax = axes[0][ax_i]
            sub = df[df["Project"] == proj].sort_values(["Model", "Test"])
            labels = [f"{r['Model']}\nT{int(r['Test'])}"
                      for _, r in sub.iterrows()]
            scores = sub["Efficiency_Score"].values
            colors = []
            for i, (_, r) in enumerate(sub.iterrows()):
                key = r["Model"].strip().lower().replace(" ", "_").replace("-", "_")
                colors.append(MODEL_COLORS.get(
                    key, DEFAULT_COLORS[i % len(DEFAULT_COLORS)]))

            x = np.arange(len(labels))
            bars = ax.bar(x, scores, width=0.7, color=colors, edgecolor="#333",
                          linewidth=0.5, alpha=0.85)
            for bar, val in zip(bars, scores):
                ax.text(bar.get_x() + bar.get_width() / 2,
                        bar.get_height() + 0.005,
                        f"{val:.3f}", ha="center", va="bottom",
                        fontsize=7, fontweight="bold")
            ax.set_title(proj, fontsize=10, fontweight="bold", pad=8)
            if ax_i == 0:
                ax.set_ylabel("Efficiency Score", fontsize=9)
            ax.spines["top"].set_visible(False)
            ax.spines["right"].set_visible(False)
            ax.spines["left"].set_linewidth(0.6)
            ax.spines["bottom"].set_linewidth(0.6)
            ax.set_ylim(0, 1.0)
            ax.set_xticks(x)
            ax.set_xticklabels(labels, fontsize=6.5, rotation=25, ha="right")

        plt.tight_layout(rect=[0, 0, 1, 0.96])
        fig.savefig(str(out_dir / f"efficiency_score_by_project.{FIG_FORMAT}"),
                    dpi=DPI, bbox_inches="tight", facecolor=FACECOLOR,
                    edgecolor="none")
        plt.close(fig)
        print(f"    ✓ efficiency_score_by_project.{FIG_FORMAT}")
        generated += 1
else:
    print("    ⚠ Sheet 'eff_score_by_project' not found — skipping")


# ──────────────────────────────────────────────────────────────────────
#  CHART 11 — Efficiency Score Aggregated
# ──────────────────────────────────────────────────────────────────────

if "eff_score_aggregated" in sheets:
    df = sheets["eff_score_aggregated"]
    if not df.empty:
        df_sorted = df.sort_values("Mean_Efficiency_Score", ascending=False)
        model_labels = df_sorted["Model"].values
        means = df_sorted["Mean_Efficiency_Score"].values
        stds = df_sorted["SD"].values
        colors = []
        for i, m in enumerate(model_labels):
            key = m.strip().lower().replace(" ", "_").replace("-", "_")
            colors.append(MODEL_COLORS.get(
                key, DEFAULT_COLORS[i % len(DEFAULT_COLORS)]))
        x = np.arange(len(model_labels))

        fig, ax = plt.subplots(figsize=(8, 5))
        bars = ax.bar(x, means, width=0.6, color=colors, edgecolor="#333",
                      linewidth=0.5, alpha=0.85, yerr=stds, capsize=5,
                      error_kw={"linewidth": 1, "color": "#555"})
        for bar, val in zip(bars, means):
            ax.text(bar.get_x() + bar.get_width() / 2,
                    bar.get_height() + 0.008,
                    f"{val:.3f}", ha="center", va="bottom",
                    fontsize=9, fontweight="bold")
        ax.set_title("Efficiency Score by Model (Mean ± SD, All Projects & Runs)",
                     fontsize=10, fontweight="bold", pad=8)
        ax.set_ylabel("Efficiency Score", fontsize=9)
        ax.spines["top"].set_visible(False)
        ax.spines["right"].set_visible(False)
        ax.spines["left"].set_linewidth(0.6)
        ax.spines["bottom"].set_linewidth(0.6)
        ax.set_ylim(0, max(means) * 1.25 + 0.05 if len(means) else 1)
        ax.set_xticks(x)
        ax.set_xticklabels(model_labels, fontsize=10, fontweight="bold")
        for i in range(len(model_labels)):
            ax.text(i, 0.01, f"#{i+1}", ha="center", fontsize=8,
                    fontweight="bold", color="white", alpha=0.9)
        plt.tight_layout()
        fig.savefig(str(out_dir / f"efficiency_score_aggregated.{FIG_FORMAT}"),
                    dpi=DPI, bbox_inches="tight", facecolor=FACECOLOR,
                    edgecolor="none")
        plt.close(fig)
        print(f"    ✓ efficiency_score_aggregated.{FIG_FORMAT}")
        generated += 1
else:
    print("    ⚠ Sheet 'eff_score_aggregated' not found — skipping")


# ──────────────────────────────────────────────────────────────────────
#  CHART 12 — Sensitivity & Specificity Dual Gold Standard
# ──────────────────────────────────────────────────────────────────────

if "sens_spec_dual_gold" in sheets:
    df = sheets["sens_spec_dual_gold"]
    if not df.empty:
        model_names = df["Model"].values
        n_m = len(model_names)
        x = np.arange(n_m)
        width = 0.2

        fig, ax = plt.subplots(figsize=(10, 5))
        vals_st = df["Sens_TIAB_pct"].fillna(0).values
        vals_sl = df["Sens_LF_pct"].fillna(0).values
        vals_spt = df["Spec_TIAB_pct"].fillna(0).values
        vals_spl = df["Spec_LF_pct"].fillna(0).values

        b1 = ax.bar(x - 1.5 * width, vals_st, width, label="Sens (TIAB)",
                    color="#4472C4", alpha=0.85, edgecolor="#333", linewidth=0.4)
        b2 = ax.bar(x - 0.5 * width, vals_sl, width, label="Sens (Listfinal)",
                    color="#27AE60", alpha=0.85, edgecolor="#333", linewidth=0.4)
        b3 = ax.bar(x + 0.5 * width, vals_spt, width, label="Spec (TIAB)",
                    color="#ED7D31", alpha=0.85, edgecolor="#333", linewidth=0.4)
        b4 = ax.bar(x + 1.5 * width, vals_spl, width, label="Spec (Listfinal)",
                    color="#FFC000", alpha=0.85, edgecolor="#333", linewidth=0.4)

        ax.axhline(y=95, color="#27AE60", linestyle="--", alpha=0.4,
                   linewidth=0.8)
        ax.set_title("Sensitivity & Specificity: TIAB vs Listfinal Gold Standard",
                     fontsize=10, fontweight="bold", pad=8)
        ax.set_ylabel("Percentage (%)", fontsize=9)
        ax.spines["top"].set_visible(False)
        ax.spines["right"].set_visible(False)
        ax.spines["left"].set_linewidth(0.6)
        ax.spines["bottom"].set_linewidth(0.6)
        ax.set_ylim(0, 110)
        ax.set_xticks(x)
        ax.set_xticklabels(model_names, fontsize=8, fontweight="bold")
        ax.legend(loc="upper center", bbox_to_anchor=(0.5, -0.08), ncol=4,
                  fontsize=8, framealpha=0.9)
        for bars in [b1, b2, b3, b4]:
            for bar in bars:
                h = bar.get_height()
                if h > 0:
                    ax.text(bar.get_x() + bar.get_width() / 2, h + 0.5,
                            f"{h:.0f}", ha="center", va="bottom", fontsize=6)
        plt.tight_layout()
        fig.savefig(str(out_dir / f"sensitivity_specificity_dual_gold.{FIG_FORMAT}"),
                    dpi=DPI, bbox_inches="tight", facecolor=FACECOLOR,
                    edgecolor="none")
        plt.close(fig)
        print(f"    ✓ sensitivity_specificity_dual_gold.{FIG_FORMAT}")
        generated += 1
else:
    print("    ⚠ Sheet 'sens_spec_dual_gold' not found — skipping")


# ──────────────────────────────────────────────────────────────────────
#  CHART 13 — Aggregated Performance Metrics (2×3 subplots)
# ──────────────────────────────────────────────────────────────────────

if "aggregated_performance" in sheets:
    df = sheets["aggregated_performance"]
    if not df.empty:
        metric_cols = [c for c in df.columns if c.endswith("_mean")]
        metric_labels = [c.replace("_mean", "") for c in metric_cols]
        model_labels = df["Model"].values
        n_models = len(model_labels)
        x = np.arange(n_models)

        nplots = len(metric_cols)
        ncols_p = min(nplots, 3)
        nrows_p = (nplots + ncols_p - 1) // ncols_p

        fig, axes = plt.subplots(nrows_p, ncols_p, figsize=(14, 8))
        if nrows_p == 1:
            axes = np.array([axes])
        fig.suptitle("Aggregated Model Performance (Mean ± SD across Projects)",
                     fontsize=12, fontweight="bold")

        for idx, (mc, label) in enumerate(zip(metric_cols, metric_labels)):
            ax = (axes[idx // ncols_p, idx % ncols_p] if nrows_p > 1
                  else axes[0, idx % ncols_p])
            sd_col = mc.replace("_mean", "_sd")
            means = df[mc].fillna(0).values
            sds = (df[sd_col].fillna(0).values if sd_col in df.columns
                   else np.zeros(n_models))
            ax.bar(x, means, yerr=sds, capsize=4, color="#5B9BD5", alpha=0.7,
                   edgecolor="#333", linewidth=0.5, width=0.6)
            ax.set_xticks(x)
            ax.set_xticklabels(model_labels, rotation=30, ha="right",
                               fontsize=7)
            ax.set_title(label, fontsize=9)
            ax.set_ylim(0, 1.08)
            ax.axhline(y=0.95, color="green", linestyle="--", alpha=0.4,
                       linewidth=0.8)
            ax.grid(axis="y", alpha=0.3)

        for idx in range(nplots, nrows_p * ncols_p):
            axes[idx // ncols_p, idx % ncols_p].set_visible(False)

        plt.tight_layout(rect=[0, 0, 1, 0.95])
        fig.savefig(str(out_dir / f"aggregated_performance_metrics.{FIG_FORMAT}"),
                    dpi=DPI, bbox_inches="tight", facecolor=FACECOLOR,
                    edgecolor="none")
        plt.close(fig)
        print(f"    ✓ aggregated_performance_metrics.{FIG_FORMAT}")
        generated += 1
else:
    print("    ⚠ Sheet 'aggregated_performance' not found — skipping")


# ──────────────────────────────────────────────────────────────────────
#  CHART 14 — F1 Score vs Cost
# ──────────────────────────────────────────────────────────────────────

if "f1_vs_cost" in sheets:
    df = sheets["f1_vs_cost"]
    if not df.empty:
        fig, ax = plt.subplots(figsize=(8, 5))
        ax.scatter(df["Avg_Cost_USD"], df["Avg_F1_LF"], s=100,
                   color="#5B9BD5", edgecolors="#333", linewidth=1, zorder=5)
        for _, row in df.iterrows():
            ax.annotate(row["Model"],
                        (row["Avg_Cost_USD"], row["Avg_F1_LF"]),
                        textcoords="offset points", xytext=(8, 8),
                        fontsize=8, fontweight="bold",
                        arrowprops=dict(arrowstyle="-", color="gray", lw=0.5))
        ax.set_xlabel("Average Cost (USD)", fontsize=10)
        ax.set_ylabel("Average F1 Score (vs Listfinal)", fontsize=10)
        ax.set_title("F1 Score (Full-Text Gold Standard) vs Cost per Model",
                     fontsize=11, fontweight="bold")
        ax.grid(True, alpha=0.3)
        ax.set_ylim(0, 1.05)
        ax.axhline(y=0.95, color="green", linestyle="--", alpha=0.5,
                   linewidth=0.8, label="F1 = 0.95")
        ax.legend(loc="lower right", fontsize=8)
        plt.tight_layout()
        fig.savefig(str(out_dir / f"f1_score_vs_cost.{FIG_FORMAT}"),
                    dpi=DPI, bbox_inches="tight", facecolor=FACECOLOR,
                    edgecolor="none")
        plt.close(fig)
        print(f"    ✓ f1_score_vs_cost.{FIG_FORMAT}")
        generated += 1
else:
    print("    ⚠ Sheet 'f1_vs_cost' not found — skipping")


# ──────────────────────────────────────────────────────────────────────
#  CHART 15 — Sensitivity vs Specificity Trade-Off (2 panels)
# ──────────────────────────────────────────────────────────────────────

if "sens_spec_tradeoff" in sheets:
    df = sheets["sens_spec_tradeoff"]
    if not df.empty:
        models = df["Model"].unique()
        projects = df["Project"].unique()
        proj_markers = {"mino": "o", "NMDA": "s", "zebra": "D"}
        default_markers_list = ["o", "s", "D", "^", "v", "P"]

        fig, (ax_tiab, ax_lf) = plt.subplots(1, 2, figsize=(12, 5.5),
                                              sharey=True)

        tiab = df[df["Gold_Standard"] == "TIAB"]
        lf = df[df["Gold_Standard"] == "Listfinal"]

        for mi, model in enumerate(models):
            key = model.strip().lower().replace(" ", "_").replace("-", "_")
            color = MODEL_COLORS.get(
                key, DEFAULT_COLORS[mi % len(DEFAULT_COLORS)])
            for pi, proj in enumerate(projects):
                marker = proj_markers.get(
                    proj,
                    default_markers_list[pi % len(default_markers_list)])
                sub_t = tiab[(tiab["Model"] == model)
                             & (tiab["Project"] == proj)]
                for _, row in sub_t.iterrows():
                    ax_tiab.scatter(row["Specificity_pct"],
                                    row["Sensitivity_pct"],
                                    c=color, marker=marker, s=60, alpha=0.8,
                                    edgecolors="#333", linewidths=0.5)
                sub_l = lf[(lf["Model"] == model)
                           & (lf["Project"] == proj)]
                for _, row in sub_l.iterrows():
                    ax_lf.scatter(row["Specificity_pct"],
                                  row["Sensitivity_pct"],
                                  c=color, marker=marker, s=60, alpha=0.8,
                                  edgecolors="#333", linewidths=0.5)

        for ax, title in [
            (ax_tiab, "Gold Standard: TIAB (Human Screening)"),
            (ax_lf, "Gold Standard: Listfinal (True Outcome)"),
        ]:
            ax.set_xlabel("Specificity (%)", fontsize=10)
            ax.set_title(title, fontsize=11, fontweight="bold")
            ax.axhline(y=95, color="#27AE60", linestyle="--", alpha=0.5,
                       linewidth=0.8)
            ax.axvline(x=50, color="#E74C3C", linestyle=":", alpha=0.4,
                       linewidth=0.8)
            ax.set_xlim(0, 105)
            ax.set_ylim(0, 105)
            ax.spines["top"].set_visible(False)
            ax.spines["right"].set_visible(False)
            ax.grid(True, alpha=0.3)
        ax_tiab.set_ylabel("Sensitivity (%)", fontsize=10)

        legend_handles = []
        for i, m in enumerate(models):
            key = m.strip().lower().replace(" ", "_").replace("-", "_")
            legend_handles.append(
                Line2D([0], [0], marker="o", color="w",
                       markerfacecolor=MODEL_COLORS.get(
                           key, DEFAULT_COLORS[i % len(DEFAULT_COLORS)]),
                       markersize=8, markeredgecolor="#333", label=m))
        for proj in projects:
            marker = proj_markers.get(proj, "o")
            legend_handles.append(
                Line2D([0], [0], marker=marker, color="w",
                       markerfacecolor="gray", markersize=8,
                       markeredgecolor="#333", label=f"Project: {proj}"))
        fig.legend(handles=legend_handles, loc="lower center",
                   ncol=len(models) + len(projects), fontsize=7,
                   framealpha=0.9, bbox_to_anchor=(0.5, -0.02))
        plt.tight_layout(rect=[0, 0.06, 1, 1])
        fig.savefig(str(out_dir / f"sensitivity_specificity_tradeoff.{FIG_FORMAT}"),
                    dpi=DPI, bbox_inches="tight", facecolor=FACECOLOR,
                    edgecolor="none")
        plt.close(fig)
        print(f"    ✓ sensitivity_specificity_tradeoff.{FIG_FORMAT}")
        generated += 1
else:
    print("    ⚠ Sheet 'sens_spec_tradeoff' not found — skipping")


# ──────────────────────────────────────────────────────────────────────
#  CHART 16 — Model Ranking Heatmap
# ──────────────────────────────────────────────────────────────────────

if "model_ranking_heatmap" in sheets:
    df = sheets["model_ranking_heatmap"]
    if not df.empty:
        metric_cols = [c for c in df.columns
                       if c not in ("Model", "Overall_Score")]
        n_models = len(df)
        n_metrics = len(metric_cols)

        matrix = df[metric_cols].values.astype(float)

        norm_matrix = np.full_like(matrix, np.nan)
        for j in range(n_metrics):
            col = matrix[:, j]
            valid = ~np.isnan(col)
            if valid.sum() < 2:
                norm_matrix[valid, j] = 0.5
                continue
            vmin, vmax = np.nanmin(col), np.nanmax(col)
            if vmax - vmin < 1e-9:
                norm_matrix[valid, j] = 0.5
            else:
                norm_matrix[:, j] = (col - vmin) / (vmax - vmin)

        fig, ax = plt.subplots(
            figsize=(10, max(3.5, n_models * 0.9)))
        cmap = plt.cm.RdYlGn

        for i in range(n_models):
            for j in range(n_metrics):
                val = matrix[i, j]
                norm_val = (norm_matrix[i, j]
                            if not np.isnan(norm_matrix[i, j]) else 0.5)
                color = cmap(norm_val)
                ax.add_patch(plt.Rectangle((j, i), 1, 1,
                                           facecolor=color,
                                           edgecolor="white", linewidth=2))
                if not np.isnan(val):
                    txt = (f"{val:.1f}%" if val > 1.5
                           else f"{val:.3f}")
                    text_color = ("white"
                                  if norm_val < 0.3 or norm_val > 0.85
                                  else "black")
                    ax.text(j + 0.5, i + 0.5, txt, ha="center",
                            va="center", fontsize=9, fontweight="bold",
                            color=text_color)

        ax.set_xlim(0, n_metrics)
        ax.set_ylim(0, n_models)
        ax.set_xticks([xi + 0.5 for xi in range(n_metrics)])
        ax.set_xticklabels(metric_cols, fontsize=8, ha="center")
        ax.set_yticks([yi + 0.5 for yi in range(n_models)])
        ax.set_yticklabels(
            [f"#{i+1} {m}" for i, m in enumerate(df["Model"].values)],
            fontsize=9, fontweight="bold")
        ax.xaxis.tick_top()
        ax.invert_yaxis()
        ax.set_title("Model Ranking Summary (Mean Across All Projects & Runs)",
                     fontsize=12, fontweight="bold", pad=30)
        for spine in ax.spines.values():
            spine.set_visible(False)
        ax.tick_params(length=0)

        scores = (df["Overall_Score"].values if "Overall_Score" in df.columns
                  else np.nanmean(norm_matrix, axis=1))
        for i, score in enumerate(scores):
            ax.text(n_metrics + 0.15, i + 0.5, f"{score:.2f}",
                    ha="left", va="center", fontsize=8, fontstyle="italic",
                    color="#555")
        ax.text(n_metrics + 0.15, -0.3, "Score", ha="left", va="center",
                fontsize=8, fontweight="bold", color="#555")

        plt.tight_layout()
        fig.savefig(str(out_dir / f"model_ranking_heatmap.{FIG_FORMAT}"),
                    dpi=DPI, bbox_inches="tight", facecolor=FACECOLOR,
                    edgecolor="none")
        plt.close(fig)
        print(f"    ✓ model_ranking_heatmap.{FIG_FORMAT}")
        generated += 1
else:
    print("    ⚠ Sheet 'model_ranking_heatmap' not found — skipping")


# ──────────────────────────────────────────────────────────────────────
#  CHART 17 — Fulltext Hours Saved (per Model, per Project)
# ──────────────────────────────────────────────────────────────────────

if "fulltext_hours_saved" in sheets:
    df = sheets["fulltext_hours_saved"]
    if not df.empty:
        # Aggregate per (Project, Model): average across tests
        agg = (df.groupby(["Project", "Model"], as_index=False)
                 .agg(Hours_Saved=("Hours_Saved_Fulltext", "mean"),
                      Articles_Saved=("Articles_Saved", "mean"),
                      Human_Inc=("Human_TIAB_Included", "mean"),
                      AI_Inc=("AI_TIAB_Included", "mean")))
        projects = sorted(agg["Project"].unique())
        models = sorted(agg["Model"].unique())

        n_projs = len(projects)
        ncols_p = min(n_projs, 3)
        nrows_p = (n_projs + ncols_p - 1) // ncols_p

        fig, axes = plt.subplots(nrows_p, ncols_p,
                                  figsize=(5.5 * ncols_p, 4.5 * nrows_p),
                                  squeeze=False)
        fig.suptitle(
            "Full-Text Reading Hours Saved per Model "
            "(2 reviewers × 5 min per article, averaged across tests)",
            fontsize=12, fontweight="bold", y=1.00)

        # Stable color per model
        model_colors = {}
        for mi, m in enumerate(models):
            key = m.strip().lower().replace(" ", "_").replace("-", "_")
            model_colors[m] = MODEL_COLORS.get(
                key, DEFAULT_COLORS[mi % len(DEFAULT_COLORS)])

        for idx, proj in enumerate(projects):
            r, c = divmod(idx, ncols_p)
            ax = axes[r][c]
            sub = agg[agg["Project"] == proj].sort_values("Hours_Saved", ascending=False)
            xs = np.arange(len(sub))
            hours = sub["Hours_Saved"].values
            colors = [model_colors[m] for m in sub["Model"].values]

            bars = ax.bar(xs, hours, color=colors, edgecolor="#333",
                          linewidth=0.5, alpha=0.85, width=0.7)

            for bar, h_val, arts in zip(bars, hours, sub["Articles_Saved"].values):
                if h_val >= 0:
                    label_y = bar.get_height() + max(abs(hours).max() * 0.02, 0.3)
                    va = "bottom"
                else:
                    label_y = bar.get_height() - max(abs(hours).max() * 0.02, 0.3)
                    va = "top"
                ax.text(bar.get_x() + bar.get_width() / 2, label_y,
                        f"{h_val:+.1f}h\n({int(arts):+d} arts)",
                        ha="center", va=va, fontsize=7, fontweight="bold")

            human_inc = sub["Human_Inc"].iloc[0] if not sub.empty else 0
            ax.set_title(f"{proj}  (human included {int(human_inc)} at TIAB)",
                         fontsize=10, fontweight="bold", pad=8)
            ax.set_xticks(xs)
            ax.set_xticklabels(sub["Model"].values, fontsize=8, rotation=20, ha="right")
            ax.axhline(y=0, color="#555", linewidth=0.8)
            if c == 0:
                ax.set_ylabel("Hours saved at full-text phase", fontsize=9)
            ax.spines["top"].set_visible(False)
            ax.spines["right"].set_visible(False)
            ax.spines["left"].set_linewidth(0.6)
            ax.spines["bottom"].set_linewidth(0.6)
            ax.grid(axis="y", alpha=0.25)
            if len(hours):
                margin = max(abs(hours).max() * 0.25, 1.0)
                ax.set_ylim(min(hours.min(), 0) - margin, hours.max() + margin)

        for idx in range(n_projs, nrows_p * ncols_p):
            r2, c2 = divmod(idx, ncols_p)
            axes[r2][c2].set_visible(False)

        plt.tight_layout(rect=[0, 0, 1, 0.96])
        fig.savefig(str(out_dir / f"fulltext_hours_saved_by_project.{FIG_FORMAT}"),
                    dpi=DPI, bbox_inches="tight", facecolor=FACECOLOR, edgecolor="none")
        plt.close(fig)
        print(f"    ✓ fulltext_hours_saved_by_project.{FIG_FORMAT}")
        generated += 1

        # ── Aggregated view: total hours saved per model across all projects ──
        agg_model = (df.groupby("Model", as_index=False)
                       .agg(Total_Hours=("Hours_Saved_Fulltext", "sum"),
                            Mean_Hours=("Hours_Saved_Fulltext", "mean"),
                            Total_Articles=("Articles_Saved", "sum")))
        agg_model = agg_model.sort_values("Total_Hours", ascending=False)

        fig2, ax2 = plt.subplots(figsize=(9, 5.5))
        x2 = np.arange(len(agg_model))
        colors2 = [model_colors.get(m, "#5B9BD5") for m in agg_model["Model"].values]
        bars2 = ax2.bar(x2, agg_model["Total_Hours"].values,
                        color=colors2, edgecolor="#333", linewidth=0.5,
                        alpha=0.85, width=0.6)
        for bar, h_val, arts in zip(bars2, agg_model["Total_Hours"].values,
                                     agg_model["Total_Articles"].values):
            offset = max(abs(agg_model["Total_Hours"]).max() * 0.02, 0.3)
            if h_val >= 0:
                ax2.text(bar.get_x() + bar.get_width() / 2, h_val + offset,
                         f"{h_val:+.1f}h\n({int(arts):+d} arts)",
                         ha="center", va="bottom", fontsize=9, fontweight="bold")
            else:
                ax2.text(bar.get_x() + bar.get_width() / 2, h_val - offset,
                         f"{h_val:+.1f}h\n({int(arts):+d} arts)",
                         ha="center", va="top", fontsize=9, fontweight="bold")

        ax2.axhline(y=0, color="#555", linewidth=0.8)
        ax2.set_title("Total Full-Text Hours Saved per Model "
                       "(sum across all projects & tests)",
                       fontsize=11, fontweight="bold", pad=8)
        ax2.set_ylabel("Hours saved at full-text phase", fontsize=10)
        ax2.set_xticks(x2)
        ax2.set_xticklabels(agg_model["Model"].values, fontsize=9, fontweight="bold")
        ax2.spines["top"].set_visible(False)
        ax2.spines["right"].set_visible(False)
        ax2.spines["left"].set_linewidth(0.6)
        ax2.spines["bottom"].set_linewidth(0.6)
        ax2.grid(axis="y", alpha=0.25)

        plt.tight_layout()
        fig2.savefig(str(out_dir / f"fulltext_hours_saved_total.{FIG_FORMAT}"),
                     dpi=DPI, bbox_inches="tight", facecolor=FACECOLOR, edgecolor="none")
        plt.close(fig2)
        print(f"    ✓ fulltext_hours_saved_total.{FIG_FORMAT}")
        generated += 1
else:
    print("    ⚠ Sheet 'fulltext_hours_saved' not found — skipping")


# ──────────────────────────────────────────────────────────────────────
#  PAPER FIGURE 2 — Sensitivity × Specificity (vs Listfinal, official-only)
# ──────────────────────────────────────────────────────────────────────

if "paper_sens_spec_lf" in sheets:
    df = sheets["paper_sens_spec_lf"]
    if not df.empty:
        fig, ax = plt.subplots(figsize=(7.5, 6))
        ax.set_facecolor("#FAFAFA")
        ax.axhspan(95, 106, xmin=0.5, xmax=1, color="#D5F5E3", alpha=0.22, zorder=0)

        model_labels = df["Model"].tolist()
        for mi, (_, row) in enumerate(df.iterrows()):
            key = row["Model"].strip().lower().replace(" ", "_").replace("-", "_")
            color = MODEL_COLORS.get(key, DEFAULT_COLORS[mi % len(DEFAULT_COLORS)])
            x_mean, x_sd = row["Sens_LF_pct_mean"], row.get("Sens_LF_pct_sd", 0) or 0
            y_mean, y_sd = row["Spec_LF_pct_mean"], row.get("Spec_LF_pct_sd", 0) or 0
            if np.isnan(x_mean) or np.isnan(y_mean):
                continue
            ax.errorbar(x_mean, y_mean, xerr=x_sd, yerr=y_sd, fmt="none",
                        ecolor=color, elinewidth=1.2, capsize=4, alpha=0.55, zorder=4)
            ax.scatter(x_mean, y_mean, s=220, color=color, edgecolors="white",
                        linewidth=2, zorder=6)
            ax.annotate(row["Model"], (x_mean, y_mean),
                        xytext=(8, 8), textcoords="offset points",
                        fontsize=9, fontweight="bold")

        ax.axhline(y=95, color="#27AE60", linestyle="--", alpha=0.55, linewidth=1)
        ax.annotate("IDEAL", xy=(97, 97), fontsize=9, fontweight="bold",
                    color="#27AE60", alpha=0.6, ha="right")

        ax.set_title("Figure 2 — Sensitivity vs Specificity against Listfinal\n"
                      "(mean ± SD across the 3 official projects)",
                      fontsize=11, fontweight="bold", pad=10)
        ax.set_xlabel("Sensitivity vs Listfinal (%)", fontsize=10)
        ax.set_ylabel("Specificity vs Listfinal (%)", fontsize=10)
        ax.set_xlim(0, 105)
        ax.set_ylim(0, 105)
        ax.spines["top"].set_visible(False)
        ax.spines["right"].set_visible(False)
        ax.grid(True, alpha=0.28)
        plt.tight_layout()
        fig.savefig(str(out_dir / f"paper_fig2_sens_spec_lf.{FIG_FORMAT}"),
                    dpi=DPI, bbox_inches="tight", facecolor=FACECOLOR, edgecolor="none")
        plt.close(fig)
        print(f"    ✓ paper_fig2_sens_spec_lf.{FIG_FORMAT}")
        generated += 1
else:
    print("    ⚠ Sheet 'paper_sens_spec_lf' not found — skipping Paper Fig. 2")


# ──────────────────────────────────────────────────────────────────────
#  PAPER FIGURE 3 — F1 (vs Listfinal) × Cost, official-only
# ──────────────────────────────────────────────────────────────────────

if "paper_f1_vs_cost_official" in sheets:
    df = sheets["paper_f1_vs_cost_official"]
    if not df.empty:
        fig, ax = plt.subplots(figsize=(7.5, 5.5))
        ax.set_facecolor("#FAFAFA")

        for mi, (_, row) in enumerate(df.iterrows()):
            key = row["Model"].strip().lower().replace(" ", "_").replace("-", "_")
            color = MODEL_COLORS.get(key, DEFAULT_COLORS[mi % len(DEFAULT_COLORS)])
            x, x_sd = row["Avg_Cost_USD"], row.get("Avg_Cost_USD_sd", 0) or 0
            y, y_sd = row["Avg_F1_LF"],   row.get("Avg_F1_LF_sd", 0) or 0
            ax.errorbar(x, y, xerr=x_sd, yerr=y_sd, fmt="none",
                        ecolor=color, elinewidth=1.2, capsize=4, alpha=0.55, zorder=4)
            ax.scatter(x, y, s=220, color=color, edgecolors="white",
                        linewidth=2, zorder=6)
            ax.annotate(row["Model"], (x, y),
                        xytext=(8, 8), textcoords="offset points",
                        fontsize=9, fontweight="bold")

        ax.axhline(y=0.95, color="#27AE60", linestyle="--", alpha=0.55,
                    linewidth=1, label="F1 = 0.95")
        ax.set_title("Figure 3 — F1 (vs Listfinal) vs Cost per Model\n"
                      "(mean ± SD across the 3 official projects)",
                      fontsize=11, fontweight="bold", pad=10)
        ax.set_xlabel("Average Cost per run (USD)", fontsize=10)
        ax.set_ylabel("F1 vs Listfinal", fontsize=10)
        ax.set_ylim(0, 1.05)
        ax.spines["top"].set_visible(False)
        ax.spines["right"].set_visible(False)
        ax.grid(True, alpha=0.28)
        ax.legend(loc="lower right", fontsize=8, framealpha=0.9)
        plt.tight_layout()
        fig.savefig(str(out_dir / f"paper_fig3_f1_vs_cost.{FIG_FORMAT}"),
                    dpi=DPI, bbox_inches="tight", facecolor=FACECOLOR, edgecolor="none")
        plt.close(fig)
        print(f"    ✓ paper_fig3_f1_vs_cost.{FIG_FORMAT}")
        generated += 1
else:
    print("    ⚠ Sheet 'paper_f1_vs_cost_official' not found — skipping Paper Fig. 3")


# ──────────────────────────────────────────────────────────────────────
#  SUMMARY
# ──────────────────────────────────────────────────────────────────────

print(f"\n  Done — {generated} charts generated in {out_dir}")
