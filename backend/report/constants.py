"""
constants.py — Paths, colors and shared configuration.
"""

from pathlib import Path

# ── Paths ────────────────────────────────────────────────────────────
SCRIPT_DIR  = Path(__file__).resolve().parent          # backend/report/
PROJECT_DIR = SCRIPT_DIR.parent.parent                 # AI-tools/
INPUT_DIR   = PROJECT_DIR / "input"
OUTPUT_DIR  = PROJECT_DIR / "output"

# ── Model colours (shared by chart_data + graphic scripts) ───────────
MODEL_COLORS = {
    "gpt-4o":     "#4472C4",
    "gpt-5_2":    "#ED7D31",
    "gpt_5_2":    "#ED7D31",
    "gpt-5-mini": "#A5A5A5",
    "gpt_5_mini": "#A5A5A5",
    "gpt-5-nano": "#FFC000",
    "gpt_5_nano": "#FFC000",
    "gpt_4o":     "#4472C4",
}

DEFAULT_COLORS = [
    "#5B9BD5", "#ED7D31", "#A5A5A5", "#FFC000", "#70AD47",
    "#9B59B6", "#E74C3C", "#1ABC9C", "#34495E", "#F39C12",
]

# ── Chart export settings ────────────────────────────────────────────
DPI        = 300
FIG_FORMAT = "png"
FACECOLOR  = "white"


# ── Paper-figure project categorisation ──────────────────────────────
# Used by the Paper Figures & Tables section of the general report and
# by the paper-only chart variants.  Match by `project_norm` — the
# lowercase alphanumeric project key derived from the file name.
#
# * PILOT_PROJECTS       — used for prompt optimisation, NOT part of the
#                          primary comparison.  Shown with a grey background
#                          to make it clear it does not count.
# * OFFICIAL_PROJECTS    — the 3 primary projects that DO count for
#                          Paper Figures 2 and 3.
# * SENSITIVITY_PROJECTS — sensitivity/robustness analyses, NOT part of the
#                          primary comparison (also greyed out).
#
# Any project not listed here defaults to "official" so the report still
# runs on new datasets.  Edit the lists below to reflect your data.

PILOT_PROJECTS = ["mino", "minociclina"]
OFFICIAL_PROJECTS = ["nmda", "zebrafish", "hfd"]
SENSITIVITY_PROJECTS = ["hfd_modified", "hfdmodified", "hfdmod"]


VALID_CATEGORIES = ("pilot", "official", "sensitivity")


def project_category(project_norm: str, overrides: dict | None = None) -> str:
    """Return 'pilot', 'official' or 'sensitivity' for a `project_norm` key.

    Resolution order:
      1. `overrides` (usually from the web UI, keyed by `project_norm`)
      2. fallback lists in this module
      3. default: 'official'
    """
    key = str(project_norm).strip().lower()
    if overrides:
        v = overrides.get(project_norm) or overrides.get(key)
        if isinstance(v, str):
            v = v.strip().lower()
            if v in VALID_CATEGORIES:
                return v
    if key in {p.lower() for p in PILOT_PROJECTS}:
        return "pilot"
    if key in {p.lower() for p in SENSITIVITY_PROJECTS}:
        return "sensitivity"
    return "official"


def project_sort_key(project_norm: str, overrides: dict | None = None) -> tuple:
    """Sort projects so pilots come first, then officials, then sensitivity.

    `overrides` follows the same contract as `project_category`.
    """
    cat = project_category(project_norm, overrides=overrides)
    order = {"pilot": 0, "official": 1, "sensitivity": 2}
    return (order.get(cat, 1), str(project_norm).lower())
