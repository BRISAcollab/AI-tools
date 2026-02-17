"""
ETAPA 1 - Pareamento de planilhas (IA vs Humano).

Faz o merge dos dois arquivos por titulo, identifica quais TIABs
nao foram pareados e salva o resultado para revisao antes da analise.

Uso:
    1. Coloque os 2 arquivos na pasta  input/
    2. Rode:  python 01_pareamento.py
    3. Revise os resultados em  output/

    Modo manual:
        python 01_pareamento.py --ai input/ia.xlsx --human input/humano.xlsx
"""

import argparse
import sys
import os
import datetime
from pathlib import Path

import pandas as pd
import numpy as np

# ------------------------------------------------------------------ paths --
SCRIPT_DIR  = Path(__file__).resolve().parent
PROJECT_DIR = SCRIPT_DIR.parent
INPUT_DIR   = PROJECT_DIR / "input"
OUTPUT_DIR  = PROJECT_DIR / "output"


# ---------------------------------------------------------------- helpers --

def ensure_folders():
    INPUT_DIR.mkdir(exist_ok=True)
    OUTPUT_DIR.mkdir(exist_ok=True)


def load_file(path: str) -> pd.DataFrame:
    ext = Path(path).suffix.lower()
    if ext == ".csv":
        try:
            return pd.read_csv(path, encoding="utf-8")
        except UnicodeDecodeError:
            return pd.read_csv(path, encoding="latin-1")
    elif ext in (".xlsx", ".xls"):
        return pd.read_excel(path)
    else:
        raise ValueError(f"Formato nao suportado: {ext}. Use .csv, .xlsx ou .xls")


def normalise_columns(df: pd.DataFrame) -> pd.DataFrame:
    df.columns = [c.strip().lower().replace(" ", "_") for c in df.columns]
    return df


def normalise_title(s) -> str:
    if pd.isna(s):
        return ""
    return str(s).strip().lower()


def normalise_decision(s) -> str:
    """Padroniza variantes de decisao: included->include, excluded->exclude."""
    if pd.isna(s):
        return ""
    d = str(s).strip().lower()
    if d == "included":
        return "include"
    if d == "excluded":
        return "exclude"
    return d


def auto_detect_files():
    """Detecta automaticamente os dois arquivos em input/."""
    valid_ext = {".csv", ".xlsx", ".xls"}
    candidates = [f for f in INPUT_DIR.iterdir() if f.suffix.lower() in valid_ext]

    if len(candidates) == 0:
        print(f"\n  ERRO: Nenhum arquivo em '{INPUT_DIR}'.")
        print("  Coloque os arquivos da IA e do humano e rode novamente.")
        sys.exit(1)
    if len(candidates) == 1:
        print(f"\n  ERRO: Apenas 1 arquivo em '{INPUT_DIR}'. Sao necessarios 2.")
        sys.exit(1)
    if len(candidates) > 2:
        print(f"  AVISO: {len(candidates)} arquivos em '{INPUT_DIR}'.")
        print("  Tentando detectar automaticamente...\n")

    ai_file = None
    human_file = None

    # Collect metadata for each candidate file
    scored = []
    for f in candidates:
        try:
            df = normalise_columns(load_file(str(f)))
            cols = set(df.columns)
            n = len(df)
        except Exception:
            continue

        scored.append({
            "path": f,
            "n": n,
            "has_screening": "screening_decision" in cols,
            "has_decision": "decision" in cols,
            "has_id": "id" in cols,
            "has_incl": "inclusion_evaluation" in cols,
        })

    # AI file: has screening_decision (prefer with id/inclusion_evaluation, largest)
    ai_candidates = [s for s in scored if s["has_screening"]]
    ai_candidates.sort(key=lambda x: (x["has_id"] or x["has_incl"], x["n"]),
                       reverse=True)
    if ai_candidates:
        ai_file = ai_candidates[0]["path"]

    # Human file: has decision, NOT the AI file, prefer LARGEST (TIAB > fulltext)
    human_candidates = [s for s in scored
                        if s["path"] != ai_file and s["has_decision"]]
    human_candidates.sort(key=lambda x: x["n"], reverse=True)
    if human_candidates:
        human_file = human_candidates[0]["path"]

    # Fallback: assign remaining files
    unassigned = [s["path"] for s in scored
                  if s["path"] != ai_file and s["path"] != human_file]
    if ai_file is None and unassigned:
        ai_file = unassigned.pop(0)
    if human_file is None and unassigned:
        human_file = unassigned.pop(0)

    if ai_file is None or human_file is None:
        print("  ERRO: Deteccao automatica falhou.")
        print("  Use: python 01_pareamento.py --ai <arq_ia> --human <arq_humano>")
        sys.exit(1)

    return str(ai_file), str(human_file)


# ----------------------------------------------------------- merge logic --

def merge_files(ai_path: str, human_path: str, output_dir: Path):
    """Faz o merge por titulo e salva os resultados."""

    ai_df = normalise_columns(load_file(ai_path))
    hu_df = normalise_columns(load_file(human_path))

    # --- Detectar colunas de decisao ---
    ai_decision_col = None
    for c in ("screening_decision", "screening", "decision_ai", "ai_decision"):
        if c in ai_df.columns:
            ai_decision_col = c
            break
    if ai_decision_col is None:
        raise KeyError(f"Coluna de decisao da IA nao encontrada. Colunas: {list(ai_df.columns)}")

    hu_decision_col = None
    for c in ("decision", "decision_human", "human_decision", "screening_decision"):
        if c in hu_df.columns:
            hu_decision_col = c
            break
    if hu_decision_col is None:
        raise KeyError(f"Coluna de decisao humana nao encontrada. Colunas: {list(hu_df.columns)}")

    if "title" not in ai_df.columns:
        raise KeyError(f"Arquivo da IA sem coluna 'title'. Colunas: {list(ai_df.columns)}")
    if "title" not in hu_df.columns:
        raise KeyError(f"Arquivo humano sem coluna 'title'. Colunas: {list(hu_df.columns)}")

    # --- Normalizar titulo para matching ---
    ai_df["_title_key"] = ai_df["title"].apply(normalise_title)
    hu_df["_title_key"] = hu_df["title"].apply(normalise_title)

    # Contagem de ocorrencia para lidar com titulos duplicados
    # Ex: se "Study X" aparece 4x em ambos, cada copia recebe _occ 0,1,2,3
    ai_df["_occ"] = ai_df.groupby("_title_key").cumcount()
    hu_df["_occ"] = hu_df.groupby("_title_key").cumcount()

    # Chave composta: titulo + numero de ocorrencia
    ai_df["_merge_key"] = ai_df["_title_key"] + "__" + ai_df["_occ"].astype(str)
    hu_df["_merge_key"] = hu_df["_title_key"] + "__" + hu_df["_occ"].astype(str)

    # --- Merge ---
    merged = pd.merge(
        ai_df,
        hu_df[["_merge_key", hu_decision_col]],
        on="_merge_key",
        how="outer",
        indicator=True,
        suffixes=("_ai", "_human"),
    )

    ai_only    = merged[merged["_merge"] == "left_only"].copy()
    human_only = merged[merged["_merge"] == "right_only"].copy()
    matched    = merged[merged["_merge"] == "both"].copy()

    # Renomear colunas de decisao
    def find_col(df, candidates):
        for c in candidates:
            if c in df.columns:
                return c
        return None

    ai_col = find_col(matched, [ai_decision_col + "_ai", ai_decision_col])
    hu_col = find_col(matched, [hu_decision_col + "_human", hu_decision_col])
    if ai_col is None:
        ai_col = ai_decision_col
    if hu_col is None:
        hu_col = hu_decision_col

    matched = matched.rename(columns={ai_col: "screening_decision", hu_col: "decision_human"})

    # Padronizar variantes de decisao (included->include, excluded->exclude)
    matched["screening_decision"] = matched["screening_decision"].apply(normalise_decision)
    matched["decision_human"] = matched["decision_human"].apply(normalise_decision)

    # ================================================================
    #  RELATORIO NO CONSOLE
    # ================================================================
    n_matched     = len(matched)
    n_ai_only     = len(ai_only)
    n_human_only  = len(human_only)

    print()
    print("=" * 62)
    print("  ETAPA 1 - PAREAMENTO DE PLANILHAS")
    print("=" * 62)
    # Verificar duplicatas
    ai_dups = ai_df[ai_df.duplicated(subset="_title_key", keep=False)]
    hu_dups = hu_df[hu_df.duplicated(subset="_title_key", keep=False)]
    n_ai_dup_titles = ai_dups["_title_key"].nunique()
    n_hu_dup_titles = hu_dups["_title_key"].nunique()

    print(f"\n  Total IA:            {len(ai_df)}")
    print(f"  Total Humano:        {len(hu_df)}")
    if n_ai_dup_titles > 0:
        print(f"  Titulos duplicados IA:   {n_ai_dup_titles} titulo(s), {len(ai_dups)} registros")
        for t in ai_dups["_title_key"].unique():
            cnt = (ai_df["_title_key"] == t).sum()
            print(f"     [{cnt}x] {t[:90]}")
    if n_hu_dup_titles > 0:
        print(f"  Titulos duplicados Hum:  {n_hu_dup_titles} titulo(s), {len(hu_dups)} registros")
        for t in hu_dups["_title_key"].unique():
            cnt = (hu_df["_title_key"] == t).sum()
            print(f"     [{cnt}x] {t[:90]}")
    print(f"  Pareados (match):    {n_matched}")
    print(f"  Somente IA:          {n_ai_only}")
    print(f"  Somente Humano:      {n_human_only}")

    if n_ai_only > 0:
        print(f"\n  ** {n_ai_only} TIAB(s) da IA SEM match no humano:")
        for _, row in ai_only.iterrows():
            t = str(row.get("title", row.get("_title_key", "?")))[:100]
            print(f"     - {t}")

    if n_human_only > 0:
        print(f"\n  ** {n_human_only} TIAB(s) do humano SEM match na IA:")
        for _, row in human_only.iterrows():
            t = str(row.get("title", row.get("_title_key", "?")))[:100]
            print(f"     - {t}")

    if n_ai_only == 0 and n_human_only == 0:
        print("\n  Todos os TIABs foram pareados com sucesso!")

    # ================================================================
    #  SALVAR ARQUIVOS
    # ================================================================
    ts = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")

    # 1) Dados pareados (arquivo principal para a Etapa 2)
    paired_df = matched[["title", "screening_decision", "decision_human"]].copy()
    paired_df.insert(0, "n", range(1, len(paired_df) + 1))
    paired_path = output_dir / "pareamento.xlsx"
    paired_df.to_excel(paired_path, index=False)

    # 2) TIABs sem pareamento
    if n_ai_only > 0 or n_human_only > 0:
        unm_rows = []
        for _, row in ai_only.iterrows():
            unm_rows.append({
                "titulo": str(row.get("title", row.get("_title_key", "")))[:200],
                "origem": "Somente IA",
            })
        for _, row in human_only.iterrows():
            unm_rows.append({
                "titulo": str(row.get("title", row.get("_title_key", "")))[:200],
                "origem": "Somente Humano",
            })
        unm_df = pd.DataFrame(unm_rows)
        unm_path = output_dir / "sem_pareamento.xlsx"
        unm_df.to_excel(unm_path, index=False)
        print(f"\n  Arquivo salvo: {unm_path.name}  ({len(unm_df)} registros)")

    print(f"  Arquivo salvo: {paired_path.name}  ({len(paired_df)} registros pareados)")

    # Status final
    if n_ai_only > 0 or n_human_only > 0:
        print(f"\n  >> Revise os TIABs nao pareados antes de prosseguir.")
        print(f"     Corrija os arquivos originais em input/ e rode novamente,")
        print(f"     ou prossiga direto para a Etapa 2 se estiver satisfeito.\n")
    else:
        print(f"\n  >> Pareamento OK! Prossiga para a Etapa 2:")
        print(f"     python 02_analise_diagnostica.py\n")

    print("=" * 62)
    return paired_path


# ------------------------------------------------------------------ main --

def main():
    parser = argparse.ArgumentParser(
        description=(
            "Etapa 1 - Pareamento de planilhas IA vs Humano.\n\n"
            "Modo automatico: coloque 2 arquivos em input/ e rode sem argumentos.\n"
            "Modo manual:     python 01_pareamento.py --ai <arq> --human <arq>"
        ),
        formatter_class=argparse.RawTextHelpFormatter,
    )
    parser.add_argument("--ai",    default=None, help="Arquivo da IA (CSV/XLSX).")
    parser.add_argument("--human", default=None, help="Arquivo do humano (CSV/XLSX).")
    args = parser.parse_args()

    ensure_folders()

    if args.ai and args.human:
        ai_path, human_path = args.ai, args.human
    elif args.ai or args.human:
        print("  ERRO: Informe ambos --ai e --human, ou nenhum.")
        sys.exit(1)
    else:
        print("\n  Modo automatico: buscando arquivos em input/ ...\n")
        ai_path, human_path = auto_detect_files()

    for label, path in [("IA", ai_path), ("Humano", human_path)]:
        if not os.path.isfile(path):
            print(f"  ERRO: arquivo {label} nao encontrado: {path}")
            sys.exit(1)

    print(f"  Arquivo IA:     {Path(ai_path).name}")
    print(f"  Arquivo Humano: {Path(human_path).name}")

    merge_files(ai_path, human_path, OUTPUT_DIR)


if __name__ == "__main__":
    main()
