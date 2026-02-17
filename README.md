# AI-Tools — Triagem por IA + Teste Diagnóstico

Plataforma para **triagem automatizada** de artigos científicos por IA (via OpenAI) e **avaliação diagnóstica** comparando as decisões da IA com revisores humanos em revisões sistemáticas.

---

## Estrutura do Projeto

```
AI-tools/
├── backend.py                  ← Backend FastAPI (triagem por IA)
├── index.html                  ← Frontend da aplicação web
├── app.js                      ← Lógica do frontend (JS)
├── style.css                   ← Estilos do frontend
├── logo.avif                   ← Logo da aplicação
├── diagnostic/                 ← Scripts de análise diagnóstica
│   ├── 01_pareamento.py        ← Etapa 1: pareamento IA vs Humano
│   ├── 02_analise_diagnostica.py ← Etapa 2: métricas + Word
│   └── 03_fulltext_check.py    ← Etapa 3: fulltext capture check
├── input/                      ← Arquivos de entrada (não versionados)
│   ├── <arquivo_da_IA>.xlsx
│   ├── <arquivo_humano_TIAB>.xlsx
│   └── <arquivo_fulltext>.xlsx
├── output/                     ← Resultados gerados (não versionados)
│   ├── pareamento.xlsx
│   ├── sem_pareamento.xlsx
│   ├── diagnostic_results_*.docx
│   ├── fulltext_check_*.docx
│   └── *.json
├── .gitignore
└── README.md
```

---

## Pré-requisitos

- **Python 3.10+**
- Criar e ativar o ambiente virtual:

```bash
python -m venv .venv
# Windows
.venv\Scripts\activate
# Linux/Mac
source .venv/bin/activate
```

- Instalar dependências:

```bash
pip install fastapi uvicorn python-dotenv requests pydantic
pip install pandas numpy openpyxl python-docx
```

> **Nota:** O pacote correto é `python-docx` (não `docx`). Se tiver conflito: `pip uninstall docx -y && pip install python-docx`.

---

# Parte 1 — Aplicação Web (Triagem por IA)

## O que faz

A aplicação web permite envio em lote de artigos científicos (título + abstract) para triagem automatizada via modelos da OpenAI. O backend processa cada artigo e retorna uma decisão de screening (`include`, `exclude` ou `maybe`) com justificativa.

## Backend (`backend.py`)

### Tecnologia

- **FastAPI** com CORS habilitado
- Comunicação com a **API da OpenAI** (modelos GPT-4o, GPT-5, etc.)
- Gerenciamento de jobs em memória com threads
- Progresso em tempo real via **Server-Sent Events (SSE)**
- Exportação de resultados em **CSV** e **XLSX**

### Como o código funciona

1. **`StartPayload`** (Pydantic model) — Recebe do frontend:
   - `model`: modelo da OpenAI (ex: `gpt-4o`, `gpt-5`)
   - `api_key`: chave da API do OpenAI
   - `study_synopsis`: sinopse/PICO do estudo
   - `inclusion_criteria` / `exclusion_criteria`: critérios de inclusão e exclusão
   - `records`: lista de artigos com `title` e `abstract`
   - `temperature`, `params`: parâmetros opcionais do modelo

2. **`build_prompt()`** — Monta o prompt de triagem com instruções detalhadas:
   - Foco em **alta sensibilidade** (errar para inclusão, não para exclusão)
   - Pede ao modelo um JSON com `decision`, `rationale`, `inclusion_evaluation`, `exclusion_evaluation`
   - Instrui o modelo a marcar como `maybe` em caso de incerteza (não excluir)

3. **`call_openai_chat()`** — Faz a requisição HTTP para a API da OpenAI:
   - Suporta tanto a API de **chat completions** (GPT-4o) quanto a API de **responses** (GPT-5)
   - **Retry com backoff exponencial** para erros 429 (rate limit) e 5xx
   - Respeita o header `Retry-After` quando presente
   - Parseia o JSON da resposta com tratamento robusto de erros

4. **`worker()`** — Thread que processa os artigos em background:
   - Itera sobre cada record, chama a API, parseia o resultado
   - Aplica **rate limiting** entre chamadas (mín. 0.6s por padrão)
   - Registra progresso em tempo real (acessível via SSE)
   - Suporta **cancelamento** do job em andamento

### Endpoints

| Método | Rota | Descrição |
|--------|------|-----------|
| `POST` | `/api/start` | Inicia um job de triagem (retorna `job_id`) |
| `POST` | `/api/cancel/{job_id}` | Cancela um job em andamento |
| `GET` | `/api/status/{job_id}` | Retorna status do job (running/done/error) |
| `GET` | `/api/progress/{job_id}` | SSE stream de progresso em tempo real |
| `GET` | `/api/partial/{job_id}` | Resultados parciais (paginados) |
| `GET` | `/api/result/{job_id}?format=csv\|xlsx` | Resultado final como CSV ou XLSX |

### Frontend (`index.html`, `app.js`, `style.css`)

Interface web para:
- Inserir API key, modelo, sinopse e critérios
- Upload de planilha (CSV/XLSX) com artigos
- Acompanhar progresso em tempo real
- Visualizar e baixar resultados

O backend serve os arquivos estáticos do frontend automaticamente (montagem do diretório raiz).

### Como rodar

```bash
uvicorn backend:app --reload --port 8000
```

Acesse: **http://localhost:8000**

### Variáveis de ambiente (.env)

| Variável | Padrão | Descrição |
|----------|--------|-----------|
| `RATE_LIMIT_MIN_INTERVAL` | `0.6` | Intervalo mínimo entre chamadas (segundos) |
| `OPENAI_MAX_RETRIES` | `5` | Máximo de tentativas em caso de erro |
| `OPENAI_BASE_BACKOFF` | `1.0` | Base do backoff exponencial (segundos) |

---

# Parte 2 — Teste Diagnóstico (IA vs Humano)

Os scripts em `diagnostic/` avaliam o desempenho da triagem automática da IA comparando com decisões de revisores humanos. Todos rodam a partir da raiz do projeto e usam as pastas `input/` e `output/`.

---

## Etapa 1 — Pareamento (`diagnostic/01_pareamento.py`)

### O que faz

- Lê os dois arquivos (IA e humano) da pasta `input/`
- Pareia os artigos por **título normalizado** (lowercase, strip)
- Lida com **títulos duplicados** usando um índice de ocorrência (`cumcount`) — se o mesmo título aparece 4x em ambos os arquivos, cada cópia é pareada separadamente
- Lista os TIABs que **não foram pareados** (somente em um dos arquivos)
- Salva em `output/pareamento.xlsx` (pareados) e `output/sem_pareamento.xlsx` (não pareados)

### Auto-detecção

Quando existem múltiplos arquivos em `input/`, o script detecta automaticamente:
- **Arquivo da IA**: contém coluna `screening_decision`, maior número de registros
- **Arquivo do Humano**: contém coluna `decision`, **maior** arquivo restante (prioriza o TIAB sobre fulltext)

### Como rodar

```bash
# Modo automático
python diagnostic/01_pareamento.py

# Modo manual
python diagnostic/01_pareamento.py --ai input/arquivo_ia.xlsx --human input/arquivo_humano.xlsx
```

### Formato dos arquivos

| Arquivo | Colunas obrigatórias |
|---------|----------------------|
| IA | `title`, `screening_decision` |
| Humano (TIAB) | `title`, `decision` |

> **Dica:** Se houver TIABs sem parear, revise e corrija os arquivos de entrada antes de prosseguir.

---

## Etapa 2 — Análise Diagnóstica (`diagnostic/02_analise_diagnostica.py`)

### O que faz

- Lê `output/pareamento.xlsx` (gerado na Etapa 1)
- **Binariza** as decisões: `include` → `maybe` (para comparação uniforme)
- Calcula a **matriz de confusão** (TP, FP, FN, TN)
- Calcula **métricas diagnósticas completas**:
  - Prevalência, Sensibilidade (Recall), Especificidade
  - VPP (Precision), VPN, Acurácia, F1
  - Likelihood Ratio + e −, Índice de Youden
- Calcula **Cohen's Kappa** com erro padrão e IC 95%
- Gera documento **Word (.docx)** com 5 tabelas formatadas para publicação

### Lógica da Binarização

A decisão humana é binária: `maybe` (inclui) ou `exclude`.
A IA pode retornar `maybe`, `include` ou `exclude`. Como `maybe` e `include` são equivalentes (artigo passa para a próxima fase):

```
include → maybe    (positivo = artigo passa)
exclude → exclude  (negativo = artigo eliminado)
```

### Tabelas geradas no Word

| Tabela | Conteúdo |
|--------|----------|
| Table 1 | Sample Characteristics |
| Table 2 | Confusion Matrix (2×2) |
| Table 3 | Diagnostic Accuracy (todas as métricas) |
| Table 4 | Inter-rater Agreement — Cohen's Kappa |
| Table 5 | Summary of Results |

As tabelas usam **Times New Roman**, bordas horizontais no estilo acadêmico, prontas para publicação.

### Como rodar

```bash
python diagnostic/02_analise_diagnostica.py
```

---

## Etapa 3 — Fulltext Capture Check (`diagnostic/03_fulltext_check.py`)

### O que faz

Verifica se os artigos **incluídos na revisão final** (após leitura completa / fulltext) teriam sido **mantidos pela IA** durante a triagem de TIAB.

- Lê o arquivo de artigos incluídos na revisão final (ex: 30 artigos)
- Procura cada artigo na base de decisões da IA (ex: 973 artigos)
- Verifica se a IA classificou como `maybe` ou `include` (passaria) vs `exclude` (perdido)
- Calcula a **taxa de captura** (capture rate) e a **taxa de perda** (miss rate)
- Gera:
  - **Word (.docx)** com tabelas de resumo, detalhamento artigo-a-artigo, e lista de artigos perdidos (com highlight em vermelho)
  - **XLSX** com a tabela detalhada
  - **JSON** com o resumo

### Auto-detecção

- **Arquivo da IA**: contém `screening_decision`, > 50 registros
- **Arquivo Fulltext**: menor arquivo restante (tipicamente 20-40 artigos)

### Como rodar

```bash
# Modo automático
python diagnostic/03_fulltext_check.py

# Modo manual
python diagnostic/03_fulltext_check.py --ai input/arquivo_ia.xlsx --fulltext input/arquivo_fulltext.xlsx
```

### Formato do arquivo fulltext

| Coluna | Descrição |
|--------|-----------|
| `title` | Título do artigo incluído |
| `abstract` | (Opcional) Abstract |
| `decision` | (Opcional) Todas `include` |

---

## Interpretação do Kappa (Landis & Koch, 1977)

| Kappa | Concordância |
|-------|-------------|
| < 0 | Poor |
| 0.00–0.20 | Slight |
| 0.21–0.40 | Fair |
| 0.41–0.60 | Moderate |
| 0.61–0.80 | Substantial |
| 0.81–1.00 | Almost Perfect |

---

## Fluxo Completo de Uso

```
1. Rodar a triagem por IA na aplicação web (backend.py)
   → Exportar resultado como .xlsx para input/

2. Colocar arquivo humano TIAB em input/

3. python diagnostic/01_pareamento.py
   → output/pareamento.xlsx

4. python diagnostic/02_analise_diagnostica.py
   → output/diagnostic_results_*.docx  (métricas, kappa, tabelas)

5. (Opcional) Colocar arquivo fulltext em input/
   python diagnostic/03_fulltext_check.py
   → output/fulltext_check_*.docx  (taxa de captura)
```
