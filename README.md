# Supply Chain PPT Summarizer

An agentic web application that summarizes Global Supply Chain Status Report PowerPoint decks (50+ slides) into two board-ready outputs: a **2–4 slide executive summary** and a **structured email status update** — both grounded in your company's own glossary of abbreviations, locations, and domain terms.

Additionally, the app generates a **Global Fulfilment Dashboard** presentation directly from the `Dashboard_Update` Excel worksheet, producing color-coded, CW-based risk heatmap slides for senior management review.

## Architecture

```
┌──────────────────────────────────────────┐
│            FastAPI Backend               │
│                                          │
│  .env       → Azure OpenAI config        │
│  glossary/  → Company glossary JSON      │
│                                          │
│  /api/upload          → PPT parse        │
│  /api/summarize       → LangGraph agent  │
│  /api/refine          → Chat refinement  │
│  /api/glossary        → Glossary CRUD    │
│  /api/tokens          → Token tracking   │
│  /api/traces          → Execution traces │
│  /api/gfd/upload      → GFD pipeline     │
│  /api/gfd/download    → GFD PPT download │
│                                          │
│  ┌────────────────────────────────────┐  │
│  │        LangGraph Workflow          │  │
│  │                                    │  │
│  │  1. Summarize Sections             │  │
│  │     (1 LLM call per section        │  │
│  │      + glossary context)           │  │
│  │              ↓                     │  │
│  │  2. Executive Slide Summary        │  │
│  │     (synthesize into 2-4 slides    │  │
│  │      + glossary context)           │  │
│  │              ↓                     │  │
│  │  3. Email Status Summary           │  │
│  │     (structured email update       │  │
│  │      + glossary context)           │  │
│  └────────────────────────────────────┘  │
│                                          │
│  ┌────────────────────────────────────┐  │
│  │   GFD Dashboard Pipeline           │  │
│  │                                    │  │
│  │  Stage 1 — Deterministic filter    │  │
│  │     Excel → merge-resolved,        │  │
│  │     date-filtered text table       │  │
│  │              ↓                     │  │
│  │  Stage 2 — LLM extraction          │  │
│  │     Text table → structured JSON   │  │
│  │     (product groups, CW integers,  │  │
│  │      customers, actions)           │  │
│  │              ↓                     │  │
│  │  Stage 3 — LLM slide spec          │  │
│  │     JSON → complete slide spec     │  │
│  │     (RAG colors, condensed text,   │  │
│  │      overview bullets, risk level) │  │
│  │              ↓                     │  │
│  │  Stage 4 — PPTX renderer           │  │
│  │     Slide spec → .pptx output      │  │
│  │     (thin paint-by-numbers,        │  │
│  │      no business logic)            │  │
│  └────────────────────────────────────┘  │
│                                          │
│  PPT Parser (python-pptx)                │
│  Glossary Loader (multi-format JSON)     │
│  GFD LLM Parser (openpyxl + LLM)        │
│  GFD LLM Slides (LLM + python-pptx)     │
└──────────────────────────────────────────┘
```

## Quick Start

### 1. Configure

Fill in your Azure OpenAI credentials in `.env`:

```bash
AZURE_OPENAI_ENDPOINT=https://your-resource.openai.azure.com/
AZURE_OPENAI_API_KEY=your-api-key
AZURE_OPENAI_DEPLOYMENT=gpt-4o
AZURE_OPENAI_API_VERSION=2024-12-01-preview
GLOSSARY_DIR=./glossary
```

### 2. Add Glossary Files

Place your company glossary `.json` files in the `glossary/` directory. Three JSON formats are auto-detected:

**Format A — Categorised** (recommended):
```json
{
  "locations": { "BHV": "Bremerhaven plant", "BOG": "Bogen plant" },
  "supply_chain": { "WOS": "Weeks of Supply", "OTD": "On-Time Delivery" }
}
```

**Format B — Flat key-value:**
```json
{ "BHV": "Bremerhaven plant", "KAM": "Key Account Manager" }
```

**Format C — Array of objects:**
```json
[
  { "abbreviation": "FM", "description": "Force Majeure", "category": "legal" },
  { "code": "WOS", "full_name": "Weeks of Supply", "type": "inventory" }
]
```

Multiple files are merged at startup. Additional files can be uploaded via the UI at runtime.

### 3. Install & Run

```bash
pip install -r requirements.txt
python main.py
```

Open http://localhost:8000

### 4. Use

**PPT Summarizer:** Upload your `.pptx` → review detected sections → generate outputs → switch between **Slide Summary** and **Email Summary** tabs → refine each independently via chat.

**GFD Dashboard:** Upload your `.xlsx` containing the `Dashboard_Update` worksheet → download the generated `.pptx` with color-coded CW risk heatmap.

---

## Global Fulfilment Dashboard (GFD) Module

### Overview

The GFD module converts the `Dashboard_Update` Excel worksheet into presentation-ready slides showing a forward-looking calendar-week risk heatmap. Each row in the Excel represents a delivery risk for a product family at a specific plant, and the generated slides show whether supply coverage extends across the next 12 weeks plus the following quarter.

The pipeline is fully LLM-driven for both content extraction and slide generation. Only the initial staleness filter is deterministic — everything else is handled by the model, which means product groupings, column semantics, customer lists, and narrative summaries are understood rather than pattern-matched.

### Slide Layout

```
┌──────────────────────────────────────────────────────────────────────────┐
│  Global Fulfilment Dashboard                              CW13/2026     │
├──────────┬─────┬──────────┬──────┬───┬───┬───┬───┬···┬───┬────┬────┬───┤
│ Component│Plant│ Customer │Suppl.│13 │14 │15 │16 │   │24 │ Q2 │Act.│FM │
│          │     │ affected │      │■■■│■■■│■■■│■■■│   │■■■│■■■ │    │   │
└──────────┴─────┴──────────┴──────┴───┴───┴───┴───┴···┴───┴────┴────┴───┘

■ GREEN = covered without mitigation
■ AMBER = covered only if mitigations succeed
■ RED   = beyond all coverage (uncovered)
```

An executive overview slide precedes the per-product-group heatmap slides, showing overall risk level, item counts by severity, and LLM-generated key-risk bullets.

### CW RAG Logic

The RAG status for each calendar-week cell is derived from two coverage boundary fields in the Excel:

- **Coverage w/o risk mitigation** (e.g., `CW15`) — supply is secured through this week without any special actions
- **Coverage w/ risk mitigation** (e.g., `CW19`) — supply is secured through this week assuming mitigations succeed

For each CW column on the slide:

| Condition | Color | Meaning |
|-----------|-------|---------|
| CW ≤ coverage w/o mitigation | GREEN | Supply secured |
| CW > w/o but ≤ w/ mitigation | AMBER | Depends on mitigation actions |
| CW > coverage w/ mitigation | RED | No supply plan in place |

The **next-quarter summary column** (e.g., Q2) shows the worst-case RAG across all weeks in that quarter. If any single week in Q2 is RED, the Q2 column shows RED.

### Four-Stage Pipeline

#### Stage 1 — Deterministic filter (`gfd_llm_parser.py`)

The only deterministic stage. Opens the workbook with openpyxl and performs three operations:

- **Merge resolution** — builds a complete master→slave map so every merged slave cell carries its group value. This is what makes product-family grouping reliable: formerly-merged cells now appear in every row of the text table.
- **Header detection** — scores each row for supply-chain keyword density (scans up to 60 rows, merge-aware). Correctly handles title blocks, logo rows, instruction text, or any other preamble above the real header — no assumption is made about which row number the header will be on.
- **Sub-header combining** — if the row immediately below the header also scores as header-like (score ≥ 2), its cells are merged into the header strings. If the row scores low (i.e. it is the first data row), combining is skipped — preventing data values from being absorbed into column names.
- **Customer column compaction** — the standard template has 32 individual OEM/customer columns (W–BB) under a merged "Customer affected" super-header. These are detected automatically via the merged super-header and collapsed into a single `Customers affected` column whose value lists only the customers flagged as affected in each row. This reduces a 64-column table to ~33 columns.
- **Staleness filter** — detects the `Last update` column (or equivalent) using a prioritised list of hints including German variants (`aktualisiert`, `last änderung`). The raw openpyxl cell value — which may be a native `datetime` object, an Excel serial number, or a lazily typed string — is parsed through multiple date format patterns covering European (`dd.mm.yyyy`), ISO, US, two-digit year, and year-less entry styles. Rows whose update date is older than `history_weeks` weeks are dropped. Falls back to scanning rows for `CW`/`KW` references if no date column is found.

The stage produces a compact pipe-delimited text table: clean headers, apostrophe-stripped values, customer columns compacted, stale rows removed.

#### Stage 2 — LLM extraction (`gfd_llm_parser.py`)

The filtered text table is sent to the LLM with a strict extraction prompt. The model is instructed to:

- Group rows by repeated product family values (formerly-merged cells that now repeat in the table)
- Extract coverage CW fields as plain integers (`"CW15"` → `15`) for precise RAG computation downstream
- Preserve all text fields verbatim — no paraphrasing
- Normalise boolean-like fields (`Customer Informed`, `Allocation Mode`) to `Yes / No / In progress / N/A`
- Note any data quality issues or structural ambiguities in `extraction_notes`

Output is a structured JSON object with `product_groups`, each containing typed `rows`. This is the single source of truth for Stage 3.

#### Stage 3 — LLM slide spec (`gfd_llm_slides.py`)

The extracted JSON is sent to a second LLM call that generates a **complete slide specification** — the model makes all content and layout decisions:

- Which rows go on which slides (paginating at 9 rows per product-group slide)
- Text condensed to fit each column's character budget (component ≤ 45 chars, action ≤ 90 chars, etc.)
- RAG color for every CW column of every row, computed from the integer coverage fields using the GREEN / AMBER / RED logic above
- Quarter worst-case RAG color
- Per-row risk level (`CRITICAL / HIGH / MEDIUM / LOW`)
- Executive overview: overall risk badge, stats panel (counts by severity), and 4–6 narrative bullets citing specific CW numbers, plant names, and customer exposure

If the slide-spec LLM call fails, a deterministic fallback computes RAG colors arithmetically from the integer CW fields the extractor produced, so the user always receives a working dashboard.

#### Stage 4 — PPTX renderer (`gfd_llm_slides.py`)

A thin "paint by numbers" renderer converts the slide spec JSON into a python-pptx `Presentation`. It contains no business logic — it maps `"GREEN"` → `RGB(0,176,80)`, `"CRITICAL"` → red badge, and so on. Every visual decision was already made in Stage 3.

### Excel File Compatibility

The parser is designed for real-world files and makes no assumptions about row positions or column order:

| Challenge | How it's handled |
|---|---|
| Title / logo rows above the header | Keyword-scoring scan up to row 60; highest-scoring row wins |
| Multi-row stacked headers | Sub-header row combined only when it itself scores as header-like |
| Merged product-group cells | Full master→slave map built before any row is read |
| 32 individual customer columns | Auto-detected via merged "Customer affected" super-header, compacted to one column |
| Leading apostrophes in cell values | Stripped in the normalisation step (common in template files) |
| Lazy date entry formats | 14 format patterns tried; native `datetime` objects used directly when available |
| No "Last update" column | Falls back to CW-number scanning across row text |
| LLM extraction failure | Deterministic fallback spec still produces a valid dashboard |

### GFD API Endpoints

| Method | Endpoint | Description |
|--------|----------|-------------|
| POST | `/api/gfd/upload` | Upload `.xlsx`, run all four pipeline stages, return session metadata |
| GET | `/api/gfd/download` | Download generated `.pptx`. Query param: `session_id` |
| GET | `/api/gfd/session/{id}` | Session metadata including product groups, warnings, overall risk, and whether the fallback renderer was used |

**Upload form fields:**

| Field | Type | Default | Description |
|---|---|---|---|
| `file` | `.xlsx` | — | The Excel workbook containing `Dashboard_Update` |
| `history_weeks` | int | `4` | Rows with `Last update` older than this many weeks are excluded |

**Upload response includes:**

| Field | Description |
|---|---|
| `current_cw` | Calendar week label used for the dashboard (e.g. `CW13/2026`) |
| `product_groups` | List of extracted product groups with row counts |
| `overall_risk` | Aggregate risk level from the executive overview slide |
| `slide_count` | Total number of slides generated |
| `extraction_notes` | LLM comments on data quality or structural ambiguities |
| `is_fallback` | `true` if the deterministic fallback renderer was used instead of the LLM slide spec |
| `warnings` | List of parser warnings (rows dropped, date column used, etc.) |

---

## Dual Output (PPT Summarizer)

The app produces two independent outputs from the same underlying section summaries:

### Slide Summary (2–4 slides)
Structured as board-presentation content with:
- Situation Overview & Key Metrics
- Supply & Production Impact
- Customer & Financial Impact
- Actions & Outlook (if applicable)

### Email Status Summary
A structured leadership email following the crisis-status template:
- Key Management Takeaways
- Overall Situation Summary
- Key Product & Customer Risks (per product group)
- Additional Product / Production Status
- Supplier Mitigation Actions
- Customer: Commercial & Legal Situation
- Logistics Situation
- Financial Exposure – Scenario Ranges

Sections are only included when substantive data exists in the source PPT — no generic filler.

Both outputs are independently refinable via the chat bar. The active tab determines which output receives the refinement instruction.

### Download

Each output can be downloaded as a formatted `.docx` Word document via a discrete button in the tab bar. The export preserves all formatting — headings, bold, italic, bullet lists, and section breaks — so the content can be directly pasted into PowerPoint or Outlook without reformatting.

## Pages

| URL | Description |
|-----|-------------|
| `/` | Main interface — upload, glossary, tabbed outputs, chat refinement |
| `/gfd` | GFD Dashboard — upload Excel, download PPTX |
| `/tracing` | Execution trace dashboard |
| `/tokens` | Token usage dashboard |

## How It Works

**Glossary Loading** — At startup, all `.json` files in `GLOSSARY_DIR` are loaded, normalised into a unified `{ABBR: {meaning, category}}` map, and rendered as a compact reference block injected into every LLM system prompt. This ensures abbreviations for locations, roles, business entities, and domain terms are correctly expanded throughout all outputs.

**PPT Parsing** — Extracts every slide's text, tables (→ Markdown), charts (→ data series), RAG color coding, and speaker notes. Auto-detects sections using the Agenda slide and keyword matching against known section types (crisis, supplier, production, fulfilment, customer, freight, cost).

**Section-by-Section Summarization** — Each section gets its own LLM call (with glossary context) to avoid context window overflow on 50+ slide decks. System prompts enforce that only facts from the source are used.

**Executive Slide Summary** — All section summaries are combined for a synthesis into 2–4 slide content with specific metrics, names, and action items.

**Email Status Summary** — The same section summaries feed a separate LLM call with a dedicated prompt following the crisis-status email template. Only sections with substantive data are included.

**GFD Dashboard Generation** — `gfd_llm_parser.py` performs a minimal deterministic read of the Excel file (merge resolution, header detection, customer column compaction, date-based staleness filtering) and converts the result to a pipe-delimited text table. This table is then passed to an LLM which extracts a fully typed JSON structure covering all product groups and risk rows. A second LLM call uses that JSON to generate a complete slide specification — RAG colors, condensed text, overview narrative, risk levels. A thin PPTX renderer then converts the spec to a `.pptx` file without applying any business logic of its own.

**Refinement** — Each output (slides or email) can be refined independently via chat. The refine endpoint accepts a `target` parameter (`slides` or `email`) and routes to the appropriate prompt, which has access to both the current output and the original section summaries.

**Observability** — Every LLM call logs prompt/completion token counts. Every graph node execution is traced with timing, inputs, and outputs. Both are viewable in dedicated dashboards.

**DOCX Export** — Either PPT summarizer output can be downloaded as a Word document. The `docx_export.py` module converts the agent's markdown into a styled `.docx` with proper heading levels, bold/italic runs, bulleted lists, and horizontal rules using `python-docx`.

## File Structure

```
supply-chain-summarizer/
├── .env                          # Azure OpenAI + app config
├── main.py                       # FastAPI app, routes, glossary endpoints
├── ppt_parser.py                 # PPT extraction & section detection
├── agent.py                      # LangGraph agent, prompts, tracing
├── glossary.py                   # Glossary loader & prompt renderer
├── docx_export.py                # Markdown → Word document converter
├── gfd_llm_parser.py             # GFD Stage 1+2: Excel filter → LLM extraction
├── gfd_llm_slides.py             # GFD Stage 3+4: LLM slide spec → PPTX renderer
├── requirements.txt
├── glossary/                     # Company glossary JSON files
│   └── _sample_glossary.json     # Example with 58 entries
├── static/
│   ├── index.html                # Main UI (tabbed: slides / email / sections)
│   ├── gfd.html                  # GFD Dashboard UI
│   ├── tracing.html              # Trace dashboard
│   └── tokens.html               # Token usage dashboard
└── uploads/                      # Uploaded PPT/XLSX files (auto-created)
```

## API Reference

### Core (PPT Summarizer)

| Method | Endpoint | Description |
|--------|----------|-------------|
| GET | `/api/health` | Health check and config status |
| POST | `/api/upload` | Upload and parse a `.pptx` file |
| POST | `/api/summarize` | Run the full summarization workflow (slides + email) |
| POST | `/api/refine` | Refine an output. Form fields: `session_id`, `instruction`, `target` (`slides` or `email`) |
| GET | `/api/download` | Download output as `.docx`. Query params: `session_id`, `target` (`slides` or `email`) |
| GET | `/api/session/{id}` | Session metadata |
| GET | `/api/sessions` | List all sessions |

### GFD Dashboard

| Method | Endpoint | Description |
|--------|----------|-------------|
| POST | `/api/gfd/upload` | Upload `.xlsx`, run four-stage pipeline, return metadata. Form fields: `file`, `history_weeks` (default `4`) |
| GET | `/api/gfd/download` | Download generated `.pptx`. Query param: `session_id` |
| GET | `/api/gfd/session/{id}` | Session metadata: product groups, warnings, overall risk, slide count, fallback flag |

### Glossary

| Method | Endpoint | Description |
|--------|----------|-------------|
| GET | `/api/glossary` | List all entries grouped by category |
| POST | `/api/glossary/upload` | Upload a new glossary JSON file |
| DELETE | `/api/glossary/{filename}` | Remove a glossary file and reload |

### Observability

| Method | Endpoint | Description |
|--------|----------|-------------|
| GET | `/api/tokens` | Token usage log (optional `?session_id=` filter) |
| GET | `/api/traces` | Execution traces (optional `?session_id=` filter) |
