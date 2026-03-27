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
│  /api/gfd/upload      → GFD Excel parse  │
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
│  │   GFD Dashboard Generator          │  │
│  │                                    │  │
│  │  1. Parse Dashboard_Update Excel   │  │
│  │     (multi-row headers, merges,    │  │
│  │      fuzzy column matching)        │  │
│  │              ↓                     │  │
│  │  2. Generate CW-based RAG slides  │  │
│  │     (12-week grid + next quarter,  │  │
│  │      auto-paginated)               │  │
│  └────────────────────────────────────┘  │
│                                          │
│  PPT Parser (python-pptx)                │
│  Glossary Loader (multi-format JSON)     │
│  GFD Excel Parser (openpyxl)             │
│  GFD Slide Generator (python-pptx)       │
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

### Slide Layout

```
┌──────────────────────────────────────────────────────────────────────────┐
│  Global Fulfilment Dashboard                              CW13/2026    │
├───────┬─────┬─────────┬──────┬───┬───┬───┬───┬···┬───┬────┬──────┬────┬───┤
│ PG    │Plant│Customer │Cover.│13 │14 │15 │16 │   │24 │ Q2 │Suppl.│Act.│FM │
│(merge)│     │         │      │■■■│■■■│■■■│■■■│   │■■■│■■■ │      │    │   │
└───────┴─────┴─────────┴──────┴───┴───┴───┴───┴···┴───┴────┴──────┴────┴───┘

■ GREEN = covered without mitigation
■ AMBER = covered only if mitigations succeed
■ RED   = beyond all coverage (uncovered)
```

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

### Excel Parser Robustness

The parser (`gfd_excel_parser.py`) is designed for real-world Excel files that are not perfectly structured:

- **Multi-row headers** — Automatically detects and flattens stacked header rows (e.g., a category row above a column-name row) using keyword-scoring heuristics
- **Headers not at row 1** — Scans the first 25 rows for the header band, skipping title rows, logos, and blank rows
- **Merged cells** — Resolves both header merges (horizontal/vertical) and data merges (e.g., product group cells spanning multiple rows)
- **Fuzzy column matching** — Three-pass matching: (1) exact normalised match, (2) keyword containment, (3) fuzzy similarity (SequenceMatcher > 0.75). Handles newlines in headers, underscores vs spaces, inconsistent casing
- **Non-data row filtering** — Automatically skips separator rows (`---`, `===`), subtotal rows, and rows with insufficient data
- **European number/date formats** — Handles `1.234,56` numbers, `DD.MM.YYYY` dates, `€` symbols
- **CW format flexibility** — Parses `CW18`, `CW18/2026`, `CW18/26`, `KW18` (German), `W18`, and bare `18`

### GFD API Endpoints

| Method | Endpoint | Description |
|--------|----------|-------------|
| POST | `/api/gfd/upload` | Upload `.xlsx`, parse `Dashboard_Update` sheet, generate slides |
| GET | `/api/gfd/download` | Download generated `.pptx`. Query param: `session_id` |
| GET | `/api/gfd/session/{id}` | Session metadata (row count, warnings, product groups) |

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
| `/tracing` | Execution trace dashboard |
| `/tokens` | Token usage dashboard |

## How It Works

**Glossary Loading** — At startup, all `.json` files in `GLOSSARY_DIR` are loaded, normalised into a unified `{ABBR: {meaning, category}}` map, and rendered as a compact reference block injected into every LLM system prompt. This ensures abbreviations for locations, roles, business entities, and domain terms are correctly expanded throughout both outputs.

**PPT Parsing** — Extracts every slide's text, tables (→ Markdown), charts (→ data series), RAG color coding, and speaker notes. Auto-detects sections using the Agenda slide and keyword matching against known section types (crisis, supplier, production, fulfilment, customer, freight, cost).

**Section-by-Section Summarization** — Each section gets its own LLM call (with glossary context) to avoid context window overflow on 50+ slide decks. System prompts enforce that only facts from the source are used.

**Executive Slide Summary** — All section summaries are combined for a synthesis into 2–4 slide content with specific metrics, names, and action items.

**Email Status Summary** — The same section summaries feed a separate LLM call with a dedicated prompt following the crisis-status email template. Only sections with substantive data are included.

**GFD Dashboard Generation** — The `gfd_excel_parser.py` module parses the `Dashboard_Update` worksheet with multi-row header detection and fuzzy column matching. Parsed rows are grouped by product family, coverage CW boundaries are extracted, and `gfd_slide_generator.py` produces widescreen slides with a 12-week + next-quarter RAG heatmap. No LLM calls are needed — this is a deterministic data-to-slide pipeline.

**Refinement** — Each output (slides or email) can be refined independently via chat. The refine endpoint accepts a `target` parameter (`slides` or `email`) and routes to the appropriate prompt, which has access to both the current output and the original section summaries.

**Observability** — Every LLM call logs prompt/completion token counts. Every graph node execution is traced with timing, inputs, and outputs. Both are viewable in dedicated dashboards.

**DOCX Export** — Either output can be downloaded as a Word document. The `docx_export.py` module converts the agent's markdown into a styled `.docx` with proper heading levels, bold/italic runs, bulleted lists, and horizontal rules using `python-docx`.

## File Structure

```
supply-chain-summarizer/
├── .env                          # Azure OpenAI + app config
├── main.py                       # FastAPI app, routes, glossary endpoints
├── ppt_parser.py                 # PPT extraction & section detection
├── agent.py                      # LangGraph agent, prompts, tracing
├── glossary.py                   # Glossary loader & prompt renderer
├── docx_export.py                # Markdown → Word document converter
├── gfd_excel_parser.py           # Dashboard_Update Excel parser
├── gfd_slide_generator.py        # GFD → PowerPoint slide generator
├── requirements.txt
├── glossary/                     # Company glossary JSON files
│   └── _sample_glossary.json     # Example with 58 entries
├── static/
│   ├── index.html                # Main UI (tabbed: slides / email / sections)
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
| POST | `/api/gfd/upload` | Upload `.xlsx` and generate dashboard `.pptx` |
| GET | `/api/gfd/download` | Download generated `.pptx`. Query param: `session_id` |
| GET | `/api/gfd/session/{id}` | Parsed data metadata and warnings |

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
