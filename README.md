# Supply Chain PPT Summarizer

An agentic web application that summarizes Global Supply Chain Status Report PowerPoint decks (50+ slides) into two board-ready outputs: a **2–4 slide executive summary** and a **structured email status update** — both grounded in your company's own glossary of abbreviations, locations, and domain terms.

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
│  PPT Parser (python-pptx)                │
│  Glossary Loader (multi-format JSON)     │
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

Upload your `.pptx` → review detected sections → generate outputs → switch between **Slide Summary** and **Email Summary** tabs → refine each independently via chat.

## Dual Output

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

**Refinement** — Each output (slides or email) can be refined independently via chat. The refine endpoint accepts a `target` parameter (`slides` or `email`) and routes to the appropriate prompt, which has access to both the current output and the original section summaries.

**Observability** — Every LLM call logs prompt/completion token counts. Every graph node execution is traced with timing, inputs, and outputs. Both are viewable in dedicated dashboards.

## File Structure

```
supply-chain-summarizer/
├── .env                          # Azure OpenAI + app config
├── main.py                       # FastAPI app, routes, glossary endpoints
├── ppt_parser.py                 # PPT extraction & section detection
├── agent.py                      # LangGraph agent, prompts, tracing
├── glossary.py                   # Glossary loader & prompt renderer
├── requirements.txt
├── glossary/                     # Company glossary JSON files
│   └── _sample_glossary.json     # Example with 58 entries
├── static/
│   ├── index.html                # Main UI (tabbed: slides / email / sections)
│   ├── tracing.html              # Trace dashboard
│   └── tokens.html               # Token usage dashboard
└── uploads/                      # Uploaded PPT files (auto-created)
```

## API Reference

### Core

| Method | Endpoint | Description |
|--------|----------|-------------|
| GET | `/api/health` | Health check and config status |
| POST | `/api/upload` | Upload and parse a `.pptx` file |
| POST | `/api/summarize` | Run the full summarization workflow (slides + email) |
| POST | `/api/refine` | Refine an output. Form fields: `session_id`, `instruction`, `target` (`slides` or `email`) |
| GET | `/api/session/{id}` | Session metadata |
| GET | `/api/sessions` | List all sessions |

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
