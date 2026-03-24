# Supply Chain PPT Summarizer

An agentic web application that summarizes Global Supply Chain Status Report PowerPoint decks (50+ slides) into crisp 2–4 slide executive summaries for senior leadership and board members.

## Architecture

```
┌────────────────────────────────────┐
│         FastAPI Backend            │
│                                    │
│  .env → Azure OpenAI config        │
│                                    │
│  /api/upload     → PPT parse       │
│  /api/summarize  → LangGraph agent │
│  /api/refine     → Chat refinement │
│  /api/tokens     → Token tracking  │
│  /api/traces     → Execution traces│
│                                    │
│  ┌──────────────────────────────┐  │
│  │      LangGraph Workflow      │  │
│  │                              │  │
│  │  Summarize Sections          │  │
│  │  (1 LLM call per section)    │  │
│  │         ↓                    │  │
│  │  Executive Summary           │  │
│  │  (synthesize all)            │  │
│  └──────────────────────────────┘  │
│                                    │
│  PPT Parser (python-pptx)          │
│  • Tables → Markdown               │
│  • Charts → data extraction        │
│  • RAG color detection             │
│  • Section auto-detection          │
└────────────────────────────────────┘
```

## Quick Start

### 1. Configure

Copy `.env` and fill in your Azure OpenAI credentials:

```bash
# .env
AZURE_OPENAI_ENDPOINT=https://your-resource.openai.azure.com/
AZURE_OPENAI_API_KEY=your-api-key
AZURE_OPENAI_DEPLOYMENT=gpt-4o
AZURE_OPENAI_API_VERSION=2024-12-01-preview
```

### 2. Install & Run

```bash
pip install -r requirements.txt
python main.py
```

Open http://localhost:8000

### 3. Use

Upload your `.pptx` → review detected sections → generate executive summary → refine via chat.

## Pages

| URL | Description |
|-----|-------------|
| `/` | Main interface — upload, summarize, refine |
| `/tracing` | Execution trace dashboard |
| `/tokens` | Token usage dashboard |

## How It Works

**PPT Parsing** — Extracts every slide's text, tables (→ Markdown), charts (→ data series), RAG color coding, and speaker notes. Auto-detects sections using the Agenda slide and keyword matching.

**Section-by-Section Summarization** — Each section gets its own LLM call to avoid context window overflow on 50+ slide decks. System prompts enforce that only facts from the source are used.

**Executive Summary** — All section summaries are combined for a final synthesis into 2–4 slide content with specific metrics, names, and action items.

**Refinement** — Chat-based iteration with access to both current summary and original section summaries.

## File Structure

```
supply-chain-summarizer/
├── .env                 # Azure OpenAI credentials (not committed)
├── main.py              # FastAPI app, routes
├── ppt_parser.py        # PPT extraction & section detection
├── agent.py             # LangGraph agent, LLM calls, tracing
├── requirements.txt
├── static/
│   ├── index.html       # Main UI
│   ├── tracing.html     # Trace dashboard
│   └── tokens.html      # Token usage dashboard
└── uploads/             # Uploaded PPT files (auto-created)
```
