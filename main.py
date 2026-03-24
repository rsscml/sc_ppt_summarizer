"""
Supply Chain PPT Summarizer - FastAPI Application
==================================================
Agentic web app for summarizing Global Supply Chain Status Reports.
Configuration loaded from .env file.
"""

import os
import uuid
import shutil
from pathlib import Path
from datetime import datetime, timezone
from dotenv import load_dotenv

from fastapi import FastAPI, UploadFile, File, Form, HTTPException
from fastapi.staticfiles import StaticFiles
from fastapi.responses import HTMLResponse

from ppt_parser import parse_presentation
from agent import (
    build_summarization_graph,
    refine_summary,
    token_usage_log,
    trace_log,
    AgentState,
)

# ─── Load environment ─────────────────────────────────────────────────

load_dotenv()

LLM_CONFIG = {
    "azure_endpoint": os.getenv("AZURE_OPENAI_ENDPOINT", ""),
    "api_key": os.getenv("AZURE_OPENAI_API_KEY", ""),
    "azure_deployment": os.getenv("AZURE_OPENAI_DEPLOYMENT", "gpt-4o"),
    "api_version": os.getenv("AZURE_OPENAI_API_VERSION", "2024-12-01-preview"),
}


def _validate_config():
    missing = [k for k, v in {
        "AZURE_OPENAI_ENDPOINT": LLM_CONFIG["azure_endpoint"],
        "AZURE_OPENAI_API_KEY": LLM_CONFIG["api_key"],
    }.items() if not v]
    if missing:
        print(f"\n⚠  Warning: Missing env vars: {', '.join(missing)}")
        print("   Set them in .env or as environment variables before using the summarizer.\n")


_validate_config()

# ─── App setup ─────────────────────────────────────────────────────────

app = FastAPI(title="Supply Chain PPT Summarizer", version="1.0.0")

STATIC_DIR = Path(__file__).parent / "static"
UPLOAD_DIR = Path(__file__).parent / "uploads"
UPLOAD_DIR.mkdir(exist_ok=True)

app.mount("/static", StaticFiles(directory=str(STATIC_DIR)), name="static")

# ─── In-memory session store ──────────────────────────────────────────

sessions: dict[str, dict] = {}


def get_session(session_id: str) -> dict:
    if session_id not in sessions:
        raise HTTPException(status_code=404, detail="Session not found. Please upload a PPT first.")
    return sessions[session_id]


# ─── Routes: Pages ────────────────────────────────────────────────────

@app.get("/", response_class=HTMLResponse)
async def index_page():
    return (STATIC_DIR / "index.html").read_text()


@app.get("/tracing", response_class=HTMLResponse)
async def tracing_page():
    return (STATIC_DIR / "tracing.html").read_text()


@app.get("/tokens", response_class=HTMLResponse)
async def tokens_page():
    return (STATIC_DIR / "tokens.html").read_text()


# ─── Routes: API ──────────────────────────────────────────────────────

@app.get("/api/health")
async def health():
    configured = bool(LLM_CONFIG["azure_endpoint"] and LLM_CONFIG["api_key"])
    return {
        "status": "ok",
        "configured": configured,
        "deployment": LLM_CONFIG["azure_deployment"],
    }


@app.post("/api/upload")
async def upload_ppt(file: UploadFile = File(...)):
    if not LLM_CONFIG["azure_endpoint"] or not LLM_CONFIG["api_key"]:
        raise HTTPException(
            status_code=500,
            detail="Azure OpenAI not configured. Set AZURE_OPENAI_ENDPOINT and AZURE_OPENAI_API_KEY in .env"
        )

    if not file.filename.endswith((".pptx", ".ppt")):
        raise HTTPException(status_code=400, detail="Please upload a .pptx file.")

    session_id = str(uuid.uuid4())
    file_path = UPLOAD_DIR / f"{session_id}_{file.filename}"
    with open(file_path, "wb") as f:
        shutil.copyfileobj(file.file, f)

    try:
        parsed = parse_presentation(str(file_path))
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Failed to parse PPT: {str(e)}")

    sessions[session_id] = {
        "session_id": session_id,
        "llm_config": LLM_CONFIG,
        "filename": file.filename,
        "file_path": str(file_path),
        "parsed_ppt": parsed,
        "section_summaries": None,
        "executive_summary": None,
        "all_summaries_text": None,
        "conversation_history": [],
        "created_at": datetime.now(timezone.utc).isoformat(),
    }

    section_overview = []
    for s in parsed["sections"]:
        section_overview.append({
            "section_name": s["section_name"],
            "slide_count": s["slide_count"],
            "slide_numbers": s["slide_numbers"],
        })

    return {
        "session_id": session_id,
        "filename": file.filename,
        "total_slides": parsed["total_slides"],
        "total_sections": parsed["total_sections"],
        "sections": section_overview,
    }


@app.post("/api/summarize")
async def summarize(session_id: str = Form(...)):
    sess = get_session(session_id)
    graph = build_summarization_graph()

    initial_state: AgentState = {
        "session_id": session_id,
        "llm_config": sess["llm_config"],
        "parsed_ppt": sess["parsed_ppt"],
        "section_summaries": [],
        "executive_summary": "",
        "all_summaries_text": "",
    }

    try:
        final_state = await graph.ainvoke(initial_state)
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Summarization failed: {str(e)}")

    sess["section_summaries"] = final_state["section_summaries"]
    sess["executive_summary"] = final_state["executive_summary"]
    sess["all_summaries_text"] = final_state["all_summaries_text"]
    sess["conversation_history"].append({
        "role": "assistant",
        "content": final_state["executive_summary"],
        "timestamp": datetime.now(timezone.utc).isoformat(),
        "type": "initial_summary"
    })

    return {
        "session_id": session_id,
        "executive_summary": final_state["executive_summary"],
        "section_summaries": [
            {"section_name": s["section_name"], "summary": s["summary"]}
            for s in final_state["section_summaries"]
        ],
    }


@app.post("/api/refine")
async def refine(session_id: str = Form(...), instruction: str = Form(...)):
    sess = get_session(session_id)

    if not sess.get("executive_summary"):
        raise HTTPException(status_code=400, detail="No summary to refine. Run /api/summarize first.")

    sess["conversation_history"].append({
        "role": "user", "content": instruction,
        "timestamp": datetime.now(timezone.utc).isoformat(),
    })

    try:
        refined = await refine_summary(
            session_id=session_id,
            llm_config=sess["llm_config"],
            current_summary=sess["executive_summary"],
            section_summaries_text=sess["all_summaries_text"],
            user_instruction=instruction,
        )
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Refinement failed: {str(e)}")

    sess["executive_summary"] = refined
    sess["conversation_history"].append({
        "role": "assistant", "content": refined,
        "timestamp": datetime.now(timezone.utc).isoformat(), "type": "refined_summary"
    })

    return {"session_id": session_id, "executive_summary": refined}


@app.get("/api/session/{session_id}")
async def get_session_info(session_id: str):
    sess = get_session(session_id)
    return {
        "session_id": session_id,
        "filename": sess["filename"],
        "total_slides": sess["parsed_ppt"]["total_slides"],
        "total_sections": sess["parsed_ppt"]["total_sections"],
        "has_summary": sess["executive_summary"] is not None,
        "conversation_turns": len(sess["conversation_history"]),
        "created_at": sess["created_at"],
    }


@app.get("/api/tokens")
async def get_token_usage(session_id: str | None = None):
    entries = [e for e in token_usage_log if e["session_id"] == session_id] if session_id else token_usage_log
    tp = sum(e["prompt_tokens"] for e in entries)
    tc = sum(e["completion_tokens"] for e in entries)
    return {"entries": entries, "totals": {"prompt_tokens": tp, "completion_tokens": tc, "total_tokens": tp + tc, "entry_count": len(entries)}}


@app.get("/api/traces")
async def get_traces(session_id: str | None = None):
    entries = [e for e in trace_log if e["session_id"] == session_id] if session_id else trace_log
    td = sum(e["duration_ms"] for e in entries)
    return {"entries": entries, "totals": {"total_duration_ms": round(td, 2), "entry_count": len(entries)}}


@app.get("/api/sessions")
async def list_sessions():
    result = []
    for k, v in sessions.items():
        if isinstance(v, dict) and "session_id" in v:
            result.append({"session_id": v["session_id"], "filename": v.get("filename"), "created_at": v.get("created_at"), "has_summary": v.get("executive_summary") is not None})
    return {"sessions": result}


if __name__ == "__main__":
    import uvicorn
    host = os.getenv("APP_HOST", "0.0.0.0")
    port = int(os.getenv("APP_PORT", "8000"))
    uvicorn.run(app, host=host, port=port)
