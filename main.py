"""
Supply Chain PPT Summarizer - FastAPI Application
==================================================
Agentic web app for summarizing Global Supply Chain Status Reports.
Configuration loaded from .env file.
"""

import os
import json
import uuid
import shutil
from pathlib import Path
from datetime import datetime, timezone
from dotenv import load_dotenv

from fastapi import FastAPI, UploadFile, File, Form, HTTPException, Query
from fastapi.staticfiles import StaticFiles
from fastapi.responses import HTMLResponse, StreamingResponse

from ppt_parser import parse_presentation
from glossary import load_glossary_dir, load_glossary_file, normalise_json, render_glossary_for_prompt
from docx_export import markdown_to_docx
from agent import (
    build_summarization_graph,
    refine_summary,
    refine_email,
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

# ─── Load glossary ────────────────────────────────────────────────────

GLOSSARY_DIR = Path(os.getenv("GLOSSARY_DIR", str(Path(__file__).parent / "glossary")))
GLOSSARY_DIR.mkdir(exist_ok=True)

_glossary_data = load_glossary_dir(str(GLOSSARY_DIR))
glossary_entries: dict = _glossary_data.get("entries", {})
glossary_meta: list = _glossary_data.get("files_loaded", [])
glossary_prompt_text: str = render_glossary_for_prompt(glossary_entries)

print(f"📖 Glossary: {len(glossary_entries)} entries loaded from {len(glossary_meta)} file(s) in {GLOSSARY_DIR}")
if glossary_entries:
    for fm in glossary_meta:
        status = f"{fm['entries']} entries" if not fm.get("error") else f"ERROR: {fm['error']}"
        print(f"   • {fm['file']}: {status}")


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
        "glossary_context": glossary_prompt_text,
        "filename": file.filename,
        "file_path": str(file_path),
        "parsed_ppt": parsed,
        "section_summaries": None,
        "executive_summary": None,
        "email_summary": None,
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
        "glossary_context": sess.get("glossary_context", ""),
        "section_summaries": [],
        "executive_summary": "",
        "email_summary": "",
        "all_summaries_text": "",
    }

    try:
        final_state = await graph.ainvoke(initial_state)
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Summarization failed: {str(e)}")

    sess["section_summaries"] = final_state["section_summaries"]
    sess["executive_summary"] = final_state["executive_summary"]
    sess["email_summary"] = final_state["email_summary"]
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
        "email_summary": final_state["email_summary"],
        "section_summaries": [
            {"section_name": s["section_name"], "summary": s["summary"]}
            for s in final_state["section_summaries"]
        ],
    }


@app.post("/api/refine")
async def refine(
    session_id: str = Form(...),
    instruction: str = Form(...),
    target: str = Form("slides"),
):
    """Refine the executive summary or email summary. target = 'slides' | 'email'."""
    sess = get_session(session_id)

    if target == "email":
        if not sess.get("email_summary"):
            raise HTTPException(status_code=400, detail="No email summary to refine.")
    else:
        if not sess.get("executive_summary"):
            raise HTTPException(status_code=400, detail="No summary to refine. Run /api/summarize first.")

    sess["conversation_history"].append({
        "role": "user", "content": f"[{target}] {instruction}",
        "timestamp": datetime.now(timezone.utc).isoformat(),
    })

    try:
        if target == "email":
            refined = await refine_email(
                session_id=session_id,
                llm_config=sess["llm_config"],
                current_email=sess["email_summary"],
                section_summaries_text=sess["all_summaries_text"],
                user_instruction=instruction,
                glossary_context=sess.get("glossary_context", ""),
            )
            sess["email_summary"] = refined
        else:
            refined = await refine_summary(
                session_id=session_id,
                llm_config=sess["llm_config"],
                current_summary=sess["executive_summary"],
                section_summaries_text=sess["all_summaries_text"],
                user_instruction=instruction,
                glossary_context=sess.get("glossary_context", ""),
            )
            sess["executive_summary"] = refined
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Refinement failed: {str(e)}")

    sess["conversation_history"].append({
        "role": "assistant", "content": refined,
        "timestamp": datetime.now(timezone.utc).isoformat(),
        "type": f"refined_{target}"
    })

    return {"session_id": session_id, "target": target, "content": refined}


@app.get("/api/download")
async def download_docx(
    session_id: str = Query(...),
    target: str = Query("slides"),
):
    """Download the slide summary or email summary as a formatted .docx file."""
    sess = get_session(session_id)

    if target == "email":
        content = sess.get("email_summary")
        if not content:
            raise HTTPException(status_code=400, detail="No email summary available.")
        title = "Crisis Status Email Summary"
        filename = "email_summary.docx"
    else:
        content = sess.get("executive_summary")
        if not content:
            raise HTTPException(status_code=400, detail="No slide summary available.")
        title = "Executive Summary — Slide Content"
        filename = "slide_summary.docx"

    # Prefix with source file info
    source_name = sess.get("filename", "")
    header = f"Source: {source_name}\n\n---\n\n" if source_name else ""

    buf = markdown_to_docx(header + content, title=title)

    return StreamingResponse(
        buf,
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        headers={"Content-Disposition": f'attachment; filename="{filename}"'},
    )


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


# ─── Routes: Glossary ────────────────────────────────────────────────

@app.get("/api/glossary")
async def get_glossary():
    """Return current glossary entries and metadata."""
    # Group by category for display
    by_category: dict[str, list] = {}
    for abbr, info in sorted(glossary_entries.items()):
        cat = info.get("category", "general")
        by_category.setdefault(cat, []).append({"abbr": abbr, "meaning": info["meaning"]})

    return {
        "total_entries": len(glossary_entries),
        "files_loaded": glossary_meta,
        "categories": {cat: len(items) for cat, items in by_category.items()},
        "entries_by_category": by_category,
    }


@app.post("/api/glossary/upload")
async def upload_glossary(file: UploadFile = File(...)):
    """Upload an additional glossary JSON file. Merges into the active glossary."""
    global glossary_entries, glossary_prompt_text, glossary_meta

    if not file.filename.endswith(".json"):
        raise HTTPException(status_code=400, detail="Please upload a .json file.")

    file_path = GLOSSARY_DIR / file.filename
    with open(file_path, "wb") as f:
        shutil.copyfileobj(file.file, f)

    try:
        new_entries = load_glossary_file(str(file_path))
    except Exception as e:
        file_path.unlink(missing_ok=True)
        raise HTTPException(status_code=400, detail=f"Failed to parse glossary JSON: {str(e)}")

    # Merge into global glossary
    added_count = len([k for k in new_entries if k not in glossary_entries])
    updated_count = len([k for k in new_entries if k in glossary_entries])
    glossary_entries.update(new_entries)
    glossary_prompt_text = render_glossary_for_prompt(glossary_entries)
    glossary_meta.append({"file": file.filename, "entries": len(new_entries)})

    # Update any active sessions with the new glossary
    for sess in sessions.values():
        if isinstance(sess, dict) and "glossary_context" in sess:
            sess["glossary_context"] = glossary_prompt_text

    return {
        "filename": file.filename,
        "new_entries": added_count,
        "updated_entries": updated_count,
        "total_entries": len(glossary_entries),
    }


@app.delete("/api/glossary/{filename}")
async def delete_glossary_file(filename: str):
    """Remove a glossary file and reload all remaining files."""
    global glossary_entries, glossary_prompt_text, glossary_meta

    file_path = GLOSSARY_DIR / filename
    if not file_path.exists():
        raise HTTPException(status_code=404, detail=f"Glossary file not found: {filename}")

    file_path.unlink()

    # Reload everything from the directory
    _data = load_glossary_dir(str(GLOSSARY_DIR))
    glossary_entries = _data.get("entries", {})
    glossary_meta = _data.get("files_loaded", [])
    glossary_prompt_text = render_glossary_for_prompt(glossary_entries)

    # Update active sessions
    for sess in sessions.values():
        if isinstance(sess, dict) and "glossary_context" in sess:
            sess["glossary_context"] = glossary_prompt_text

    return {"deleted": filename, "total_entries": len(glossary_entries)}


if __name__ == "__main__":
    import uvicorn
    host = os.getenv("APP_HOST", "0.0.0.0")
    port = int(os.getenv("APP_PORT", "8000"))
    uvicorn.run(app, host=host, port=port)
