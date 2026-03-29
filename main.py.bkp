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
from gfd_llm_parser import parse_gfd_with_llm
from gfd_llm_slides import generate_gfd_dashboard
from gfd_docx_export import gfd_spec_to_docx
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
gfd_sessions: dict[str, dict] = {}


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


@app.get("/gfd", response_class=HTMLResponse)
async def gfd_page():
    return (STATIC_DIR / "gfd.html").read_text()


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

# ─── Routes: GFD Dashboard ──────────────────────────────────────────

@app.post("/api/gfd/upload")
async def upload_gfd(file: UploadFile = File(...), history_weeks: int = Form(4)):
    """
    Upload a Dashboard_Update Excel file — runs Stages 1 + 2 only.

    Stage 1 — deterministic: Excel → filtered pipe-delimited text table
    Stage 2 — LLM extraction: text table → structured JSON (product groups, CW integers)

    Returns the extracted JSON for user review before slide generation.
    """
    if not file.filename.endswith((".xlsx", ".xls", ".xlsm")):
        raise HTTPException(status_code=400, detail="Please upload an .xlsx file.")

    session_id = str(uuid.uuid4())
    file_path = UPLOAD_DIR / f"{session_id}_{file.filename}"
    with open(file_path, "wb") as f:
        shutil.copyfileobj(file.file, f)

    # ── Stages 1 + 2: Excel → LLM-extracted JSON ─────────────────────
    try:
        extracted = await parse_gfd_with_llm(
            filepath=str(file_path),
            llm_config=LLM_CONFIG,
            session_id=session_id,
            history_weeks=history_weeks,
            glossary_context=glossary_prompt_text,
        )
    except Exception as e:
        raise HTTPException(
            status_code=500,
            detail=f"Excel parsing / LLM extraction failed: {str(e)}"
        )

    # ── Store session (no slide_spec yet) ─────────────────────────────
    gfd_sessions[session_id] = {
        "session_id": session_id,
        "filename": file.filename,
        "file_path": str(file_path),
        "pptx_path": None,
        "extracted": extracted,
        "slide_spec": None,
        "created_at": datetime.now(timezone.utc).isoformat(),
    }

    # ── Build response with full extracted data for review ────────────
    meta = extracted.get("_meta", {})

    # Build a clean copy of the extracted JSON for frontend display
    # (exclude internal _meta, keep everything the LLM produced)
    extracted_for_display = {k: v for k, v in extracted.items() if k != "_meta"}

    return {
        "session_id": session_id,
        "filename": file.filename,
        "current_cw": extracted.get("current_cw", ""),
        "total_rows_in_file": meta.get("total_rows_in_file", 0),
        "rows_after_filter": meta.get("rows_after_filter", 0),
        "rows_extracted_by_llm": meta.get("rows_extracted_by_llm", 0),
        "extraction_chunks": meta.get("extraction_chunks", 1),
        "extraction_chunks_failed": meta.get("extraction_chunks_failed", 0),
        "product_groups": extracted.get("product_groups", []),
        "extraction_notes": extracted.get("extraction_notes", ""),
        "warnings": extracted.get("warnings", []),
        "extracted_json": extracted_for_display,
    }


@app.post("/api/gfd/generate")
async def generate_gfd(session_id: str = Form(...), format: str = Form("pptx")):
    """
    Generate the GFD dashboard from previously extracted data (Stage 3 + 4).

    Runs the LLM slide-spec generation and PPTX/DOCX rendering.
    Must be called after /api/gfd/upload.

    Form fields:
      session_id : from the upload response
      format     : "pptx" or "docx" (default: "pptx")
    """
    if session_id not in gfd_sessions:
        raise HTTPException(status_code=404, detail="GFD session not found. Please upload first.")

    sess = gfd_sessions[session_id]
    extracted = sess.get("extracted")
    if not extracted:
        raise HTTPException(status_code=400, detail="No extracted data in session.")

    # ── Run Stages 3 + 4 if not already generated ────────────────────
    if sess.get("slide_spec") is None:
        pptx_path = str(UPLOAD_DIR / f"{session_id}_gfd_dashboard.pptx")
        try:
            buf, slide_spec = await generate_gfd_dashboard(
                extracted_data=extracted,
                llm_config=LLM_CONFIG,
                session_id=session_id,
                output_path=pptx_path,
                glossary_context=glossary_prompt_text,
            )
        except Exception as e:
            raise HTTPException(
                status_code=500,
                detail=f"Slide generation failed: {str(e)}"
            )
        sess["slide_spec"] = slide_spec
        sess["pptx_path"] = pptx_path

    slide_spec = sess["slide_spec"]
    meta = extracted.get("_meta", {})

    pg_overview = [
        {
            "product_family": pg.get("product_family_desc", ""),
            "code": pg.get("product_family_code", ""),
            "risk_items": len(pg.get("rows", [])),
        }
        for pg in extracted.get("product_groups", [])
    ]

    return {
        "session_id": session_id,
        "filename": sess["filename"],
        "current_cw": extracted.get("current_cw", ""),
        "product_groups": pg_overview,
        "overall_risk": slide_spec.get("overall_risk", ""),
        "slide_count": slide_spec.get("slide_count", len(slide_spec.get("slides", []))),
        "is_fallback": slide_spec.get("_fallback", False),
    }


@app.get("/api/gfd/download")
async def download_gfd_pptx(session_id: str = Query(...)):
    """Download the generated GFD dashboard as a .pptx file."""
    if session_id not in gfd_sessions:
        raise HTTPException(status_code=404, detail="GFD session not found.")

    sess = gfd_sessions[session_id]
    if not sess.get("pptx_path"):
        raise HTTPException(status_code=400, detail="Dashboard not yet generated. Call /api/gfd/generate first.")

    pptx_path = Path(sess["pptx_path"])

    if not pptx_path.exists():
        raise HTTPException(status_code=500, detail="Generated PPTX file not found.")

    source_name = sess.get("filename", "dashboard")
    download_name = f"GFD_Dashboard_{source_name.replace('.xlsx', '')}.pptx"

    return StreamingResponse(
        open(pptx_path, "rb"),
        media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        headers={"Content-Disposition": f'attachment; filename="{download_name}"'},
    )


@app.get("/api/gfd/download/docx")
async def download_gfd_docx(session_id: str = Query(...)):
    """Download the GFD dashboard as a formatted .docx Word document."""
    if session_id not in gfd_sessions:
        raise HTTPException(status_code=404, detail="GFD session not found.")

    sess = gfd_sessions[session_id]
    slide_spec = sess.get("slide_spec")
    if not slide_spec:
        raise HTTPException(status_code=400, detail="Dashboard not yet generated. Call /api/gfd/generate first.")

    buf = gfd_spec_to_docx(slide_spec)

    source_name = sess.get("filename", "dashboard")
    download_name = f"GFD_Dashboard_{source_name.replace('.xlsx', '')}.docx"

    return StreamingResponse(
        buf,
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        headers={"Content-Disposition": f'attachment; filename="{download_name}"'},
    )


@app.get("/api/gfd/session/{session_id}")
async def get_gfd_session(session_id: str):
    """Return metadata for a GFD session."""
    if session_id not in gfd_sessions:
        raise HTTPException(status_code=404, detail="GFD session not found.")

    sess = gfd_sessions[session_id]
    extracted = sess.get("extracted", {})
    slide_spec = sess.get("slide_spec") or {}
    meta = extracted.get("_meta", {})

    return {
        "session_id": session_id,
        "filename": sess["filename"],
        "created_at": sess["created_at"],
        "current_cw": extracted.get("current_cw", ""),
        "extraction_notes": extracted.get("extraction_notes", ""),
        "warnings": extracted.get("warnings", []),
        "overall_risk": slide_spec.get("overall_risk", ""),
        "slide_count": slide_spec.get("slide_count", 0),
        "is_fallback": slide_spec.get("_fallback", False),
        "is_generated": sess.get("slide_spec") is not None,
        "product_groups": [
            {
                "product_family": pg.get("product_family_desc", ""),
                "code": pg.get("product_family_code", ""),
                "risk_items": len(pg.get("rows", [])),
            }
            for pg in extracted.get("product_groups", [])
        ],
        "metadata": {
            "total_rows_in_file": meta.get("total_rows_in_file", 0),
            "rows_after_filter": meta.get("rows_after_filter", 0),
            "rows_extracted_by_llm": meta.get("rows_extracted_by_llm", 0),
            "extraction_chunks": meta.get("extraction_chunks", 1),
            "extraction_chunks_failed": meta.get("extraction_chunks_failed", 0),
            "sheet_used": meta.get("sheet_used", ""),
            "headers_detected": len(meta.get("headers", [])),
        },
    }


if __name__ == "__main__":
    import uvicorn
    host = os.getenv("APP_HOST", "0.0.0.0")
    port = int(os.getenv("APP_PORT", "8000"))
    uvicorn.run(app, host=host, port=port)
