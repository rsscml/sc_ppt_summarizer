"""
main.py — CHANGED SECTIONS ONLY
================================
Replace the indicated blocks in your existing main.py.
Everything else (PPT summarizer routes, glossary routes, token/trace routes) stays unchanged.

CHANGE 1 ─ Top-of-file imports
CHANGE 2 ─ /api/gfd/upload route
CHANGE 3 ─ /api/gfd/session/{session_id} route

════════════════════════════════════════════════════════════════════════
CHANGE 1: Replace the old GFD import block with these two lines.
Remove the old:
    from gfd_excel_parser import parse_dashboard_update, summarise_for_prompt, filter_by_recency
    from gfd_slide_generator import generate_gfd_slides
    from gfd_agent import run_gfd_pipeline          ← if present
════════════════════════════════════════════════════════════════════════
"""

from gfd_llm_parser import parse_gfd_with_llm
from gfd_llm_slides import generate_gfd_dashboard


# ════════════════════════════════════════════════════════════════════════
# CHANGE 2: Replace the entire /api/gfd/upload route with this.
# ════════════════════════════════════════════════════════════════════════

@app.post("/api/gfd/upload")
async def upload_gfd(file: UploadFile = File(...), history_weeks: int = Form(4)):
    """
    Upload a Dashboard_Update Excel file.

    Pipeline:
      Stage 1 — deterministic: Excel → filtered pipe-delimited text table
      Stage 2 — LLM extraction: text table → structured JSON (product groups, CW integers)
      Stage 3 — LLM slide spec: extracted JSON → complete slide specification
      Stage 4 — PPTX renderer: slide spec → .pptx file
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

    # ── Stages 3 + 4: LLM slide spec → PPTX ─────────────────────────
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

    # ── Store session ─────────────────────────────────────────────────
    gfd_sessions[session_id] = {
        "session_id":  session_id,
        "filename":    file.filename,
        "file_path":   str(file_path),
        "pptx_path":   pptx_path,
        "extracted":   extracted,
        "slide_spec":  slide_spec,
        "created_at":  datetime.now(timezone.utc).isoformat(),
    }

    # ── Build response summary ────────────────────────────────────────
    meta = extracted.get("_meta", {})
    pg_overview = [
        {
            "product_family": pg.get("product_family_desc", ""),
            "code":           pg.get("product_family_code", ""),
            "risk_items":     len(pg.get("rows", [])),
        }
        for pg in extracted.get("product_groups", [])
    ]

    overview_slide = next(
        (s for s in slide_spec.get("slides", []) if s.get("type") == "overview"),
        {}
    )

    return {
        "session_id":        session_id,
        "filename":          file.filename,
        "current_cw":        extracted.get("current_cw", ""),
        "total_rows_in_file": meta.get("total_rows_in_file", 0),
        "rows_after_filter": meta.get("rows_after_filter", 0),
        "history_weeks":     history_weeks,
        "product_groups":    pg_overview,
        "overall_risk":      overview_slide.get("overall_risk", ""),
        "slide_count":       len(slide_spec.get("slides", [])),
        "warnings":          extracted.get("warnings", []),
        "extraction_notes":  extracted.get("extraction_notes", ""),
        "is_fallback":       slide_spec.get("_fallback", False),
    }


# ════════════════════════════════════════════════════════════════════════
# CHANGE 3: Replace the /api/gfd/session/{session_id} route with this.
# ════════════════════════════════════════════════════════════════════════

@app.get("/api/gfd/session/{session_id}")
async def get_gfd_session(session_id: str):
    """Return metadata for a GFD session."""
    if session_id not in gfd_sessions:
        raise HTTPException(status_code=404, detail="GFD session not found.")

    sess       = gfd_sessions[session_id]
    extracted  = sess.get("extracted", {})
    slide_spec = sess.get("slide_spec", {})
    meta       = extracted.get("_meta", {})

    overview_slide = next(
        (s for s in slide_spec.get("slides", []) if s.get("type") == "overview"),
        {}
    )

    return {
        "session_id":        session_id,
        "filename":          sess["filename"],
        "created_at":        sess["created_at"],
        "current_cw":        extracted.get("current_cw", ""),
        "extraction_notes":  extracted.get("extraction_notes", ""),
        "warnings":          extracted.get("warnings", []),
        "overall_risk":      overview_slide.get("overall_risk"),
        "slide_count":       len(slide_spec.get("slides", [])),
        "is_fallback":       slide_spec.get("_fallback", False),
        "product_groups": [
            {
                "product_family": pg.get("product_family_desc", ""),
                "code":           pg.get("product_family_code", ""),
                "risk_items":     len(pg.get("rows", [])),
            }
            for pg in extracted.get("product_groups", [])
        ],
        "metadata": {
            "total_rows_in_file": meta.get("total_rows_in_file", 0),
            "rows_after_filter":  meta.get("rows_after_filter", 0),
            "sheet_used":         meta.get("sheet_used", ""),
            "headers_detected":   len(meta.get("headers", [])),
        },
    }


# ════════════════════════════════════════════════════════════════════════
# /api/gfd/download stays UNCHANGED — it reads pptx_path from the session,
# which is set identically in the new upload route above.
# ════════════════════════════════════════════════════════════════════════
