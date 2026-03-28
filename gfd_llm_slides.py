"""
GFD LLM Slide Generator
=======================
Two-stage pipeline:

Stage 1 — LLM slide spec:
  Receives the structured JSON from gfd_llm_parser and generates a complete
  slide specification — RAG colors, condensed cell text, groupings, overview
  bullets, risk assessments. The LLM makes ALL content and layout decisions.

Stage 2 — PPTX renderer:
  A thin "paint by numbers" renderer that converts the slide spec JSON into a
  python-pptx Presentation object. No business logic lives here — the renderer
  just maps JSON fields to shapes, colors, and text runs.

If Stage 1 fails, a deterministic fallback computes RAG colors arithmetically
from the integer coverage CW fields produced by the extractor, so the user
always gets a working dashboard even on LLM error.
"""

from __future__ import annotations

import io
import json
import time
from datetime import date, datetime
from typing import Any

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.oxml.ns import qn
from lxml import etree

from langchain_openai import AzureChatOpenAI
from langchain_core.messages import HumanMessage, SystemMessage

from agent import log_tokens, log_trace


# ─── LLM factory (local, higher max_tokens for large JSON output) ─────

def _create_llm(config: dict, max_tokens: int = 8192) -> AzureChatOpenAI:
    return AzureChatOpenAI(
        azure_deployment=config["azure_deployment"],
        azure_endpoint=config["azure_endpoint"],
        api_key=config["api_key"],
        api_version=config.get("api_version", "2024-12-01-preview"),
        temperature=0.15,
        max_tokens=max_tokens,
    )


# ─── Color palette ────────────────────────────────────────────────────

_C: dict[str, RGBColor] = {
    "GREEN":      RGBColor(0x00, 0xB0, 0x50),
    "AMBER":      RGBColor(0xFF, 0xC0, 0x00),
    "RED":        RGBColor(0xC0, 0x00, 0x00),
    "GREY":       RGBColor(0xA0, 0xA0, 0xA0),
    "WHITE":      RGBColor(0xFF, 0xFF, 0xFF),
    "BLACK":      RGBColor(0x1A, 0x1A, 0x1A),
    "NAVY":       RGBColor(0x1E, 0x27, 0x61),
    "BLUE":       RGBColor(0x2C, 0x52, 0x82),
    "LIGHT_BLUE": RGBColor(0xBF, 0xD7, 0xED),
    "LIGHT_GREY": RGBColor(0xF4, 0xF4, 0xF6),
    "MEDIUM_ORANGE": RGBColor(0xFF, 0x99, 0x00),
}

_RAG_FG: dict[str, RGBColor] = {
    "GREEN": _C["WHITE"],
    "AMBER": _C["BLACK"],
    "RED":   _C["WHITE"],
    "GREY":  _C["WHITE"],
}

_RISK_BG: dict[str, RGBColor] = {
    "CRITICAL": _C["RED"],
    "HIGH":     _C["AMBER"],
    "MEDIUM":   _C["MEDIUM_ORANGE"],
    "LOW":      _C["GREEN"],
}
_RISK_FG: dict[str, RGBColor] = {
    "CRITICAL": _C["WHITE"],
    "HIGH":     _C["BLACK"],
    "MEDIUM":   _C["BLACK"],
    "LOW":      _C["WHITE"],
}

# Slide dimensions — widescreen 16:9
_W  = Inches(13.333)
_H  = Inches(7.5)
_M  = Inches(0.28)     # side margin
_HH = Inches(0.62)     # header bar height

FONT = "Calibri"


# ─── Low-level python-pptx helpers ───────────────────────────────────

def _solid(shape, rgb: RGBColor) -> None:
    """Solid fill + no border on a shape."""
    shape.fill.solid()
    shape.fill.fore_color.rgb = rgb
    shape.line.fill.background()


def _rect(slide, x, y, w, h, rgb: RGBColor):
    """Add a filled rectangle (MSO_AUTO_SHAPE_TYPE.RECTANGLE = 1)."""
    s = slide.shapes.add_shape(1, int(x), int(y), int(w), int(h))
    _solid(s, rgb)
    return s


def _tb(slide, x, y, w, h, text: str, *,
        size: int = 10, bold: bool = False,
        color: RGBColor = None, align=PP_ALIGN.LEFT,
        bg: RGBColor = None) -> Any:
    """Add a textbox. bg=None → transparent fill."""
    tb = slide.shapes.add_textbox(int(x), int(y), int(w), int(h))
    if bg:
        tb.fill.solid()
        tb.fill.fore_color.rgb = bg
    else:
        tb.fill.background()
    tf = tb.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.alignment = align
    run = p.add_run()
    run.text = text or ""
    run.font.name = FONT
    run.font.size = Pt(size)
    run.font.bold = bold
    if color:
        run.font.color.rgb = color
    return tb


def _cell_fill(cell, rgb: RGBColor) -> None:
    """Set solid background fill on a table cell via XML."""
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    # Remove any existing fill elements
    for child in list(tcPr):
        local = child.tag.split("}")[-1] if "}" in child.tag else child.tag
        if local in ("solidFill", "gradFill", "noFill", "pattFill", "blipFill"):
            tcPr.remove(child)
    sf = etree.SubElement(tcPr, qn("a:solidFill"))
    clr = etree.SubElement(sf, qn("a:srgbClr"))
    clr.set("val", f"{rgb[0]:02X}{rgb[1]:02X}{rgb[2]:02X}")


def _cell_write(cell, text: str, *,
                size: int = 7, bold: bool = False,
                color: RGBColor = None, align=PP_ALIGN.LEFT) -> None:
    """Write text into a table cell, clearing previous content."""
    tf = cell.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.alignment = align
    # Remove extra runs from first paragraph
    for r in list(p.runs):
        p._p.remove(r._r)
    # Remove extra paragraphs
    for extra in list(tf.paragraphs)[1:]:
        extra._p.getparent().remove(extra._p)
    run = p.add_run()
    run.text = str(text) if text is not None else ""
    run.font.name = FONT
    run.font.size = Pt(size)
    run.font.bold = bold
    if color:
        run.font.color.rgb = color


# ─── Stage 1: LLM slide spec ─────────────────────────────────────────

_SLIDES_SYSTEM = """\
You are a Chief Supply Chain Officer designing a Global Fulfilment Dashboard PowerPoint
presentation for a board-level audience. You are given JSON extracted from the
Dashboard_Update worksheet. Generate a COMPLETE slide specification in JSON.

═══ SLIDE DESIGN RULES ═══

SLIDE 1 — Executive Overview (type: "overview")
  • overall_risk: worst-case across ALL rows (CRITICAL > HIGH > MEDIUM > LOW).
  • 4–6 bullet points (max 110 chars each). Lead with the most critical items.
    Every bullet must cite specific facts: CW numbers, plant names, counts.
  • stats: count risk items by risk level.

SLIDES 2+ — One slide per product group (type: "product_group")
  • If a product group has >9 rows, split it into multiple slides with a
    "(cont.)" suffix on the title.
  • title: "Product Group Desc (CODE)" — use actual codes and descriptions.
  • headline: 1-sentence situation summary (max 80 chars). Be specific.
  • cw_columns: list of exactly 12 consecutive CW integers starting from
    the current CW number (e.g. if current is CW13 → [13,14,15,16,17,18,19,20,21,22,23,24]).
    Handle year-end wrap correctly (e.g. CW51,52,1,2,...).
  • quarter_label: next quarter label (e.g. "Q2/2026").

RAG COLOR COMPUTATION (per row, per CW column):
  Given coverage_without_mitigation_cw (W) and coverage_with_mitigation_cw (M):
    • "GREEN" if CW ≤ W  (supply secured without mitigation)
    • "AMBER" if W < CW ≤ M  (supply depends on mitigation actions)
    • "RED"   if CW > M  (no supply plan in place)
    • If M is null but W is set: "GREEN" if CW ≤ W, else "RED"
    • If both are null: "GREY"
  quarter_color = worst RAG across all 13 weeks of that quarter
  (RED > AMBER > GREEN > GREY).

RISK LEVEL (per row):
  CRITICAL = any CW in cw_columns is RED
  HIGH     = at least one AMBER, no RED
  MEDIUM   = all GREEN but ≤ 3 week safety margin beyond the last CW column
  LOW      = comfortably all GREEN

TEXT CONDENSING for slide display (applied to each row field):
  component ≤ 45 chars, plant ≤ 15 chars, customer ≤ 40 chars,
  supplier ≤ 30 chars, action ≤ 90 chars, fm_status ≤ 15 chars.

═══ OUTPUT FORMAT ═══

Respond with ONLY valid JSON — no markdown fences, no explanation:

{{
  "presentation_title": "Global Fulfilment Dashboard",
  "current_cw": "CW13/2026",
  "generated_date": "{today}",
  "slides": [
    {{
      "type": "overview",
      "title": "Situation Overview — CW13/2026",
      "overall_risk": "HIGH",
      "bullets": ["3 plants at RED from CW16 — BMW Group coverage expires CW15", "..."],
      "stats": {{
        "total_items": 12,
        "critical_count": 2,
        "high_count": 4,
        "medium_count": 4,
        "low_count": 2,
        "product_groups_count": 3
      }}
    }},
    {{
      "type": "product_group",
      "title": "Sensors / Radar (SEN)",
      "headline": "NXP wafer shortage — BHV at RED from CW16, mitigation active",
      "cw_columns": [13,14,15,16,17,18,19,20,21,22,23,24],
      "quarter_label": "Q2/2026",
      "rows": [
        {{
          "plant": "BHV",
          "customer": "BMW Group, Mercedes-Benz",
          "component": "NXP S32K wafer",
          "supplier": "NXP Semiconductors",
          "coverage_without": 15,
          "coverage_with": 19,
          "action": "Dual source by CW14; air freight bridge until CW18",
          "fm_status": "In progress",
          "risk_level": "HIGH",
          "cw_colors": {{"13":"GREEN","14":"GREEN","15":"GREEN","16":"AMBER","17":"AMBER","18":"AMBER","19":"AMBER","20":"RED","21":"RED","22":"RED","23":"RED","24":"RED"}},
          "quarter_color": "RED"
        }}
      ]
    }}
  ]
}}
{glossary_block}"""

_SLIDES_USER = """\
Today: {today}
Current calendar week: {current_cw}

Generate the complete dashboard slide specification from this extracted data:

{extracted_json}"""


async def llm_generate_slide_spec(
    extracted: dict,
    llm_config: dict,
    session_id: str,
    glossary_context: str = "",
) -> dict:
    """
    Stage 1: LLM generates a complete slide specification from extracted data.

    Returns a dict with a "slides" list. Falls back to a deterministic spec on error.
    """
    llm = _create_llm(llm_config, max_tokens=8192)
    t0 = time.time()
    today = date.today().isoformat()
    current_cw = extracted.get("current_cw", "CW??/????")

    glossary_block = (
        f"\n\nCOMPANY GLOSSARY:\n{glossary_context}" if glossary_context else ""
    )

    # Exclude internal metadata keys from what the LLM sees
    clean = {k: v for k, v in extracted.items() if not k.startswith("_")}
    extracted_json = json.dumps(clean, indent=2, ensure_ascii=False)

    messages = [
        SystemMessage(content=_SLIDES_SYSTEM.format(
            today=today,
            glossary_block=glossary_block,
        )),
        HumanMessage(content=_SLIDES_USER.format(
            today=today,
            current_cw=current_cw,
            extracted_json=extracted_json,
        )),
    ]

    try:
        response = await llm.ainvoke(messages)
        raw = response.content.strip()

        if raw.startswith("```"):
            raw = "\n".join(raw.split("\n")[1:])
        if raw.endswith("```"):
            raw = "\n".join(raw.split("\n")[:-1])

        spec: dict = json.loads(raw.strip())

        usage = response.response_metadata.get("token_usage", {})
        log_tokens(session_id, "gfd_llm_slide_spec", usage, llm_config.get("azure_deployment", ""))

        n_slides = len(spec.get("slides", []))
        duration = (time.time() - t0) * 1000
        log_trace(
            session_id, "gfd_llm_slide_spec",
            f"Input: {len(extracted.get('product_groups', []))} product groups",
            f"Generated {n_slides} slide specs",
            duration,
        )
        return spec

    except json.JSONDecodeError as exc:
        duration = (time.time() - t0) * 1000
        log_trace(session_id, "gfd_llm_slide_spec",
                  "Generating slide spec",
                  f"JSON PARSE ERROR — using deterministic fallback: {str(exc)[:100]}",
                  duration, {"error": True, "fallback": True})
        return _deterministic_fallback_spec(extracted)

    except Exception as exc:
        duration = (time.time() - t0) * 1000
        log_trace(session_id, "gfd_llm_slide_spec",
                  "Generating slide spec",
                  f"ERROR — using deterministic fallback: {str(exc)[:100]}",
                  duration, {"error": True, "fallback": True})
        return _deterministic_fallback_spec(extracted)


# ─── Deterministic fallback spec ─────────────────────────────────────
# Used only when the slide-spec LLM call fails. Computes RAG colors
# arithmetically from the integer coverage CW fields the extractor produced.

def _cw_color(cw_num: int, cov_wo: int | None, cov_w: int | None) -> str:
    if cov_wo is None and cov_w is None:
        return "GREY"
    boundary_w = cov_w or cov_wo
    boundary_wo = cov_wo or cov_w
    if cw_num <= (boundary_wo or 0):
        return "GREEN"
    if cw_num <= (boundary_w or 0):
        return "AMBER"
    return "RED"


def _worst_rag(colors: list[str]) -> str:
    rank = {"GREY": 0, "GREEN": 1, "AMBER": 2, "RED": 3}
    return max(colors, key=lambda c: rank.get(c, 0), default="GREY")


def _deterministic_fallback_spec(extracted: dict) -> dict:
    """
    Fallback: compute slide spec without LLM using arithmetic RAG logic.
    Produces a correct dashboard but with less polished text than the LLM version.
    """
    from datetime import datetime as _dt

    current_cw_str = extracted.get("current_cw", "CW1/2026")
    m = __import__("re").match(r"CW(\d+)/(\d+)", current_cw_str)
    if m:
        cw_start, cw_year = int(m.group(1)), int(m.group(2))
    else:
        iso = _dt.now().isocalendar()
        cw_start, cw_year = iso.week, iso.year

    # 12 consecutive CW numbers with year-end wrap
    cw_columns: list[int] = []
    for i in range(12):
        w = ((cw_start - 1 + i) % 52) + 1
        cw_columns.append(w)

    # Quarter for the last CW in the range
    last_cw = cw_columns[-1]
    q_num = (last_cw - 1) // 13 + 2  # next quarter after current
    q_label = f"Q{min(q_num, 4)}/{cw_year}"

    all_risk_levels = []
    pg_slides = []

    for pg in extracted.get("product_groups", []):
        code = pg.get("product_family_code", "")
        desc = pg.get("product_family_desc", "Unknown")
        title = f"{desc} ({code})" if code else desc

        slide_rows = []
        for row in pg.get("rows", []):
            cov_wo: int | None = row.get("coverage_without_mitigation_cw")
            cov_w:  int | None = row.get("coverage_with_mitigation_cw")

            cw_colors = {str(cw): _cw_color(cw, cov_wo, cov_w) for cw in cw_columns}
            all_colors = list(cw_colors.values())
            quarter_colors = [_cw_color(w, cov_wo, cov_w)
                              for w in range((q_num - 2) * 13 + 1, (q_num - 1) * 13 + 1)
                              if 1 <= w <= 52]
            quarter_color = _worst_rag(quarter_colors) if quarter_colors else _worst_rag(all_colors)

            worst = _worst_rag(all_colors)
            risk_level = {"RED": "CRITICAL", "AMBER": "HIGH", "GREEN": "LOW", "GREY": "LOW"}.get(worst, "MEDIUM")
            all_risk_levels.append(risk_level)

            slide_rows.append({
                "plant":           (row.get("plant_location") or "")[:15],
                "customer":        (row.get("customer_affected") or "")[:40],
                "component":       (row.get("critical_component") or "")[:45],
                "supplier":        (row.get("supplier_text") or "")[:30],
                "coverage_without": cov_wo,
                "coverage_with":    cov_w,
                "action":          (row.get("action_comment") or "")[:90],
                "fm_status":       str(row.get("customer_informed") or "N/A")[:15],
                "risk_level":      risk_level,
                "cw_colors":       cw_colors,
                "quarter_color":   quarter_color,
            })

        # Paginate at 9 rows per slide
        for page_i in range(0, max(len(slide_rows), 1), 9):
            chunk = slide_rows[page_i: page_i + 9]
            suffix = f" (cont. {page_i // 9 + 1})" if page_i > 0 else ""
            pg_slides.append({
                "type":          "product_group",
                "title":         title + suffix,
                "headline":      f"{len(chunk)} risk item(s) — see CW coverage grid",
                "cw_columns":    cw_columns,
                "quarter_label": q_label,
                "rows":          chunk,
            })

    # Build overview stats
    risk_rank = {"LOW": 0, "MEDIUM": 1, "HIGH": 2, "CRITICAL": 3}
    overall_risk = max(all_risk_levels, key=lambda r: risk_rank.get(r, 0), default="HIGH")

    total = sum(len(pg.get("rows", [])) for pg in extracted.get("product_groups", []))

    overview = {
        "type":         "overview",
        "title":        f"Situation Overview — {current_cw_str}",
        "overall_risk": overall_risk,
        "bullets": [
            f"{len(extracted.get('product_groups', []))} product group(s) with active fulfilment risks",
            f"{total} risk items tracked — see per-group slides for CW coverage detail",
            "(LLM slide generation used deterministic fallback — reprocess to get AI narrative)",
        ],
        "stats": {
            "total_items":          total,
            "critical_count":       all_risk_levels.count("CRITICAL"),
            "high_count":           all_risk_levels.count("HIGH"),
            "medium_count":         all_risk_levels.count("MEDIUM"),
            "low_count":            all_risk_levels.count("LOW"),
            "product_groups_count": len(extracted.get("product_groups", [])),
        },
    }

    return {
        "presentation_title": "Global Fulfilment Dashboard",
        "current_cw":         current_cw_str,
        "generated_date":     date.today().isoformat(),
        "_fallback":          True,
        "slides":             [overview] + pg_slides,
    }


# ─── Stage 2: PPTX renderer ──────────────────────────────────────────

def _render_overview(prs: Presentation, spec: dict) -> None:
    """Render the executive overview slide."""
    slide = prs.slides.add_slide(prs.slide_layouts[6])  # blank

    title        = spec.get("title", "Global Fulfilment Dashboard")
    overall_risk = spec.get("overall_risk", "HIGH").upper()
    bullets      = spec.get("bullets", [])
    stats        = spec.get("stats", {})

    # ── Header bar ────────────────────────────────────────────────────
    _rect(slide, 0, 0, _W, _HH, _C["NAVY"])
    _tb(slide, _M, Inches(0.1), Inches(10.5), Inches(0.45),
        title, size=20, bold=True, color=_C["WHITE"])

    # Risk badge (top-right corner of header)
    badge_w = Inches(1.75)
    badge_x = _W - badge_w - _M
    _rect(slide, badge_x, Inches(0.09), badge_w, Inches(0.44),
          _RISK_BG.get(overall_risk, _C["AMBER"]))
    _tb(slide, badge_x, Inches(0.09), badge_w, Inches(0.44),
        f"● {overall_risk} RISK",
        size=11, bold=True,
        color=_RISK_FG.get(overall_risk, _C["BLACK"]),
        align=PP_ALIGN.CENTER)

    # ── Stats bar ─────────────────────────────────────────────────────
    bar_y  = _HH + Inches(0.1)
    bar_h  = Inches(0.68)
    stat_w = Inches(2.38)
    sx     = _M

    for key, label, bg, fg in [
        ("critical_count", "CRITICAL", _C["RED"],          _C["WHITE"]),
        ("high_count",     "HIGH",     _C["AMBER"],        _C["BLACK"]),
        ("medium_count",   "MEDIUM",   _C["MEDIUM_ORANGE"],_C["BLACK"]),
        ("low_count",      "LOW",      _C["GREEN"],        _C["WHITE"]),
    ]:
        count = stats.get(key, 0)
        _rect(slide, sx, bar_y, stat_w, bar_h, bg)
        _tb(slide, sx, bar_y, stat_w, bar_h,
            f"{count}  {label}",
            size=13, bold=True, color=fg, align=PP_ALIGN.CENTER)
        sx += stat_w + Inches(0.12)

    # Item / PG summary (right-aligned in stats bar)
    total = stats.get("total_items", 0)
    pgs   = stats.get("product_groups_count", 0)
    _tb(slide, Inches(10.4), bar_y + Inches(0.19), Inches(2.6), Inches(0.32),
        f"{total} items  ·  {pgs} product groups",
        size=8, color=_C["BLACK"], align=PP_ALIGN.RIGHT)

    # ── Section label ─────────────────────────────────────────────────
    sub_y = bar_y + bar_h + Inches(0.14)
    _tb(slide, _M, sub_y, Inches(12.5), Inches(0.26),
        "KEY RISK HIGHLIGHTS",
        size=8, bold=True, color=_C["BLUE"])
    sub_y += Inches(0.29)

    # ── Bullets ───────────────────────────────────────────────────────
    remaining = _H - sub_y - _M
    per_bullet = min(remaining / max(len(bullets), 1), Inches(0.65))
    for bullet in bullets:
        _tb(slide, Inches(0.55), sub_y, Inches(12.1), per_bullet,
            f"▸  {bullet}", size=11, color=_C["BLACK"])
        sub_y += per_bullet

    # ── Footer ────────────────────────────────────────────────────────
    gen_date = date.today().isoformat()
    _tb(slide, Inches(10.0), _H - Inches(0.28), Inches(3.0), Inches(0.24),
        f"Generated {gen_date}", size=7,
        color=_C["GREY"], align=PP_ALIGN.RIGHT)


def _render_product_group(prs: Presentation, spec: dict) -> None:
    """Render a single product-group RAG heatmap slide."""
    slide = prs.slides.add_slide(prs.slide_layouts[6])

    title         = spec.get("title", "Product Group")
    headline      = spec.get("headline", "")
    cw_columns: list[int] = spec.get("cw_columns", [])
    quarter_label = spec.get("quarter_label", "Q?")
    rows: list[dict] = spec.get("rows", [])

    # ── Header ────────────────────────────────────────────────────────
    _rect(slide, 0, 0, _W, _HH, _C["NAVY"])
    _tb(slide, _M, Inches(0.06), Inches(11.5), Inches(0.30),
        title, size=14, bold=True, color=_C["WHITE"])
    _tb(slide, _M, Inches(0.36), Inches(11.5), Inches(0.22),
        headline, size=8, color=_C["LIGHT_BLUE"])

    if not rows:
        _tb(slide, _M, Inches(1.2), Inches(12.0), Inches(0.4),
            "No data rows for this product group.", size=11, color=_C["BLACK"])
        return

    # ── Table geometry ────────────────────────────────────────────────
    # Fixed columns:  Component | Plant | Customer | Supplier | Action | FM | Risk
    FIXED_LABELS  = ["Component",  "Plant",      "Customer",   "Supplier",
                     "Action / Comment",          "FM",         "Risk"]
    FIXED_WIDTHS  = [Inches(1.70), Inches(0.62), Inches(1.12), Inches(0.92),
                     Inches(1.80),               Inches(0.44), Inches(0.53)]

    CW_W = Inches(0.365)
    Q_W  = Inches(0.53)

    n_cw = len(cw_columns)
    total_fixed = sum(FIXED_WIDTHS)
    total_cw    = n_cw * CW_W + Q_W
    available   = _W - 2 * _M

    # Scale CW columns proportionally if they don't fit
    if total_fixed + total_cw > available:
        scale = (available - total_fixed) / total_cw
        CW_W = int(CW_W * scale)
        Q_W  = int(Q_W  * scale)

    all_widths = FIXED_WIDTHS + [CW_W] * n_cw + [Q_W]
    n_cols     = len(all_widths)
    n_rows_tbl = len(rows) + 1  # +1 for header

    tbl_y  = _HH + Inches(0.1)
    tbl_h  = _H - tbl_y - Inches(0.28)   # space for legend
    row_h  = max(int(tbl_h / n_rows_tbl), int(Inches(0.31)))

    tbl_shape = slide.shapes.add_table(
        n_rows_tbl, n_cols,
        int(_M), int(tbl_y),
        sum(int(w) for w in all_widths),
        row_h * n_rows_tbl,
    )
    tbl = tbl_shape.table

    # Apply column widths and row heights
    for i, w in enumerate(all_widths):
        tbl.columns[i].width = int(w)
    for r in range(n_rows_tbl):
        tbl.rows[r].height = row_h

    # ── Header row ────────────────────────────────────────────────────
    hdr_labels = FIXED_LABELS + [f"CW{c}" for c in cw_columns] + [quarter_label]
    for ci, label in enumerate(hdr_labels):
        cell = tbl.cell(0, ci)
        _cell_fill(cell, _C["NAVY"])
        is_cw_col = ci >= len(FIXED_LABELS)
        _cell_write(cell, label, size=7, bold=True, color=_C["WHITE"],
                    align=PP_ALIGN.CENTER if is_cw_col else PP_ALIGN.LEFT)

    # ── Data rows ─────────────────────────────────────────────────────
    for ri, row in enumerate(rows, start=1):
        alt = _C["LIGHT_GREY"] if ri % 2 == 0 else _C["WHITE"]

        # Fixed text columns
        fixed_vals = [
            row.get("component", ""),
            row.get("plant", ""),
            row.get("customer", ""),
            row.get("supplier", ""),
            row.get("action", ""),
            row.get("fm_status", ""),
        ]
        for ci, val in enumerate(fixed_vals):
            cell = tbl.cell(ri, ci)
            _cell_fill(cell, alt)
            center = ci in (1, 5)   # Plant and FM columns → center
            _cell_write(cell, str(val) if val else "", size=7,
                        align=PP_ALIGN.CENTER if center else PP_ALIGN.LEFT,
                        color=_C["BLACK"])

        # Risk level badge (column index 6)
        risk = str(row.get("risk_level", "")).upper()
        rc = tbl.cell(ri, 6)
        _cell_fill(rc, _RISK_BG.get(risk, alt))
        _cell_write(rc, risk[:4], size=7, bold=True,
                    color=_RISK_FG.get(risk, _C["BLACK"]),
                    align=PP_ALIGN.CENTER)

        # CW colored cells
        cw_colors: dict = row.get("cw_colors", {})
        for cw_i, cw_num in enumerate(cw_columns):
            ci = len(FIXED_LABELS) + cw_i
            rag = str(cw_colors.get(str(cw_num), "GREY")).upper()
            cell = tbl.cell(ri, ci)
            _cell_fill(cell, _C.get(rag, _C["GREY"]))
            _cell_write(cell, "", size=5)   # empty — color carries the meaning

        # Quarter column
        q_ci  = len(FIXED_LABELS) + n_cw
        q_rag = str(row.get("quarter_color", "GREY")).upper()
        q_cell = tbl.cell(ri, q_ci)
        _cell_fill(q_cell, _C.get(q_rag, _C["GREY"]))
        _cell_write(q_cell, "", size=5)

    # ── Color legend ──────────────────────────────────────────────────
    legend_y = tbl_y + row_h * n_rows_tbl + Inches(0.07)
    if legend_y + Inches(0.2) <= _H:
        lx = _M
        for label, color in [
            ("■ Covered (w/o mitigation)", _C["GREEN"]),
            ("■ Covered w/ mitigation",    _C["AMBER"]),
            ("■ Uncovered",                _C["RED"]),
            ("■ Data unavailable",         _C["GREY"]),
        ]:
            _tb(slide, lx, legend_y, Inches(2.85), Inches(0.2),
                label, size=6, color=color)
            lx += Inches(2.9)


def render_pptx_from_spec(slide_spec: dict, output_path: str | None = None) -> io.BytesIO:
    """
    Stage 2 (deterministic renderer): Convert an LLM-generated slide spec dict
    into a python-pptx Presentation.

    Parameters
    ----------
    slide_spec  : dict produced by llm_generate_slide_spec (or the fallback)
    output_path : optional path to save the file to disk

    Returns
    -------
    io.BytesIO buffer (seeked to 0) containing the .pptx bytes
    """
    prs = Presentation()
    prs.slide_width  = _W
    prs.slide_height = _H

    for slide_def in slide_spec.get("slides", []):
        stype = slide_def.get("type", "")
        if stype == "overview":
            _render_overview(prs, slide_def)
        elif stype == "product_group":
            _render_product_group(prs, slide_def)
        # Silently skip unknown types

    buf = io.BytesIO()
    prs.save(buf)
    buf.seek(0)

    if output_path:
        with open(output_path, "wb") as fh:
            fh.write(buf.read())
        buf.seek(0)

    return buf


# ─── Combined entry point ─────────────────────────────────────────────

async def generate_gfd_dashboard(
    extracted_data: dict,
    llm_config: dict,
    session_id: str,
    output_path: str | None = None,
    glossary_context: str = "",
) -> tuple[io.BytesIO, dict]:
    """
    Full GFD dashboard generation pipeline.

    1. LLM generates a complete slide specification from extracted_data.
    2. Thin PPTX renderer converts the spec to a python-pptx file.

    Parameters
    ----------
    extracted_data : output of parse_gfd_with_llm()
    llm_config     : Azure OpenAI config dict
    session_id     : for token / trace logging
    output_path    : optional path to save the .pptx file
    glossary_context : rendered glossary string for LLM context

    Returns
    -------
    (buf: BytesIO, slide_spec: dict)
    """
    spec = await llm_generate_slide_spec(
        extracted_data, llm_config, session_id, glossary_context
    )
    buf = render_pptx_from_spec(spec, output_path=output_path)
    return buf, spec
