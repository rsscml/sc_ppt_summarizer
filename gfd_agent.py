"""
GFD Agent Module
=================
LangGraph workflow that interprets parsed Dashboard_Update data through an LLM.

Pipeline:
  1. Parse Excel (gfd_excel_parser) → raw structured data
  2. Interpret per product group (this module) → condensed, slide-ready content
  3. Generate executive overview (this module) → narrative summary slide
  4. Generate slides (gfd_slide_generator) → .pptx output

The LLM condenses verbose free-text fields (Action/Comment, Root Cause,
Constraint), resolves abbreviations using the glossary, and produces
tight slide-ready text. CW coverage boundaries and RAG computation
remain deterministic — the LLM does not alter those.
"""

import json
import time
from typing import Any, TypedDict

from langchain_openai import AzureChatOpenAI
from langchain_core.messages import HumanMessage, SystemMessage
from langgraph.graph import StateGraph, END

from agent import create_llm, log_tokens, log_trace
from gfd_excel_parser import get_current_cw, summarise_for_prompt


# ─── Agent State ─────────────────────────────────────────────────────

class GFDAgentState(TypedDict):
    """State for the GFD interpretation workflow."""
    session_id: str
    llm_config: dict
    parsed_data: dict
    glossary_context: str
    interpreted_groups: list        # LLM-processed product groups
    executive_overview: str         # narrative summary for overview slide
    current_cw_label: str           # e.g. "CW13/2026"


# ─── Prompts ─────────────────────────────────────────────────────────

INTERPRET_PG_SYSTEM = """You are a supply chain operations analyst preparing content for a Global Fulfilment Dashboard slide deck for senior automotive executives.

You are given raw data for one product group from the Dashboard_Update worksheet. Your job is to **condense and interpret** each risk row into tight, slide-ready text.

RULES:
1. ONLY use facts from the data provided. Do NOT invent figures, dates, or supplier names.
2. Use the company glossary below to expand abbreviations on first mention.
3. Keep all CW numbers, dates, and percentages exactly as given — do not round or estimate.
4. Be concise: executives scan, not read. Every word must earn its place.

FOR EACH ROW, produce:
- **customer_summary**: Who is affected. Combine from "Customer affected" and customer impact flags. Max 40 chars.
- **root_cause_short**: 1-line root cause. Max 60 chars. Be specific (not "supply issue" — say "NXP wafer allocation at 70%").
- **action_condensed**: Condensed action/comment. Max 100 chars. Prioritize: what is being done, by when, and current status. Drop generic filler.
- **supplier_short**: Supplier name, trimmed. Max 30 chars.
- **constraint_short**: Constraint type, trimmed. Max 30 chars.
- **fm_status**: Force Majeure / customer informed status. "Yes", "No", "In progress", or "N/A". Max 15 chars.
- **risk_level**: Your assessment: "CRITICAL", "HIGH", "MEDIUM", or "LOW" — based on coverage gap, customer exposure, and mitigation maturity.

ALSO produce for the entire product group:
- **pg_headline**: A 1-line headline summarizing the product group's situation. Max 80 chars.

OUTPUT FORMAT — respond with ONLY valid JSON, no markdown fences:
{{
  "pg_headline": "...",
  "rows": [
    {{
      "row_index": 0,
      "customer_summary": "...",
      "root_cause_short": "...",
      "action_condensed": "...",
      "supplier_short": "...",
      "constraint_short": "...",
      "fm_status": "...",
      "risk_level": "..."
    }}
  ]
}}
{glossary_block}"""

INTERPRET_PG_USER = """Product Group: {pg_name}
Number of risk items: {row_count}
Current calendar week: {current_cw}

Raw data for each risk row:

{rows_text}"""

EXECUTIVE_OVERVIEW_SYSTEM = """You are a Chief Supply Chain Officer writing a brief executive overview for the first slide of a Global Fulfilment Dashboard presentation.

You are given LLM-interpreted summaries of all product groups in the dashboard. Write a concise narrative overview.

RULES:
1. ONLY use facts from the summaries provided.
2. Lead with the most critical risks.
3. Quantify: how many product groups, how many CRITICAL/HIGH items, which customers are most exposed.
4. Mention specific CW coverage gaps and key mitigation actions.
5. Use the company glossary to expand abbreviations where helpful.
6. Maximum 5-6 bullet points. Each bullet max 120 characters.

OUTPUT FORMAT — respond with ONLY valid JSON, no markdown fences:
{{
  "title": "Situation Overview — CW{cw_number}/{cw_year}",
  "bullets": [
    "bullet 1 text",
    "bullet 2 text"
  ],
  "overall_risk": "CRITICAL|HIGH|MEDIUM|LOW"
}}
{glossary_block}"""

EXECUTIVE_OVERVIEW_USER = """Dashboard Summary — {total_rows} risk items across {total_pgs} product groups.
Current CW: {current_cw}

{all_pg_summaries}"""


# ─── Row Formatting for Prompt ───────────────────────────────────────

def _format_row_for_prompt(row: dict, index: int) -> str:
    """Format a single parsed row as a readable text block for the LLM."""
    lines = [f"--- Row {index} ---"]

    fields = [
        ("Plant", "plant_location"),
        ("Region", "region"),
        ("Customer affected", "customer_affected"),
        ("Critical Component", "critical_component"),
        ("Constraint / Task Force", "constraint_task_force"),
        ("Root Cause", "root_cause"),
        ("Supplier", "supplier_text"),
        ("Supplier Type", "supplier_type"),
        ("Supplier Region", "supplier_region"),
        ("Coverage w/o mitigation", "coverage_without_mitigation"),
        ("Coverage w/ mitigation", "coverage_with_mitigation"),
        ("Fulfillment current Q", "fulfillment_current_q"),
        ("Fulfillment Q+1", "fulfillment_q_plus_1"),
        ("Fulfillment Q+2", "fulfillment_q_plus_2"),
        ("Recovery Week", "recovery_week"),
        ("OPS Capacity Risk", "ops_capacity_risk"),
        ("Strategic Capacity Risk", "strategic_capacity_risk"),
        ("Allocation Mode", "allocation_mode"),
        ("Customer Informed", "customer_informed"),
        ("Action / Comment", "action_comment"),
        ("Task Force Leader", "task_force_leader"),
        ("Special Freight Cost €", "special_freight_cost"),
        ("Special Freight", "special_freight_remarks"),
    ]

    for label, key in fields:
        val = row.get(key)
        if val is not None and str(val).strip():
            # Format booleans and floats nicely
            if isinstance(val, bool):
                val = "Yes" if val else "No"
            elif isinstance(val, float) and key.startswith("fulfillment"):
                val = f"{val:.0%}"
            lines.append(f"  {label}: {val}")

    # Customer impact flags
    impacts = row.get("customer_impact", {})
    affected = [k for k, v in impacts.items() if v]
    if affected:
        lines.append(f"  Customer impact flags: {', '.join(affected)}")

    return "\n".join(lines)


def _format_pg_for_prompt(pg: dict, current_cw: str) -> str:
    """Format an entire product group for the LLM prompt."""
    rows_text = []
    for i, row in enumerate(pg["rows"]):
        rows_text.append(_format_row_for_prompt(row, i))
    return "\n\n".join(rows_text)


# ─── Graph Nodes ─────────────────────────────────────────────────────

async def interpret_product_groups_node(state: GFDAgentState) -> dict:
    """
    Node: Interpret each product group through the LLM.
    One LLM call per product group (parallel-safe, avoids token overflow).
    """
    config = state["llm_config"]
    llm = create_llm(config)
    session_id = state["session_id"]
    parsed = state["parsed_data"]
    glossary = state.get("glossary_context", "")
    glossary_block = f"\n\nCOMPANY GLOSSARY:\n{glossary}" if glossary else ""

    current_cw = get_current_cw()
    current_cw_label = f"CW{current_cw[1]}/{current_cw[0]}"

    interpreted_groups = []

    for pg_idx, pg in enumerate(parsed["product_groups"]):
        t0 = time.time()
        pg_name = pg["product_family_desc"]
        pg_code = pg["product_family_code"]
        pg_display = f"{pg_name} ({pg_code})" if pg_code else pg_name

        rows_text = _format_pg_for_prompt(pg, current_cw_label)

        messages = [
            SystemMessage(content=INTERPRET_PG_SYSTEM.format(
                glossary_block=glossary_block
            )),
            HumanMessage(content=INTERPRET_PG_USER.format(
                pg_name=pg_display,
                row_count=len(pg["rows"]),
                current_cw=current_cw_label,
                rows_text=rows_text,
            ))
        ]

        try:
            response = await llm.ainvoke(messages)
            raw_text = response.content.strip()

            # Parse JSON from response (strip markdown fences if present)
            json_text = raw_text
            if json_text.startswith("```"):
                json_text = "\n".join(json_text.split("\n")[1:])
            if json_text.endswith("```"):
                json_text = "\n".join(json_text.split("\n")[:-1])
            json_text = json_text.strip()

            llm_result = json.loads(json_text)

            # Log tokens
            usage = response.response_metadata.get("token_usage", {})
            log_tokens(session_id, f"gfd_interpret_pg_{pg_idx}_{pg_name}",
                       usage, config.get("azure_deployment", ""))

            # Merge LLM interpretations back onto the raw rows
            interpreted_rows = []
            llm_rows = llm_result.get("rows", [])
            for i, raw_row in enumerate(pg["rows"]):
                # Find the matching LLM row by index
                llm_row = next(
                    (r for r in llm_rows if r.get("row_index") == i),
                    llm_rows[i] if i < len(llm_rows) else {}
                )
                merged = {
                    **raw_row,  # preserve all raw data (coverage CWs, etc.)
                    "llm_customer_summary": llm_row.get("customer_summary", ""),
                    "llm_root_cause": llm_row.get("root_cause_short", ""),
                    "llm_action": llm_row.get("action_condensed", ""),
                    "llm_supplier": llm_row.get("supplier_short", ""),
                    "llm_constraint": llm_row.get("constraint_short", ""),
                    "llm_fm_status": llm_row.get("fm_status", ""),
                    "llm_risk_level": llm_row.get("risk_level", ""),
                }
                interpreted_rows.append(merged)

            interpreted_groups.append({
                "product_family_desc": pg["product_family_desc"],
                "product_family_code": pg["product_family_code"],
                "pg_headline": llm_result.get("pg_headline", pg_display),
                "rows": interpreted_rows,
            })

            duration = (time.time() - t0) * 1000
            log_trace(session_id, "gfd_interpret_pg",
                      f"PG: {pg_display} ({len(pg['rows'])} rows)",
                      llm_result.get("pg_headline", "")[:200],
                      duration, {"pg_index": pg_idx})

        except Exception as e:
            # On failure, fall back to raw data with no LLM enrichment
            fallback_rows = []
            for raw_row in pg["rows"]:
                fallback_rows.append({
                    **raw_row,
                    "llm_customer_summary": raw_row.get("customer_affected", ""),
                    "llm_root_cause": raw_row.get("root_cause", ""),
                    "llm_action": (raw_row.get("action_comment") or "")[:100],
                    "llm_supplier": raw_row.get("supplier_text", ""),
                    "llm_constraint": raw_row.get("constraint_task_force", ""),
                    "llm_fm_status": "Yes" if raw_row.get("customer_informed") else "No",
                    "llm_risk_level": raw_row.get("ops_capacity_risk", ""),
                })
            interpreted_groups.append({
                "product_family_desc": pg["product_family_desc"],
                "product_family_code": pg["product_family_code"],
                "pg_headline": pg_display,
                "rows": fallback_rows,
            })
            log_trace(session_id, "gfd_interpret_pg",
                      f"PG: {pg_display}", f"FALLBACK (error: {str(e)[:100]})",
                      (time.time() - t0) * 1000, {"error": True})

    return {
        "interpreted_groups": interpreted_groups,
        "current_cw_label": current_cw_label,
    }


async def generate_executive_overview_node(state: GFDAgentState) -> dict:
    """
    Node: Generate executive overview narrative from all interpreted groups.
    """
    config = state["llm_config"]
    llm = create_llm(config)
    session_id = state["session_id"]
    parsed = state["parsed_data"]
    interpreted_groups = state["interpreted_groups"]
    glossary = state.get("glossary_context", "")
    glossary_block = f"\n\nCOMPANY GLOSSARY:\n{glossary}" if glossary else ""
    current_cw_label = state["current_cw_label"]

    t0 = time.time()

    # Build summary text from interpreted groups
    pg_summaries = []
    total_rows = 0
    for pg in interpreted_groups:
        desc = pg["product_family_desc"]
        code = pg["product_family_code"]
        headline = pg.get("pg_headline", desc)
        pg_display = f"{desc} ({code})" if code else desc
        total_rows += len(pg["rows"])

        lines = [f"### {pg_display}", f"Headline: {headline}"]
        for row in pg["rows"]:
            plant = row.get("plant_location", "?")
            cov_wo = row.get("coverage_without_mitigation", "?")
            cov_w = row.get("coverage_with_mitigation", "?")
            risk = row.get("llm_risk_level", "?")
            customer = row.get("llm_customer_summary", "?")
            action = row.get("llm_action", "?")
            root = row.get("llm_root_cause", "?")

            lines.append(
                f"  [{risk}] {plant}: {customer} | Coverage: {cov_wo}→{cov_w} | "
                f"Root cause: {root} | Action: {action}"
            )
        pg_summaries.append("\n".join(lines))

    all_summaries_text = "\n\n".join(pg_summaries)

    # Parse CW for prompt
    cw = get_current_cw()

    messages = [
        SystemMessage(content=EXECUTIVE_OVERVIEW_SYSTEM.format(
            cw_number=cw[1], cw_year=cw[0],
            glossary_block=glossary_block,
        )),
        HumanMessage(content=EXECUTIVE_OVERVIEW_USER.format(
            total_rows=total_rows,
            total_pgs=len(interpreted_groups),
            current_cw=current_cw_label,
            all_pg_summaries=all_summaries_text,
        ))
    ]

    try:
        response = await llm.ainvoke(messages)
        raw_text = response.content.strip()

        json_text = raw_text
        if json_text.startswith("```"):
            json_text = "\n".join(json_text.split("\n")[1:])
        if json_text.endswith("```"):
            json_text = "\n".join(json_text.split("\n")[:-1])

        overview = json.loads(json_text.strip())
        overview_text = json.dumps(overview)

        usage = response.response_metadata.get("token_usage", {})
        log_tokens(session_id, "gfd_executive_overview", usage,
                   config.get("azure_deployment", ""))

    except Exception as e:
        # Fallback: generate a simple overview without LLM
        overview = {
            "title": f"Fulfilment Risk Overview — {current_cw_label}",
            "bullets": [
                f"{len(interpreted_groups)} product groups with active fulfilment risks",
                f"{total_rows} risk items tracked across all product groups",
                "See detailed CW coverage grid on following slides",
            ],
            "overall_risk": "HIGH",
        }
        overview_text = json.dumps(overview)
        log_trace(session_id, "gfd_executive_overview",
                  "Generating overview", f"FALLBACK (error: {str(e)[:100]})",
                  (time.time() - t0) * 1000, {"error": True})

    duration = (time.time() - t0) * 1000
    log_trace(session_id, "gfd_executive_overview",
              f"Generating overview for {len(interpreted_groups)} PGs",
              overview_text[:300], duration)

    return {"executive_overview": overview_text}


# ─── Build the Graph ─────────────────────────────────────────────────

def build_gfd_graph():
    """Build the LangGraph for GFD interpretation workflow."""
    workflow = StateGraph(GFDAgentState)

    workflow.add_node("interpret_product_groups", interpret_product_groups_node)
    workflow.add_node("generate_executive_overview", generate_executive_overview_node)

    workflow.set_entry_point("interpret_product_groups")
    workflow.add_edge("interpret_product_groups", "generate_executive_overview")
    workflow.add_edge("generate_executive_overview", END)

    return workflow.compile()


# ─── Convenience Entry Point ─────────────────────────────────────────

async def run_gfd_pipeline(session_id: str, llm_config: dict,
                           parsed_data: dict,
                           glossary_context: str = "") -> dict:
    """
    Run the full GFD interpretation pipeline.

    Returns:
        {
            "interpreted_groups": list,   — LLM-enriched product groups
            "executive_overview": dict,   — parsed JSON overview
            "current_cw_label": str,
        }
    """
    graph = build_gfd_graph()

    initial_state: GFDAgentState = {
        "session_id": session_id,
        "llm_config": llm_config,
        "parsed_data": parsed_data,
        "glossary_context": glossary_context,
        "interpreted_groups": [],
        "executive_overview": "",
        "current_cw_label": "",
    }

    final_state = await graph.ainvoke(initial_state)

    # Parse the executive overview JSON
    try:
        overview = json.loads(final_state["executive_overview"])
    except (json.JSONDecodeError, TypeError):
        overview = {
            "title": "Fulfilment Risk Overview",
            "bullets": ["See detailed slides for coverage status."],
            "overall_risk": "HIGH",
        }

    return {
        "interpreted_groups": final_state["interpreted_groups"],
        "executive_overview": overview,
        "current_cw_label": final_state["current_cw_label"],
    }
