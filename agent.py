"""
LangGraph Agent Module
======================
Agentic workflow for section-by-section summarization
with conversation-based refinement.
"""

import time
import uuid
from datetime import datetime, timezone
from typing import Any, TypedDict

from langchain_openai import AzureChatOpenAI
from langchain_core.messages import HumanMessage, SystemMessage
from langgraph.graph import StateGraph, END


# ─── In-memory stores ────────────────────────────────────────────────

token_usage_log: list[dict] = []
trace_log: list[dict] = []


def log_tokens(session_id: str, step: str, usage: dict, model: str = ""):
    """Record token usage for a given step."""
    entry = {
        "id": str(uuid.uuid4()),
        "session_id": session_id,
        "timestamp": datetime.now(timezone.utc).isoformat(),
        "step": step,
        "model": model,
        "prompt_tokens": usage.get("prompt_tokens", 0),
        "completion_tokens": usage.get("completion_tokens", 0),
        "total_tokens": usage.get("prompt_tokens", 0) + usage.get("completion_tokens", 0),
    }
    token_usage_log.append(entry)
    return entry


def log_trace(session_id: str, node: str, input_summary: str, output_summary: str,
              duration_ms: float, metadata: dict | None = None):
    """Record a trace entry for a graph node execution."""
    entry = {
        "id": str(uuid.uuid4()),
        "session_id": session_id,
        "timestamp": datetime.now(timezone.utc).isoformat(),
        "node": node,
        "input_summary": input_summary[:500],
        "output_summary": output_summary[:500],
        "duration_ms": round(duration_ms, 2),
        "metadata": metadata or {}
    }
    trace_log.append(entry)
    return entry


# ─── Agent State ──────────────────────────────────────────────────────

class AgentState(TypedDict):
    """State for the summarization agent graph."""
    session_id: str
    llm_config: dict
    parsed_ppt: dict
    glossary_context: str
    section_summaries: list
    executive_summary: str
    email_summary: str
    all_summaries_text: str


# ─── LLM Setup ───────────────────────────────────────────────────────

def create_llm(config: dict) -> AzureChatOpenAI:
    """Create AzureChatOpenAI instance from config."""
    return AzureChatOpenAI(
        azure_deployment=config["azure_deployment"],
        azure_endpoint=config["azure_endpoint"],
        api_key=config["api_key"],
        api_version=config.get("api_version", "2024-12-01-preview"),
        temperature=0.2,
        max_tokens=4096,
    )


# ─── Prompts ─────────────────────────────────────────────────────────

SECTION_SUMMARY_SYSTEM = """You are an expert supply chain analyst summarizing a section of a Global Supply Chain Status Report for a senior automotive parts manufacturing executive.

CRITICAL RULES:
1. ONLY use facts, figures, percentages, dates, names, and numbers that are EXPLICITLY present in the source content provided.
2. NEVER invent or estimate any numbers. If a figure is unclear, say "figure unclear in source".
3. Preserve ALL specific metrics: percentages, Euro amounts, counts, week numbers, days, supplier names, customer names, plant names, KAM names.
4. Note any RAG (Red/Amber/Green) status indicators and their meaning.
5. Highlight critical action items, deadlines, and responsible persons.
6. Keep the summary dense with facts - no filler or generic statements.
7. When abbreviations appear in the source, use the company glossary below to expand them. On first mention write "Full Name (ABBR)", thereafter just the abbreviation is fine.

Output format:
- Start with a 1-line section headline
- Use bullet points with specific data
- Group related facts together
- End with "KEY RISKS" and "ACTIONS" sub-sections if applicable
- If charts/tables contain data, summarize the key data points and trends
{glossary_block}"""

SECTION_SUMMARY_USER = """Summarize the following section from the supply chain status report. 
Extract ALL critical facts, figures, names, and action items.

{section_content}"""

EXECUTIVE_SUMMARY_SYSTEM = """You are a Chief Supply Chain Officer preparing a 2-4 slide executive summary for the Board of Directors of a global automotive parts manufacturer.

CRITICAL RULES:
1. ONLY use facts and figures from the section summaries provided - DO NOT invent anything.
2. The summary must be structured for 2-4 presentation slides.
3. Prioritize: immediate risks, financial impact, customer impact, and required decisions.
4. Every bullet point must contain a specific fact, number, or name from the source material.
5. Use concise, executive-level language.
6. Where abbreviations appear, use the company glossary to ensure full names are used alongside abbreviations for clarity, especially for locations, business units, and key roles.

OUTPUT FORMAT - Structure as slide content:

## Slide 1: Situation Overview & Key Metrics
(Crisis status, headline numbers, overall impact metrics)

## Slide 2: Supply & Production Impact  
(Supplier status, production backlogs, capacity issues, inventory)

## Slide 3: Customer & Financial Impact
(Customer fulfillment, demand changes, cost impact, FM letters)

## Slide 4: Actions & Outlook (if needed)
(Critical actions, deadlines, responsible persons, outlook)

For each slide, provide:
- A clear slide title
- 4-6 bullet points with SPECIFIC data
- Any suggested chart data (as simple key:value pairs) if a visualization would help

If information for a slide is sparse, merge slides. Minimum 2 slides, maximum 4.
{glossary_block}"""

EXECUTIVE_SUMMARY_USER = """Based on the following section-by-section summaries of our Global Supply Chain Status Report ({total_slides} slides total, {total_sections} sections), create the executive summary slides.

{all_summaries}"""

EMAIL_SUMMARY_SYSTEM = """You are a senior supply chain executive drafting a concise crisis‑status email update for leadership and cross‑functional stakeholders of a global automotive parts manufacturer.

CRITICAL RULES:
1. ONLY use facts, figures, names, and data that appear in the section summaries provided. DO NOT invent.
2. Where abbreviations appear, use the company glossary to expand them on first mention.
3. Write in professional, direct email prose — no slide‑style bullets unless a section naturally calls for a short list.
4. Include ONLY sections for which substantive information exists in the source material. Skip any section heading if no relevant data was found.
5. Every claim must be traceable to the source summaries.

STRUCTURE — use the following section headings where information is available:

### Key Management Takeaways
(3‑5 high‑level bullet points: material risks, strategic constraints, planning‑baseline guidance, governance priorities)

### Overall Situation Summary
(Paragraph or short bullets: duration outlook, sourcing constraints, supplier status, customer risk, logistics volatility)

### Key Product & Customer Risks
(Sub‑section per affected product group. For each: coverage status with week/date references, critical supplier dependencies, quality issues, air‑freight or mitigation options, Force Majeure communication status. Only include product groups with material information.)

### Additional Product / Production Status
(Brief overview of other product groups under monitoring, risk horizon, inventory coverage)

### Supplier Mitigation Actions
(Active and planned mitigation: alternative suppliers, fallback activation, dual‑sourcing, fuel switching, structural constraints)

### Customer: Commercial & Legal Situation
(Pricing governance, FM warning letter status, cost pass‑through, Sales/Legal coordination)

### Logistics Situation
(Freight rates, transit times, capacity, regional pauses, revenue timing impact)

### Financial Exposure – Scenario Ranges
(Quantified exposure if available: top‑line risk, cost impact ranges, scenario assumptions)

FORMATTING:
- Start with "Dear all," and a brief one‑line context sentence.
- Use markdown headings (###) for sections.
- Keep the total length appropriate for an email — thorough but scannable.
- End with a brief sign‑off line.
{glossary_block}"""

EMAIL_SUMMARY_USER = """Based on the following section‑by‑section summaries of our Global Supply Chain Status Report ({total_slides} slides total, {total_sections} sections), draft the leadership email update. Include only sections for which data exists.

{all_summaries}"""

REFINE_SYSTEM = """You are helping refine an executive summary of a Global Supply Chain Status Report. 
The user will provide instructions for changes. Apply them precisely.

RULES:
1. Only use facts from the original section summaries provided as context.
2. Do not invent new figures or data points.
3. Maintain the slide-structured format.
4. If asked to add detail, pull from the section summaries context.
5. Use the company glossary below to expand any abbreviations correctly.

Current executive summary:
{current_summary}

Original section summaries for reference:
{section_summaries}
{glossary_block}"""

REFINE_EMAIL_SYSTEM = """You are helping refine a crisis‑status email summary of a Global Supply Chain Status Report.
The user will provide instructions for changes. Apply them precisely.

RULES:
1. Only use facts from the original section summaries provided as context.
2. Do not invent new figures or data points.
3. Maintain the email format with section headings.
4. If asked to add detail, pull from the section summaries context.
5. Use the company glossary below to expand any abbreviations correctly.

Current email summary:
{current_email}

Original section summaries for reference:
{section_summaries}
{glossary_block}"""


# ─── Graph Nodes ──────────────────────────────────────────────────────

async def summarize_sections_node(state: AgentState) -> dict:
    """Node: Summarize each section independently."""
    config = state["llm_config"]
    llm = create_llm(config)
    parsed = state["parsed_ppt"]
    session_id = state["session_id"]
    glossary = state.get("glossary_context", "")
    glossary_block = f"\n\n{glossary}" if glossary else ""
    section_summaries = []

    for i, section in enumerate(parsed["sections"]):
        t0 = time.time()
        section_name = section["section_name"]
        content = section["formatted_content"]

        # Truncate very long sections to avoid token limits per call
        if len(content) > 15000:
            content = content[:15000] + "\n\n[Content truncated for processing - additional slides in section]"

        messages = [
            SystemMessage(content=SECTION_SUMMARY_SYSTEM.format(glossary_block=glossary_block)),
            HumanMessage(content=SECTION_SUMMARY_USER.format(section_content=content))
        ]

        try:
            response = await llm.ainvoke(messages)
            summary_text = response.content

            # Log tokens
            usage = response.response_metadata.get("token_usage", {})
            log_tokens(session_id, f"section_summary_{i}_{section_name}", usage,
                       config.get("azure_deployment", ""))

            section_summaries.append({
                "section_name": section_name,
                "slide_count": section["slide_count"],
                "slide_numbers": section["slide_numbers"],
                "summary": summary_text
            })

            duration = (time.time() - t0) * 1000
            log_trace(session_id, "summarize_section", 
                      f"Section: {section_name} ({section['slide_count']} slides)",
                      summary_text[:200], duration,
                      {"section_index": i, "input_chars": len(content)})
        except Exception as e:
            section_summaries.append({
                "section_name": section_name,
                "slide_count": section["slide_count"],
                "slide_numbers": section["slide_numbers"],
                "summary": f"[Error summarizing section: {str(e)}]"
            })
            log_trace(session_id, "summarize_section",
                      f"Section: {section_name}", f"ERROR: {str(e)}",
                      (time.time() - t0) * 1000, {"error": True})

    return {"section_summaries": section_summaries}


async def generate_executive_summary_node(state: AgentState) -> dict:
    """Node: Generate final executive summary from all section summaries."""
    config = state["llm_config"]
    llm = create_llm(config)
    session_id = state["session_id"]
    parsed = state["parsed_ppt"]
    section_summaries = state["section_summaries"]
    glossary = state.get("glossary_context", "")
    glossary_block = f"\n\n{glossary}" if glossary else ""

    t0 = time.time()

    # Combine all section summaries
    all_summaries_text = ""
    for ss in section_summaries:
        all_summaries_text += f"\n\n### {ss['section_name']} (Slides {ss['slide_numbers'][0]}-{ss['slide_numbers'][-1]})\n"
        all_summaries_text += ss["summary"]

    messages = [
        SystemMessage(content=EXECUTIVE_SUMMARY_SYSTEM.format(glossary_block=glossary_block)),
        HumanMessage(content=EXECUTIVE_SUMMARY_USER.format(
            total_slides=parsed["total_slides"],
            total_sections=parsed["total_sections"],
            all_summaries=all_summaries_text
        ))
    ]

    response = await llm.ainvoke(messages)
    exec_summary = response.content

    usage = response.response_metadata.get("token_usage", {})
    log_tokens(session_id, "executive_summary", usage, config.get("azure_deployment", ""))

    duration = (time.time() - t0) * 1000
    log_trace(session_id, "generate_executive_summary",
              f"Combined {len(section_summaries)} section summaries",
              exec_summary[:300], duration)

    return {"executive_summary": exec_summary, "all_summaries_text": all_summaries_text}


async def generate_email_summary_node(state: AgentState) -> dict:
    """Node: Generate email status summary from section summaries."""
    config = state["llm_config"]
    llm = create_llm(config)
    session_id = state["session_id"]
    parsed = state["parsed_ppt"]
    all_summaries_text = state["all_summaries_text"]
    glossary = state.get("glossary_context", "")
    glossary_block = f"\n\n{glossary}" if glossary else ""

    t0 = time.time()

    messages = [
        SystemMessage(content=EMAIL_SUMMARY_SYSTEM.format(glossary_block=glossary_block)),
        HumanMessage(content=EMAIL_SUMMARY_USER.format(
            total_slides=parsed["total_slides"],
            total_sections=parsed["total_sections"],
            all_summaries=all_summaries_text
        ))
    ]

    response = await llm.ainvoke(messages)
    email_text = response.content

    usage = response.response_metadata.get("token_usage", {})
    log_tokens(session_id, "email_summary", usage, config.get("azure_deployment", ""))

    duration = (time.time() - t0) * 1000
    log_trace(session_id, "generate_email_summary",
              f"Generating email from {parsed['total_sections']} section summaries",
              email_text[:300], duration)

    return {"email_summary": email_text}


# ─── Build the Graph ──────────────────────────────────────────────────

def build_summarization_graph():
    """Build the LangGraph for the summarization workflow."""
    workflow = StateGraph(AgentState)

    workflow.add_node("summarize_sections", summarize_sections_node)
    workflow.add_node("generate_executive_summary", generate_executive_summary_node)
    workflow.add_node("generate_email_summary", generate_email_summary_node)

    workflow.set_entry_point("summarize_sections")
    workflow.add_edge("summarize_sections", "generate_executive_summary")
    workflow.add_edge("generate_executive_summary", "generate_email_summary")
    workflow.add_edge("generate_email_summary", END)

    return workflow.compile()


# ─── Refinement (outside graph, simple chain) ─────────────────────────

async def refine_summary(session_id: str, llm_config: dict,
                         current_summary: str, section_summaries_text: str,
                         user_instruction: str, glossary_context: str = "") -> str:
    """Refine the executive slide summary based on user instructions."""
    llm = create_llm(llm_config)
    t0 = time.time()
    glossary_block = f"\n\n{glossary_context}" if glossary_context else ""

    messages = [
        SystemMessage(content=REFINE_SYSTEM.format(
            current_summary=current_summary,
            section_summaries=section_summaries_text,
            glossary_block=glossary_block,
        )),
        HumanMessage(content=user_instruction)
    ]

    response = await llm.ainvoke(messages)
    refined = response.content

    usage = response.response_metadata.get("token_usage", {})
    log_tokens(session_id, "refine_summary", usage, llm_config.get("azure_deployment", ""))

    duration = (time.time() - t0) * 1000
    log_trace(session_id, "refine_summary",
              f"User instruction: {user_instruction[:200]}",
              refined[:300], duration)

    return refined


async def refine_email(session_id: str, llm_config: dict,
                       current_email: str, section_summaries_text: str,
                       user_instruction: str, glossary_context: str = "") -> str:
    """Refine the email summary based on user instructions."""
    llm = create_llm(llm_config)
    t0 = time.time()
    glossary_block = f"\n\n{glossary_context}" if glossary_context else ""

    messages = [
        SystemMessage(content=REFINE_EMAIL_SYSTEM.format(
            current_email=current_email,
            section_summaries=section_summaries_text,
            glossary_block=glossary_block,
        )),
        HumanMessage(content=user_instruction)
    ]

    response = await llm.ainvoke(messages)
    refined = response.content

    usage = response.response_metadata.get("token_usage", {})
    log_tokens(session_id, "refine_email", usage, llm_config.get("azure_deployment", ""))

    duration = (time.time() - t0) * 1000
    log_trace(session_id, "refine_email",
              f"User instruction: {user_instruction[:200]}",
              refined[:300], duration)

    return refined
