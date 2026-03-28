"""
GFD LLM Parser
==============
Two-stage pipeline that replaces the deterministic gfd_excel_parser for content extraction.

Stage 1 — Minimal deterministic read (the ONLY deterministic part):
  • Open the workbook and find the Dashboard_Update sheet
  • Detect the header row via keyword scoring (scans up to 60 rows, merge-aware)
  • Resolve merged cells so repeated group values appear in every row
  • Convert all rows to a pipe-delimited text table
  • Filter stale rows using the "Last updated" date column (preferred) — any
    common date format is handled; openpyxl native datetime objects are used
    directly when available, strings are tried against multiple format patterns.
    Falls back to CW-number scanning if no date column can be identified.

Stage 2 — LLM extraction:
  • Send the filtered text table to the LLM
  • LLM understands column semantics, infers product-family groupings
    (repeated values from formerly-merged cells), and returns a precise JSON object
  • No brittle column-name fuzzy matching; no schema hard-coding

The output JSON is the single source of truth consumed by gfd_llm_slides.py.
"""

from __future__ import annotations

import re
import json
import time
from datetime import datetime, date, timedelta
from pathlib import Path
from typing import Any

import openpyxl
from langchain_openai import AzureChatOpenAI
from langchain_core.messages import HumanMessage, SystemMessage

from agent import log_tokens, log_trace


# ─── CW utilities ────────────────────────────────────────────────────

def _get_current_cw() -> tuple[int, int]:
    """Return (year, iso_week) for today."""
    iso = datetime.now().isocalendar()
    return (iso.year, iso.week)


# ─── LLM factory (local, allows custom max_tokens) ───────────────────

def _create_llm(config: dict, max_tokens: int = 64000) -> AzureChatOpenAI:
    return AzureChatOpenAI(
        azure_deployment=config["azure_deployment"],
        azure_endpoint=config["azure_endpoint"],
        api_key=config["api_key"],
        api_version=config.get("api_version", "2024-12-01-preview"),
        #temperature=0.0,   # extraction must be deterministic
        max_tokens=max_tokens,
    )


# ─── Stage 1: Excel → filtered pipe-delimited text table ─────────────

_HEADER_KEYWORDS = {
    "plant", "product", "family", "coverage", "customer", "supplier",
    "root", "cause", "action", "comment", "mitigation", "risk", "cw", "kw",
    "constraint", "freight", "recovery", "fulfil", "fulfillment",
    "region", "component", "allocation", "informed", "task", "force",
}


def _cell_str(value: Any) -> str:
    """Normalise an openpyxl cell value to a clean string.

    Also strips a leading apostrophe character: in some Excel files (including
    the standard GFD template) cell values are stored with a literal leading
    apostrophe that was used as a text-prefix escape marker.  openpyxl reads
    this as part of the value string, so we remove it here.
    """
    if value is None:
        return ""
    if isinstance(value, datetime):
        return value.strftime("%d.%m.%Y")
    s = re.sub(r"[\r\n]+", " ", str(value).strip())
    # Strip leading/trailing apostrophe(s) that Excel text-prefix escape leaves behind.
    # The standard GFD template stores all cell values as 'value' with surrounding quotes.
    if s.startswith("'") or s.endswith("'"):
        s = s.strip("'").strip()
    return s


def _header_score(cells: list[str]) -> int:
    """Return keyword-match score for a row being the header row."""
    text = " ".join(cells).lower()
    return sum(1 for kw in _HEADER_KEYWORDS if kw in text)


def _find_header_row(ws, merge_map: dict | None = None, max_scan: int = 60) -> int:
    """
    Return the 1-based row index of the best header-row candidate.

    Scans up to `max_scan` rows (default 60 — enough for any realistic title
    block, logo row, or instruction section above the actual data header).

    Uses the merge map when provided so that merged slave cells contribute
    their master's value to the keyword score rather than scoring as empty.

    Falls back to row 1 only when no row scores above zero (completely
    blank or unrecognisable sheet).
    """
    if merge_map is None:
        merge_map = {}

    max_col = ws.max_column or 1
    max_row = ws.max_row or 1
    best_idx, best_score = 1, 0

    for row_idx in range(1, min(max_scan + 1, max_row + 1)):
        cells = []
        for col_idx in range(1, max_col + 1):
            raw = merge_map.get((row_idx, col_idx), ws.cell(row_idx, col_idx).value)
            cells.append(_cell_str(raw))
        score = _header_score(cells)
        if score > best_score:
            best_score, best_idx = score, row_idx

    return best_idx


def _build_merge_map(ws) -> dict[tuple[int, int], Any]:
    """
    Map every merged slave cell (row, col) → the master cell's value.
    This is the core fix for product-group grouping: formerly-merged cells
    now carry the group value in every row, letting the LLM see the grouping.
    """
    merge_map: dict[tuple[int, int], Any] = {}
    try:
        for merge in ws.merged_cells.ranges:
            master_val = ws.cell(merge.min_row, merge.min_col).value
            for r in range(merge.min_row, merge.max_row + 1):
                for c in range(merge.min_col, merge.max_col + 1):
                    if (r, c) != (merge.min_row, merge.min_col):
                        merge_map[(r, c)] = master_val
    except Exception:
        pass  # read_only mode has no merged_cells; fall back gracefully
    return merge_map


# ─── Staleness filtering ─────────────────────────────────────────────

# Substrings that identify the "Last updated" column (lower-cased, order matters).
_DATE_COL_HINTS = [
    "last update", "last änderung", "last change", "aktualisiert",
    "updated", "update date", "datum", "date",
]

# Date format strings tried in order when the cell value is a plain string.
# Covers the most common lazy-entry styles seen in German/European supply-chain Excel files.
_DATE_FORMATS = [
    "%d.%m.%Y",   # 15.03.2026  (European standard)
    "%d.%m.%y",   # 15.03.26
    "%d/%m/%Y",   # 15/03/2026
    "%d/%m/%y",   # 15/03/26
    "%m/%d/%Y",   # 03/15/2026  (US style, less common)
    "%m/%d/%y",   # 03/15/26
    "%Y-%m-%d",   # 2026-03-15  (ISO)
    "%d-%m-%Y",   # 15-03-2026
    "%d %b %Y",   # 15 Mar 2026
    "%d %B %Y",   # 15 March 2026
    "%b %d, %Y",  # Mar 15, 2026
    "%B %d, %Y",  # March 15, 2026
    "%d.%m",      # 15.03  (no year → current year assumed)
    "%d/%m",      # 15/03  (no year → current year assumed)
]


def _find_date_col(headers: list[str]) -> int | None:
    """
    Return the 0-based index of the "Last updated" column, or None if not found.
    Checks for substring matches against _DATE_COL_HINTS in priority order.
    """
    lowered = [h.lower().replace("\n", " ").strip() for h in headers]
    for hint in _DATE_COL_HINTS:
        for i, h in enumerate(lowered):
            if hint in h:
                return i
    return None


def _parse_date_value(raw: Any) -> date | None:
    """
    Convert a raw openpyxl cell value to a Python date.

    openpyxl returns proper Excel date cells as datetime objects already.
    Lazily-typed strings are tried against _DATE_FORMATS. Returns None if
    unparseable (caller should keep the row).
    """
    from datetime import date as _date, timedelta as _td

    if raw is None:
        return None

    # openpyxl native datetime / date (most reliable path)
    if isinstance(raw, datetime):
        return raw.date()
    if isinstance(raw, _date):
        return raw

    # Numeric: Excel stores dates as serial numbers; openpyxl normally converts
    # them, but with data_only=True on some files they may arrive as floats.
    if isinstance(raw, (int, float)):
        try:
            # Excel epoch: 1900-01-00 (with the off-by-two leap-year bug)
            base = _date(1899, 12, 30)
            return base + _td(days=int(raw))
        except (ValueError, OverflowError):
            return None

    # String — try all known formats
    s = str(raw).strip()
    if not s:
        return None

    # Remove weekday prefixes like "Mon, ", "Mo. " etc.
    s = re.sub(r"^[A-Za-z]{2,3}[,.\s]+", "", s).strip()

    current_year = datetime.now().year
    for fmt in _DATE_FORMATS:
        try:
            dt = datetime.strptime(s, fmt)
            # If the format has no year component, attach current year
            if "%Y" not in fmt and "%y" not in fmt:
                dt = dt.replace(year=current_year)
            return dt.date()
        except ValueError:
            continue

    return None  # unparseable → caller keeps the row


def _is_stale_by_date(raw_date_val: Any, cutoff: date) -> bool:
    """
    Return True if the parsed date is strictly before cutoff.
    Returns False (keep) when the value is None / unparseable.
    """
    from datetime import date as _date
    parsed = _parse_date_value(raw_date_val)
    if parsed is None:
        return False           # can't read the date → benefit of the doubt
    return parsed < cutoff


def _is_stale_by_cw(cells: list[str], current_week: int, history_weeks: int) -> bool:
    """
    Fallback: return True only when the row contains CW/KW references AND
    every one of them is older than (current_week − history_weeks).
    Rows with no CW references are kept (benefit of the doubt).
    """
    cutoff_week = current_week - history_weeks
    row_text = " ".join(cells)
    cw_nums = [int(m) for m in re.findall(r"\b(?:CW|KW|W)(\d{1,2})\b", row_text, re.IGNORECASE)]
    if not cw_nums:
        return False
    return all(n < cutoff_week for n in cw_nums)


def _find_customer_col_range(ws) -> tuple[int, int] | None:
    """
    Return the (start_col_idx, end_col_idx) range (0-based, inclusive) of the
    individual customer-name columns, identified by a merged super-header cell
    whose value contains both "customer" and "affect" (e.g. "Customer affected").

    Returns None if no such merged range is found.
    """
    try:
        for merge in ws.merged_cells.ranges:
            master_val = ws.cell(merge.min_row, merge.min_col).value or ""
            lowered = str(master_val).lower()
            if "customer" in lowered and "affect" in lowered:
                return (merge.min_col - 1, merge.max_col - 1)   # convert to 0-based
    except Exception:
        pass
    return None


def _compact_customer_cols(
    headers: list[str],
    rows: list[list[str]],
    cust_range: tuple[int, int] | None,
) -> tuple[list[str], list[list[str]]]:
    """
    Collapse the individual customer columns (W–BB in the standard template)
    into a single synthetic column called "Customers affected".

    For each data row the customer columns that carry a non-empty, non-"N"
    value are collected; their header names become a comma-separated list.
    This reduces a 64-column table to ~33 columns and makes the LLM prompt
    far more readable.

    If cust_range is None the headers and rows are returned unchanged.
    """
    if cust_range is None:
        return headers, rows

    start, end = cust_range
    customer_names = headers[start : end + 1]   # names of each customer column

    new_headers = headers[:start] + ["Customers affected"] + headers[end + 1 :]

    new_rows: list[list[str]] = []
    for row in rows:
        affected: list[str] = []
        for rel_i, cname in enumerate(customer_names):
            col_i = start + rel_i
            if col_i >= len(row):
                break
            val = row[col_i].strip()
            # Any non-empty, non-negative-looking value counts as "affected"
            if val and val.lower() not in ("", "n", "no", "0", "-", "·", "false"):
                affected.append(cname)
        customers_str = ", ".join(affected) if affected else ""
        tail = row[end + 1 :] if end + 1 < len(row) else []
        new_rows.append(row[:start] + [customers_str] + tail)

    return new_headers, new_rows


def _build_text_table(headers: list[str], rows: list[list[str]]) -> str:
    """Render headers + rows as a compact pipe-delimited markdown table."""
    # Cap column display width so the table stays manageable
    MAX_COL_W = 500
    col_widths = [min(max(len(h), 4), MAX_COL_W) for h in headers]
    for row in rows:
        for i, cell in enumerate(row[: len(headers)]):
            if i < len(col_widths):
                col_widths[i] = min(max(col_widths[i], len(cell)), MAX_COL_W)

    def fmt(cells: list[str]) -> str:
        padded = []
        for i, h in enumerate(headers):
            val = (cells[i] if i < len(cells) else "")[: MAX_COL_W]
            padded.append(val.ljust(col_widths[i]))
        return "| " + " | ".join(padded) + " |"

    sep = "|-" + "-|-".join("-" * w for w in col_widths) + "-|"
    lines = [fmt(headers), sep] + [fmt(row) for row in rows]
    return "\n".join(lines)


def excel_to_text_table(filepath: str, history_weeks: int = 4) -> dict:
    """
    Stage 1 (deterministic): Open the Excel workbook and produce a
    pipe-delimited text table with stale rows removed.

    Staleness is determined by the "Last updated" column (preferred): any row
    whose date is older than (today − history_weeks) is dropped. If no date
    column is identified, falls back to scanning rows for CW/KW references.

    Returns
    -------
    {
      text_table      : pipe-delimited table string (headers + data rows)
      headers         : list of column header strings
      current_cw      : "CW{week}/{year}"
      total_rows      : row count before filtering
      kept_rows       : row count after filtering
      sheet_used      : actual sheet name
      date_col_name   : header of the detected date column, or None
      warnings        : list of warning strings
    }
    """
    warnings: list[str] = []
    year, week = _get_current_cw()
    current_cw = f"CW{week}/{year}"

    # ── Open workbook (not read_only so merged_cells works) ──────────
    wb = openpyxl.load_workbook(filepath, data_only=True)

    # ── Locate the Dashboard_Update sheet ────────────────────────────
    sheet_used: str | None = None
    for name in wb.sheetnames:
        n = name.lower()
        if "dashboard" in n and "update" in n:
            sheet_used = name
            break
    if not sheet_used:
        for name in wb.sheetnames:
            if "dashboard" in name.lower() or "update" in name.lower():
                sheet_used = name
                break
    if not sheet_used:
        sheet_used = wb.sheetnames[0]
        warnings.append(f"'Dashboard_Update' sheet not found; using '{sheet_used}'")

    ws = wb[sheet_used]

    # ── Build merge map before anything else ─────────────────────────
    merge_map = _build_merge_map(ws)

    def cell_val(row: int, col: int) -> Any:
        return merge_map.get((row, col), ws.cell(row, col).value)

    max_col = ws.max_column or 1
    max_row = ws.max_row or 1

    # ── Detect header row (merge-aware, scans up to 60 rows) ─────────
    hdr_idx = _find_header_row(ws, merge_map=merge_map)

    # ── Score the row immediately below the header to decide whether it
    #    is a genuine sub-header continuation (e.g. a second header row
    #    with units or sub-labels) or the first data row.
    #    Only combine when the next row itself scores as header-like (≥ 2).
    #    This prevents absorbing data values (e.g. supplier names in the
    #    first data row) into column header strings.
    next_row_raw = [
        _cell_str(cell_val(hdr_idx + 1, c))
        for c in range(1, min(max_col, 60) + 1)
    ] if hdr_idx + 1 <= max_row else []
    next_row_is_subhdr = _header_score(next_row_raw) >= 2

    # ── Extract headers ───────────────────────────────────────────────
    headers: list[str] = []
    for col in range(1, max_col + 1):
        top = _cell_str(cell_val(hdr_idx, col))
        if next_row_is_subhdr:
            bot = next_row_raw[col - 1] if col - 1 < len(next_row_raw) else ""
            # Combine only when bot adds new information
            if bot and bot != top and not re.search(r"\d{2}\.\d{2}\.\d{4}", bot) \
                    and bot.lower() not in ("y", "n", "yes", "no"):
                combined = f"{top} {bot}".strip()
            else:
                combined = top
        else:
            combined = top
        headers.append(combined)

    # Trim trailing empty headers
    while headers and not headers[-1]:
        headers.pop()
    max_col = len(headers)

    # data starts on the row after the (possibly two-row) header band
    data_start = hdr_idx + 2 if next_row_is_subhdr else hdr_idx + 1

    # ── Detect customer column range (from merged "Customer affected" super-header) ──
    cust_range = _find_customer_col_range(ws)

    # ── Detect "Last updated" column ─────────────────────────────────
    date_col_idx: int | None = _find_date_col(headers)
    date_col_name: str | None = headers[date_col_idx] if date_col_idx is not None else None
    if date_col_name:
        warnings.append(f"Recency filter: using '{date_col_name}' column for date-based filtering.")
    else:
        warnings.append("Recency filter: no 'Last updated' column found — falling back to CW-number scanning.")

    # ── Extract and clean data rows ───────────────────────────────────
    # Each entry is (stringified_cells, raw_date_value) so we can filter
    # on the native openpyxl value (datetime object or raw string) before
    # discarding it.
    all_rows: list[tuple[list[str], Any]] = []
    for row_idx in range(data_start, max_row + 1):
        raw_cells = [cell_val(row_idx, col) for col in range(1, max_col + 1)]
        cells = [_cell_str(v) for v in raw_cells]

        # Skip empty rows
        if not any(cells):
            continue
        # Skip separator rows (---, ===, etc.)
        non_empty = [c for c in cells if c]
        if all(re.match(r"^[-=_\s]+$", c) for c in non_empty):
            continue
        # Skip rows with fewer than 3 non-empty cells (spacers / subtotals)
        if len(non_empty) < 3:
            continue

        # Extract raw date value before stringification is lost
        raw_date = raw_cells[date_col_idx] if date_col_idx is not None and date_col_idx < len(raw_cells) else None
        all_rows.append((cells[:max_col], raw_date))

    wb.close()

    total_rows = len(all_rows)
    cutoff_date = date.today() - timedelta(weeks=history_weeks)

    # ── Filter stale rows ─────────────────────────────────────────────
    kept: list[list[str]] = []
    skipped = 0

    for cells, raw_date in all_rows:
        if date_col_idx is not None:
            # Primary: date-column comparison
            if _is_stale_by_date(raw_date, cutoff_date):
                skipped += 1
            else:
                kept.append(cells)
        else:
            # Fallback: CW-number scan
            if _is_stale_by_cw(cells, week, history_weeks):
                skipped += 1
            else:
                kept.append(cells)

    if skipped:
        if date_col_idx is not None:
            warnings.append(
                f"Recency filter: removed {skipped} row(s) with '{date_col_name}' "
                f"older than {cutoff_date.isoformat()} ({history_weeks}-week window)."
            )
        else:
            warnings.append(
                f"Recency filter (CW fallback): removed {skipped} row(s) with all "
                f"CW references before CW{max(1, week - history_weeks)}."
            )
    if not kept:
        warnings.append("No rows remain after recency filtering — consider increasing history_weeks.")

    # ── Compact the 32 individual customer columns → one synthetic column ──
    #    This reduces table width from 64 → ~33 columns, keeping the LLM
    #    prompt manageable. The actual customer names become a comma-separated
    #    value when they are marked as affected in the row.
    compacted_headers, kept = _compact_customer_cols(headers, kept, cust_range)

    text_table = _build_text_table(compacted_headers, kept)

    return {
        "text_table":     text_table,
        "headers":        compacted_headers,
        "original_headers": headers,
        "current_cw":     current_cw,
        "total_rows":     total_rows,
        "kept_rows":      len(kept),
        "sheet_used":     sheet_used,
        "date_col_name":  date_col_name,
        "customer_col_range": cust_range,
        "warnings":       warnings,
    }


# ─── Stage 2: LLM extraction ─────────────────────────────────────────

_EXTRACT_SYSTEM = """\
You are a supply chain data extraction specialist for an automotive parts manufacturer.

You will receive a pipe-delimited table exported from the "Dashboard_Update" worksheet of
a Global Fulfilment Dashboard Excel file. Your task is to extract EVERY row into a
structured JSON object.

═══ EXTRACTION RULES ═══

1. PRODUCT GROUP IDENTIFICATION
   Product family / group cells are often merged vertically in Excel, meaning the same
   value appears on multiple consecutive rows in our table. Group all consecutive rows
   that share the same product family code or description into one product_group entry.
   If no product family column exists, infer groupings from context (e.g. shared
   component type or shared root cause block). Use "Unknown" only as a last resort.

2. COVERAGE CW FIELDS
   For "Coverage w/o mitigation" and "Coverage w/ mitigation" (and any synonyms like
   "Coverage without", "Coverage with", "Deckung ohne", "Deckung mit"):
   — Extract the CW number as a plain INTEGER (e.g. "CW15" → 15, "KW15/2026" → 15).
   — If the cell is empty or unparseable, use null.

3. TEXT FIELDS — preserve verbatim. Do NOT paraphrase, shorten, or infer missing info.

4. BOOLEAN-LIKE FIELDS (Customer Informed, Allocation Mode, Force Majeure / FM):
   Normalise to exactly one of: "Yes", "No", "In progress", "N/A".

5. PERCENTAGE FIELDS — keep original string (e.g. "85%").

6. CURRENCY AMOUNTS — extract as a float (e.g. "€ 45.000,00" → 45000.0).

7. EMPTY / N/A CELLS — use null.

8. EVERY DATA ROW must appear in the output — skip nothing.

═══ OUTPUT FORMAT ═══

Respond with ONLY valid JSON — no markdown fences, no explanation text.

{{
  "current_cw": "CW13/2026",
  "sheet_name": "Dashboard_Update",
  "extraction_notes": "brief note on data quality, ambiguities, or assumptions",
  "product_groups": [
    {{
      "product_family_code": "SEN",
      "product_family_desc": "Sensors / Radar",
      "rows": [
        {{
          "plant_location": "BHV",
          "region": "Europe",
          "customer_affected": "BMW Group, Mercedes-Benz AG",
          "critical_component": "NXP S32K wafer",
          "constraint_task_force": "Task Force Alpha",
          "root_cause": "NXP wafer fab allocation cut by 30% following process incident",
          "supplier_text": "NXP Semiconductors",
          "supplier_type": "Tier 1",
          "supplier_region": "Netherlands",
          "coverage_without_mitigation_cw": 15,
          "coverage_with_mitigation_cw": 19,
          "fulfillment_current_q": "85%",
          "fulfillment_q_plus_1": "70%",
          "fulfillment_q_plus_2": null,
          "recovery_week": "CW22",
          "ops_capacity_risk": "Yes",
          "strategic_capacity_risk": "No",
          "allocation_mode": "Yes",
          "customer_informed": "In progress",
          "action_comment": "Activating dual source at Plant Hannover by CW14; air freight bridge for 3 weeks",
          "task_force_leader": "J. Mueller",
          "special_freight_cost_eur": 45000.0,
          "special_freight_remarks": "Air freight BHV→HAN weekly until CW18"
        }}
      ]
    }}
  ]
}}
{glossary_block}"""

_EXTRACT_USER = """\
Current calendar week: {current_cw}
Sheet name: {sheet_name}
Data rows after recency filter: {kept_rows} (of {total_rows} total)

Extract all product groups and risk rows from the table below.
Remember: rows with the same repeating product family value belong to the same product group.

{text_table}"""


async def llm_extract_gfd_data(
    stage1: dict,
    llm_config: dict,
    session_id: str,
    glossary_context: str = "",
) -> dict:
    """
    Stage 2 (LLM): Send the filtered text table to the LLM and extract structured JSON.

    Returns the parsed extraction dict. On JSON parse failure, returns a minimal
    fallback structure with the error noted in warnings.
    """
    llm = _create_llm(llm_config, max_tokens=8192)
    t0 = time.time()

    glossary_block = (
        f"\n\nCOMPANY GLOSSARY — use these to expand abbreviations correctly:\n{glossary_context}"
        if glossary_context else ""
    )

    messages = [
        SystemMessage(content=_EXTRACT_SYSTEM.format(glossary_block=glossary_block)),
        HumanMessage(content=_EXTRACT_USER.format(
            current_cw=stage1["current_cw"],
            sheet_name=stage1["sheet_used"],
            kept_rows=stage1["kept_rows"],
            total_rows=stage1["total_rows"],
            text_table=stage1["text_table"],
        )),
    ]

    try:
        response = await llm.ainvoke(messages)
        raw = response.content.strip()

        # Strip markdown fences if the model disobeyed instructions
        if raw.startswith("```"):
            raw = "\n".join(raw.split("\n")[1:])
        if raw.endswith("```"):
            raw = "\n".join(raw.split("\n")[:-1])

        extracted: dict = json.loads(raw.strip())

        usage = response.response_metadata.get("token_usage", {})
        log_tokens(session_id, "gfd_llm_extract", usage, llm_config.get("azure_deployment", ""))

        n_pgs = len(extracted.get("product_groups", []))
        n_rows = sum(len(pg.get("rows", [])) for pg in extracted.get("product_groups", []))
        duration = (time.time() - t0) * 1000
        log_trace(
            session_id, "gfd_llm_extract",
            f"Input: {stage1['kept_rows']} rows, {len(stage1['headers'])} columns",
            f"Extracted {n_pgs} product groups, {n_rows} rows",
            duration,
        )

        # Attach Stage 1 metadata and merge warnings
        extracted.setdefault("warnings", [])
        extracted["warnings"] = stage1["warnings"] + extracted["warnings"]
        extracted["_meta"] = {
            "headers": stage1["headers"],
            "original_headers": stage1.get("original_headers", stage1["headers"]),
            "total_rows_in_file": stage1["total_rows"],
            "rows_after_filter": stage1["kept_rows"],
            "sheet_used": stage1["sheet_used"],
            "date_col_name": stage1.get("date_col_name"),
            "customer_col_range": stage1.get("customer_col_range"),
        }
        return extracted

    except json.JSONDecodeError as exc:
        duration = (time.time() - t0) * 1000
        log_trace(session_id, "gfd_llm_extract",
                  f"Input: {stage1['kept_rows']} rows",
                  f"JSON PARSE ERROR: {str(exc)[:120]}", duration, {"error": True})
        return _extraction_fallback(stage1, f"LLM returned unparseable JSON: {str(exc)[:200]}")

    except Exception as exc:
        duration = (time.time() - t0) * 1000
        log_trace(session_id, "gfd_llm_extract",
                  f"Input: {stage1['kept_rows']} rows",
                  f"ERROR: {str(exc)[:120]}", duration, {"error": True})
        return _extraction_fallback(stage1, f"LLM extraction failed: {str(exc)[:200]}")


def _extraction_fallback(stage1: dict, error_msg: str) -> dict:
    """Minimal structure returned when LLM extraction fails entirely."""
    return {
        "current_cw": stage1["current_cw"],
        "sheet_name": stage1["sheet_used"],
        "extraction_notes": f"FALLBACK — {error_msg}",
        "product_groups": [],
        "warnings": stage1["warnings"] + [error_msg],
        "_meta": {
            "headers": stage1["headers"],
            "original_headers": stage1.get("original_headers", stage1["headers"]),
            "total_rows_in_file": stage1["total_rows"],
            "rows_after_filter": stage1["kept_rows"],
            "sheet_used": stage1["sheet_used"],
            "date_col_name": stage1.get("date_col_name"),
            "customer_col_range": stage1.get("customer_col_range"),
        },
    }


# ─── Combined entry point ─────────────────────────────────────────────

async def parse_gfd_with_llm(
    filepath: str,
    llm_config: dict,
    session_id: str,
    history_weeks: int = 4,
    glossary_context: str = "",
) -> dict:
    """
    Full two-stage GFD parsing pipeline.

    Stage 1: deterministic Excel → filtered text table
    Stage 2: LLM text table → structured JSON

    Returns the LLM-extracted JSON dict (see _EXTRACT_SYSTEM for full schema).
    """
    stage1 = excel_to_text_table(filepath, history_weeks=history_weeks)
    return await llm_extract_gfd_data(stage1, llm_config, session_id, glossary_context)
