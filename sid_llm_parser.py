"""
SID LLM Parser
==============
Two-stage pipeline for extracting structured data from Supplier Impact Tracking
Excel files (KB Input worksheet).

Stage 1 — Pandas-based deterministic extraction:
  • Open the workbook and find the KB Input sheet
  • Detect the header row via token-based scoring (robust to title rows,
    logos, merged headers)
  • Extract a clean DataFrame with deduplicated column names
  • Filter stale rows using the "Date" column (current + previous month)
  • Convert the filtered DataFrame to a pipe-delimited text table

Stage 2 — LLM extraction:
  • Send the filtered text table to the LLM in chunks
  • LLM understands column semantics, normalises fields, and returns
    a precise JSON object per chunk
  • Results are merged across chunks

The output JSON is consumed by sid_llm_slides.py.
"""

from __future__ import annotations

import re
import json
import time
from datetime import datetime
from pathlib import Path
from typing import Any

import numpy as np
import pandas as pd
import openpyxl
from langchain_openai import AzureChatOpenAI
from langchain_core.messages import HumanMessage, SystemMessage

from agent import log_tokens, log_trace


# ─── JSON repair (shared with GFD) ──────────────────────────────────

def _repair_llm_json(raw: str, debug_label: str = "") -> str:
    """
    Attempt to clean common JSON defects produced by LLMs before parsing.

    Handles:
      1. Markdown fences (```json ... ```)
      2. Trailing commas before } or ]
      3. Unescaped literal newlines / tabs inside JSON string values
    """
    s = raw.strip()

    # 1. Strip markdown fences
    if s.startswith("```"):
        first_nl = s.index("\n") if "\n" in s else len(s)
        s = s[first_nl + 1:]
    if s.endswith("```"):
        s = s[: s.rfind("```")]
    s = s.strip()

    # 2. Trailing commas
    s = re.sub(r",\s*([}\]])", r"\1", s)

    # 3. Unescaped control characters inside string values
    result = []
    in_string = False
    i = 0
    while i < len(s):
        ch = s[i]
        if ch == '"' and (i == 0 or s[i - 1] != '\\'):
            in_string = not in_string
            result.append(ch)
        elif in_string:
            if ch == '\n':
                result.append('\\n')
            elif ch == '\r':
                result.append('\\r')
            elif ch == '\t':
                result.append('\\t')
            else:
                result.append(ch)
        else:
            result.append(ch)
        i += 1
    s = "".join(result)

    if debug_label:
        print(f"[SID DEBUG] {debug_label}: JSON repair applied "
              f"({len(raw)} → {len(s)} chars)")
    return s


def _parse_llm_json(raw: str, session_id: str = "", attempt: int = 0) -> dict:
    """Parse LLM JSON output with safety-net repair."""
    label = f"attempt{attempt}" if attempt else ""
    repaired = _repair_llm_json(raw, debug_label=label)
    try:
        return json.loads(repaired)
    except json.JSONDecodeError as exc:
        pos = exc.pos or 0
        start = max(0, pos - 120)
        end = min(len(repaired), pos + 120)
        context = repaired[start:end]
        pointer_offset = pos - start
        print(f"[SID ERROR] JSON parse failed — "
              f"line {exc.lineno}, col {exc.colno}, char {pos}")
        print(f"[SID ERROR] Context: ...{context}...")
        print(f"  {' ' * (pointer_offset + 5)}^ error here")
        raise


# ─── CW / date utilities ────────────────────────────────────────────

def _get_current_cw() -> tuple[int, int]:
    """Return (year, iso_week) for today."""
    iso = datetime.now().isocalendar()
    return (iso.year, iso.week)


# ─── LLM factory ────────────────────────────────────────────────────

def _create_llm(config: dict, max_tokens: int = 32000) -> AzureChatOpenAI:
    return AzureChatOpenAI(
        azure_deployment=config["azure_deployment"],
        azure_endpoint=config["azure_endpoint"],
        api_key=config["api_key"],
        api_version=config.get("api_version", "2024-12-01-preview"),
        max_tokens=max_tokens,
    )


# ═══════════════════════════════════════════════════════════════════════
#  Stage 1: Pandas-based Excel → filtered pipe-delimited text table
# ═══════════════════════════════════════════════════════════════════════

# Tokens that strongly indicate the real header row for KB Input
_HEADER_TOKENS = {
    "sn", "cat", "vendor code", "vendor", "category buyer",
    "part description", "part decription", "process impacted", "location",
    "formal notice available ?", "formal notice available?",
    "remarks", "severity", "date", "coverage",
    "fm rejection sent?", "fm rejection sent",
    "supplier email sent", "fuel/gas being used",
    "current fuel coverage", "device / product line",
    "dom / ico", "dom customer name", "ico customer name",
}


def _normalize_cell(x) -> str:
    """Normalize a single cell value for header detection."""
    if x is None or (isinstance(x, float) and np.isnan(x)):
        return ""
    if pd.isna(x):
        return ""
    s = str(x).strip().lower()
    s = re.sub(r"\s+", " ", s)
    return s


def _cellwise_map(df: pd.DataFrame, func) -> pd.DataFrame:
    """Safe element-wise map across pandas versions."""
    if hasattr(df, "map") and callable(getattr(df, "map")):
        try:
            return df.map(func)
        except Exception:
            pass
    if hasattr(df, "applymap") and callable(getattr(df, "applymap")):
        try:
            return df.applymap(func)
        except Exception:
            pass
    vals = df.to_numpy()
    vfunc = np.vectorize(func, otypes=[object])
    return pd.DataFrame(vfunc(vals), index=df.index, columns=df.columns)


def _find_header_row(raw_df: pd.DataFrame,
                     header_tokens: set[str] | None = None,
                     min_hits: int = 4) -> tuple[int, int]:
    """
    Scan each row and score it by how many header_tokens appear.
    Return (row_index, score).
    """
    if header_tokens is None:
        header_tokens = _HEADER_TOKENS

    best_idx = 0
    best_score = -1

    for i in range(len(raw_df)):
        row_vals = [_normalize_cell(v) for v in raw_df.iloc[i].tolist()]
        row_set = {v for v in row_vals if v}
        # Token containment — also check substring match for flexible headers
        score = 0
        for t in header_tokens:
            if t in row_set:
                score += 1
            elif any(t in cell for cell in row_set):
                score += 0.5
        if score > best_score:
            best_score = score
            best_idx = i

    if best_score < min_hits:
        # Fallback: row with most non-empty cells
        non_empty_counts = raw_df.apply(
            lambda r: sum(1 for v in r.values if _normalize_cell(v)), axis=1
        )
        best_idx = int(non_empty_counts.idxmax())
        best_score = 0
        print(f"[SID WARNING] Low header score ({best_score}); using most-populated row {best_idx}")

    return best_idx, int(best_score)


def _detect_kb_input_sheet(xlsx_path: str) -> tuple[str, list[str]]:
    """Find the best sheet name matching 'KB Input'."""
    warnings: list[str] = []
    wb = openpyxl.load_workbook(xlsx_path, read_only=True)
    names = wb.sheetnames
    wb.close()

    # Exact substring match
    for name in names:
        n = name.lower().replace("_", " ")
        if "kb" in n and "input" in n:
            return name, warnings

    # Partial match
    for name in names:
        n = name.lower()
        if "kb" in n or "input" in n or "supplier" in n or "tracking" in n:
            return name, warnings

    warnings.append(f"'KB Input' sheet not found; using '{names[0]}'")
    return names[0], warnings


def _extract_kb_input_dataframe(
    xlsx_path: str,
    sheet_name: str,
) -> tuple[pd.DataFrame, int]:
    """
    Extract a clean DataFrame from a KB Input-style Excel sheet.
    Returns (df, header_row_idx).
    """
    raw = pd.read_excel(
        xlsx_path, sheet_name=sheet_name,
        header=None, engine="openpyxl", dtype=str,
    )
    raw = raw.dropna(how="all")

    header_idx, score = _find_header_row(raw)
    print(f"[SID DEBUG] Header detected at row {header_idx} (score={score})")

    # Build header from detected row
    header = raw.iloc[header_idx].astype(str).str.strip()
    header = header.str.replace(r"\s+", " ", regex=True)

    # Deduplicate headers
    cleaned: list[str] = []
    seen: dict[str, int] = {}
    for i, h in enumerate(header):
        h = h.strip() if isinstance(h, str) and h != "nan" else ""
        if not h or h == "nan":
            h = f"Unnamed_{i}"
        key = h.lower()
        if key in seen:
            seen[key] += 1
            h = f"{h}__{seen[key]}"
        else:
            seen[key] = 0
        cleaned.append(h)

    body = raw.iloc[header_idx + 1:].copy()
    body.columns = cleaned

    # Drop fully-empty columns and rows
    body = body.dropna(axis=1, how="all")
    body = body.dropna(axis=0, how="all")

    # Filter out footer/legend rows using Sn column (serial number)
    sn_candidates = [c for c in body.columns if c.strip().lower() in ("sn", "s/n", "s.no", "sr", "sr.", "no")]
    if sn_candidates:
        sn_col = sn_candidates[0]
        sn_numeric = pd.to_numeric(body[sn_col], errors="coerce")
        non_null_count = body.notna().sum(axis=1)
        body = body[(sn_numeric.notna()) | ((sn_numeric.isna()) & (non_null_count >= 4))]

    # Normalise whitespace in string columns
    for col in body.columns:
        if body[col].dtype == object:
            body[col] = body[col].apply(
                lambda x: re.sub(r"\s+", " ", str(x)).strip() if pd.notna(x) and str(x).strip().lower() != "nan" else x
            )

    return body, header_idx


# ─── Date-based recency filtering ────────────────────────────────────

_DATE_COL_HINTS = [
    "date", "last update", "last change", "updated", "update date", "datum",
]


def _find_date_column(columns: list[str]) -> str | None:
    """Return the column name matching a date hint, or None."""
    for hint in _DATE_COL_HINTS:
        for col in columns:
            if hint in col.lower().replace("\n", " ").strip():
                return col
    return None


def _filter_by_recent_months(
    df: pd.DataFrame,
    date_col: str,
) -> tuple[pd.DataFrame, int, str]:
    """
    Keep rows whose date_col falls within current month or previous month.
    Rows with unparseable dates are kept (benefit of the doubt).
    """
    dates = pd.to_datetime(df[date_col], errors="coerce")

    today = pd.Timestamp.today().normalize()
    start_current_month = today.replace(day=1)
    start_previous_month = start_current_month - pd.offsets.MonthBegin(1)
    end_current_month = start_current_month + pd.offsets.MonthEnd(1)

    prev_label = start_previous_month.strftime("%b %Y")
    curr_label = start_current_month.strftime("%b %Y")
    desc = f"{prev_label} – {curr_label}"

    in_window = (dates >= start_previous_month) & (dates < end_current_month)
    # Keep NaT rows (unparseable or missing)
    keep_mask = in_window | dates.isna()
    filtered = df[keep_mask]
    num_removed = int((~keep_mask).sum())

    return filtered, num_removed, desc


# ─── Text table builder ─────────────────────────────────────────────

def _cell_str(value: Any) -> str:
    """Convert a DataFrame cell value to a clean string."""
    if value is None or (isinstance(value, float) and np.isnan(value)):
        return ""
    if isinstance(value, datetime):
        return value.strftime("%d.%m.%Y")
    if isinstance(value, pd.Timestamp):
        return value.strftime("%d.%m.%Y")
    s = re.sub(r"[\r\n]+", " ", str(value).strip())
    if s.lower() == "nan":
        return ""
    # Strip leading/trailing apostrophes
    if s.startswith("'") or s.endswith("'"):
        s = s.strip("'").strip()
    return s


def _build_text_table(headers: list[str], rows: list[list[str]]) -> str:
    """Render headers + rows as a compact pipe-delimited table with row labels."""
    MAX_PAD_W = 80
    ROW_NUM_W = 4

    col_widths = [min(max(len(h), 4), MAX_PAD_W) for h in headers]
    for row in rows:
        for i, cell in enumerate(row[:len(headers)]):
            if i < len(col_widths):
                col_widths[i] = min(max(col_widths[i], len(cell)), MAX_PAD_W)

    def fmt(cells: list[str], row_label: str = "") -> str:
        padded = []
        if row_label is not None:
            padded.append(row_label.ljust(ROW_NUM_W))
        for i, h in enumerate(headers):
            val = cells[i] if i < len(cells) else ""
            padded.append(val.ljust(col_widths[i]))
        return "| " + " | ".join(padded) + " |"

    hdr_line = fmt(headers, row_label="#")
    sep_widths = [ROW_NUM_W] + col_widths
    sep = "|-" + "-|-".join("-" * w for w in sep_widths) + "-|"
    data_lines = [fmt(row, row_label=f"R{i+1:03d}") for i, row in enumerate(rows)]
    return "\n".join([hdr_line, sep] + data_lines)


# ─── Main Stage 1 entry point ───────────────────────────────────────

def excel_to_text_table(filepath: str, **_kwargs) -> dict:
    """
    Stage 1 (deterministic): Open the Excel workbook and produce a
    pipe-delimited text table with rows filtered to the current and
    previous calendar month.

    Returns dict with text_table, headers, metadata, warnings, etc.
    """
    year, week = _get_current_cw()
    current_cw = f"CW{week}/{year}"
    today_str = datetime.now().strftime("%d.%m.%Y")

    print(f"\n[SID DEBUG] ═══ Stage 1: Excel → text table ═══")
    print(f"[SID DEBUG] File: {filepath}")
    print(f"[SID DEBUG] Current CW: {current_cw}")

    # Detect sheet name
    sheet_used, warnings = _detect_kb_input_sheet(filepath)
    print(f"[SID DEBUG] Sheet detected: '{sheet_used}'")

    # Extract clean DataFrame
    try:
        df, header_idx = _extract_kb_input_dataframe(filepath, sheet_name=sheet_used)
    except Exception as exc:
        raise RuntimeError(
            f"Failed to extract DataFrame from sheet '{sheet_used}': {exc}"
        ) from exc

    original_headers = list(df.columns)
    total_rows = len(df)
    print(f"[SID DEBUG] Extracted: {total_rows} rows × {len(original_headers)} cols "
          f"(header at row {header_idx})")

    if total_rows == 0:
        warnings.append("No data rows found below the detected header row.")

    # Write extracted CSV (full dataset before filtering)
    csv_path = str(Path(filepath).with_suffix(".csv"))
    try:
        df.to_csv(csv_path, index=False, encoding="utf-8-sig")
        warnings.append(f"Extracted CSV saved: {csv_path}")
    except Exception as csv_exc:
        warnings.append(f"Could not save extracted CSV: {csv_exc}")
        csv_path = None

    # Filter rows by recency
    date_col_name = _find_date_column(list(df.columns))
    print(f"[SID DEBUG] Date column: {date_col_name or '(none found)'}")

    if date_col_name:
        pre_filter = len(df)
        df, skipped, window_desc = _filter_by_recent_months(df, date_col_name)
        warnings.append(
            f"Recency filter: keeping rows from {window_desc} "
            f"(based on '{date_col_name}' column)."
        )
        if skipped:
            warnings.append(
                f"Recency filter: removed {skipped} row(s) outside the "
                f"{window_desc} window."
            )
        print(f"[SID DEBUG] Recency filter: {pre_filter} → {len(df)} rows "
              f"(removed {skipped}, window: {window_desc})")
    else:
        warnings.append(
            "Recency filter: no 'Date' column found — "
            "all rows retained (no filtering applied)."
        )

    kept_rows = len(df)
    if kept_rows == 0:
        warnings.append(
            "No rows remain after recency filtering — "
            "check that the 'Date' column has dates in the "
            "current or previous month."
        )

    # Save filtered CSV
    filtered_csv_path = str(Path(filepath).with_name(
        Path(filepath).stem + "_filtered.csv"
    ))
    try:
        df.to_csv(filtered_csv_path, index=False, encoding="utf-8-sig")
        warnings.append(f"Filtered CSV saved: {filtered_csv_path} ({kept_rows} rows)")
    except Exception as fcsv_exc:
        warnings.append(f"Could not save filtered CSV: {fcsv_exc}")
        filtered_csv_path = None

    # Build pipe-delimited text table
    headers = list(df.columns)
    rows: list[list[str]] = []
    for _, row in df.iterrows():
        rows.append([_cell_str(v) for v in row.values])

    text_table = _build_text_table(headers, rows)

    text_table_chars = len(text_table)
    text_table_lines = text_table.count("\n") + 1
    est_input_tokens = text_table_chars // 4
    print(f"[SID DEBUG] Text table: {kept_rows} data rows, {text_table_lines} lines, "
          f"{text_table_chars:,} chars (~{est_input_tokens:,} tokens est.)")
    print(f"[SID DEBUG] ═══ Stage 1 complete ═══\n")

    return {
        "text_table":         text_table,
        "headers":            headers,
        "original_headers":   original_headers,
        "current_cw":         current_cw,
        "today":              today_str,
        "total_rows":         total_rows,
        "kept_rows":          kept_rows,
        "sheet_used":         sheet_used,
        "date_col_name":      date_col_name,
        "csv_path":           csv_path,
        "filtered_csv_path":  filtered_csv_path,
        "warnings":           warnings,
    }


# ═══════════════════════════════════════════════════════════════════════
#  Stage 2: LLM extraction
# ═══════════════════════════════════════════════════════════════════════

_EXTRACT_SYSTEM = """\
You are a supply chain data extraction specialist for an automotive parts manufacturer.

You will receive a pipe-delimited table exported from the "KB Input" worksheet of
a Supplier Impact Tracking Excel file. Your task is to extract EVERY row into a
structured JSON object.

═══ EXTRACTION RULES ═══

1. FIELD MAPPING
   Map each column to the corresponding JSON field. Column headers may vary
   slightly (e.g. "Part Decription" vs "Part Description") — use semantic
   matching to identify the correct column.

2. TEXT FIELDS — preserve content verbatim but flatten to single-line strings.
   Replace line breaks, tabs, and control characters with "; " or " ".
   Example: a cell containing
       "Supplier confirmed shortage
        Alternative sourcing started"
   should become: "Supplier confirmed shortage; Alternative sourcing started"

3. DATE FIELDS — preserve in their original format (e.g. "25.03.2026").

4. SEVERITY — normalise to exactly one of: "R" (Red), "Y" (Yellow), "G" (Green),
   or null if missing / unclear.

5. BOOLEAN-LIKE FIELDS (Formal Notice, FM Rejection, Supplier Email, etc.):
   Normalise to: "Yes", "No", "In progress", "N/A", or null.

6. NUMERIC FIELDS (coverage days, Sn, Vendor Code):
   Extract as numbers where possible. Use null for empty/unparseable.

7. EMPTY / N/A CELLS — use null.

8. ROW COMPLETENESS — CRITICAL
   — The first column of each data row is a sequential row number (R001, R002, …).
     Your output must contain EXACTLY as many rows as there are numbered data rows
     in the input table.  The user message states the exact count.
   — NEVER merge, deduplicate, or summarise similar-looking rows.
   — After generating your JSON, mentally verify the row count.

═══ OUTPUT FORMAT ═══

Respond with ONLY valid JSON — no markdown fences, no explanation text.

CRITICAL JSON RULES:
  • Every string value must be a single line — no literal newline, tab, or
    carriage-return characters inside strings.
  • No trailing commas before }} or ]].
  • The output must be parseable by a strict JSON parser (Python json.loads).

{{
  "extraction_notes": "brief note on data quality or assumptions",
  "suppliers": [
    {{
      "sn": 1,
      "cat": "SC",
      "vendor_code": "12345",
      "vendor": "Supplier Name",
      "category_buyer": "Buyer Name",
      "part_description": "Part XYZ",
      "process_impacted": "Casting, Machining",
      "location": "City, Country",
      "formal_notice_available": "Yes",
      "fm_rejection_sent": "No",
      "supplier_email_sent": "Yes",
      "fuel_gas_used": "Natural Gas",
      "date": "25.03.2026",
      "current_fuel_coverage": "5 days",
      "al_other_rm_coverage_days": 10,
      "total_coverage_fg_days": 15,
      "severity": "R",
      "supplier_integrator_na": null,
      "remarks": "Root cause description; mitigation steps",
      "device_product_line": "Product A",
      "dom_ico": "DOM",
      "gso": null,
      "ico_location": null,
      "ico_customer_name": null,
      "dom_customer_name": "Customer X",
      "receiving_plant_ico": null,
      "receiving_plant_gso": null
    }}
  ]
}}
{glossary_block}"""

_EXTRACT_USER = """\
Current date: {today}
Current calendar week: {current_cw}
Sheet name: {sheet_name}

The table below contains EXACTLY {kept_rows} data rows ({first_row_label}–{last_row_label}), filtered from {total_rows} total.
Your JSON output MUST contain exactly {kept_rows} supplier entries — no more, no fewer.
Do NOT merge, skip, or summarise any rows even if they look similar.

{text_table}"""


def _count_extracted_rows(extracted: dict) -> int:
    """Count supplier rows in the LLM output."""
    return len(extracted.get("suppliers", []))


# ─── Chunked extraction ─────────────────────────────────────────────

_CHUNK_SIZE = 5


def _chunk_text_table(text_table: str, chunk_size: int = _CHUNK_SIZE,
                      ) -> list[dict]:
    """Split a pipe-delimited text table into chunks."""
    lines = text_table.split("\n")
    header_line = lines[0]
    sep_line    = lines[1]
    data_lines  = lines[2:]

    if not data_lines:
        return [{"text_table": text_table, "row_count": 0,
                 "first_label": "", "last_label": ""}]

    chunks: list[dict] = []
    for start in range(0, len(data_lines), chunk_size):
        batch = data_lines[start : start + chunk_size]
        chunk_table = "\n".join([header_line, sep_line] + batch)
        first_label = batch[0].split("|")[1].strip() if batch else ""
        last_label  = batch[-1].split("|")[1].strip() if batch else ""
        chunks.append({
            "text_table": chunk_table,
            "row_count":  len(batch),
            "first_label": first_label,
            "last_label":  last_label,
        })
    return chunks


async def _extract_single_chunk(
    chunk: dict,
    stage1: dict,
    llm: AzureChatOpenAI,
    system_prompt: str,
    session_id: str,
    llm_config: dict,
    chunk_idx: int,
    total_chunks: int,
) -> dict | None:
    """Send a single text-table chunk to the LLM and return parsed JSON."""
    chunk_label = f"chunk {chunk_idx + 1}/{total_chunks}"
    row_count  = chunk["row_count"]
    first_lbl  = chunk["first_label"]
    last_lbl   = chunk["last_label"]

    user_msg = _EXTRACT_USER.format(
        today=stage1["today"],
        current_cw=stage1["current_cw"],
        sheet_name=stage1["sheet_used"],
        kept_rows=row_count,
        total_rows=stage1["total_rows"],
        first_row_label=first_lbl,
        last_row_label=last_lbl,
        text_table=chunk["text_table"],
    )

    messages = [
        SystemMessage(content=system_prompt),
        HumanMessage(content=user_msg),
    ]

    t0 = time.time()
    try:
        response = await llm.ainvoke(messages)
        raw = response.content.strip()

        # Dump raw response for debugging
        try:
            csv_path = stage1.get("csv_path")
            dump_dir = Path(csv_path).parent if csv_path else Path(".")
            dump_path = dump_dir / f"{session_id}_sid_extract_{chunk_label.replace(' ', '_').replace('/', 'of')}_raw.txt"
            dump_path.write_text(raw, encoding="utf-8")
        except Exception:
            pass

        extracted = _parse_llm_json(raw, session_id=session_id, attempt=chunk_idx)

        usage = response.response_metadata.get("token_usage", {})
        completion_tokens = usage.get("completion_tokens", 0)
        log_tokens(session_id, f"sid_llm_extract_{chunk_label}", usage,
                   llm_config.get("azure_deployment", ""))

        n_rows_out = _count_extracted_rows(extracted)
        duration = (time.time() - t0) * 1000

        print(f"[SID DEBUG] {chunk_label} ({first_lbl}–{last_lbl}): "
              f"sent {row_count} rows → got {n_rows_out} rows back "
              f"({completion_tokens} tokens, {duration:.0f}ms)")

        if n_rows_out < row_count:
            print(f"[SID WARNING] {chunk_label}: lost {row_count - n_rows_out} rows")

        log_trace(session_id, f"sid_llm_extract_{chunk_label}",
                  f"Input: {row_count} rows ({first_lbl}–{last_lbl})",
                  f"Extracted {n_rows_out} rows", duration)
        return extracted

    except Exception as exc:
        duration = (time.time() - t0) * 1000
        print(f"[SID ERROR] {chunk_label}: {str(exc)[:120]}")
        log_trace(session_id, f"sid_llm_extract_{chunk_label}",
                  f"Input: {row_count} rows ({first_lbl}–{last_lbl})",
                  f"ERROR: {str(exc)[:120]}", duration, {"error": True})
        return None


def _merge_extracted_chunks(
    chunk_results: list[dict | None],
    current_cw: str,
    sheet_name: str,
    today: str,
) -> dict:
    """Merge supplier lists from multiple chunk extraction results."""
    all_suppliers: list[dict] = []
    all_notes: list[str] = []
    all_warnings: list[str] = []

    for chunk_result in chunk_results:
        if chunk_result is None:
            continue
        notes = chunk_result.get("extraction_notes", "")
        if notes:
            all_notes.append(notes)
        for w in chunk_result.get("warnings", []):
            if w not in all_warnings:
                all_warnings.append(w)
        all_suppliers.extend(chunk_result.get("suppliers", []))

    return {
        "current_cw": current_cw,
        "today": today,
        "sheet_name": sheet_name,
        "extraction_notes": "; ".join(all_notes) if all_notes else "",
        "suppliers": all_suppliers,
        "warnings": all_warnings,
    }


# ─── Main Stage 2 entry point ───────────────────────────────────────

async def llm_extract_sid_data(
    stage1: dict,
    llm_config: dict,
    session_id: str,
    glossary_context: str = "",
    chunk_size: int = _CHUNK_SIZE,
) -> dict:
    """
    Stage 2 (LLM): Extract structured JSON from the filtered text table
    using chunked extraction.
    """
    input_rows = stage1["kept_rows"]
    llm = _create_llm(llm_config, max_tokens=32000)
    t0 = time.time()

    print(f"\n[SID DEBUG] ═══ Stage 2: LLM extraction (chunked, {chunk_size} rows/chunk) ═══")
    print(f"[SID DEBUG] Input: {input_rows} data rows, {len(stage1['headers'])} columns")

    glossary_block = (
        f"\n\nCOMPANY GLOSSARY — use these to expand abbreviations correctly:\n{glossary_context}"
        if glossary_context else ""
    )

    # Save text table for debugging
    text_table_path = None
    if stage1.get("csv_path"):
        text_table_path = str(Path(stage1["csv_path"]).with_name(
            Path(stage1["csv_path"]).stem + "_llm_input.txt"
        ))
        try:
            with open(text_table_path, "w", encoding="utf-8") as f:
                f.write(f"Session: {session_id}\n")
                f.write(f"Current CW: {stage1['current_cw']}\n")
                f.write(f"Sheet: {stage1['sheet_used']}\n")
                f.write(f"Input rows: {input_rows}  (of {stage1['total_rows']} total)\n")
                f.write(f"Columns: {stage1['headers']}\n")
                f.write("─" * 80 + "\n\n")
                f.write(stage1["text_table"])
            print(f"[SID DEBUG] LLM input text table saved: {text_table_path}")
        except Exception as e:
            print(f"[SID DEBUG] Could not save text table debug file: {e}")
            text_table_path = None

    system_prompt = _EXTRACT_SYSTEM.format(glossary_block=glossary_block)

    # ── Split into chunks ────────────────────────────────────────────
    chunks = _chunk_text_table(stage1["text_table"], chunk_size=chunk_size)
    total_chunks = len(chunks)
    print(f"[SID DEBUG] Split into {total_chunks} chunk(s): "
          + ", ".join(f"{c['first_label']}–{c['last_label']} ({c['row_count']})"
                      for c in chunks))

    # Extract each chunk
    chunk_results: list[dict | None] = []
    total_rows_extracted = 0

    for i, chunk in enumerate(chunks):
        result = await _extract_single_chunk(
            chunk, stage1, llm, system_prompt,
            session_id, llm_config, chunk_idx=i, total_chunks=total_chunks,
        )
        chunk_results.append(result)
        if result is not None:
            total_rows_extracted += _count_extracted_rows(result)

    # Merge results
    merged = _merge_extracted_chunks(
        chunk_results,
        current_cw=stage1["current_cw"],
        sheet_name=stage1["sheet_used"],
        today=stage1["today"],
    )

    n_suppliers = len(merged.get("suppliers", []))
    total_duration = (time.time() - t0) * 1000

    failed_chunks = sum(1 for r in chunk_results if r is None)
    if failed_chunks:
        merged.setdefault("warnings", []).append(
            f"{failed_chunks} of {total_chunks} chunk(s) failed LLM extraction"
        )

    if input_rows > 0 and n_suppliers < input_rows:
        loss_pct = (1 - n_suppliers / input_rows) * 100
        loss_msg = (f"Row count mismatch: extracted {n_suppliers} of "
                    f"{input_rows} input rows ({loss_pct:.0f}% loss)")
        print(f"[SID WARNING] {loss_msg}")
        merged.setdefault("warnings", []).append(loss_msg)

    print(f"[SID DEBUG] Merged: {n_suppliers}/{input_rows} suppliers "
          f"({total_chunks} chunks, {failed_chunks} failed, "
          f"{total_duration:.0f}ms total)")

    # Attach metadata
    merged.setdefault("warnings", [])
    merged["warnings"] = stage1["warnings"] + merged["warnings"]
    merged["_meta"] = {
        "headers": stage1["headers"],
        "original_headers": stage1.get("original_headers", stage1["headers"]),
        "total_rows_in_file": stage1["total_rows"],
        "rows_after_filter": stage1["kept_rows"],
        "rows_extracted_by_llm": n_suppliers,
        "extraction_chunks": total_chunks,
        "extraction_chunks_failed": failed_chunks,
        "sheet_used": stage1["sheet_used"],
        "date_col_name": stage1.get("date_col_name"),
        "csv_path": stage1.get("csv_path"),
        "filtered_csv_path": stage1.get("filtered_csv_path"),
        "text_table_path": text_table_path,
    }

    log_trace(
        session_id, "sid_llm_extract",
        f"Input: {input_rows} rows, {len(stage1['headers'])} columns",
        f"Extracted {n_suppliers} suppliers "
        f"({total_chunks} chunks, {total_duration:.0f}ms)",
        total_duration,
    )
    print(f"[SID DEBUG] ═══ Stage 2 complete ({n_suppliers}/{input_rows} rows, "
          f"{total_chunks} chunks) ═══\n")

    return merged


def _extraction_fallback(stage1: dict, error_msg: str) -> dict:
    """Minimal structure returned when LLM extraction fails entirely."""
    print(f"[SID FALLBACK] {error_msg}")
    return {
        "current_cw": stage1["current_cw"],
        "today": stage1.get("today", ""),
        "sheet_name": stage1["sheet_used"],
        "extraction_notes": f"FALLBACK — {error_msg}",
        "suppliers": [],
        "warnings": stage1["warnings"] + [error_msg],
        "_meta": {
            "headers": stage1["headers"],
            "original_headers": stage1.get("original_headers", stage1["headers"]),
            "total_rows_in_file": stage1["total_rows"],
            "rows_after_filter": stage1["kept_rows"],
            "rows_extracted_by_llm": 0,
            "extraction_chunks": 0,
            "extraction_chunks_failed": 0,
            "sheet_used": stage1["sheet_used"],
            "date_col_name": stage1.get("date_col_name"),
            "csv_path": stage1.get("csv_path"),
            "filtered_csv_path": stage1.get("filtered_csv_path"),
        },
    }


# ─── Combined entry point ───────────────────────────────────────────

async def parse_sid_with_llm(
    filepath: str,
    llm_config: dict,
    session_id: str,
    glossary_context: str = "",
) -> dict:
    """
    Full two-stage SID parsing pipeline.

    Stage 1: pandas-based Excel → filtered text table
    Stage 2: LLM text table → structured JSON
    """
    stage1 = excel_to_text_table(filepath)
    return await llm_extract_sid_data(stage1, llm_config, session_id, glossary_context)
