"""
GFD LLM Parser
==============
Two-stage pipeline for extracting structured data from Dashboard_Update Excel files.

Stage 1 — Pandas-based deterministic extraction:
  • Open the workbook and find the Dashboard_Update sheet
  • Detect the header row via anchor-based scoring (robust to title rows,
    logos, merged headers, and multi-row header bands)
  • Extract a clean DataFrame with deduplicated column names
  • Forward-fill merged-cell columns (product groups, regions, etc.)
  • Compact individual customer flag columns into a single text column
  • Filter stale rows using the "Last update" date column (preferred) —
    falls back to CW-number scanning if no date column is found
  • Convert the filtered DataFrame to a pipe-delimited text table

Stage 2 — LLM extraction:
  • Send the filtered text table to the LLM
  • LLM understands column semantics, infers product-family groupings,
    and returns a precise JSON object
  • No brittle column-name fuzzy matching; no schema hard-coding

The output JSON is the single source of truth consumed by gfd_llm_slides.py.
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
        max_tokens=max_tokens,
    )


# ═══════════════════════════════════════════════════════════════════════
#  Stage 1: Pandas-based Excel → filtered pipe-delimited text table
# ═══════════════════════════════════════════════════════════════════════

# ─── Pandas extractor (adapted from gfd_extractor_script.py) ─────────

def _normalize_cell(x) -> str:
    """Normalize a single cell value for header detection."""
    if x is None or (isinstance(x, float) and np.isnan(x)):
        return ""
    s = str(x).strip()
    s = re.sub(r"\s+", " ", s)
    return s


def _cellwise_map_dataframe(df: pd.DataFrame, func) -> pd.DataFrame:
    """
    Safe element-wise map that works across pandas versions.
    Tries DataFrame.map (pandas ≥ 2.1), then applymap (classic),
    then numpy fallback.
    """
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
    out = vfunc(vals)
    return pd.DataFrame(out, index=df.index, columns=df.columns)


# Default anchor terms expected in the GFD template header row.
_DEFAULT_ANCHORS = [
    "Customer affected",
    "Region",
    "Plant / Location",
    "Root Cause",
    "Action / Comment",
    "Task Force Leader",
    "Last update",
]


def _extract_dashboard_dataframe(
    xlsx_path: str,
    sheet_name: str = "Dashboard_Update",
    required_anchors: list[str] | None = None,
    min_header_hits: int = 2,
    min_non_empty_in_header: int = 8,
    stop_when_blank_streak: int = 20,
) -> tuple[pd.DataFrame, int]:
    """
    Extract a clean DataFrame from a Dashboard_Update-style Excel sheet.

    Uses anchor-based scoring to detect the header row, handles duplicate
    and blank column names, trims trailing blank rows, and normalises
    whitespace in text cells.

    Returns
    -------
    (df, header_row_idx) where df has clean column names and body data.
    """
    if required_anchors is None:
        required_anchors = _DEFAULT_ANCHORS

    raw = pd.read_excel(xlsx_path, sheet_name=sheet_name, header=None, engine="openpyxl")
    norm = _cellwise_map_dataframe(raw, _normalize_cell)

    def row_score(row_values):
        cells = [c for c in row_values if c]
        if not cells:
            return -1, 0, 0
        anchors_lower = {a.lower() for a in required_anchors}
        hits = sum(c.lower() in anchors_lower for c in cells)
        non_empty = len(cells)
        uniq_ratio = len(set(cells)) / max(1, len(cells))
        numeric_like = sum(bool(re.fullmatch(r"[-+]?\d+(\.\d+)?", c)) for c in cells)
        numeric_penalty = numeric_like / max(1, non_empty)
        score = (hits * 10) + (non_empty * 0.5) + (uniq_ratio * 2) - (numeric_penalty * 5)
        return score, hits, non_empty

    best = (-1e9, None, None, None)
    for idx in range(len(norm)):
        score, hits, non_empty = row_score(list(norm.iloc[idx].values))
        if hits >= min_header_hits and non_empty >= min_non_empty_in_header:
            if score > best[0]:
                best = (score, idx, hits, non_empty)

    if best[1] is None:
        non_empty_counts = norm.apply(lambda r: sum(bool(x) for x in r.values), axis=1)
        header_row_idx = int(non_empty_counts.idxmax())
    else:
        header_row_idx = int(best[1])

    # Build deduplicated headers from detected row
    header = list(norm.iloc[header_row_idx].values)
    cleaned_header: list[str] = []
    seen: dict[str, int] = {}
    for i, h in enumerate(header):
        h = h.strip() if isinstance(h, str) else ""
        if not h:
            h = f"Unnamed_{i}"
        key = h.lower()
        if key in seen:
            seen[key] += 1
            h = f"{h}__{seen[key]}"
        else:
            seen[key] = 0
        cleaned_header.append(h)

    # Extract body below header row
    body = raw.iloc[header_row_idx + 1:].copy()
    body.columns = cleaned_header

    # Drop fully-empty columns and rows
    body = body.dropna(axis=1, how="all")
    body = body.dropna(axis=0, how="all")

    # Trim after a long blank streak
    empty_streak = 0
    keep_idx: list = []
    for i, row in body.iterrows():
        if row.isna().all():
            empty_streak += 1
        else:
            empty_streak = 0
        keep_idx.append(i)
        if empty_streak >= stop_when_blank_streak:
            keep_idx = keep_idx[:-stop_when_blank_streak]
            break
    body = body.loc[keep_idx].copy()

    # Normalise whitespace in string columns
    for col in body.columns:
        if body[col].dtype == object:
            body[col] = body[col].apply(
                lambda x: re.sub(r"\s+", " ", str(x)).strip() if pd.notna(x) else x
            )

    return body, header_row_idx


# ─── Sheet-name detection ────────────────────────────────────────────

def _detect_dashboard_sheet(xlsx_path: str) -> tuple[str, list[str]]:
    """
    Find the best sheet name matching 'Dashboard_Update'.

    Opens the workbook in read-only mode briefly, returns (sheet_name, warnings).
    """
    warnings: list[str] = []
    wb = openpyxl.load_workbook(xlsx_path, read_only=True)
    names = wb.sheetnames
    wb.close()

    # Exact substring match: "dashboard" AND "update"
    for name in names:
        n = name.lower()
        if "dashboard" in n and "update" in n:
            return name, warnings

    # Partial match
    for name in names:
        if "dashboard" in name.lower() or "update" in name.lower():
            return name, warnings

    # Fallback: first sheet
    warnings.append(f"'Dashboard_Update' sheet not found; using '{names[0]}'")
    return names[0], warnings


# ─── Forward-fill merged-cell columns ────────────────────────────────

# Keywords that identify columns likely to originate from vertically
# merged cells (product groups, regions, etc.).
_FFILL_COL_KEYWORDS = [
    "product", "family", "gruppe", "group", "component", "region",
]


def _forward_fill_merged_columns(df: pd.DataFrame) -> None:
    """
    Forward-fill columns whose header name suggests they were merged
    in the original Excel (e.g. product group, region).

    Modifies the DataFrame in place.
    """
    for col in df.columns:
        lowered = col.lower()
        if any(kw in lowered for kw in _FFILL_COL_KEYWORDS):
            df[col] = df[col].ffill()


# ─── Customer column compaction ──────────────────────────────────────

# Tokens that indicate a cell is "not affected" (case-insensitive).
_CUST_NEGATIVE = {"", "n", "no", "0", "-", "·", "false", "nan", "none"}


def _compact_customer_columns(df: pd.DataFrame) -> tuple[pd.DataFrame, bool]:
    """
    Detect and compact individual customer-flag columns into a single
    "Customers affected" text column.

    Detection heuristic: find a contiguous block of ≥ 5 columns where
    > 70 % of non-NaN values are single-character flags (X, Y, N, …) or
    empty strings.  These are the per-customer flag columns present in
    the standard GFD template (typically 32 columns, one per OEM).

    Returns (df, was_compacted).
    """
    flag_col_indices: list[int] = []

    for i, col in enumerate(df.columns):
        vals = df[col].dropna().astype(str).str.strip().str.lower()
        if len(vals) == 0:
            # Entirely empty column — could be part of the customer block
            flag_col_indices.append(i)
            continue
        flag_like = vals.apply(lambda v: len(v) <= 1 or v in _CUST_NEGATIVE)
        if flag_like.mean() >= 0.70:
            flag_col_indices.append(i)

    if len(flag_col_indices) < 5:
        return df, False

    # Find longest contiguous run among the flagged indices
    runs: list[tuple[int, int]] = []
    start = flag_col_indices[0]
    prev = start
    for idx in flag_col_indices[1:]:
        if idx == prev + 1:
            prev = idx
        else:
            runs.append((start, prev))
            start = idx
            prev = idx
    runs.append((start, prev))

    longest = max(runs, key=lambda r: r[1] - r[0] + 1)
    span = longest[1] - longest[0] + 1
    if span < 5:
        return df, False

    start_idx, end_idx = longest
    customer_col_names = list(df.columns[start_idx : end_idx + 1])

    # Build the compacted value per row
    def _compact_row(row):
        affected: list[str] = []
        for cname in customer_col_names:
            val = str(row.get(cname, "")).strip().lower()
            if val and val not in _CUST_NEGATIVE:
                affected.append(cname)
        return ", ".join(affected)

    compacted_vals = df.apply(_compact_row, axis=1)

    # Rebuild DataFrame: columns before | Customers affected | columns after
    cols_before = list(df.columns[:start_idx])
    cols_after = list(df.columns[end_idx + 1 :])

    new_df = df[cols_before].copy()
    new_df["Customers affected"] = compacted_vals.values
    for c in cols_after:
        new_df[c] = df[c].values

    return new_df, True


# ─── Date-based recency filtering ────────────────────────────────────

# Substrings that identify the "Last updated" column (lower-cased).
_DATE_COL_HINTS = [
    "last update", "last änderung", "last change", "aktualisiert",
    "updated", "update date", "datum", "date",
]


def _find_date_column(columns: list[str]) -> str | None:
    """Return the column name matching a 'Last updated' hint, or None."""
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
    Keep only rows whose date_col falls within the current calendar
    month or the previous calendar month.

    Uses a two-pass pd.to_datetime approach:
      Pass 1: pd.to_datetime(errors='coerce') — handles native datetime
              objects, Timestamps, ISO strings, and most common formats.
      Pass 2: for any remaining NaTs, try common European dot-separated
              formats (dd.mm.yyyy, dd.mm.yy) that pd.to_datetime misses.

    Rows where the date is still NaT after both passes are KEPT
    (benefit of the doubt — unparseable or missing).

    Returns (filtered_df, num_removed, window_description).
    """
    # ── Two-pass date parsing ────────────────────────────────────────
    # Pass 1: let pandas auto-detect
    dates = pd.to_datetime(df[date_col], errors="coerce")

    # ── Compute the month window ─────────────────────────────────────
    today = pd.Timestamp.today().normalize()
    start_current_month = today.replace(day=1)
    start_previous_month = start_current_month - pd.offsets.MonthBegin(1)
    end_current_month = start_current_month + pd.offsets.MonthEnd(1)

    prev_label = start_previous_month.strftime("%b %Y")
    curr_label = start_current_month.strftime("%b %Y")
    desc = f"{prev_label} – {curr_label}"

    # ── Apply filter ─────────────────────────────────────────────────
    # Keep rows where date is within window OR date is NaT (unparseable)
    in_window = (dates >= start_previous_month) & (dates < end_current_month)
    keep_mask = in_window

    filtered = df[keep_mask]
    num_removed = int((~keep_mask).sum())

    return filtered, num_removed, desc


# ─── Text table builder ─────────────────────────────────────────────

def _cell_str(value: Any) -> str:
    """Convert a DataFrame cell value to a clean string for the text table."""
    if value is None or (isinstance(value, float) and np.isnan(value)):
        return ""
    if isinstance(value, datetime):
        return value.strftime("%d.%m.%Y")
    if isinstance(value, pd.Timestamp):
        return value.strftime("%d.%m.%Y")
    s = re.sub(r"[\r\n]+", " ", str(value).strip())
    # Strip leading/trailing apostrophes (Excel text-prefix escape)
    if s.startswith("'") or s.endswith("'"):
        s = s.strip("'").strip()
    return s


def _build_text_table(headers: list[str], rows: list[list[str]]) -> str:
    """Render headers + rows as a compact pipe-delimited markdown table.

    Each data row is prefixed with a sequential row number (R001, R002, …)
    so the LLM can individually track and account for every row, preventing
    it from silently merging or skipping similar-looking rows.

    Padding width is capped at 80 characters so a single long outlier cell
    (e.g. a verbose action/comment) doesn't inflate the whitespace of every
    other row in that column.  Cell *content* is never truncated — only the
    ljust padding is capped, so free-text columns are preserved in full.
    """
    MAX_PAD_W = 80
    ROW_NUM_W = 4                          # "R001"

    col_widths = [min(max(len(h), 4), MAX_PAD_W) for h in headers]
    for row in rows:
        for i, cell in enumerate(row[: len(headers)]):
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

    Uses the pandas-based extractor for robust header detection and
    data extraction, then applies month-based date filtering (solely
    on the 'Last update' column) and customer column compaction
    before building the text table.

    Returns
    -------
    {
      text_table          : pipe-delimited table string
      headers             : list of column header strings
      original_headers    : headers before compaction
      current_cw          : "CW{week}/{year}"
      total_rows          : row count before filtering
      kept_rows           : row count after filtering
      sheet_used          : actual sheet name
      date_col_name       : header of the detected date column, or None
      customer_col_range  : None (legacy field, kept for compatibility)
      csv_path            : path to the extracted CSV (before filtering), or None
      filtered_csv_path   : path to the filtered CSV (after filtering), or None
      warnings            : list of warning strings
    }
    """
    year, week = _get_current_cw()
    current_cw = f"CW{week}/{year}"
    print(f"\n[GFD DEBUG] ═══ Stage 1: Excel → text table ═══")
    print(f"[GFD DEBUG] File: {filepath}")
    print(f"[GFD DEBUG] Current CW: {current_cw}")

    # ── Detect sheet name ────────────────────────────────────────────
    sheet_used, warnings = _detect_dashboard_sheet(filepath)
    print(f"[GFD DEBUG] Sheet detected: '{sheet_used}'")

    # ── Extract clean DataFrame using pandas-based extractor ─────────
    try:
        df, header_idx = _extract_dashboard_dataframe(filepath, sheet_name=sheet_used)
    except Exception as exc:
        raise RuntimeError(
            f"Failed to extract DataFrame from sheet '{sheet_used}': {exc}"
        ) from exc

    original_headers = list(df.columns)
    total_rows = len(df)
    print(f"[GFD DEBUG] Extracted: {total_rows} rows × {len(original_headers)} cols "
          f"(header at row {header_idx})")

    if total_rows == 0:
        warnings.append("No data rows found below the detected header row.")

    # ── Forward-fill columns that were merged in Excel ───────────────
    _forward_fill_merged_columns(df)

    # ── Write extracted CSV (full dataset before any filtering) ──────
    #    This is the clean intermediate artifact: messy Excel → tidy CSV.
    #    Downstream filtering and compaction operate on the DataFrame
    #    in memory, but this CSV is saved alongside the uploaded file so
    #    the user can inspect / debug the extraction independently.
    csv_path = str(Path(filepath).with_suffix(".csv"))
    try:
        df.to_csv(csv_path, index=False, encoding="utf-8-sig")
        warnings.append(f"Extracted CSV saved: {csv_path}")
    except Exception as csv_exc:
        warnings.append(f"Could not save extracted CSV: {csv_exc}")
        csv_path = None

    # ── Compact individual customer columns → single text column ─────
    df, was_compacted = _compact_customer_columns(df)
    if was_compacted:
        n_compacted = len(original_headers) - len(df.columns) + 1
        warnings.append(
            f"Compacted {n_compacted} individual customer columns into "
            f"'Customers affected'."
        )
        print(f"[GFD DEBUG] Customer compaction: {n_compacted} cols → 1 "
              f"(now {len(df.columns)} cols)")
    else:
        print(f"[GFD DEBUG] Customer compaction: not triggered")

    # ── Filter rows by recency (current month + previous month) ────────
    date_col_name = _find_date_column(list(df.columns))
    print(f"[GFD DEBUG] Date column: {date_col_name or '(none found)'}")

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
        print(f"[GFD DEBUG] Recency filter: {pre_filter} → {len(df)} rows "
              f"(removed {skipped}, window: {window_desc})")
    else:
        warnings.append(
            "Recency filter: no 'Last update' column found — "
            "all rows retained (no filtering applied)."
        )

    kept_rows = len(df)
    if kept_rows == 0:
        warnings.append(
            "No rows remain after recency filtering — "
            "check that the 'Last update' column has dates in the "
            "current or previous month."
        )

    # ── Save filtered rows to a second CSV ───────────────────────────
    filtered_csv_path = str(Path(filepath).with_name(
        Path(filepath).stem + "_filtered.csv"
    ))
    try:
        df.to_csv(filtered_csv_path, index=False, encoding="utf-8-sig")
        warnings.append(f"Filtered CSV saved: {filtered_csv_path} ({kept_rows} rows)")
    except Exception as fcsv_exc:
        warnings.append(f"Could not save filtered CSV: {fcsv_exc}")
        filtered_csv_path = None

    # ── Build pipe-delimited text table ──────────────────────────────
    headers = list(df.columns)
    rows: list[list[str]] = []
    for _, row in df.iterrows():
        rows.append([_cell_str(v) for v in row.values])

    text_table = _build_text_table(headers, rows)

    text_table_chars = len(text_table)
    text_table_lines = text_table.count("\n") + 1
    est_input_tokens = text_table_chars // 4
    print(f"[GFD DEBUG] Text table: {kept_rows} data rows, {text_table_lines} lines, "
          f"{text_table_chars:,} chars (~{est_input_tokens:,} tokens est.)")
    print(f"[GFD DEBUG] ═══ Stage 1 complete ═══\n")

    return {
        "text_table":         text_table,
        "headers":            headers,
        "original_headers":   original_headers,
        "current_cw":         current_cw,
        "total_rows":         total_rows,
        "kept_rows":          kept_rows,
        "sheet_used":         sheet_used,
        "date_col_name":      date_col_name,
        "customer_col_range": None,    # legacy field kept for Stage 2 compat
        "csv_path":           csv_path,
        "filtered_csv_path":  filtered_csv_path,
        "warnings":           warnings,
    }


# ═══════════════════════════════════════════════════════════════════════
#  Stage 2: LLM extraction  (unchanged)
# ═══════════════════════════════════════════════════════════════════════

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

8. ROW COMPLETENESS — CRITICAL
   — The first column of each data row is a sequential row number (R001, R002, …).
     Your output must contain EXACTLY as many rows as there are numbered data rows
     in the input table.  The user message states the exact count.
   — NEVER merge, deduplicate, or summarise similar-looking rows.  Even if two rows
     have identical product group, plant, and root cause, they represent SEPARATE
     risk items and must each appear individually in your output.
   — After generating your JSON, mentally verify: does the number of rows across
     all product_groups sum to the exact count stated in the user message?
     If not, you have missed rows — go back and include them.

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

The table below contains EXACTLY {kept_rows} data rows (R001–R{kept_rows:03d}), filtered from {total_rows} total.
Your JSON output MUST contain exactly {kept_rows} rows across all product_groups combined — no more, no fewer.
Do NOT merge, skip, or summarise any rows even if they look similar.

{text_table}"""


def _count_extracted_rows(extracted: dict) -> int:
    """Count total data rows across all product_groups in the LLM output."""
    return sum(
        len(pg.get("rows", []))
        for pg in extracted.get("product_groups", [])
    )


async def llm_extract_gfd_data(
    stage1: dict,
    llm_config: dict,
    session_id: str,
    glossary_context: str = "",
    max_retries: int = 2,
) -> dict:
    """
    Stage 2 (LLM): Send the filtered text table to the LLM and extract structured JSON.

    Validates that the number of output rows matches the input row count.
    If the LLM drops more than 10 % of rows, retries with an explicit reminder.

    Returns the parsed extraction dict. On JSON parse failure, returns a minimal
    fallback structure with the error noted in warnings.
    """
    input_rows = stage1["kept_rows"]
    llm = _create_llm(llm_config, max_tokens=64000)
    t0 = time.time()

    print(f"[GFD DEBUG] ═══ Stage 2: LLM extraction ═══")
    print(f"[GFD DEBUG] Input: {input_rows} data rows, "
          f"{len(stage1['headers'])} columns, max_tokens=64000")

    glossary_block = (
        f"\n\nCOMPANY GLOSSARY — use these to expand abbreviations correctly:\n{glossary_context}"
        if glossary_context else ""
    )

    # ── Save the text table to disk for debugging ────────────────────
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
            print(f"[GFD DEBUG] LLM input text table saved: {text_table_path}")
        except Exception as e:
            print(f"[GFD DEBUG] Could not save text table debug file: {e}")
            text_table_path = None

    base_system = _EXTRACT_SYSTEM.format(glossary_block=glossary_block)
    base_user = _EXTRACT_USER.format(
        current_cw=stage1["current_cw"],
        sheet_name=stage1["sheet_used"],
        kept_rows=stage1["kept_rows"],
        total_rows=stage1["total_rows"],
        text_table=stage1["text_table"],
    )

    attempt = 0
    extracted: dict = {}

    while attempt <= max_retries:
        attempt += 1
        attempt_t0 = time.time()

        # On retry, append an explicit list of what's missing
        if attempt == 1:
            user_msg = base_user
        else:
            prev_count = _count_extracted_rows(extracted)
            user_msg = (
                base_user
                + f"\n\n⚠ CRITICAL: Your previous response contained only "
                  f"{prev_count} rows, but the input has {input_rows} data "
                  f"rows (R001–R{input_rows:03d}).  You have LOST "
                  f"{input_rows - prev_count} rows.  Go through the table "
                  f"row by row and include every single R-numbered row."
            )
            print(f"[GFD DEBUG] Retry {attempt}: previous attempt had "
                  f"{prev_count}/{input_rows} rows")

        messages = [
            SystemMessage(content=base_system),
            HumanMessage(content=user_msg),
        ]

        try:
            response = await llm.ainvoke(messages)
            raw = response.content.strip()

            # Strip markdown fences if the model disobeyed instructions
            if raw.startswith("```"):
                raw = "\n".join(raw.split("\n")[1:])
            if raw.endswith("```"):
                raw = "\n".join(raw.split("\n")[:-1])

            extracted = json.loads(raw.strip())

            usage = response.response_metadata.get("token_usage", {})
            completion_tokens = usage.get("completion_tokens", 0)
            log_tokens(session_id, f"gfd_llm_extract_attempt{attempt}", usage,
                       llm_config.get("azure_deployment", ""))

            n_pgs = len(extracted.get("product_groups", []))
            n_rows_out = _count_extracted_rows(extracted)
            duration = (time.time() - attempt_t0) * 1000

            print(f"[GFD DEBUG] Attempt {attempt}: LLM returned {n_rows_out}/{input_rows} rows "
                  f"across {n_pgs} product groups  "
                  f"(completion_tokens={completion_tokens}, duration={duration:.0f}ms)")

            log_trace(
                session_id, f"gfd_llm_extract_attempt{attempt}",
                f"Input: {input_rows} rows, {len(stage1['headers'])} columns",
                f"Extracted {n_pgs} product groups, {n_rows_out} rows "
                f"(completion_tokens={completion_tokens})",
                duration,
            )

            # ── Row-count validation ─────────────────────────────────
            if input_rows > 0 and n_rows_out < input_rows:
                loss_pct = (1 - n_rows_out / input_rows) * 100
                loss_msg = (
                    f"Row count mismatch: LLM returned {n_rows_out} of "
                    f"{input_rows} input rows ({loss_pct:.0f}% loss, "
                    f"attempt {attempt})"
                )
                print(f"[GFD WARNING] {loss_msg}")

                # Tolerate small loss (≤10%) — the LLM may legitimately
                # skip fully-empty rows; otherwise retry
                if loss_pct > 10 and attempt <= max_retries:
                    print(f"[GFD DEBUG] Loss exceeds 10% — will retry")
                    continue

                extracted.setdefault("warnings", [])
                extracted["warnings"].append(loss_msg)

            # ── Success — attach Stage 1 metadata ────────────────────
            extracted.setdefault("warnings", [])
            extracted["warnings"] = stage1["warnings"] + extracted["warnings"]
            extracted["_meta"] = {
                "headers": stage1["headers"],
                "original_headers": stage1.get("original_headers", stage1["headers"]),
                "total_rows_in_file": stage1["total_rows"],
                "rows_after_filter": stage1["kept_rows"],
                "rows_extracted_by_llm": n_rows_out,
                "extraction_attempts": attempt,
                "completion_tokens_used": completion_tokens,
                "sheet_used": stage1["sheet_used"],
                "date_col_name": stage1.get("date_col_name"),
                "customer_col_range": stage1.get("customer_col_range"),
                "csv_path": stage1.get("csv_path"),
                "filtered_csv_path": stage1.get("filtered_csv_path"),
                "text_table_path": text_table_path,
            }

            total_duration = (time.time() - t0) * 1000
            log_trace(
                session_id, "gfd_llm_extract",
                f"Input: {input_rows} rows, {len(stage1['headers'])} columns",
                f"Final: {n_pgs} product groups, {n_rows_out} rows "
                f"(attempts={attempt}, total={total_duration:.0f}ms)",
                total_duration,
            )
            print(f"[GFD DEBUG] ═══ Stage 2 complete ({n_rows_out}/{input_rows} rows, "
                  f"{attempt} attempt(s)) ═══\n")
            return extracted

        except json.JSONDecodeError as exc:
            duration = (time.time() - attempt_t0) * 1000
            print(f"[GFD ERROR] Attempt {attempt}: JSON parse error — {str(exc)[:120]}")
            log_trace(session_id, f"gfd_llm_extract_attempt{attempt}",
                      f"Input: {input_rows} rows",
                      f"JSON PARSE ERROR: {str(exc)[:120]}", duration, {"error": True})
            if attempt <= max_retries:
                continue
            return _extraction_fallback(stage1, f"LLM returned unparseable JSON: {str(exc)[:200]}")

        except Exception as exc:
            duration = (time.time() - attempt_t0) * 1000
            print(f"[GFD ERROR] Attempt {attempt}: {str(exc)[:120]}")
            log_trace(session_id, f"gfd_llm_extract_attempt{attempt}",
                      f"Input: {input_rows} rows",
                      f"ERROR: {str(exc)[:120]}", duration, {"error": True})
            if attempt <= max_retries:
                continue
            return _extraction_fallback(stage1, f"LLM extraction failed: {str(exc)[:200]}")

    # Safety net — should not reach here
    return _extraction_fallback(stage1, f"All {max_retries + 1} extraction attempts failed")


def _extraction_fallback(stage1: dict, error_msg: str) -> dict:
    """Minimal structure returned when LLM extraction fails entirely."""
    print(f"[GFD FALLBACK] {error_msg}")
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
            "rows_extracted_by_llm": 0,
            "extraction_attempts": 0,
            "completion_tokens_used": 0,
            "sheet_used": stage1["sheet_used"],
            "date_col_name": stage1.get("date_col_name"),
            "customer_col_range": stage1.get("customer_col_range"),
            "csv_path": stage1.get("csv_path"),
            "filtered_csv_path": stage1.get("filtered_csv_path"),
        },
    }


# ─── Combined entry point ─────────────────────────────────────────────

async def parse_gfd_with_llm(
    filepath: str,
    llm_config: dict,
    session_id: str,
    history_weeks: int = 4,       # accepted for API compat; not used
    glossary_context: str = "",
) -> dict:
    """
    Full two-stage GFD parsing pipeline.

    Stage 1: pandas-based Excel → filtered text table (current + previous month)
    Stage 2: LLM text table → structured JSON

    Returns the LLM-extracted JSON dict (see _EXTRACT_SYSTEM for full schema).
    """
    stage1 = excel_to_text_table(filepath)
    return await llm_extract_gfd_data(stage1, llm_config, session_id, glossary_context)
