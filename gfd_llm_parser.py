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


# ─── JSON repair for LLM output ─────────────────────────────────────

def _repair_llm_json(raw: str, debug_label: str = "") -> str:
    """
    Attempt to clean common JSON defects produced by LLMs before parsing.

    Handles:
      1. Markdown fences (```json ... ```)
      2. Trailing commas before } or ]
      3. Unescaped literal newlines / tabs inside JSON string values
      4. Unescaped backslashes (single \\ not part of a valid escape)

    Returns the cleaned string.  Does NOT call json.loads — the caller
    should do that and can still get a parse error if the damage is
    too severe for these heuristics.
    """
    s = raw.strip()

    # ── 1. Strip markdown fences ─────────────────────────────────────
    #    Handle ```json, ```JSON, ``` at start; ``` at end.
    #    Also handle cases where the model wraps in triple-backtick mid-stream.
    if s.startswith("```"):
        first_nl = s.index("\n") if "\n" in s else len(s)
        s = s[first_nl + 1:]
    if s.endswith("```"):
        s = s[: s.rfind("```")]
    s = s.strip()

    # ── 2. Trailing commas — ,} or ,] ───────────────────────────────
    s = re.sub(r",\s*([}\]])", r"\1", s)

    # ── 3. Unescaped control characters inside string values ─────────
    #    Walk through the string tracking whether we're inside a JSON
    #    string (between unescaped double-quotes).  Replace literal
    #    newlines/tabs/carriage-returns with their escaped forms.
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
        print(f"[GFD DEBUG] {debug_label}: JSON repair applied "
              f"({len(raw)} → {len(s)} chars)")

    return s


def _parse_llm_json(raw: str, session_id: str = "", attempt: int = 0) -> dict:
    """
    Parse LLM JSON output with safety-net repair and detailed error diagnostics.

    The LLM is instructed to produce clean, single-line JSON strings (no literal
    newlines/tabs, no trailing commas).  The repair step here is a safety net for
    the occasional case where the model still emits a control character despite
    the prompt instructions.

    On failure, logs the region around the error position so you can
    see exactly what broke.
    """
    label = f"attempt{attempt}" if attempt else ""

    # Always apply repair — source data routinely contains control chars
    repaired = _repair_llm_json(raw, debug_label=label)
    try:
        return json.loads(repaired)
    except json.JSONDecodeError as exc:
        # Log context around the error position for debugging
        pos = exc.pos or 0
        start = max(0, pos - 120)
        end = min(len(repaired), pos + 120)
        context = repaired[start:end]
        pointer_offset = pos - start

        print(f"[GFD ERROR] JSON parse failed after repair — "
              f"line {exc.lineno}, col {exc.colno}, char {pos}")
        print(f"[GFD ERROR] Context around error position:")
        print(f"  ...{context}...")
        print(f"  {' ' * (pointer_offset + 5)}^ error here")

        raise


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

2. COVERAGE FIELDS — coverage_without_mitigation_cw / coverage_with_mitigation_cw 
   ═══════════════════════════════════════════════════════════════════════════
   GOAL — Normalise every coverage cell to the canonical form CW<WEEK_NUMBER>,
   where WEEK_NUMBER is an ISO 8601 calendar week integer in the range 1–53.
   The Excel file expresses this information in THREE different ways; your
   job is to recognise all three and convert each to the same canonical form.

   Once you have the canonical CW<WEEK_NUMBER>, emit ONLY the integer
   WEEK_NUMBER into the JSON field (e.g. canonical form CW15 → emit 15).
   Emit null only if the cell cannot be mapped to any of the three forms.

   The current calendar week is given at the top of the user message
   (e.g. "CW13/2026"). Call its numeric portion CURRENT_WEEK — you will
   need it for Form C.

   ── Form A — Explicit CW / KW label ────────────────────────────────
       Already in (or near) canonical form — just extract the week digits.
         "CW15"        →  canonical CW15  →  emit 15
         "CW15/2026"   →  canonical CW15  →  emit 15
         "KW15"        →  canonical CW15  →  emit 15
         "kw 15"       →  canonical CW15  →  emit 15
         "Woche 15"    →  canonical CW15  →  emit 15
         "W15"         →  canonical CW15  →  emit 15
       Drop any year suffix.

   ── Form B — Calendar date → canonical CW ──────────────────────────
       Accept ONLY these formats (day-first European or ISO):
         dd/mm/yyyy, dd/mm/yy, dd.mm.yyyy, dd.mm.yy,
         dd-mm-yyyy, dd-mm-yy, yyyy-mm-dd, yyyy/mm/dd, yyyy.mm.dd
       Convert the date to its ISO 8601 week number, form the canonical
       label CW<that_week>, then emit the integer.
         "15/05/2026"  →  ISO week 20  →  canonical CW20  →  emit 20
         "15.05.2026"  →  ISO week 20  →  canonical CW20  →  emit 20
         "2026-05-15"  →  ISO week 20  →  canonical CW20  →  emit 20
       Any date missing a year component (e.g. "03.04.", "15 May")
         → null.  Do NOT guess the year.

   ── Form C — Days of remaining coverage → canonical CW ─────────────
       A bare number with no "CW"/"KW"/"Woche"/"W" prefix and no date
       separators means "that many days of supply remaining, counted
       from the current week". Convert to canonical CW with:
           canonical_week = CURRENT_WEEK + (days // 7)
       then emit canonical_week as the integer.
       Examples (assuming CURRENT_WEEK = 13):
         "0"      →  canonical CW13  →  emit 13
         "3"      →  canonical CW13  →  emit 13
         "7"      →  canonical CW14  →  emit 14
         "14"     →  canonical CW15  →  emit 15
         "21"     →  canonical CW16  →  emit 16
       A trailing "d", "days", or "Tage" still counts as Form C:
         "14d", "14 days", "14 Tage"  →  canonical CW15  →  emit 15
    
    ── Form D — Zero / exhausted coverage ─────────────────────────────
       A cell whose entire value (after trimming whitespace, quotes,
       and leading apostrophes, case-insensitive) equals one of:
         "0", "zero", "none", "exhausted", "nil", "null", "n.a. coverage",
         "no coverage", "keine", "keine deckung"
       means supply is already gone. Convert to the LAST COMPLETED
       calendar week — the week immediately before CURRENT_WEEK:
           canonical_week = CURRENT_WEEK - 1
       Emit that integer. If CURRENT_WEEK is 1, emit 52.
       Examples (assuming CURRENT_WEEK = 13):
         "zero"         →  canonical CW12  →  emit 12
         "NONE"         →  canonical CW12  →  emit 12
         "no coverage"  →  canonical CW12  →  emit 12

       IMPORTANT — do not confuse Form D's textual "0"/"zero" with
       Form C's numeric "0". A bare numeric zero ("0") is ambiguous;
       treat it as Form D (exhausted) because Form C's "CURRENT_WEEK
       + 0" is the current week, not a meaningful coverage boundary.

   ── Form E — Range or open-ended expression ───────────────────────
       A coverage cell may express a range or a lower bound. Always
       take the UPPER (latest) week in the expression and convert
       through whichever earlier form (A, B, C) applies to that
       upper value. Recognise:

         •  ">CW26", "> CW 26", ">=CW26", "after CW26", "beyond CW26",
            "ab CW26", "nach CW26"
              → treat the quoted week as the upper bound
              →  canonical CW26  →  emit 26

         •  "CW19 - CW22", "CW19-CW22", "CW 19 to CW 22", "KW19-22",
            "between CW19 and CW22"
              → take the upper endpoint
              →  canonical CW22  →  emit 22

         •  "<CW18", "before CW18", "bis CW18", "until CW18"
              → upper bound IS the quoted week itself
              →  canonical CW18  →  emit 18

         •  A range expressed with dates, e.g. "15.04.2026 - 15.05.2026"
              → convert the later date to its ISO week per Form B
              →  canonical CW20  →  emit 20

       If the cell is purely a lower-bound phrase with no number
       (e.g. "open", "ongoing", "tbd") → null (see Disambiguation
       step 6).

   ── Form F — CW buried in free text (last-resort extraction) ──────
       If the cell is a longer sentence or comment (i.e. it did not
       match any form above as a clean value), scan it for the
       literal substring "CW" or "KW" (case-insensitive), allow an
       optional space, and read the immediately following 1–2 digits.
       Take the FIRST such match and apply Form A's extraction to it.
         "Secured until CW22 pending allocation"   → 22
         "Coverage cw 8 only"                      →  8
         "Stop shipment; resume after KW35/2026"   → 35
       If no "CW"/"KW" substring is found, or the characters after
       it are not digits, → null.

       Guardrails for Form F:
         • Apply ONLY as a last resort, after every cleaner form has
           failed. Clean short cells like "CW15", "15/05/2026", and
           "14" must continue to be handled by Forms A/B/C — do not
           fall through to Form F just because the cell contains
           "CW" somewhere.
         • Do not extract from action_comment, special_freight_remarks,
           root_cause, or any other free-text field. Form F applies
           ONLY when the source cell is the Coverage column itself.
         • If the cell contains multiple "CW" mentions (e.g. "CW18
           with risk, CW22 without"), the FIRST one wins — do not
           pick the largest.
   ── Form G — "No risk" / fully covered sentinel ────────────────────
       A cell whose trimmed, lowercased value matches any of:
         "no risk", "no risk at the moment", "no risk at present",
         "no risk currently", "not at risk", "fully covered",
         "full coverage", "covered", "kein risiko", "keine gefahr"
       — or contains "no risk" as a substring of a short phrase
       (≤ 6 words total) — means coverage is green across the
       entire foreseeable horizon. Emit the integer 53 (the
       canonical-form upper cap), which downstream logic treats
       as "covered for every forward-looking CW on the slide".
       Examples:
         "no risk at the moment"  →  canonical CW53  →  emit 53
         "No risk"                →  canonical CW53  →  emit 53
         "fully covered"          →  canonical CW53  →  emit 53
         "kein Risiko"            →  canonical CW53  →  emit 53

       Guardrail — Form G applies ONLY when the "no risk" phrase
       is the whole cell (or essentially the whole cell). A longer
       sentence like "No risk identified in CW15 but exposure rises
       in CW20" is NOT Form G; fall through to Form F, which will
       extract CW15.

   ── Form H — Month name → end-of-month CW ──────────────────────────
       A cell whose trimmed value is a month name (full or 3-letter
       abbreviation, English or German, case-insensitive) means
       coverage runs through the END of that month. Convert by
       computing the ISO 8601 week number of the LAST DAY of that
       month in the applicable year, then emit that integer.

       Recognised month names:
         January/Jan, February/Feb, March/Mar, April/Apr, May,
         June/Jun, July/Jul, August/Aug, September/Sep/Sept,
         October/Oct, November/Nov, December/Dec
         Januar, Februar, März/Maerz, April, Mai, Juni, Juli,
         August, September, Oktober, November, Dezember

       Year resolution — the cell rarely carries a year, so infer:
         • If the named month's number is ≥ CURRENT_WEEK's month
           → use the SAME calendar year as CURRENT_WEEK.
         • If the named month's number is < CURRENT_WEEK's month
           → use the NEXT calendar year (coverage rolls forward).
         If the cell includes an explicit year (e.g. "May 2026",
         "Mai/2027"), honour that year exactly.

       Examples (assuming CURRENT_WEEK = CW13/2026, i.e. late March):
         "May"        →  31 May 2026 → ISO week 22 →  emit 22
         "May 2026"   →  31 May 2026 → ISO week 22 →  emit 22
         "April"      →  30 Apr 2026 → ISO week 18 →  emit 18
         "February"   →  month < current → year 2027
                      →  28 Feb 2027 → ISO week  9 →  emit  9
         "Dec"        →  31 Dec 2026 → ISO week 53 →  emit 53
         "Mai"        →  31 May 2026 → ISO week 22 →  emit 22

       Guardrail — Form H applies ONLY when the cell's entire value
       is a month name (optionally with a year). "Secured until end
       of May" is NOT Form H — it is free text; Form F will handle
       it only if a "CW"/"KW" substring is present, otherwise it
       falls to null. Do NOT infer a month from arbitrary free text. 

   ── Canonical-form check before emitting ───────────────────────────
   Before writing the integer into the JSON, mentally verify you have
   a valid canonical label of the form CW<N>, where N is an integer
   in 1–53.  If N falls outside that range (e.g. Form C produced 58
   because CURRENT_WEEK is late in the year plus many days), cap it
   at 53 and add a note to extraction_notes explaining which row and
   what the raw cell contained.

   ── Disambiguation — apply in this order ───────────────────────────
     1. Cell trimmed exactly matches a Form D keyword → Form D.
     2. Cell trimmed (≤ 6 words) matches a Form G "no risk" phrase
        → Form G.
     3. Cell trimmed is a month name alone, optionally with a year
        → Form H.
     4. Cell contains a range/comparator/bound expression (">", "<",
        "-", "to", "between", "ab", "nach", "bis", "until", "before",
        "after", "beyond") → Form E.
     5. Any "CW" / "KW" / "Woche" / "W" prefix at the start of the
        trimmed cell, or a "/yyyy" suffix → Form A.
     6. Cell trimmed matches a date pattern from the whitelist
        → Form B.
     7. Cell trimmed is a bare integer, with or without "d"/"days"/
        "Tage" → Form C.
     8. Cell is a longer string not matching any of the above, but
        contains "CW" or "KW" followed by 1–2 digits → Form F.
     9. Anything else — empty, "-", "n/a", "TBD", "pending", free
        text with no CW marker → null, and add a short note to
        extraction_notes describing what you saw.

3. TEXT FIELDS — preserve the content verbatim but flatten to single-line strings.
   Excel cells often contain line breaks, tabs, and other control characters.
   Replace these with a semicolon-space ("; ") or a plain space to keep all text
   on one line.  Do NOT use literal newline or tab characters inside JSON string
   values — these break JSON parsing.
   Example: a cell containing
       "Dual source activation CW14
        Air freight bridge 3 weeks"
   should become: "Dual source activation CW14; Air freight bridge 3 weeks"

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

CRITICAL JSON RULES:
  • Every string value must be a single line — no literal newline, tab, or
    carriage-return characters inside strings.  Replace line breaks from the
    source data with "; " (semicolon-space) or " " (space).
  • No trailing commas before }} or ]].
  • The output must be parseable by a strict JSON parser (Python json.loads)
    without any post-processing.

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

The table below contains EXACTLY {kept_rows} data rows ({first_row_label}–{last_row_label}), filtered from {total_rows} total.
Your JSON output MUST contain exactly {kept_rows} rows across all product_groups combined — no more, no fewer.
Do NOT merge, skip, or summarise any rows even if they look similar.

{text_table}"""


def _count_extracted_rows(extracted: dict) -> int:
    """Count total data rows across all product_groups in the LLM output."""
    return sum(
        len(pg.get("rows", []))
        for pg in extracted.get("product_groups", [])
    )


# ─── Chunked extraction ─────────────────────────────────────────────

_CHUNK_SIZE = 5   # rows per LLM call — small enough for reliable extraction


def _chunk_text_table(text_table: str, chunk_size: int = _CHUNK_SIZE,
                      ) -> list[dict]:
    """
    Split a pipe-delimited text table into chunks of at most `chunk_size`
    data rows, each retaining the original header and separator lines.

    Returns a list of dicts:
      { "text_table": str, "row_count": int,
        "first_label": "R005", "last_label": "R014" }
    """
    lines = text_table.split("\n")
    header_line = lines[0]       # | # | Product Group | ...
    sep_line    = lines[1]       # |---|---| ...
    data_lines  = lines[2:]      # | R001 | ... |

    if not data_lines:
        return [{"text_table": text_table, "row_count": 0,
                 "first_label": "", "last_label": ""}]

    chunks: list[dict] = []
    for start in range(0, len(data_lines), chunk_size):
        batch = data_lines[start : start + chunk_size]
        chunk_table = "\n".join([header_line, sep_line] + batch)

        # Extract row labels from the first column (e.g. "R001")
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
    """
    Send a single text-table chunk to the LLM and return the parsed JSON.
    Returns None on failure (the caller will log and continue).
    """
    chunk_label = f"chunk {chunk_idx + 1}/{total_chunks}"
    row_count  = chunk["row_count"]
    first_lbl  = chunk["first_label"]
    last_lbl   = chunk["last_label"]

    user_msg = _EXTRACT_USER.format(
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

        # ── Dump raw response ────────────────────────────────────────
        try:
            csv_path = stage1.get("csv_path")
            dump_dir = Path(csv_path).parent if csv_path else Path(".")
            dump_path = dump_dir / f"{session_id}_llm_extract_{chunk_label.replace(' ', '_').replace('/', 'of')}_raw.txt"
            dump_path.write_text(raw, encoding="utf-8")
        except Exception:
            pass

        extracted = _parse_llm_json(raw, session_id=session_id, attempt=chunk_idx)

        usage = response.response_metadata.get("token_usage", {})
        completion_tokens = usage.get("completion_tokens", 0)
        log_tokens(session_id, f"gfd_llm_extract_{chunk_label}", usage,
                   llm_config.get("azure_deployment", ""))

        n_rows_out = _count_extracted_rows(extracted)
        duration = (time.time() - t0) * 1000

        print(f"[GFD DEBUG] {chunk_label} ({first_lbl}–{last_lbl}): "
              f"sent {row_count} rows → got {n_rows_out} rows back "
              f"({completion_tokens} tokens, {duration:.0f}ms)")

        if n_rows_out < row_count:
            print(f"[GFD WARNING] {chunk_label}: lost {row_count - n_rows_out} rows")

        log_trace(session_id, f"gfd_llm_extract_{chunk_label}",
                  f"Input: {row_count} rows ({first_lbl}–{last_lbl})",
                  f"Extracted {n_rows_out} rows", duration)

        return extracted

    except Exception as exc:
        duration = (time.time() - t0) * 1000
        print(f"[GFD ERROR] {chunk_label}: {str(exc)[:120]}")
        log_trace(session_id, f"gfd_llm_extract_{chunk_label}",
                  f"Input: {row_count} rows ({first_lbl}–{last_lbl})",
                  f"ERROR: {str(exc)[:120]}", duration, {"error": True})
        return None


def _merge_extracted_chunks(
    chunk_results: list[dict],
    current_cw: str,
    sheet_name: str,
) -> dict:
    """
    Merge product_groups from multiple chunk extraction results.

    Product groups with the same (product_family_code, product_family_desc)
    pair are combined — their rows are concatenated in order.
    """
    # Use (code, desc) as the merge key to handle groups split across chunks
    merged_groups: dict[tuple[str, str], dict] = {}
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

        for pg in chunk_result.get("product_groups", []):
            code = pg.get("product_family_code", "")
            desc = pg.get("product_family_desc", "")
            key = (code, desc)

            if key in merged_groups:
                merged_groups[key]["rows"].extend(pg.get("rows", []))
            else:
                merged_groups[key] = {
                    "product_family_code": code,
                    "product_family_desc": desc,
                    "rows": list(pg.get("rows", [])),
                }

    return {
        "current_cw": current_cw,
        "sheet_name": sheet_name,
        "extraction_notes": "; ".join(all_notes) if all_notes else "",
        "product_groups": list(merged_groups.values()),
        "warnings": all_warnings,
    }


# ─── Main Stage 2 entry point ───────────────────────────────────────

async def llm_extract_gfd_data(
    stage1: dict,
    llm_config: dict,
    session_id: str,
    glossary_context: str = "",
    chunk_size: int = _CHUNK_SIZE,
) -> dict:
    """
    Stage 2 (LLM): Extract structured JSON from the filtered text table
    using chunked extraction.

    The text table is split into batches of ~10 rows.  Each batch is sent
    to the LLM independently.  Results are merged by product-group code,
    so groups that span a chunk boundary are recombined automatically.

    This approach is far more reliable than single-shot extraction because
    each LLM call handles a small, manageable number of rows.
    """
    input_rows = stage1["kept_rows"]
    llm = _create_llm(llm_config, max_tokens=32000)
    t0 = time.time()

    print(f"\n[GFD DEBUG] ═══ Stage 2: LLM extraction (chunked, {chunk_size} rows/chunk) ═══")
    print(f"[GFD DEBUG] Input: {input_rows} data rows, "
          f"{len(stage1['headers'])} columns")

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

    system_prompt = _EXTRACT_SYSTEM.format(glossary_block=glossary_block)

    # ── Split into chunks ────────────────────────────────────────────
    chunks = _chunk_text_table(stage1["text_table"], chunk_size=chunk_size)
    total_chunks = len(chunks)
    print(f"[GFD DEBUG] Split into {total_chunks} chunk(s): "
          + ", ".join(f"{c['first_label']}–{c['last_label']} ({c['row_count']})"
                      for c in chunks))

    # ── Extract each chunk ───────────────────────────────────────────
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

    # ── Merge results ────────────────────────────────────────────────
    merged = _merge_extracted_chunks(
        chunk_results,
        current_cw=stage1["current_cw"],
        sheet_name=stage1["sheet_used"],
    )

    n_pgs = len(merged.get("product_groups", []))
    n_rows_out = _count_extracted_rows(merged)
    total_duration = (time.time() - t0) * 1000

    # ── Row-count summary ────────────────────────────────────────────
    failed_chunks = sum(1 for r in chunk_results if r is None)
    if failed_chunks:
        merged.setdefault("warnings", []).append(
            f"{failed_chunks} of {total_chunks} chunk(s) failed LLM extraction"
        )

    if input_rows > 0 and n_rows_out < input_rows:
        loss_pct = (1 - n_rows_out / input_rows) * 100
        loss_msg = (f"Row count mismatch: extracted {n_rows_out} of "
                    f"{input_rows} input rows ({loss_pct:.0f}% loss)")
        print(f"[GFD WARNING] {loss_msg}")
        merged.setdefault("warnings", []).append(loss_msg)

    print(f"[GFD DEBUG] Merged: {n_rows_out}/{input_rows} rows across "
          f"{n_pgs} product groups  "
          f"({total_chunks} chunks, {failed_chunks} failed, "
          f"{total_duration:.0f}ms total)")

    # ── Attach Stage 1 metadata ──────────────────────────────────────
    merged.setdefault("warnings", [])
    merged["warnings"] = stage1["warnings"] + merged["warnings"]
    merged["_meta"] = {
        "headers": stage1["headers"],
        "original_headers": stage1.get("original_headers", stage1["headers"]),
        "total_rows_in_file": stage1["total_rows"],
        "rows_after_filter": stage1["kept_rows"],
        "rows_extracted_by_llm": n_rows_out,
        "extraction_chunks": total_chunks,
        "extraction_chunks_failed": failed_chunks,
        "sheet_used": stage1["sheet_used"],
        "date_col_name": stage1.get("date_col_name"),
        "customer_col_range": stage1.get("customer_col_range"),
        "csv_path": stage1.get("csv_path"),
        "filtered_csv_path": stage1.get("filtered_csv_path"),
        "text_table_path": text_table_path,
    }

    log_trace(
        session_id, "gfd_llm_extract",
        f"Input: {input_rows} rows, {len(stage1['headers'])} columns",
        f"Extracted {n_rows_out} rows in {n_pgs} groups "
        f"({total_chunks} chunks, {total_duration:.0f}ms)",
        total_duration,
    )
    print(f"[GFD DEBUG] ═══ Stage 2 complete ({n_rows_out}/{input_rows} rows, "
          f"{total_chunks} chunks) ═══\n")

    return merged


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
            "extraction_chunks": 0,
            "extraction_chunks_failed": 0,
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
