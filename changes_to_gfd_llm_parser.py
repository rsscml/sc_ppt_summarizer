

# Replace _filter_by_recent_months and add a small helper. This is a drop-in patch — no other call sites change.

# ─── Robust multi-format date parsing ──────────────────────────────

# Explicit patterns for lazy / locale-specific date entry.
# Ordered from most-specific to least-specific so earlier matches
# are not accidentally overridden by later, looser patterns.
_DATE_FORMATS = [
    "%d/%m/%Y", "%d/%m/%y",
    "%d.%m.%Y", "%d.%m.%y",
    "%d-%m-%Y", "%d-%m-%y",
    "%Y-%m-%d", "%Y/%m/%d", "%Y.%m.%d",
    "%d %b %Y", "%d %B %Y",
    "%d-%b-%Y", "%d-%b-%y",
    "%b %d, %Y", "%B %d, %Y",
]


def _parse_date_series(series: pd.Series) -> pd.Series:
    """
    Robust multi-pass date parser designed for messy user-entered xlsx files.

    Pass 1 — pandas auto-detect with dayfirst=True
             Handles native datetime/Timestamp objects (Excel real dates),
             ISO strings, and unambiguous dd/mm/yyyy strings correctly.
    Pass 2 — explicit format sweep on cells still NaT after Pass 1
             Covers dd.mm.yyyy, dd.mm.yy, dd-mm-yyyy, and
             text-month variants that pandas misses.
    Pass 3 — dateutil fallback (dayfirst) on anything still unparsed
             Catches one-offs like "15 Oct 26", "2026/10/15", extra spaces,
             leading apostrophes, etc.

    Cells that remain NaT after all three passes are left as NaT so the
    caller can decide what to do with them.
    """
    # Pass 1 — pandas, with explicit day-first hint
    out = pd.to_datetime(series, errors="coerce", dayfirst=True)

    # Normalise string values ONCE for Pass 2 / Pass 3
    #  - coerce everything to str so .str methods work
    #  - strip whitespace
    #  - strip leading apostrophes that Excel uses for text escape
    as_str = (
        series.astype(str)
        .str.strip()
        .str.lstrip("'")
        .str.replace(r"\s+", " ", regex=True)
    )

    # Pass 2 — sweep explicit formats on remaining NaT cells
    remaining = out.isna() & series.notna()
    for fmt in _DATE_FORMATS:
        if not remaining.any():
            break
        try:
            parsed = pd.to_datetime(
                as_str.where(remaining),
                format=fmt,
                errors="coerce",
            )
        except Exception:
            continue
        hit = parsed.notna() & remaining
        if hit.any():
            out.loc[hit] = parsed[hit]
            remaining = out.isna() & series.notna()

    # Pass 3 — per-cell dateutil fallback for anything still unparsed
    if remaining.any():
        try:
            from dateutil import parser as _du
        except ImportError:
            _du = None
        if _du is not None:
            for idx in series.index[remaining]:
                raw = as_str.loc[idx]
                if not raw or raw.lower() in {"nan", "nat", "none"}:
                    continue
                try:
                    out.loc[idx] = _du.parse(raw, dayfirst=True)
                except (ValueError, TypeError, OverflowError):
                    continue

    return out


def _filter_by_recent_months(
    df: pd.DataFrame,
    date_col: str,
) -> tuple[pd.DataFrame, int, str]:
    """
    Keep only rows whose date_col falls within the current or previous
    calendar month. Rows whose date cannot be parsed after three passes
    are KEPT (benefit of the doubt), matching the behaviour documented
    in the module README.

    Returns (filtered_df, num_removed, window_description).
    """
    dates = _parse_date_series(df[date_col])

    today                 = pd.Timestamp.today().normalize()
    start_current_month   = today.replace(day=1)
    start_previous_month  = start_current_month - pd.offsets.MonthBegin(1)
    end_current_month     = start_current_month + pd.offsets.MonthEnd(1)

    prev_label = start_previous_month.strftime("%b %Y")
    curr_label = start_current_month.strftime("%b %Y")
    desc = f"{prev_label} – {curr_label}"

    in_window      = (dates >= start_previous_month) & (dates < end_current_month)
    is_unparseable = dates.isna()            # survive-at-all-costs rows
    keep_mask      = in_window | is_unparseable

    filtered    = df[keep_mask]
    num_removed = int((~keep_mask).sum())

    # Handy observability: log the parse success rate so you can see
    # at a glance if a new format showed up that still isn't handled.
    total     = len(df)
    unparsed  = int(is_unparseable.sum())
    in_wnd    = int(in_window.sum())
    print(
        f"[GFD DEBUG] Date parse: {total} rows  → "
        f"{total - unparsed} parsed, {unparsed} unparsed (kept), "
        f"{in_wnd} within window, {num_removed} removed."
    )

    return filtered, num_removed, desc
    

# Optional hardening for _find_date_column    

# Be explicit that the primary anchor is "last update";
# "date"/"updated" are only accepted if nothing stronger matches.
_DATE_COL_HINTS = [
    "last update", "last änderung", "last change", "last updated",
    "aktualisiert am", "aktualisiert",
    "update date", "updated on", "updated",
    "datum", "date",
]