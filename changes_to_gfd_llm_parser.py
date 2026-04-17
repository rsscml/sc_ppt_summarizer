

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

################# Option 2 #################

# Fast regex precheck for obvious year-less patterns: "dd.mm", "dd/mm",
# "dd.mm.", "d-m", etc. Two numeric groups with no trailing year.
_YEARLESS_RE = re.compile(
    r"""^\s*'?                       # optional leading apostrophe
        \d{1,2}                      # day
        [.\-/\s]                     # separator
        \d{1,2}                      # month
        [.\-/\s]?                    # optional trailing separator
        \s*$                         # end — no year component
    """,
    re.VERBOSE,
)


def _parse_date_series(series: pd.Series) -> tuple[pd.Series, pd.Series]:
    """
    Robust multi-pass date parser for messy user-entered xlsx files.

    Returns
    -------
    dates : pd.Series of datetime64 — NaT where the value is truly
            unparseable garbage (should be kept on benefit-of-the-doubt).
    yearless : pd.Series of bool — True where the input was date-like but
            the year was missing and had to be inferred (e.g., "03.04.").
            The caller should DROP these rows, since the real year could
            be years in the past.
    """
    out      = pd.to_datetime(series, errors="coerce", dayfirst=True)
    yearless = pd.Series(False, index=series.index)

    # Normalise strings once
    as_str = (
        series.astype(str)
        .str.strip()
        .str.lstrip("'")
        .str.replace(r"\s+", " ", regex=True)
    )

    # ── Pre-screen: obvious year-less patterns ──────────────────────
    # Do this BEFORE Pass 1's output is trusted, because pandas will
    # sometimes silently accept "3/4" as current-year March 4th.
    yl_regex = as_str.str.match(_YEARLESS_RE, na=False) & series.notna()
    if yl_regex.any():
        yearless.loc[yl_regex] = True
        out.loc[yl_regex] = pd.NaT     # invalidate any Pass-1 false positive

    # ── Pass 2: explicit formats (all include a year component) ─────
    remaining = out.isna() & series.notna() & ~yearless
    for fmt in _DATE_FORMATS:
        if not remaining.any():
            break
        try:
            parsed = pd.to_datetime(
                as_str.where(remaining), format=fmt, errors="coerce"
            )
        except Exception:
            continue
        hit = parsed.notna() & remaining
        if hit.any():
            out.loc[hit] = parsed[hit]
            remaining = out.isna() & series.notna() & ~yearless

    # ── Pass 3: dateutil with TWO sentinel defaults ─────────────────
    # If the parsed year differs between the two sentinels, the year
    # was supplied by the default (i.e. missing from the input) →
    # flag as year-less and DO NOT trust the date.
    if remaining.any():
        try:
            from dateutil import parser as _du
        except ImportError:
            _du = None
        if _du is not None:
            SENTINEL_A = datetime(1900, 1, 1)
            SENTINEL_B = datetime(1800, 6, 15)
            for idx in series.index[remaining]:
                raw = as_str.loc[idx]
                if not raw or raw.lower() in {"nan", "nat", "none"}:
                    continue
                try:
                    p_a = _du.parse(raw, dayfirst=True, default=SENTINEL_A)
                    p_b = _du.parse(raw, dayfirst=True, default=SENTINEL_B)
                except (ValueError, TypeError, OverflowError):
                    continue
                if p_a.year != p_b.year:
                    # Year came from the default → input was year-less
                    yearless.loc[idx] = True
                    continue
                out.loc[idx] = p_a

    return out, yearless


def _filter_by_recent_months(
    df: pd.DataFrame,
    date_col: str,
) -> tuple[pd.DataFrame, int, str]:
    """
    Keep rows whose date falls within the current or previous calendar
    month. Rows with truly unparseable dates are KEPT (benefit of the
    doubt). Rows whose date was date-like but year-less are DROPPED —
    their true year cannot be determined and they may be very old.
    """
    dates, yearless = _parse_date_series(df[date_col])

    today                = pd.Timestamp.today().normalize()
    start_current_month  = today.replace(day=1)
    start_previous_month = start_current_month - pd.offsets.MonthBegin(1)
    end_current_month    = start_current_month + pd.offsets.MonthEnd(1)

    prev_label = start_previous_month.strftime("%b %Y")
    curr_label = start_current_month.strftime("%b %Y")
    desc = f"{prev_label} – {curr_label}"

    in_window      = (dates >= start_previous_month) & (dates < end_current_month)
    is_unparseable = dates.isna() & ~yearless            # truly garbage → kept
    keep_mask      = in_window | is_unparseable          # yearless excluded

    filtered    = df[keep_mask]
    num_removed = int((~keep_mask).sum())

    print(
        f"[GFD DEBUG] Date parse: {len(df)} rows → "
        f"{int(in_window.sum())} in window, "
        f"{int(is_unparseable.sum())} unparseable (kept), "
        f"{int(yearless.sum())} year-less (dropped), "
        f"{num_removed} removed total."
    )

    return filtered, num_removed, desc

###################### Option 3 #################################

# Whitelist of acceptable date formats. Anything not matching one of
# these (after stripping whitespace and leading apostrophes) is treated
# as unparseable and the row is DROPPED.
_DATE_FORMATS = [
    "%d/%m/%Y", "%d/%m/%y",
    "%d.%m.%Y", "%d.%m.%y",
    "%d-%m-%Y", "%d-%m-%y",
    "%Y-%m-%d", "%Y/%m/%d", "%Y.%m.%d",
]


def _parse_date_series(series: pd.Series) -> pd.Series:
    """
    Strict date parser: returns a datetime64 Series where only values
    matching one of the whitelisted formats (or already-typed datetime
    cells from Excel) are populated. Everything else is NaT.

    Strategy:
      • Cells that are already datetime / Timestamp → used as-is.
        (Excel real-date cells come through this path.)
      • String cells are tested against _DATE_FORMATS one by one.
        The first format that parses cleanly wins.
      • Nothing else is accepted — no dayfirst auto-detect, no
        dateutil fallback, no year inference.
    """
    out = pd.Series(pd.NaT, index=series.index, dtype="datetime64[ns]")

    # ── Branch 1: cells that are already real datetime objects ──────
    # Covers pd.Timestamp, python datetime, and numpy datetime64.
    is_native = series.apply(
        lambda v: isinstance(v, (pd.Timestamp, datetime))
                  and not (isinstance(v, float) and pd.isna(v))
    )
    if is_native.any():
        out.loc[is_native] = pd.to_datetime(
            series[is_native], errors="coerce"
        )

    # ── Branch 2: string cells — strict format whitelist ────────────
    as_str = (
        series.where(~is_native)
        .astype(str)
        .str.strip()
        .str.lstrip("'")
        .str.strip()
    )
    # Blank / placeholder strings are not worth trying
    is_blank = as_str.isin({"", "nan", "NaN", "NaT", "None", "none"})
    remaining = out.isna() & ~is_native & ~is_blank

    for fmt in _DATE_FORMATS:
        if not remaining.any():
            break
        parsed = pd.to_datetime(
            as_str.where(remaining), format=fmt, errors="coerce"
        )
        hit = parsed.notna() & remaining
        if hit.any():
            out.loc[hit] = parsed[hit]
            remaining = out.isna() & ~is_native & ~is_blank

    return out


def _filter_by_recent_months(
    df: pd.DataFrame,
    date_col: str,
) -> tuple[pd.DataFrame, int, str]:
    """
    Keep only rows whose date falls within the current or previous
    calendar month. Rows whose date does not match one of the nine
    whitelisted formats (and is not a native Excel date cell) are
    DROPPED.

    Returns (filtered_df, num_removed, window_description).
    """
    dates = _parse_date_series(df[date_col])

    today                = pd.Timestamp.today().normalize()
    start_current_month  = today.replace(day=1)
    start_previous_month = start_current_month - pd.offsets.MonthBegin(1)
    end_current_month    = start_current_month + pd.offsets.MonthEnd(1)

    prev_label = start_previous_month.strftime("%b %Y")
    curr_label = start_current_month.strftime("%b %Y")
    desc = f"{prev_label} – {curr_label}"

    in_window   = (dates >= start_previous_month) & (dates < end_current_month)
    keep_mask   = in_window.fillna(False)        # NaT comparisons → False → dropped

    filtered    = df[keep_mask]
    num_removed = int((~keep_mask).sum())

    print(
        f"[GFD DEBUG] Date parse: {len(df)} rows → "
        f"{int(dates.notna().sum())} parsed, "
        f"{int(dates.isna().sum())} unparseable (dropped), "
        f"{int(in_window.sum())} in window, "
        f"{num_removed} removed total."
    )

    return filtered, num_removed, desc