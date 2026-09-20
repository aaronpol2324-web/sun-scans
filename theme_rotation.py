# =====================================================================
# THEME ROTATION DASHBOARD
# ---------------------------------------------------------------------
# A separate, standalone script from scans.py -- its own file, its own
# .xlsx output -- built to salvage the two things that actually worked
# in an earlier, abandoned web-dashboard project ("Sector Rotation
# Map") without dragging along the part that never did.
#
# WHAT CAME FROM THAT OLD PROJECT:
#   - mega_theme_structure.json: the full taxonomy of 41 mega themes,
#     181 sub-themes, and ~3,500 constituent tickers (Banks -> Universal
#     Banks / Regional Banks, Semiconductors & Hardware -> Memory /
#     Foundry / RF Semiconductors / ..., and so on). This is real,
#     useful, hand-built research and is used here completely unchanged.
#   - The equal-weighted theme-aggregation approach (normalize each
#     constituent to its own base=100 index off its first available
#     close, then average High/Low/Close across constituents per date
#     to get one synthetic "theme index" OHLC series) -- ported from
#     that project's get_candles()/get_theme_rankings() endpoints,
#     which worked fine. See _build_theme_aggregate_ohlc below.
#
# WHAT DID NOT COME FROM THAT OLD PROJECT:
#   - Its RS/Thrust scoring (an annualized-information-ratio-style
#     "composite_rs", separately calibrated 0-100 "rs_score" bands, and
#     a from-scratch RRG ratio/momentum calculation) was the one piece
#     that never got working end to end. Rather than debug a second,
#     parallel RS methodology, every mega theme, sub-theme, and
#     individual ticker here is scored with the EXACT SAME Real
#     Relative Strength (RRS) engine already built, tested, and running
#     in scans.py's RS_Groups/RS_Indices_Sectors tabs -- imported
#     directly (_rs_metrics_for, _rs_row, write_rs_flat_sheet), not
#     reimplemented. A theme's synthetic OHLC series (or a real ticker's
#     own OHLC series, at the Tickers level) is just handed to the same
#     code that scores a ticker on the other two tabs; nothing about the
#     RRS math changes for a theme aggregate versus a single ticker.
#
# LAYOUT: three tabs -- Mega_Themes, Sub_Themes, Tickers -- each ONE flat,
# fully sortable/filterable Excel Table (scans.write_rs_flat_sheet), not
# split into a separate mini-table per mega theme the way
# RS_Indices_Sectors is. The whole point is to see every mega theme (or
# every sub-theme, or every ticker) together at once and sort the WHOLE
# list by RS Thrust/1-Mth RRS to surface the leading and emerging ones,
# rather than having to scroll block by block. Drilling down from one
# level to the next is filtering, not scrolling: every column (including
# the "Category"/"Sub-Theme" identifying columns) gets its own header
# dropdown, so spotting a promising mega theme on Mega_Themes, then
# filtering Sub_Themes' "Category" column to just that name, then
# filtering Tickers' "Sub-Theme" column to one promising sub-theme, is
# how you go from "which themes are hot" down to "which tickers." That
# reuse is direct (this file calls scans.write_rs_flat_sheet itself), not
# a visual imitation of it, so any future fix to that rendering pipeline
# in scans.py applies here too without touching this file.
#
# ONE HONEST CAVEAT: "Price" for a mega-theme or sub-theme row is NOT a
# real dollar price -- it's a synthetic base-100 index level, drifting
# from 100 as of the earliest date the fetch window covers (only the
# Tickers tab's "Price" is a real price, since that's a real ticker's own
# history). "RS Thrust" and "1-Mth RRS" mean exactly what they mean on
# the other two scans.py tabs (excess move vs. SPY, in units of the
# aggregate's own synthetic ATR) at every level here -- there is nothing
# theme-specific about the scoring itself, only about what feeds it.
#
# CACHING: unlike scans.py's ~86-ticker fetch, this pulls ~3,500 tickers in
# ~24 chunks, which takes real time. Since these are daily bars, a second
# run on the same calendar day can't get anything new from yfinance no
# matter what -- so results are cached to theme_data_cache.pkl (next to
# this file) and reused for any ticker already fetched today. The cache
# resets itself automatically once the date rolls over. See
# fetch_theme_universe_history / CACHE_FILE / FORCE_REFRESH below.
# =====================================================================

import json
import logging
import os
import pickle
import time
from datetime import date

import pandas as pd
import xlsxwriter
import yfinance as yf

import scans  # scan-running code lives behind `if __name__ == "__main__"` in
              # scans.py, so importing it here just gets us its functions/
              # constants -- no side effects, no network calls on import.

TAXONOMY_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "mega_theme_structure.json")

# Daily on-disk cache: these are daily OHLC bars, so within a single calendar
# day nothing new is coming from yfinance no matter how many times this
# script runs -- a same-day rerun should just reuse what the first run
# already fetched instead of re-pulling all ~3,500 tickers in ~24 chunks
# again. The moment the calendar date rolls over, every ticker needs a new
# bar anyway, so there's no partial "some tickers are staler than others"
# case worth tracking -- the cache is either "from today" (usable) or "not
# from today" (treated as empty and fully refetched).
CACHE_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "theme_data_cache.pkl")

# Set True to ignore any existing cache and force a fresh fetch of every
# ticker, e.g. if you want an intraday re-pull despite the daily-cache logic
# above, or suspect the cached data is bad.
FORCE_REFRESH = False

# Fetching the full taxonomy is ~3,500 tickers, not the ~86 the two existing
# RS tabs pull -- one single yf.download() call for that many names is exactly
# the kind of oversized batch request that gets unreliable well before you'd
# expect, so this goes in ~150-ticker chunks instead, each independently
# wrapped so one bad chunk (a delisted symbol, a rate-limit blip) only costs
# that chunk's tickers rather than the whole run.
FETCH_CHUNK_SIZE = 150
CHUNK_PAUSE_SECONDS = 1.0  # a small courtesy pause between ~24 chunks

# A sub-theme needs at least this many of ITS OWN listed tickers to actually
# price before we bother building a synthetic aggregate for it at all -- 0
# priced tickers means there's nothing to aggregate, full stop.
MIN_CONSTITUENTS = 1

# Within a sub-theme's aggregate, a given DATE needs at least this fraction of
# that sub-theme's priced constituents present to count -- otherwise a data
# gap affecting most of the group (a holiday quirk, a late listing) would let
# one or two remaining names silently stand in for "the whole theme's return"
# that day.
MIN_DATE_COVERAGE_FRACTION = 0.5


def load_taxonomy() -> dict:
    """Load the mega-theme -> sub-theme -> [tickers] structure, unchanged from
    the old project's own mega_theme_structure.json."""
    with open(TAXONOMY_FILE, "r") as f:
        return json.load(f)


def _all_taxonomy_tickers(taxonomy: dict) -> set:
    tickers = set()
    for sub_dict in taxonomy.values():
        for tick_list in sub_dict.values():
            tickers.update(t.upper().strip() for t in (tick_list or []) if t)
    tickers.add("SPY")  # the benchmark every sub-theme is measured against
    return tickers


def _load_cache() -> dict:
    """Returns {"date": "YYYY-MM-DD", "data": {ticker: DataFrame}}. A
    missing, corrupt, or otherwise unreadable cache file is just treated as
    an empty one -- a cache miss should only ever cost time, never break
    the run."""
    try:
        with open(CACHE_FILE, "rb") as f:
            cache = pickle.load(f)
        if isinstance(cache, dict) and "date" in cache and "data" in cache:
            return cache
    except Exception:
        pass
    return {"date": None, "data": {}}


def _save_cache(cache: dict):
    try:
        with open(CACHE_FILE, "wb") as f:
            pickle.dump(cache, f)
    except Exception as e:
        print(f"  [warning] could not write cache file ({e}) -- next run will re-fetch everything.")


def _fetch_chunked(tickers) -> dict:
    """The actual network call: chunked yfinance download of High/Low/Close
    for the given tickers. Per-ticker extraction logic mirrors scans.py's
    own _fetch_rs_history (same dropna/min-length rule), just wrapped in
    chunks with per-chunk error isolation since a request this large can't
    safely go in one call the way the ~86-ticker RS dashboard's fetch does.

    Returns {ticker: DataFrame} for whatever priced successfully -- missing
    tickers just aren't in the dict, so callers see it as a blank/degraded
    row rather than the whole run failing."""
    logger = yf.utils.get_yf_logger()
    original_level = logger.level
    logger.setLevel(logging.CRITICAL)
    fetched = {}
    tickers = sorted(set(tickers))
    n_chunks = (len(tickers) + FETCH_CHUNK_SIZE - 1) // FETCH_CHUNK_SIZE
    try:
        for i in range(0, len(tickers), FETCH_CHUNK_SIZE):
            chunk = tickers[i:i + FETCH_CHUNK_SIZE]
            chunk_num = i // FETCH_CHUNK_SIZE + 1
            print(f"  Fetching chunk {chunk_num}/{n_chunks} ({len(chunk)} tickers)...")
            try:
                data = yf.download(chunk, period="2y", progress=False, group_by='ticker')
            except Exception as e:
                print(f"    chunk {chunk_num} failed entirely ({e}); skipping its {len(chunk)} tickers.")
                continue
            for t in chunk:
                try:
                    df = data if len(chunk) == 1 else (data[t] if t in data.columns.levels[0] else pd.DataFrame())
                    df = df.dropna(subset=['Close'])
                    if len(df) >= scans.RS_ONE_MONTH_WINDOW + 1:
                        fetched[t] = df
                except Exception:
                    continue
            if chunk_num < n_chunks:
                time.sleep(CHUNK_PAUSE_SECONDS)
    finally:
        logger.setLevel(original_level)
    return fetched


def fetch_theme_universe_history(tickers, use_cache: bool = True) -> dict:
    """Cache-aware wrapper around _fetch_chunked. These are daily bars, so
    a rerun on the SAME calendar day can't get anything new from yfinance
    no matter how many times it asks -- if today's cache already has a
    ticker, this reuses it and only hits the network for tickers that are
    missing or new (e.g. just added to the taxonomy). The moment the date
    rolls over, the whole cache is one day stale for every ticker alike
    (there's no such thing as "half-stale" daily data), so it's treated as
    empty and everything gets refetched fresh -- see the CACHE_FILE/
    FORCE_REFRESH comments above.

    Returns {ticker: DataFrame}, same shape/contract as _fetch_chunked."""
    tickers = sorted(set(tickers))
    if not use_cache:
        return _fetch_chunked(tickers)

    today_str = date.today().isoformat()
    cache = _load_cache()
    cache_is_fresh = (cache["date"] == today_str) and not FORCE_REFRESH
    cached_data = cache["data"] if cache_is_fresh else {}

    to_fetch = [t for t in tickers if t not in cached_data]
    if not to_fetch:
        print(f"  Using today's cache for all {len(tickers)} tickers -- no fetch needed.")
        return {t: cached_data[t] for t in tickers}

    if cached_data:
        print(f"  {len(tickers) - len(to_fetch)}/{len(tickers)} tickers already cached from earlier today; "
              f"fetching the other {len(to_fetch)}.")
    freshly_fetched = _fetch_chunked(to_fetch)

    merged = dict(cached_data)
    merged.update(freshly_fetched)
    _save_cache({"date": today_str, "data": merged})

    return {t: merged[t] for t in tickers if t in merged}


def _build_theme_aggregate_ohlc(constituent_dfs: dict) -> pd.DataFrame:
    """Equal-weighted synthetic OHLC for a theme: each constituent's own
    High/Low/Close normalized to its OWN base=100 (its first available close
    in this fetch window), then averaged across constituents per date --
    ported from the old project's candle/theme-ranking aggregation, which is
    the part of it that worked fine. This is an equal-weighted RETURN index,
    not a real tradable OHLC series -- a $500 stock and a $5 stock both start
    at 100 and contribute equally to the average from there.

    A date only counts toward the aggregate if at least
    MIN_DATE_COVERAGE_FRACTION of the theme's priced constituents have data
    on it, so a data gap in most of the group can't let one or two remaining
    names quietly stand in for "the whole theme" that day."""
    if not constituent_dfs:
        return pd.DataFrame(columns=['High', 'Low', 'Close'])

    closes = pd.DataFrame({t: df['Close'] for t, df in constituent_dfs.items()})
    highs = pd.DataFrame({t: df['High'] for t, df in constituent_dfs.items()})
    lows = pd.DataFrame({t: df['Low'] for t, df in constituent_dfs.items()})

    # bfill().iloc[0] picks up each column's own first REAL value (any leading
    # gap before a ticker's earliest date gets backfilled from that first
    # value), which is exactly "normalize to this ticker's own first
    # available close" -- just computed vectorized across the whole group
    # instead of ticker-by-ticker.
    base = closes.bfill().iloc[0]
    valid_cols = base[(base > 0) & base.notna()].index
    if len(valid_cols) == 0:
        return pd.DataFrame(columns=['High', 'Low', 'Close'])

    norm_close = closes[valid_cols] / base[valid_cols] * 100
    norm_high = highs[valid_cols] / base[valid_cols] * 100
    norm_low = lows[valid_cols] / base[valid_cols] * 100

    coverage = norm_close.notna().sum(axis=1)
    min_required = max(1, int(len(valid_cols) * MIN_DATE_COVERAGE_FRACTION))
    keep = coverage >= min_required

    agg = pd.DataFrame({
        'High': norm_high.mean(axis=1),
        'Low': norm_low.mean(axis=1),
        'Close': norm_close.mean(axis=1),
    })[keep].dropna(how='any')
    return agg


def _mega_theme_tickers(sub_dict: dict) -> set:
    """Every ticker across ALL of one mega theme's sub-themes, deduped --
    a ticker tagged under more than one sub-theme of the same mega theme
    (e.g. a diversified name relevant to two of a mega theme's sub-groups)
    only counts once when building that mega theme's own aggregate."""
    tickers = set()
    for tick_list in sub_dict.values():
        tickers.update(t.upper().strip() for t in (tick_list or []) if t)
    return tickers


def run_theme_dashboard():
    """Builds all three levels of the theme rotation view -- mega theme,
    sub-theme, and individual ticker -- each scored by the exact same Real
    Relative Strength engine as scans.py's RS_Groups/RS_Indices_Sectors
    tabs (see the header comment above for why). Returns
    (mega_df, sub_df, ticker_df), one DataFrame per level, each ready for
    scans.write_rs_flat_sheet:

      mega_df:   one row per mega theme -- ALL of that mega theme's
                 tickers (across every one of its sub-themes) aggregated
                 into one synthetic index and scored, same way a
                 sub-theme is. No group_cols -- there's no level above
                 "mega theme."
      sub_df:    one row per sub-theme (what run_theme_dashboard used to
                 return on its own) -- "Category" = mega theme, so
                 write_rs_flat_sheet's group_cols=["Category"] makes the
                 mega theme a filterable column instead of a banner row,
                 keeping every sub-theme in ONE sortable list.
      ticker_df: one row per (mega theme, sub-theme, ticker) MEMBERSHIP,
                 each individual ticker scored on its own real price
                 history (not the aggregate) -- a ticker tagged under two
                 sub-themes gets two rows, one per membership, so
                 filtering this tab down to a single sub-theme shows
                 exactly (and only) that sub-theme's constituents.
                 group_cols=["Category", "Sub-Theme"].

    All three share the same Category/Ticker/Name/Price/RS Thrust/
    1-Mth RRS/1D %/1M %/Off 52W High %/_PriceSpark/_RSSpark schema (mega_df
    just doesn't use "Category")."""
    taxonomy = load_taxonomy()
    all_tickers = _all_taxonomy_tickers(taxonomy)
    n_sub_configured = sum(len(v) for v in taxonomy.values())

    print(f"Theme Rotation: {len(taxonomy)} mega themes, {n_sub_configured} sub-themes, "
          f"{len(all_tickers)} unique tickers to fetch.")

    history = fetch_theme_universe_history(all_tickers)
    spy_df = history.get("SPY")
    if spy_df is None:
        raise RuntimeError("Could not fetch SPY -- nothing can be scored without the benchmark.")

    # Individual tickers are scored once each (on their own real price
    # history, via scans._rs_metrics_for -- exactly like a RS_Groups row)
    # and cached here so a ticker tagged under several sub-themes doesn't
    # get recomputed -- or, worse, risk drifting -- for each membership.
    ticker_rec_cache = {}

    def _ticker_rec(t):
        if t not in ticker_rec_cache:
            ticker_rec_cache[t] = scans._rs_metrics_for(history[t], spy_df, t) if t in history else None
        return ticker_rec_cache[t]

    mega_rows, sub_rows, ticker_rows = [], [], []
    n_sub_scored = 0

    for mega, sub_dict in taxonomy.items():
        # --- mega theme level: aggregate ALL of this mega theme's tickers ---
        mega_clean = sorted(_mega_theme_tickers(sub_dict))
        mega_priced = {t: history[t] for t in mega_clean if t in history}
        if len(mega_priced) >= MIN_CONSTITUENTS:
            mega_agg = _build_theme_aggregate_ohlc(mega_priced)
            if len(mega_agg) >= scans.RS_ONE_MONTH_WINDOW + 1:
                rec = scans._rs_metrics_for(mega_agg, spy_df, mega)
                mega_rows.append(scans._rs_row(
                    {"Ticker": mega, "Name": f"{len(mega_priced)}/{len(mega_clean)} tickers priced"}, rec))

        for sub, tick_list in sub_dict.items():
            clean = sorted({t.upper().strip() for t in (tick_list or []) if t})
            priced = {t: history[t] for t in clean if t in history}

            # --- ticker level: one row per (mega, sub, ticker) membership ---
            for t in clean:
                if t in history:
                    ticker_rows.append(scans._rs_row(
                        {"Category": mega, "Sub-Theme": sub, "Ticker": t, "Name": ""}, _ticker_rec(t)))

            # --- sub-theme level: same aggregate approach as mega theme,
            # scoped to just this sub-theme's own constituents ---
            if len(priced) < MIN_CONSTITUENTS:
                continue
            agg_df = _build_theme_aggregate_ohlc(priced)
            if len(agg_df) < scans.RS_ONE_MONTH_WINDOW + 1:
                continue

            # `sub` (a sub-theme name like "Regional Banks") is never
            # literally "SPY", so this always takes _rs_metrics_for's real
            # RRS-vs-SPY path -- never the self-momentum fallback that
            # RS_Indices_Sectors' own SPY row needs.
            rec = scans._rs_metrics_for(agg_df, spy_df, sub)
            sub_rows.append(scans._rs_row(
                {"Category": mega, "Ticker": sub, "Name": f"{len(priced)}/{len(clean)} tickers priced"}, rec))
            n_sub_scored += 1

    print(f"Finished Theme Rotation: {len(mega_rows)}/{len(taxonomy)} mega themes, "
          f"{n_sub_scored}/{n_sub_configured} sub-themes, {len(ticker_rows)} ticker memberships scored.")
    return pd.DataFrame(mega_rows), pd.DataFrame(sub_rows), pd.DataFrame(ticker_rows)


def write_theme_workbook(mega_df: pd.DataFrame, sub_df: pd.DataFrame, ticker_df: pd.DataFrame, path: str):
    """Three tabs, each one flat, fully sortable/filterable Excel Table
    (scans.write_rs_flat_sheet) rather than the banner-block-per-category
    layout write_rs_indices_sheet uses -- the whole point here is to see
    EVERY mega theme (or sub-theme, or ticker membership) together at
    once and sort/surface the leading and emerging ones across the whole
    list, then drill down by filtering a column rather than by scrolling
    to a separate block:
      Mega_Themes -> filter/sort freely
      Sub_Themes  -> filter its "Category" column to one mega theme's name
      Tickers     -> filter its "Sub-Theme" column to one sub-theme's name
                     (or "Category" for a whole mega theme's tickers)."""
    workbook = xlsxwriter.Workbook(path)
    fmt_cache = {}

    ws_mega = scans.write_rs_flat_sheet(workbook, fmt_cache, mega_df, sheet_name="Mega_Themes")
    ws_sub = scans.write_rs_flat_sheet(workbook, fmt_cache, sub_df, sheet_name="Sub_Themes", group_cols=["Category"])
    ws_ticker = scans.write_rs_flat_sheet(workbook, fmt_cache, ticker_df, sheet_name="Tickers",
                                           group_cols=["Category", "Sub-Theme"])

    # Mega/sub-theme names ("Semiconductors & Hardware", "Precious Metals
    # Royalty", ...) run longer than scans.py sized "Ticker" for (a 4-5
    # character real ticker symbol) -- widen it on every tab where "Ticker"
    # holds a theme name rather than an actual symbol. (The Tickers tab's
    # own "Ticker" column holds real symbols, so it's left at its normal
    # width; only its "Category"/"Sub-Theme" group columns are theme names,
    # already widened by write_rs_flat_sheet's own group_cols handling.)
    if ws_mega is not None:
        ws_mega.set_column(scans.RS_COL_TICKER, scans.RS_COL_TICKER, 30)
    if ws_sub is not None:
        ws_sub.set_column(1 + scans.RS_COL_TICKER, 1 + scans.RS_COL_TICKER, 26)

    workbook.close()


if __name__ == "__main__":
    mega_df, sub_df, ticker_df = run_theme_dashboard()
    out_path = r"C:\TradingScans\Theme_Rotation.xlsx"
    try:
        write_theme_workbook(mega_df, sub_df, ticker_df, out_path)
        print(f"Done! Saved to: {out_path}")
    except PermissionError:
        alt_path = r"C:\TradingScans\Theme_Rotation_NEW.xlsx"
        write_theme_workbook(mega_df, sub_df, ticker_df, alt_path)
        print(f"\n[WARNING] Excel file was locked. Saved to: {alt_path}")
