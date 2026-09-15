import os
import json
import time
import hashlib
from datetime import datetime, timedelta
import pandas as pd
import yfinance as yf
from fastapi import FastAPI, HTTPException, Query
from fastapi.middleware.cors import CORSMiddleware
from fastapi.staticfiles import StaticFiles
import logging
from fastapi.responses import Response
import os
import time

# The name of the file that will live on your hard drive. Separate file per timeframe -
# this used to be a single shared file/column-name cache with no timeframe awareness,
# which meant whichever timeframe fetched a symbol FIRST (e.g. /api/rrg's weekly default)
# would "poison" that symbol's cache entry for every other timeframe that asked for it
# afterward (e.g. daily stats requests silently getting back weekly-resampled data).
def _local_cache_file(timeframe):
    tf = (timeframe or "1D").upper()
    return "market_data_cache.json" if tf == "1D" else f"market_data_cache_{tf}.json"


def get_local_cache(symbols, timeframe="1D"):
    """Loads whatever historical price data we already have saved on disk for this timeframe."""
    cache_file = _local_cache_file(timeframe)
    if os.path.exists(cache_file):
        if time.time() - os.path.getmtime(cache_file) < 86400:
            try:
                df = pd.read_json(cache_file, orient="split")
                if not isinstance(df.index, pd.DatetimeIndex):
                    df.index = pd.to_datetime(df.index)
                return df
            except Exception:
                pass
    return None


def save_local_cache(df, timeframe="1D"):
    """Saves the DataFrame directly using Pandas built-in JSON handler, to this timeframe's file."""
    cache_file = _local_cache_file(timeframe)
    print(f"--> Saving fresh {timeframe} data to local computer cache ({cache_file})...")
    try:
        df.to_json(cache_file, orient="split", date_format="iso")
    except Exception as e:
        print(f"Could not save local cache: {e}")

# Suppress yfinance internal warning prints completely
logging.getLogger('yfinance').setLevel(logging.CRITICAL)

app = FastAPI()
@app.get("/favicon.ico", include_in_schema=False)
async def favicon():
    return Response(status_code=204)

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"],
)

SECTOR_ETFS = {
    "XLK": {"name": "Technology", "color": "#00bcd4"},
    "XLF": {"name": "Financials", "color": "#3f51b5"},
    "XLV": {"name": "Health Care", "color": "#4caf50"},
    "XLE": {"name": "Energy", "color": "#ff9800"},
    "XLP": {"name": "Consumer Staples", "color": "#e91e63"},
    "XLB": {"name": "Materials", "color": "#9c27b0"},
    "XLRE": {"name": "Real Estate", "color": "#009688"},
    "XLI": {"name": "Industrials", "color": "#795548"},
    "XLU": {"name": "Utilities", "color": "#607d8b"},
    "XLY": {"name": "Discretionary", "color": "#cddc39"},
    "XLC": {"name": "Comm Services", "color": "#ffeb3b"},
}

import time
import hashlib
import requests
import pandas as pd
from datetime import datetime, timedelta

_price_cache = {}
CACHE_TTL = 300

def _cache_key(symbols, start_date, end_date, timeframe):
    raw = f"{sorted(symbols)}:{start_date}:{end_date}:{timeframe}"
    return hashlib.md5(raw.encode()).hexdigest()


def fetch_prices(symbols, start_date=None, end_date=None, timeframe="1D"):
    if end_date is None:
        end_date = (datetime.now() + timedelta(days=1)).strftime('%Y-%m-%d')
    if start_date is None:
        start_date = (datetime.now() - timedelta(days=730)).strftime('%Y-%m-%d')

    symbols = list(symbols)

    # 1. Load what we already have from disk cache (for THIS timeframe specifically)
    saved_dataframe = get_local_cache(symbols, timeframe)
    missing_symbols = []

    if saved_dataframe is not None:
        missing_symbols = [s for s in symbols if s not in saved_dataframe.columns]
        if not missing_symbols:
            print("--> Loading 100% from local computer cache (Instant)")
            return saved_dataframe
    else:
        missing_symbols = symbols

    # If everything is already cached, we wouldn't reach here.
    # Now we only fetch data for the actual missing symbols instead of the whole portfolio!
    symbols_to_download = missing_symbols if saved_dataframe is not None else symbols

    def _extract_close(df, sym):
        if df is None or df.empty:
            return None

        close_col = None
        if isinstance(df.columns, pd.MultiIndex):
            if ('Adj Close', sym) in df.columns:
                close_col = df[('Adj Close', sym)]
            elif ('Close', sym) in df.columns:
                close_col = df[('Close', sym)]
            else:
                return None
        else:
            if 'Adj Close' in df.columns:
                close_col = df['Adj Close']
            elif 'Close' in df.columns:
                close_col = df['Close']
            else:
                return None

        if not isinstance(close_col.index, pd.DatetimeIndex):
            close_col.index = pd.to_datetime(close_col.index)
        if close_col.index.tz is not None:
            close_col.index = close_col.index.tz_localize(None)

        return close_col if not close_col.dropna().empty else None

    data_dict = {}
    CHUNK_SIZE = 150

    for i in range(0, len(symbols_to_download), CHUNK_SIZE):
        chunk = symbols_to_download[i:i + CHUNK_SIZE]
        try:
            downloaded_df = yf.download(chunk, start=start_date, end=end_date, progress=False, auto_adjust=True)
        except Exception:
            continue

        if downloaded_df is None or downloaded_df.empty:
            continue

        if len(chunk) == 1:
            close_col = _extract_close(downloaded_df, chunk[0])
            if close_col is not None:
                data_dict[chunk[0]] = close_col
        else:
            for sym in chunk:
                close_col = _extract_close(downloaded_df, sym)
                if close_col is not None:
                    data_dict[sym] = close_col

    new_resampled = None
    if data_dict:
        all_close = pd.DataFrame(data_dict)
        if timeframe == "1D":
            new_resampled = all_close.dropna(how="all")
        elif timeframe == "1W":
            new_resampled = all_close.resample("W-FRI").last().dropna(how="all")
        elif timeframe == "1M":
            new_resampled = all_close.resample("ME").last().dropna(how="all")
        else:
            new_resampled = all_close.resample("1D").last().dropna(how="all")

    # 2. Merge newly fetched missing symbols with existing disk cache
    if saved_dataframe is not None:
        if new_resampled is not None and not new_resampled.empty:
            combined_df = pd.concat([saved_dataframe, new_resampled], axis=1)
            combined_df = combined_df.loc[:, ~combined_df.columns.duplicated()]
            save_local_cache(combined_df, timeframe)
            return combined_df
        return saved_dataframe
    else:
        if new_resampled is not None:
            save_local_cache(new_resampled, timeframe)
            return new_resampled
        else:
            raise HTTPException(status_code=404, detail="No data returned for given symbols")

LOCAL_OHLC_CACHE_FILE = "ohlc_data_cache.json"
OHLC_DISK_CACHE_MAX_AGE = 86400  # 24h, matching the existing price disk cache
OHLC_DISK_MAX_BARS = 260         # cache ~1yr of daily bars per symbol; requests slice .tail(days) from this

_ohlc_cache = {}
OHLC_CACHE_TTL = 300


def _load_ohlc_disk_cache():
    """Loads whatever OHLC history we already have saved on disk, as a plain dict
    (not pandas' to_json/read_json - MultiIndex-like per-symbol OHLC round-trips more
    reliably through plain json than through pandas' 'split' orient)."""
    if not os.path.exists(LOCAL_OHLC_CACHE_FILE):
        return {}
    try:
        if time.time() - os.path.getmtime(LOCAL_OHLC_CACHE_FILE) > OHLC_DISK_CACHE_MAX_AGE:
            return {}
        with open(LOCAL_OHLC_CACHE_FILE, "r") as f:
            return json.load(f)
    except Exception:
        return {}


def _save_ohlc_disk_cache(cache_dict):
    try:
        with open(LOCAL_OHLC_CACHE_FILE, "w") as f:
            json.dump(cache_dict, f)
    except Exception as e:
        print(f"Could not save OHLC disk cache: {e}")


def fetch_ohlc(symbols, days=90):
    """Batch-download Open/High/Low/Close (not just Close) for a list of tickers,
    for use by the candlestick grid view. Separate from fetch_prices/_price_cache
    since that path only ever kept Close.

    Three layers, fastest first:
      1. In-memory cache (5 min) - survives repeated clicks in the same session.
      2. Disk cache (up to 24h old, like market_data_cache.json) - survives server restarts.
      3. Live yfinance fetch for whatever's still missing, which then gets written to disk.

    Caches a generous ~1yr window per symbol regardless of the requested `days`, so
    different lookback requests for the same symbol reuse the same cached history
    instead of fragmenting the cache by day-count.
    """
    symbols = sorted(set(symbols))
    if not symbols:
        return {}

    now = time.time()
    mem_key = hashlib.md5(f"{symbols}".encode()).hexdigest()
    if mem_key in _ohlc_cache and (now - _ohlc_cache[mem_key]['ts']) < OHLC_CACHE_TTL:
        return {sym: df.tail(days) for sym, df in _ohlc_cache[mem_key]['data'].items()}

    disk_cache = _load_ohlc_disk_cache()
    result = {}
    missing = []

    for sym in symbols:
        entry = disk_cache.get(sym)
        if entry and len(entry.get("dates", [])) >= min(days, 30):
            try:
                result[sym] = pd.DataFrame(
                    {"Open": entry["open"], "High": entry["high"], "Low": entry["low"], "Close": entry["close"]},
                    index=pd.to_datetime(entry["dates"]),
                )
                continue
            except Exception:
                pass
        missing.append(sym)

    if missing:
        start_date = (datetime.now() - timedelta(days=int(OHLC_DISK_MAX_BARS * 1.6) + 10)).strftime('%Y-%m-%d')
        end_date = (datetime.now() + timedelta(days=1)).strftime('%Y-%m-%d')

        CHUNK_SIZE = 100
        for i in range(0, len(missing), CHUNK_SIZE):
            chunk = missing[i:i + CHUNK_SIZE]
            try:
                df = yf.download(chunk, start=start_date, end=end_date, progress=False, auto_adjust=True)
            except Exception:
                continue
            if df is None or df.empty:
                continue

            for sym in chunk:
                try:
                    if isinstance(df.columns, pd.MultiIndex):
                        if ('Close', sym) not in df.columns:
                            continue
                        sub = pd.DataFrame({
                            'Open': df[('Open', sym)],
                            'High': df[('High', sym)],
                            'Low': df[('Low', sym)],
                            'Close': df[('Close', sym)],
                        })
                    else:
                        if 'Close' not in df.columns:
                            continue
                        sub = df[['Open', 'High', 'Low', 'Close']]

                    sub = sub.dropna()
                    if not isinstance(sub.index, pd.DatetimeIndex):
                        sub.index = pd.to_datetime(sub.index)
                    if sub.index.tz is not None:
                        sub.index = sub.index.tz_localize(None)
                    if sub.empty:
                        continue

                    sub = sub.tail(OHLC_DISK_MAX_BARS)
                    result[sym] = sub
                    disk_cache[sym] = {
                        "dates": [d.strftime("%Y-%m-%d") for d in sub.index],
                        "open": sub["Open"].round(4).tolist(),
                        "high": sub["High"].round(4).tolist(),
                        "low": sub["Low"].round(4).tolist(),
                        "close": sub["Close"].round(4).tolist(),
                    }
                except Exception:
                    continue

        _save_ohlc_disk_cache(disk_cache)

    _ohlc_cache[mem_key] = {'data': result, 'ts': now}
    return {sym: df.tail(days) for sym, df in result.items()}


def load_taxonomy():
    """Load the complete mega-theme, theme, and ticker taxonomy from the single JSON structure."""
    json_path = os.path.join(os.path.dirname(__file__), "mega_theme_structure.json")
    with open(json_path, "r") as f:
        return json.load(f)

@app.get("/api/theme-structure")
def get_theme_structure():
    return load_taxonomy()


@app.get("/api/theme-rankings")
def get_theme_rankings(benchmark: str = Query("SPY")):
    benchmark = benchmark.upper().strip()
    structure = load_taxonomy()  # <-- Updated from load_excel_taxonomy() to load_taxonomy()

    all_tickers = {benchmark}
    subtheme_map = []

    for mega_name, sub_dict in structure.items():
        for sub_name, tickers in sub_dict.items():
            if tickers:
                clean_tickers = [t.upper().strip() for t in tickers if t]
                if clean_tickers:
                    subtheme_map.append({
                        "mega": mega_name,
                        "sub": sub_name,
                        "tickers": clean_tickers
                    })
                    all_tickers.update(clean_tickers)

    prices_df = fetch_prices(list(all_tickers), timeframe="1D")
    if benchmark not in prices_df.columns:
        raise HTTPException(status_code=404, detail=f"Benchmark {benchmark} not found")

    bench_closes = prices_df[benchmark]
    rankings = []

    for item in subtheme_map:
        valid_tickers = [t for t in item["tickers"] if t in prices_df.columns]
        if not valid_tickers:
            continue

        sub_df = prices_df[valid_tickers]
        # Normalize each constituent to its own base=100 index before averaging, so the
        # theme aggregate reflects equal-weighted RETURNS. Averaging raw prices instead
        # would let whichever constituent has the highest dollar price dominate the
        # theme's dollar-level swings regardless of its actual % move.
        normalized = sub_df / sub_df.bfill().iloc[0] * 100
        sub_closes = normalized.mean(axis=1).dropna()
        combined = pd.concat([sub_closes, bench_closes], axis=1).dropna()
        if len(combined) < 50:
            continue

        sub_c = combined.iloc[:, 0]
        bench_c = combined.iloc[:, 1]

        rs_1m = _vol_adj_rs_component(sub_c, bench_c, 21)
        rs_3m = _vol_adj_rs_component(sub_c, bench_c, 63)
        rs_6m = _vol_adj_rs_component(sub_c, bench_c, 126)
        rs_1y = _vol_adj_rs_component(sub_c, bench_c, 252)
        # Composite goes through the shared _raw_composite_rs (not a local recompute of the
        # same 40/30/20/10 blend) so it also picks up the 52-week high/low term - keeping this
        # endpoint, get_group_stats, and get_full_universe_rs_scores all measuring RS the same
        # way. rs_1m/rs_3m/rs_6m/rs_1y above stay pure per-timeframe numbers for display.
        composite_rs = _raw_composite_rs(sub_c, bench_c)

        high_52w = sub_c.tail(min(252, len(sub_c))).max()
        current_price = sub_c.iloc[-1]
        dist_from_high = float((current_price / high_52w) - 1) if high_52w > 0 else 0.0

        rankings.append({
            "mega_theme": item["mega"],
            "sub_theme": item["sub"],
            "tickers": valid_tickers,
            "current_price": round(float(current_price), 2),
            "rs_1m": round(rs_1m * 100, 2),
            "rs_3m": round(rs_3m * 100, 2),
            "rs_6m": round(rs_6m * 100, 2),
            "rs_1y": round(rs_1y * 100, 2),
            "composite_rs": round(composite_rs * 100, 2),
            "dist_from_high": round(dist_from_high * 100, 2),
        })

    rankings.sort(key=lambda x: x["composite_rs"], reverse=True)
    return {"benchmark": benchmark, "rankings": rankings}


@app.get("/api/dashboard-stats")
def get_dashboard_stats():
    """Restricted strictly to SPY and core sector ETFs for the bottom stats table."""
    tracked_tickers = ["SPY"] + list(SECTOR_ETFS.keys())

    prices_df = fetch_prices(tracked_tickers, timeframe="1D")

    volumes_df = pd.DataFrame()
    try:
        raw_vol = yf.download(tracked_tickers, period="3y", progress=False)
        if isinstance(raw_vol.columns, pd.MultiIndex) and "Volume" in raw_vol.columns.levels[0]:
            volumes_df = raw_vol["Volume"]
        elif "Volume" in raw_vol.columns:
            volumes_df = raw_vol[["Volume"]]
    except Exception:
        pass

    stats_list = []

    for symbol in tracked_tickers:
        if symbol not in prices_df.columns:
            continue
        series = prices_df[symbol].dropna()

        if len(series) < 20:
            continue

        current_price = float(series.iloc[-1])
        vol = 0
        avg_vol = 0
        if not volumes_df.empty:
            if symbol in volumes_df.columns:
                v_ser = volumes_df[symbol].dropna()
                if not v_ser.empty:
                    vol = int(v_ser.iloc[-1])
                    avg_vol = int(v_ser.tail(20).mean())

        def safe_chg(bars):
            """
            Pure trading-session lookback: compares the current close to the close
            exactly `bars` trading sessions ago. Matches standard platform convention
            and is consistent with the other stats endpoints in this app.
            """
            if series is None or len(series) <= bars:
                return 0.0

            base_price = series.iloc[-1 - bars]
            curr_price = series.iloc[-1]

            if not base_price:
                return 0.0

            return float((curr_price / base_price) - 1.0)

        sma_20 = series.tail(20).mean() if len(series) >= 20 else current_price
        sma_50 = series.tail(50).mean() if len(series) >= 50 else current_price
        sma_200 = series.tail(200).mean() if len(series) >= 200 else current_price

        if symbol == "SPY":
            desc = "SPDR S&P 500 ETF Trust"
        elif symbol in SECTOR_ETFS:
            desc = f"{SECTOR_ETFS.get(symbol, {}).get('name', symbol)} Sector ETF"
        else:
            desc = f"Equity / Asset: {symbol}"

        stats_list.append({
            "symbol": symbol,
            "description": SECTOR_ETFS.get(symbol, {}).get("name", "Benchmark"),
            "price": current_price,
            "chg_1d": round(safe_chg(1), 4),
            "chg_1w": round(safe_chg(5), 4),
            "chg_2w": round(safe_chg(10), 4),
            "chg_1m": round(safe_chg(21), 4),
            "chg_2m": round(safe_chg(42), 4),
            "chg_3m": round(safe_chg(63), 4),
            "chg_6m": round(safe_chg(126), 4),
            "chg_1y": round(safe_chg(252), 4),
            "volume": vol,
            "avg_volume": avg_vol,
            # ... other indicators ...
        })

    return {"stats": stats_list}


_STATS_CACHE = []
_CACHE_LOADED = False


@app.get("/api/theme-ticker-stats")
def get_theme_ticker_stats():
    """Cached stats endpoint exclusively for sidebar theme tickers."""
    global _STATS_CACHE, _CACHE_LOADED
    if _CACHE_LOADED and _STATS_CACHE:
        return _STATS_CACHE

    try:
        structure = load_taxonomy()
        theme_tickers = set()
        if structure and isinstance(structure, dict):
            for mega_name, sub_dict in structure.items():
                if isinstance(sub_dict, dict):
                    for sub_name, tickers in sub_dict.items():
                        if isinstance(tickers, list):
                            for t in tickers:
                                if t and isinstance(t, str):
                                    theme_tickers.add(t.upper().strip())

        tracked_tickers = list(theme_tickers)
        if not tracked_tickers:
            return []

        # Uses your existing fetch_prices helper
        prices_df = fetch_prices(tracked_tickers, timeframe="1D")
        stats_list = []

        for symbol in tracked_tickers:
            if symbol not in prices_df.columns:
                continue
            series = prices_df[symbol].dropna()
            if len(series) < 5:
                continue

            current_price = float(series.iloc[-1])

            def safe_chg(bars):
                if series is None or len(series) < 2:
                    return 0.0

                # For 1W (bars=5), anchor strictly to the previous Friday's close
                if bars == 5:
                    # Find all Friday bars in the index
                    fridays = series[series.index.weekday == 4]
                    if not fridays.empty:
                        # Get the last Friday that is strictly before the current week/day
                        current_date = series.index[-1]
                        past_fridays = fridays[fridays.index < current_date]
                        if not past_fridays.empty:
                            base_price = past_fridays.iloc[-1]
                            return float((series.iloc[-1] / base_price) - 1.0)

                # Fallback to standard session lookback for other timeframes
                if len(series) > bars:
                    return float((series.iloc[-1] / series.iloc[-1 - bars]) - 1.0)
                elif len(series) > 1:
                    return float((series.iloc[-1] / series.iloc[0]) - 1.0)
                return 0.0
    except Exception as e:
        print(f"Error in theme-ticker-stats: {e}")
        return []

def compute_single_rrg(sector_prices, benchmark_prices, tail_length=8):
    """Calculates RRG ratio and momentum with strict date alignment."""
    # 1. Align the theme prices and benchmark prices by date first
    combined = pd.concat([sector_prices, benchmark_prices], axis=1).dropna()
    if len(combined) < 52:
        return None

    sec_c = combined.iloc[:, 0]
    bench_c = combined.iloc[:, 1]

    # 2. Run the math on cleanly aligned series
    raw_rs = (sec_c / bench_c) * 100
    rs_smoothed = raw_rs.ewm(span=10, adjust=False).mean()

    rolling_mean = rs_smoothed.rolling(window=52, min_periods=20).mean()
    rolling_std = rs_smoothed.rolling(window=52, min_periods=20).std()

    rolling_std = rolling_std.replace(0, float('nan'))
    rs_ratio = 100 + ((rs_smoothed - rolling_mean) / rolling_std) * 2

    rs_momentum_raw = rs_ratio - rs_ratio.shift(1)
    mom_smoothed = rs_momentum_raw.ewm(span=5, adjust=False).mean()
    mom_mean = mom_smoothed.rolling(window=52, min_periods=20).mean()
    mom_std = mom_smoothed.rolling(window=52, min_periods=20).std()
    mom_std = mom_std.replace(0, float('nan'))
    rs_momentum = 100 + ((mom_smoothed - mom_mean) / mom_std) * 2

    valid = rs_ratio.notna() & rs_momentum.notna()
    rs_r = rs_ratio[valid]
    rs_m = rs_momentum[valid]

    if len(rs_r) == 0:
        return None

    n = min(tail_length, len(rs_r))
    tail = []
    for i in range(-n, 0):
        tail.append({
            "rs_ratio": round(float(rs_r.iloc[i]), 2),
            "rs_momentum": round(float(rs_m.iloc[i]), 2),
        })

    current = tail[-1] if tail else None
    if current:
        r, m = current["rs_ratio"], current["rs_momentum"]
        if r >= 100 and m >= 100:
            quadrant = "Leading"
        elif r >= 100 and m < 100:
            quadrant = "Weakening"
        elif r < 100 and m >= 100:
            quadrant = "Improving"
        else:
            quadrant = "Lagging"
    else:
        quadrant = "Unknown"

    return {
        "tail": tail,
        "quadrant": quadrant,
        "rs_ratio": current["rs_ratio"] if current else 100.0,
        "rs_momentum": current["rs_momentum"] if current else 100.0
    }


@app.get("/api/rrg")
def get_rrg(
    benchmark: str = "SPY",
    tail: int = Query(8, ge=4, le=60),
    timeframe: str = "1W",
    group: str = Query("sectors"),
    mega_theme: str = Query(None),
    sub_theme: str = Query(None),
    extra: str = Query(None)
):
    try:
        benchmark = benchmark.upper().strip()
        timeframe = timeframe.upper().strip()
        group = group.lower().strip()
        structure = load_taxonomy() or {}

        tickers_to_fetch = [benchmark]
        target_mapping = []

        # 1. Resolve symbols dynamically based on table context (Sectors, Mega, or Sub-themes)
        if sub_theme:
            sub_theme = sub_theme.strip()
            for m_name, sub_dict in structure.items():
                if sub_theme in sub_dict:
                    tickers = sub_dict[sub_theme]
                    for t in tickers:
                        if t:
                            clean_t = t.upper().strip()
                            target_mapping.append({"symbol": clean_t, "description": f"Component of {sub_theme}"})
                            tickers_to_fetch.append(clean_t)
        elif mega_theme:
            mega_theme = mega_theme.strip()
            if mega_theme in structure:
                sub_dict = structure[mega_theme]
                for s_name, t_list in sub_dict.items():
                    valid_t = [t.upper().strip() for t in t_list if t]
                    if valid_t:
                        # Represent sub-themes as group nodes on the RRG chart
                        target_mapping.append({"symbol": s_name, "description": f"Sub-Theme under {mega_theme}", "tickers": valid_t})
                        tickers_to_fetch.extend(valid_t)
        elif group == "sectors":
            for sym, info in SECTOR_ETFS.items():
                target_mapping.append({"symbol": sym, "description": info["name"], "color": info["color"], "tickers": [sym]})
                tickers_to_fetch.append(sym)
        elif group == "mega_themes":
            for m_name, sub_dict in structure.items():
                all_m_tickers = []
                for s_name, t_list in sub_dict.items():
                    if t_list:
                        all_m_tickers.extend([t.upper().strip() for t in t_list if t])
                if all_m_tickers:
                    target_mapping.append({"symbol": m_name, "description": "Mega-Theme Group", "tickers": list(set(all_m_tickers))})
                    tickers_to_fetch.extend(all_m_tickers)
        elif group == "sub_themes":
            for m_name, sub_dict in structure.items():
                for s_name, t_list in sub_dict.items():
                    valid_t = [t.upper().strip() for t in t_list if t]
                    if valid_t:
                        target_mapping.append({"symbol": s_name, "description": f"Sub-Theme ({m_name})", "tickers": valid_t})
                        tickers_to_fetch.extend(valid_t)

        # Layer on any individually-added extra tickers (e.g. clicked from the theme
        # sidebar) as their own standalone RRG tails, on top of whatever group/drill-down
        # is already active - skipping anything already present so we don't duplicate.
        if extra:
            existing_symbols = {item["symbol"] for item in target_mapping}
            for t in extra.split(","):
                clean_t = t.upper().strip()
                if clean_t and clean_t != benchmark and clean_t not in existing_symbols:
                    target_mapping.append({"symbol": clean_t, "description": f"Ticker: {clean_t}"})
                    tickers_to_fetch.append(clean_t)
                    existing_symbols.add(clean_t)

        # 2. Fetch prices for the resolved universe using your local cache
        prices_df = fetch_prices(list(set(tickers_to_fetch)), timeframe=timeframe)
        if prices_df is None or benchmark not in prices_df.columns:
            return {"benchmark": benchmark, "sectors": {}}

        bench_closes = prices_df[benchmark]
        sectors_data = {}

        # 3. Compute RRG tails matching the current table view
        for item in target_mapping:
            sym_key = item["symbol"]
            if sym_key == benchmark:
                continue

            if "tickers" in item:
                comp_ts = [t for t in item["tickers"] if t in prices_df.columns]
                if not comp_ts:
                    continue
                series_to_use = prices_df[comp_ts].mean(axis=1).dropna() if len(comp_ts) > 1 else prices_df[comp_ts[0]].dropna()
            else:
                if sym_key not in prices_df.columns:
                    continue
                series_to_use = prices_df[sym_key].dropna()

            if len(series_to_use) < 5:
                continue

            rrg_res = compute_single_rrg(series_to_use, bench_closes, tail_length=tail)
            if rrg_res:
                color = SECTOR_ETFS.get(sym_key, {}).get("color", "#00bcd4")
                sectors_data[sym_key] = {
                    "symbol": sym_key,
                    "description": item["description"],
                    "color": color,
                    "quadrant": rrg_res["quadrant"],
                    "rs_ratio": rrg_res["rs_ratio"],
                    "rs_momentum": rrg_res["rs_momentum"],
                    "tail": rrg_res["tail"]
                }

        return {"benchmark": benchmark, "sectors": sectors_data}

    except Exception as e:
        print(f"CRITICAL ERROR in get_rrg: {e}")
        return {"benchmark": benchmark, "sectors": {}}


@app.get("/api/candles")
def get_candles(
    group: str = Query("sectors"),
    mega_theme: str = Query(None),
    sub_theme: str = Query(None),
    symbols: str = Query(None),
    days: int = Query(90, ge=20, le=252)
):
    """Candlestick data for the chart-grid view. Mirrors the same universe resolution as
    /api/group-stats and /api/rrg (sectors / mega_themes / sub_themes / drill-down), then
    optionally filters down to a specific `symbols` list (e.g. whatever the stats table
    currently shows, respecting an active state filter).

    For real tickers this returns actual OHLC. For theme aggregates (rows with multiple
    constituents), it synthesizes a blended candle the same way theme price/RS aggregates
    are built elsewhere in this app: each constituent is normalized to its own base=100
    index, then O/H/L/C are averaged across constituents per date - an equal-weighted
    return-index candle, not a real tradable OHLC.
    """
    try:
        structure = load_taxonomy() or {}
        target_mapping = []

        if sub_theme:
            sub_theme = sub_theme.strip()
            for m_name, sub_dict in structure.items():
                if sub_theme in sub_dict:
                    for t in sub_dict[sub_theme]:
                        if t:
                            clean_t = t.upper().strip()
                            target_mapping.append({"symbol": clean_t, "description": f"Component of {sub_theme}"})
        elif mega_theme:
            mega_theme = mega_theme.strip()
            if mega_theme in structure:
                for s_name, t_list in structure[mega_theme].items():
                    valid_t = [t.upper().strip() for t in t_list if t]
                    if valid_t:
                        target_mapping.append({"symbol": s_name, "description": f"Sub-Theme under {mega_theme}", "tickers": valid_t})
        elif group == "sectors":
            for sym, info in SECTOR_ETFS.items():
                target_mapping.append({"symbol": sym, "description": info["name"]})
        elif group == "mega_themes":
            for m_name, sub_dict in structure.items():
                all_m_tickers = []
                for s_name, t_list in sub_dict.items():
                    if t_list:
                        all_m_tickers.extend([t.upper().strip() for t in t_list if t])
                if all_m_tickers:
                    target_mapping.append({"symbol": m_name, "description": "Mega-Theme Group", "tickers": list(set(all_m_tickers))})
        elif group == "sub_themes":
            for m_name, sub_dict in structure.items():
                for s_name, t_list in sub_dict.items():
                    valid_t = [t.upper().strip() for t in t_list if t]
                    if valid_t:
                        target_mapping.append({"symbol": s_name, "description": f"Sub-Theme ({m_name})", "tickers": valid_t})

        if symbols:
            # Parsed as JSON rather than comma-split, since some theme names contain
            # literal commas (e.g. "Pipelines, LNG & Refining"), which would otherwise
            # get corrupted into fragments that match nothing.
            try:
                wanted = {s.strip() for s in json.loads(symbols) if s and s.strip()}
            except (json.JSONDecodeError, TypeError):
                # Fallback for any old caller still sending a plain comma-separated string
                wanted = {s.strip() for s in symbols.split(",") if s.strip()}
            target_mapping = [item for item in target_mapping if item["symbol"] in wanted]

        if not target_mapping:
            return {}

        all_tickers = set()
        for item in target_mapping:
            all_tickers.update(item["tickers"] if "tickers" in item else [item["symbol"]])

        ohlc_data = fetch_ohlc(list(all_tickers), days=days)

        result = {}
        for item in target_mapping:
            sym_key = item["symbol"]

            if "tickers" in item:
                comp_frames = []
                for t in item["tickers"]:
                    df = ohlc_data.get(t)
                    if df is None or df.empty:
                        continue
                    base = df['Close'].iloc[0]
                    if not base:
                        continue
                    comp_frames.append(df / base * 100)

                if not comp_frames:
                    continue

                combined = pd.concat(comp_frames, axis=1, keys=range(len(comp_frames)))
                blended = pd.DataFrame({
                    'Open': combined.xs('Open', axis=1, level=1).mean(axis=1),
                    'High': combined.xs('High', axis=1, level=1).mean(axis=1),
                    'Low': combined.xs('Low', axis=1, level=1).mean(axis=1),
                    'Close': combined.xs('Close', axis=1, level=1).mean(axis=1),
                }).dropna()
                candle_df = blended.tail(days)
            else:
                candle_df = ohlc_data.get(sym_key)
                if candle_df is None or candle_df.empty:
                    continue

            candles = [
                {
                    "t": idx.strftime("%Y-%m-%d"),
                    "o": round(float(row.Open), 2),
                    "h": round(float(row.High), 2),
                    "l": round(float(row.Low), 2),
                    "c": round(float(row.Close), 2),
                }
                for idx, row in candle_df.iterrows()
            ]
            if candles:
                result[sym_key] = {"description": item["description"], "candles": candles}

        return result

    except Exception as e:
        print(f"CRITICAL ERROR in get_candles: {e}")
        return {}



def _vol_adj_rs_component(sym_series, bench_series, bars):
    """Annualized information-ratio-style relative strength for ONE lookback window: mean
    daily excess return over the benchmark, divided by the volatility of that daily excess
    return, annualized by sqrt(252).

    This is scale-invariant across different lookback windows (unlike raw cumulative excess
    return, where a 1-year window's cumulative return is inherently much larger than a
    1-month window's at the same volatility - blending those directly would let the longest
    window dominate for the wrong reason). It also directly rewards STEADY outperformance
    over erratic/choppy outperformance of the same total magnitude, matching how TradersLab
    describes their group/theme ranking: "a steadier group isn't outranked by a more erratic
    one purely on raw performance."
    """
    if bench_series is None or sym_series is None:
        return 0.0
    if len(sym_series) <= bars or len(bench_series) <= bars:
        return 0.0
    try:
        sym_window = sym_series.iloc[-bars - 1:]
        bench_window = bench_series.iloc[-bars - 1:]
        sym_ret = sym_window.pct_change().dropna()
        bench_ret = bench_window.pct_change().dropna()
        n = min(len(sym_ret), len(bench_ret))
        if n < 5:
            return 0.0
        rel_daily = sym_ret.values[-n:] - bench_ret.values[-n:]
        vol_daily = rel_daily.std(ddof=1)
        # Guard against near-zero volatility - low-vol instruments (preferred stocks, some
        # closed-end funds) can have daily relative-return volatility so tiny that dividing
        # by it explodes this ratio to an extreme, meaningless value. 1e-4 (0.01% daily vol)
        # is well below even the calmest real securities' typical volatility.
        if not vol_daily or pd.isna(vol_daily) or vol_daily < 1e-4:
            return 0.0
        mean_daily = rel_daily.mean()
        ir = (mean_daily / vol_daily) * (252 ** 0.5)
        # Clip to a sane range - genuine, sustained annualized information ratios rarely
        # exceed a few points; anything beyond this is almost certainly numerical noise
        # from a low-volatility instrument rather than a real signal, and left unclipped it
        # can single-handedly blow the whole blended composite_rs out to the scale's extreme.
        return float(max(-5.0, min(5.0, ir)))
    except Exception:
        return 0.0


# Weight given to 52-week high/low positioning inside composite_rs. TradersLab's own docs
# describe their RS Rank as factoring in "a stock's performance over multiple timeframes...
# along with the stock's distance from its 52-week high and low" - but they don't publish an
# exact weighting. This is a deliberately modest weight, chosen so it nudges composite_rs
# without disturbing the multi-timeframe blend that the Leading/Emerging/Fading/Lagging/
# Breaking Down thresholds in get_group_stats were calibrated against (via direct comparison
# to TradersLab's live Themes Lab table) - since percentile ranking is rank-based, a small
# weight here is very unlikely to reorder tickers enough to invalidate that calibration.
# Revisit if real TradersLab RS Rank values become available to calibrate against directly.
HIGH_LOW_WEIGHT = 0.15


def _high_low_position_score(series, bars=252):
    """Where the current price sits within its trailing ~52-week high/low range, expressed on
    a scale comparable to the vol-adjusted IR components below (roughly -2 at the 52-week low,
    +2 at the 52-week high, 0 at the midpoint of the range)."""
    if series is None or series.empty:
        return 0.0
    window = series.tail(min(bars, len(series)))
    if window.empty:
        return 0.0
    hi = window.max()
    lo = window.min()
    if hi == lo or pd.isna(hi) or pd.isna(lo):
        return 0.0
    current = series.iloc[-1]
    pct_of_range = (current - lo) / (hi - lo)
    return float((pct_of_range - 0.5) * 4.0)


def _raw_composite_rs(sym_series, bench_series):
    """Volatility-adjusted composite RS: front-loaded blend (40/30/20/10) of the annualized
    information ratio over 1M/3M/6M/1Y windows, blended with a 52-week high/low positioning
    term (see HIGH_LOW_WEIGHT above). Single shared implementation - used for theme/sector
    aggregates, individual constituent tickers (breadth scoring), and the theme-rankings
    sidebar, so all of them measure relative strength the same way."""
    r1m = _vol_adj_rs_component(sym_series, bench_series, 21)
    r3m = _vol_adj_rs_component(sym_series, bench_series, 63)
    r6m = _vol_adj_rs_component(sym_series, bench_series, 126)
    r1y = _vol_adj_rs_component(sym_series, bench_series, 252)
    timeframe_blend = (r1m * 0.4) + (r3m * 0.3) + (r6m * 0.2) + (r1y * 0.1)
    hl_score = _high_low_position_score(sym_series)
    return (timeframe_blend * (1 - HIGH_LOW_WEIGHT)) + (hl_score * HIGH_LOW_WEIGHT)


# How far composite_rs has to move, in either direction, before the 0-100 score below hits
# its floor/ceiling. composite_rs is now an annualized-information-ratio-style number
# (typically roughly -3 to +3 for real securities; |IR| > 1 is already a fairly strong,
# consistent outperformance), NOT a raw percentage - this range was recalibrated for that
# new scale. Still a judgment call - tune it if scores cluster too tightly around 50 or pin
# too often at 0/100.
RS_SCORE_CLAMP_RANGE = 2.0


def rs_score_from_composite(composite_rs):
    """Maps a SPY-relative composite_rs value onto a FIXED 0-100 scale, anchored so that
    composite_rs == 0 (performing exactly in line with SPY) always maps to 50 - regardless
    of what else happens to be in the current table view.

    This replaces an earlier version that percentile-ranked each row against whatever
    else was currently displayed (other sectors, other mega themes, etc). That made scores
    incomparable across views and dependent on the current peer set rather than on actual
    performance relative to SPY - the same underlying performance could score a 90 one day
    (weak peers) and a 20 another day (strong peers), which isn't what "relative strength"
    is supposed to mean.
    """
    if composite_rs is None:
        return None
    scaled = 50 + (composite_rs / RS_SCORE_CLAMP_RANGE) * 50
    return round(max(0, min(100, scaled)))


BREADTH_RS_THRESHOLD = 80  # a constituent ticker counts toward "breadth" at/above this 0-100 RS score

_full_universe_rs_cache = {"ts": 0, "now": {}, "past": {}}
FULL_UNIVERSE_RS_CACHE_TTL = 1800  # 30 min - expensive to recompute (thousands of tickers), doesn't need to be real-time

RS_ROC_BARS = 40  # 8 trading weeks


def get_full_universe_rs_scores():
    """Each ticker's RS score, 0-100 on the fixed SPY-anchored scale (rs_score_from_composite
    applied directly to that ticker's own composite_rs vs SPY) - the SAME scale sectors,
    mega-themes, and the theme-rankings sidebar all use. No ranking against any peer pool.

    This replaced TWO earlier designs, in order, both of which ranked each ticker's score
    against a peer pool instead of scoring it absolutely - and both broke for the same
    underlying reason (a peer-relative rank isn't a measurement of the ticker itself, it's
    a measurement of the ticker relative to whoever else happens to be in the pool):

    1. Percentile rank against the full ~3-4k ticker taxonomy. Produced much more extreme
       swings than TradersLab's own numbers on a direct side-by-side (their Oil Tankers:
       RS 75, +27 trajectory; ours: RS 98, +49pp) - comparing a tanker stock to random
       biotech and software names is too wide/incoherent a pool.

    2. Percentile rank scoped to just the ticker's own mega-theme (e.g. all "Pipelines,
       LNG & Refining" names together). Fixed the incoherent-pool problem, but caused two
       distinct bugs once get_group_stats averaged these percentiles for a theme's
       constituents:
         a) Mega-theme (Group) rows average ALL of a mega-theme's tickers - which is the
            ENTIRE pool those percentiles were ranked within. The average percentile of a
            population ranked against itself is mathematically always ~50, regardless of
            the underlying data (verified analytically and empirically) - every Mega Theme
            row showed Comp RS=50, ROC=0.00pp, State=Neutral, reported 2026-09-12.
         b) Sub-theme rows average a proper SUBSET of their mega-theme's pool, so this
            isn't degenerate in the same way - but it's still relative to whichever other
            sub-themes happen to share that mega-theme, which can flip the sign of the
            result. Confirmed 2026-09-12: TradersLab shows "Oil & Gas Refining" (VLO, PSX,
            MPC, DINO, DK) clearly Leading (RS 92, +9 trajectory) - genuinely outperforming
            SPY - but our percentile-within-mega-theme approach showed it as Lagging,
            because its mega-theme sibling "Oil Tankers" (INSW, ECO, TNK, FRO, TRMD) was
            up +45pp and dragged the WHOLE POOL's percentile distribution so high that
            Oil & Gas Refining's own genuinely-strong tickers ranked in the bottom half of
            it. Same mechanism flipped "Content & Document Mgmt" and "Universal Banks" the
            wrong direction in that same comparison. A theme's classification should not
            depend on how hot its sibling theme happens to be.

    The absolute score has neither problem: it isn't a rank, so averaging it over an entire
    pool isn't degenerate, and it isn't relative to any peer group, so one theme's numbers
    can't leak into a sibling theme's classification. This also matches how TradersLab
    describes RS Rank in their own docs - performance across timeframes plus 52-week
    high/low positioning, i.e. an absolute measure, not a rank against an arbitrary
    peer subset. get_group_stats uses this same absolute score uniformly for mega-theme
    rows, sub-theme rows, breadth counting, top-tickers, and the individual-ticker
    drill-down - one methodology everywhere, matching what sectors and mega-themes already
    used for composite_rs.

    Cached since scoring thousands of tickers on every request would be wasteful when
    nothing's changed.
    """
    now = time.time()
    if _full_universe_rs_cache["now"] and (now - _full_universe_rs_cache["ts"]) < FULL_UNIVERSE_RS_CACHE_TTL:
        return _full_universe_rs_cache["now"], _full_universe_rs_cache["past"]

    structure = load_taxonomy() or {}

    all_tickers = set()
    for m_name, sub_dict in structure.items():
        for s_name, t_list in (sub_dict or {}).items():
            for t in (t_list or []):
                if t:
                    all_tickers.add(t.upper().strip())

    fetch_list = list(all_tickers | {"SPY"})
    prices_df = fetch_prices(fetch_list, timeframe="1D") if fetch_list else pd.DataFrame()
    bench_series = prices_df["SPY"].dropna() if "SPY" in prices_df.columns else None

    now_scores, past_scores = {}, {}
    if bench_series is not None:
        for t in all_tickers:
            if t not in prices_df.columns:
                continue
            t_series = prices_df[t].dropna()
            if len(t_series) < 30:
                continue
            now_scores[t] = rs_score_from_composite(_raw_composite_rs(t_series, bench_series))
            if len(t_series) >= 30 + RS_ROC_BARS:
                past_raw = _raw_composite_rs(t_series.iloc[:-RS_ROC_BARS], bench_series.iloc[:-RS_ROC_BARS])
                past_scores[t] = rs_score_from_composite(past_raw)

    _full_universe_rs_cache["now"] = now_scores
    _full_universe_rs_cache["past"] = past_scores
    _full_universe_rs_cache["ts"] = now

    return now_scores, past_scores


@app.get("/api/group-stats")
@app.get("/api/group-stats")
def get_group_stats(
        group: str = Query("sectors"),
        mega_theme: str = Query(None),
        sub_theme: str = Query(None)
):
    group = group.lower().strip()
    structure = load_taxonomy()

    tickers_to_fetch = ["SPY"]  # Always ensure SPY is present for benchmark comparison
    target_mapping = []
    group_mode = "sectors"

    if sub_theme:
        sub_theme = sub_theme.strip()
        for m_name, sub_dict in structure.items():
            if sub_theme in sub_dict:
                tickers = sub_dict[sub_theme]
                for t in tickers:
                    if t:
                        clean_t = t.upper().strip()
                        target_mapping.append(
                            {"symbol": clean_t, "description": f"Component of {sub_theme}", "is_group": False})
                        tickers_to_fetch.append(clean_t)
        group_mode = "tickers"

    elif mega_theme:
        mega_theme = mega_theme.strip()
        if mega_theme in structure:
            sub_dict = structure[mega_theme]
            for s_name, t_list in sub_dict.items():
                valid_t = [t.upper().strip() for t in t_list if t]
                if valid_t:
                    target_mapping.append({
                        "symbol": s_name,
                        "description": f"Sub-Theme under {mega_theme}",
                        "tickers": valid_t,
                        "is_group": True,
                        "next_level": "sub"
                    })
                    tickers_to_fetch.extend(valid_t)
        group_mode = "sub_themes_filtered"

    elif group == "sectors":
        tickers_to_fetch.extend(list(SECTOR_ETFS.keys()))
        group_mode = "sectors"

    elif group == "mega_themes":
        for m_name, sub_dict in structure.items():
            all_m_tickers = []
            for s_name, t_list in sub_dict.items():
                if t_list:
                    all_m_tickers.extend([t.upper().strip() for t in t_list if t])
            if all_m_tickers:
                target_mapping.append({
                    "symbol": m_name,
                    "description": "Mega-Theme Group",
                    "tickers": list(set(all_m_tickers)),
                    "is_group": True,
                    "next_level": "mega"
                })
                tickers_to_fetch.extend(all_m_tickers)
        group_mode = "mega_themes"

    elif group == "sub_themes":
        for m_name, sub_dict in structure.items():
            for s_name, t_list in sub_dict.items():
                valid_t = [t.upper().strip() for t in t_list if t]
                if valid_t:
                    target_mapping.append({
                        "symbol": s_name,
                        "description": f"Sub-Theme ({m_name})",
                        "tickers": valid_t,
                        "is_group": True,
                        "next_level": "sub"
                    })
                    tickers_to_fetch.extend(valid_t)
        group_mode = "sub_themes"

    prices_df = fetch_prices(list(set(tickers_to_fetch)), timeframe="1D") if tickers_to_fetch else pd.DataFrame()
    bench_series = prices_df["SPY"].dropna() if "SPY" in prices_df.columns else None

    # Individual constituent tickers are scored on the same absolute, SPY-anchored 0-100
    # scale as sectors/mega-themes/sub-themes (see get_full_universe_rs_scores() for why
    # this replaced peer-pool percentile ranking - it caused both a mega-theme averaging
    # degeneracy and a sub-theme cross-contamination bug where one theme's classification
    # could flip depending on how hot a SIBLING theme in the same mega-theme was doing).
    ticker_rs_scores = {}
    ticker_rs_scores_8w_ago = {}
    if group_mode in ("mega_themes", "sub_themes", "sub_themes_filtered"):
        full_now, full_past = get_full_universe_rs_scores()
        all_constituents = sorted({t for item in target_mapping for t in item.get("tickers", [])})
        for t in all_constituents:
            if t in full_now:
                ticker_rs_scores[t] = full_now[t]
            if t in full_past:
                ticker_rs_scores_8w_ago[t] = full_past[t]

    stats_list = []

    def calculate_metrics(series, symbol, desc, is_grp=False, n_lvl="none"):
        current_price = float(series.iloc[-1])

        def get_rs(bars):
            return _vol_adj_rs_component(series, bench_series, bars)

        def safe_chg(bars):
            # Keep fallback for short-term bar-based lookups if needed,
            # or handle months/years via calendar offset if bars >= 126
            if series is None or len(series) < 2:
                return 0.0

            end_date = series.index[-1]

            # Map long-term bar checks to calendar offsets
            if bars == 126:  # 6 Months
                target_date = end_date - pd.DateOffset(months=6)
            elif bars == 252:  # 1 Year
                target_date = end_date - pd.DateOffset(years=1)
            else:
                # Standard bar-based check for short terms (1D, 1W, etc.)
                if len(series) <= bars:
                    return 0.0
                past_val = series.iloc[-1 - bars]
                curr_val = series.iloc[-1]
                if past_val == 0 or pd.isna(past_val) or pd.isna(curr_val):
                    return 0.0
                return (curr_val - past_val) / past_val

            # Calendar offset lookup for 6M and 1Y
            valid_subset = series[series.index <= target_date]
            past_val = valid_subset.iloc[-1] if not valid_subset.empty else series.iloc[0]
            curr_val = series.iloc[-1]

            if past_val == 0 or pd.isna(past_val) or pd.isna(curr_val):
                return 0.0

            return (curr_val - past_val) / past_val

        rs_1m = get_rs(21)
        rs_3m = get_rs(63)
        rs_6m = get_rs(126)
        rs_1y = get_rs(252)
        # Goes through the shared _raw_composite_rs (same helper used by get_theme_rankings
        # and get_full_universe_rs_scores) rather than a local recompute, so it also picks up
        # the 52-week high/low term instead of duplicating (and risking drifting from) the
        # front-loaded 40/30/20/10 timeframe blend.
        composite_rs = _raw_composite_rs(series, bench_series)

        sma_20 = series.tail(20).mean() if len(series) >= 20 else current_price
        sma_50 = series.tail(50).mean() if len(series) >= 50 else current_price
        sma_200 = series.tail(200).mean() if len(series) >= 200 else current_price

        # 8-week (40 trading session) rate of change of the RS score itself, i.e. how much
        # relative strength has moved over the last 2 months, not how much price has moved.
        # Expressed as a point change on the same 0-100 rs_score scale used everywhere else
        # (so it's directly comparable to theme-level rs_roc_8w, which is also computed in
        # 0-100-scale points via constituent averaging).
        #
        # NOTE (2026-09-12): this is exactly "RS score now minus RS score 8 weeks ago" -
        # nothing fancier. An earlier attempt to smooth each side over a few nearby trading
        # days (to fight the wild swings reported the same day) was reverted: simulation
        # showed it barely reduced the noise (a handful of adjacent daily windows share
        # ~90%+ of the same underlying data, so averaging them isn't real diversification).
        # The actual cause is structural - composite_rs is 40% weighted on a 21-trading-day
        # annualized information ratio, and IR estimated from a 21-day sample has a standard
        # error of roughly sqrt(252/21) =~ 3.5 IR units from ordinary noise alone, which
        # RS_SCORE_CLAMP_RANGE=2.0 (25 score-points per 1.0 IR unit) turns into routine
        # 25-90 point swings for a theme whose true relative strength hasn't materially
        # changed. That's a property of the scale the score is measured on, not of how the
        # two snapshots are sampled - see the message to the user for the fix options.
        RS_ROC_BARS = 40
        if len(series) > RS_ROC_BARS + 252 and bench_series is not None and len(bench_series) > RS_ROC_BARS + 252:
            composite_rs_8w_ago = _raw_composite_rs(series.iloc[:-RS_ROC_BARS], bench_series.iloc[:-RS_ROC_BARS])
            score_now = rs_score_from_composite(composite_rs)
            score_8w_ago = rs_score_from_composite(composite_rs_8w_ago)
            rs_roc_8w = round(score_now - score_8w_ago, 2)
        else:
            composite_rs_8w_ago = None
            rs_roc_8w = None

        return {
            "symbol": symbol,
            "description": desc,
            "price": round(current_price, 2),
            "chg_1d": round(safe_chg(1), 4),
            "chg_1w": round(safe_chg(5), 4),
            "chg_2w": round(safe_chg(10), 4),
            "chg_1m": round(safe_chg(21), 4),
            "chg_2m": round(safe_chg(42), 4),
            "chg_3m": round(safe_chg(63), 4),
            "chg_6m": round(safe_chg(126), 4),
            "chg_1y": round(safe_chg(252), 4),
            "rs_1m": round(rs_1m * 100, 2),
            "rs_3m": round(rs_3m * 100, 2),
            "rs_6m": round(rs_6m * 100, 2),
            "rs_1y": round(rs_1y * 100, 2),
            "composite_rs": round(composite_rs * 100, 2),
            "composite_rs_8w_ago": composite_rs_8w_ago,
            "rs_roc_8w": round(rs_roc_8w, 2) if rs_roc_8w is not None else None,
            "volume": 0,
            "avg_vol": 0,
            "vs_20sma": round(float((current_price / sma_20) - 1), 4) if sma_20 else 0.0,
            "vs_50sma": round(float((current_price / sma_50) - 1), 4) if sma_50 else 0.0,
            "vs_200sma": round(float((current_price / sma_200) - 1), 4) if sma_200 else 0.0,
            "is_group": is_grp,
            "next_level": n_lvl
        }

    if group_mode in ["sectors", "tickers"]:
        symbols_to_loop = [s for s in (tickers_to_fetch if group_mode == "sectors" else [m["symbol"] for m in target_mapping]) if s != "SPY"]
        for sym in symbols_to_loop:
            if sym not in prices_df.columns:
                continue
            series = prices_df[sym].dropna()
            if len(series) < 5:
                continue
            desc = SECTOR_ETFS.get(sym, {}).get("name", f"Asset: {sym}") if group_mode == "sectors" else f"Ticker: {sym}"
            row = calculate_metrics(series, sym, desc, is_grp=False)
            row["breadth_count"] = None
            row["breadth_total"] = None
            row["top_tickers"] = []
            stats_list.append(row)
    else:
        for item in target_mapping:
            comp_tickers = [t for t in item["tickers"] if t in prices_df.columns]
            if not comp_tickers:
                continue
            comp_df = prices_df[comp_tickers]
            # Same fix as the theme-rankings endpoint: normalize each constituent to its
            # own base=100 index before averaging, so this reflects equal-weighted returns
            # rather than being skewed by whichever constituent has the highest dollar price.
            normalized = comp_df / comp_df.bfill().iloc[0] * 100
            sub_closes = normalized.mean(axis=1).dropna()
            if len(sub_closes) < 5:
                continue
            row = calculate_metrics(
                sub_closes,
                item["symbol"],
                item["description"],
                is_grp=True,
                n_lvl=item.get("next_level", "sub")
            )

            # Breadth: how many of this theme's constituents individually score >= threshold,
            # plus the strongest few by name, mirroring the "N names above RS X" framing.
            constituent_scores = [
                (t, ticker_rs_scores[t]) for t in item["tickers"] if t in ticker_rs_scores
            ]
            if constituent_scores:
                row["breadth_count"] = sum(1 for _, s in constituent_scores if s >= BREADTH_RS_THRESHOLD)
                row["breadth_total"] = len(constituent_scores)
                top3 = sorted(constituent_scores, key=lambda x: x[1], reverse=True)[:3]
                row["top_tickers"] = [{"symbol": t, "score": s} for t, s in top3]

                # Theme RS score = average of individual constituent RS scores (confirmed
                # against TradersLab: summing their displayed per-ticker scores for a theme
                # and dividing by constituent count reproduces their theme-level score
                # exactly). This replaces the blended-price-index approach for this field.
                #
                # ticker_rs_scores is now each ticker's ABSOLUTE, SPY-anchored score (see
                # get_full_universe_rs_scores()), not a peer-pool percentile rank - so this
                # average is safe and correct uniformly for mega-theme rows (averaging the
                # theme's ENTIRE ticker set), sub-theme rows (averaging a subset), and
                # everything in between. Two earlier percentile-based designs each broke
                # differently here: ranking within the whole mega-theme pool made a
                # mega-theme's own average degenerate (always ~50, since it's the entire
                # ranked population averaging itself - "every Mega Theme shows Comp RS=50"
                # reported 2026-09-12), and even for sub-themes (a proper subset, so not
                # degenerate) it could still flip a theme's classification depending on how
                # hot a SIBLING theme in the same mega-theme was doing (confirmed 2026-09-12:
                # TradersLab showed "Oil & Gas Refining" clearly Leading, but our
                # percentile-within-mega-theme version showed it Lagging, because sibling
                # theme "Oil Tankers" being up +45pp dragged the whole pool's percentile
                # distribution high enough to rank Oil & Gas Refining's own strong tickers
                # in the bottom half of it). An absolute score isn't ranked against anything,
                # so neither failure mode applies - see get_full_universe_rs_scores() for
                # the full history, including the intermediate blended-price-index attempt
                # that fixed the degeneracy but caused a separate 100-or-0 saturation bug.
                row["rs_score"] = round(sum(s for _, s in constituent_scores) / len(constituent_scores))

                if item["symbol"] == "Oil Tankers":
                    print(f"DEBUG Oil Tankers constituents:")
                    for t, s_now in constituent_scores:
                        s_past = ticker_rs_scores_8w_ago.get(t)
                        print(f"    {t}: now_abs={s_now}, past_abs={s_past}")

                # rs_roc_8w = literally how many points this same score has moved over the
                # last 8 weeks - i.e. (current rs_score) minus (rs_score as of 8 weeks ago),
                # computed the identical way (average of constituent absolute scores) so
                # it's actually measuring the change in what's displayed, not a different
                # underlying calculation.
                past_scores = [ticker_rs_scores_8w_ago[t] for t, _ in constituent_scores if t in ticker_rs_scores_8w_ago]
                if past_scores:
                    avg_now = sum(s for _, s in constituent_scores) / len(constituent_scores)
                    avg_past = sum(past_scores) / len(past_scores)
                    row["rs_roc_8w"] = round(avg_now - avg_past, 2)
            else:
                row["breadth_count"] = None
                row["breadth_total"] = None
                row["top_tickers"] = []

            stats_list.append(row)

    # Sectors and individual tickers are NOT rated 0-100 like Groups/Themes are - see below.
    if group_mode == "sectors":
        # Sectors get a simple ORDINAL RANK (1 = strongest), not a 0-100 rating - matching
        # TradersLab, where sectors are just ranked 1-N among themselves. That's fine here
        # specifically because the sector set is small, fixed, and complete (there's no
        # larger "sector universe" to measure against - these ARE the whole population),
        # unlike themes or tickers where ranking against an arbitrary in-view subset was
        # the actual problem we fixed earlier.
        ranked_sectors = sorted(stats_list, key=lambda r: r["composite_rs"], reverse=True)
        for rank, row in enumerate(ranked_sectors, start=1):
            row["rs_score"] = rank

        # rs_roc_8w for sectors = how many rank positions improved over 8 weeks, using the
        # exact same ranking method (so it's a change in the displayed number, not a
        # different underlying calculation). Inverted (past rank minus current rank) since
        # rank 1 is best - moving from rank 5 to rank 2 is an improvement and should show
        # as positive, consistent with every other "positive = good" column in this table.
        past_ranked = sorted(
            [r for r in stats_list if r.get("composite_rs_8w_ago") is not None],
            key=lambda r: r["composite_rs_8w_ago"],
            reverse=True
        )
        past_rank_by_symbol = {r["symbol"]: rank for rank, r in enumerate(past_ranked, start=1)}
        for row in stats_list:
            past_rank = past_rank_by_symbol.get(row["symbol"])
            row["rs_roc_8w"] = (past_rank - row["rs_score"]) if past_rank is not None else None
    elif group_mode == "tickers":
        # Drilled into a sub-theme's individual constituent tickers - score each one
        # against the full taxonomy universe, same as breadth scoring does elsewhere, so
        # a given ticker's score means the same thing here as it does anywhere else.
        full_now, full_past = get_full_universe_rs_scores()
        for row in stats_list:
            row["rs_score"] = full_now.get(row["symbol"])
            past = full_past.get(row["symbol"])
            row["rs_roc_8w"] = round(row["rs_score"] - past, 2) if row["rs_score"] is not None and past is not None else None
    else:
        # Rare-edge-case fallback: mega-theme and sub-theme rows already got rs_score/
        # rs_roc_8w from averaging ticker_rs_scores (constituent_scores) above; this only
        # catches a theme whose constituent_scores came back empty entirely (e.g. every
        # ticker missing price data), falling back to the SPY-anchored composite_rs scale
        # sectors also use.
        for row in stats_list:
            if row.get("rs_score") is None:
                row["rs_score"] = rs_score_from_composite(row["composite_rs"])

    # State classification (Leading/Emerging/Neutral/Fading/Lagging/Breaking Down), based on
    # RS level (rs_score) and RS trajectory (rs_roc_8w) - the same two axes as an RRG quadrant.
    #
    # Recalibrated 2026-09-12 against ~50 sub-themes ("Theme" rows) pulled directly from
    # TradersLab's live Themes Lab table (RS score + 8-week RS trajectory + displayed State),
    # not from published docs - TradersLab doesn't publish these thresholds anywhere. Findings
    # from that side-by-side:
    #   - Neutral holds for roc in [-3, +3] regardless of RS level (e.g. RS47/roc+3 stayed
    #     Neutral). roc == +4 was the smallest observed Emerging (RS53 "Specialty Pharma"),
    #     roc == -4 the smallest observed Lagging (RS58 "Electronics Distribution") - a hard
    #     +-4 cutoff, not the old asymmetric +0/-5.
    #   - LEADING_RS_FLOOR=64 is an exact boundary: RS64/roc+1 ("Drug Discovery AI") showed as
    #     Leading, while RS63/roc+43 ("Precious Metals Royalty" - the single most extreme
    #     positive trajectory in the whole sample) was still only Emerging. RS<64 never reached
    #     Leading no matter how strongly positive the trajectory got.
    #   - Fading requires BOTH RS>=64 AND a strongly negative trajectory: RS65/roc-9 ("Staffing
    #     + IT Services") showed Fading, but RS65/roc-7 ("Universal Banks") was still plain
    #     Lagging - a mildly negative trajectory doesn't "demote" a high-RS theme into Fading,
    #     it drops straight to Lagging like anywhere else; only roc <= ~-8 flips it to Fading.
    #     (Sample doesn't pin the exact edge between -7 and -9 - using -8, the midpoint.)
    #   - Breaking Down never appeared anywhere in the live data (0 of 183 themes), even at RS
    #     as low as 22 with roc as low as -24 ("Photonic ICs", "Semiconductor Equipment") -
    #     those were just "Lagging". The old thresholds (RS<35, roc<=-15) would have
    #     misclassified several observed themes as Breaking Down when TradersLab shows them as
    #     Lagging. Floor/roc below are tightened to be consistent with every observed
    #     non-trigger; no real Breaking Down example was available to pin the exact cutoff, so
    #     revisit if/when one shows up live.
    LEADING_RS_FLOOR = 64
    EMERGING_ROC_MIN = 4
    LAGGING_ROC_MAX = -4
    FADING_ROC_MAX = -8
    BREAKING_DOWN_RS_FLOOR = 20
    BREAKING_DOWN_ROC_MAX = -25

    def _classify_state(rs_score, roc):
        if rs_score is None:
            return None
        if roc is None:
            return "Leading" if rs_score >= LEADING_RS_FLOOR else "Neutral"

        if rs_score >= LEADING_RS_FLOOR:
            if roc <= FADING_ROC_MAX:
                return "Fading"
            if roc <= LAGGING_ROC_MAX:
                return "Lagging"
            return "Leading"

        if rs_score < BREAKING_DOWN_RS_FLOOR and roc <= BREAKING_DOWN_ROC_MAX:
            return "Breaking Down"
        if roc >= EMERGING_ROC_MIN:
            return "Emerging"
        if roc <= LAGGING_ROC_MAX:
            return "Lagging"
        return "Neutral"

    for row in stats_list:
        # Sectors are ranked 1-N (ordinal), not rated 0-100 - the Leading/Emerging/etc
        # thresholds below are calibrated for a 0-100 scale and don't mean anything applied
        # to a rank, so sectors simply don't get a state label (shows as "—" in the table).
        row["state"] = None if group_mode == "sectors" else _classify_state(row.get("rs_score"), row.get("rs_roc_8w"))

    return stats_list

@app.get("/api/holdings")
def get_holdings():
    return {}

from fastapi.responses import FileResponse

@app.get("/")
def serve_index():
    return FileResponse("index.html")

@app.get("/app.js")
def serve_js():
    return FileResponse("app.js", media_type="application/javascript")

@app.get("/style.css")
def serve_css():
    return FileResponse("style.css", media_type="text/css")

if __name__ == "__main__":
    import uvicorn
    uvicorn.run("api_server:app", host="127.0.0.1", port=8000, reload=True)